import React, {
  useState,
  useMemo,
  useRef,
  useEffect,
  useCallback,
} from 'react';
import * as XLSX from 'xlsx';

// ============================================================
// SUPABASE CLIENT
// ============================================================
const SUPABASE_URL = 'https://phpzvynjccegalkezlnb.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBocHp2eW5qY2NlZ2Fsa2V6bG5iIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk1Mjk2MTAsImV4cCI6MjA5NTEwNTYxMH0.fHZOlPnPa-oEyx5KUDsiV_cpVVsriWZqB6R2H50HFoE';

const supa = (() => {
  const headers = {
    'apikey': SUPABASE_KEY,
    'Authorization': `Bearer ${SUPABASE_KEY}`,
    'Content-Type': 'application/json',
    'Prefer': 'return=representation',
  };
  const url = (table, params = '') => `${SUPABASE_URL}/rest/v1/${table}${params}`;

  return {
    // SELECT
    from: (table) => ({
      select: async (params = '') => {
        const queryStr = params ? `?${params}&order=created_at.asc` : '?order=created_at.asc';
        const res = await fetch(url(table, queryStr), { headers });
        if (!res.ok) throw new Error(await res.text());
        return res.json();
      },
      selectOne: async (col, val) => {
        const res = await fetch(url(table, `?${col}=eq.${encodeURIComponent(val)}`), { headers });
        if (!res.ok) throw new Error(await res.text());
        const rows = await res.json();
        return rows[0] || null;
      },
    }),
    // INSERT
    insert: async (table, data) => {
      const res = await fetch(url(table), {
        method: 'POST',
        headers,
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    // UPSERT
    upsert: async (table, data) => {
      const res = await fetch(url(table), {
        method: 'POST',
        headers: { ...headers, 'Prefer': 'resolution=merge-duplicates,return=representation' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    // UPDATE
    update: async (table, col, val, data) => {
      const res = await fetch(url(table, `?${col}=eq.${encodeURIComponent(val)}`), {
        method: 'PATCH',
        headers,
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    // DELETE
    delete: async (table, col, val) => {
      const res = await fetch(url(table, `?${col}=eq.${encodeURIComponent(val)}`), {
        method: 'DELETE',
        headers,
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    // REALTIME via SSE (Supabase realtime channel)
    channel: (name) => {
      let ws = null;
      let handlers = {};
      const connect = () => {
        const wsUrl = SUPABASE_URL.replace('https://', 'wss://') + '/realtime/v1/websocket?apikey=' + SUPABASE_KEY + '&vsn=1.0.0';
        try {
          ws = new WebSocket(wsUrl);
          ws.onopen = () => {
            ws.send(JSON.stringify({ topic: `realtime:${name}`, event: 'phx_join', payload: { config: { broadcast: { ack: false }, presence: { key: '' }, postgres_changes: Object.keys(handlers).map(t => ({ event: '*', schema: 'public', table: t })) } }, ref: '1' }));
          };
          ws.onmessage = (e) => {
            try {
              const msg = JSON.parse(e.data);
              if (msg.event === 'postgres_changes') {
                const { table, eventType, new: newRow, old: oldRow } = msg.payload.data || {};
                if (handlers[table]) handlers[table].forEach(h => h({ eventType, newRow, oldRow }));
              }
            } catch {}
          };
          ws.onerror = () => {};
          ws.onclose = () => { setTimeout(connect, 3000); };
        } catch {}
      };
      return {
        on: (table, handler) => {
          if (!handlers[table]) handlers[table] = [];
          handlers[table].push(handler);
          return this;
        },
        subscribe: () => { connect(); return { unsubscribe: () => { if (ws) ws.close(); } }; },
      };
    },
  };
})();

// ============================================================
// DB HELPERS: convert DB row <-> App format
// ============================================================
const dbToPartner = (r) => ({
  id: r.id, name: r.name, region: r.region, subregion: r.subregion,
  country: r.country, type: r.type, tier: r.tier,
  allocated: Number(r.allocated || 0), spent: Number(r.spent || 0), pending: Number(r.pending || 0),
  status: r.status || 'Active', note: r.note || '',
  contactName: r.contact_name, contactEmail: r.contact_email,
  accountManager: r.account_manager,
  portalEmail: r.portal_email,
  portalPassword: r.portal_password_hash, // SECURITY TODO (BUG-007): must be bcrypt hashed in production
});

const partnerToDB = (p) => ({
  id: p.id, name: p.name, region: p.region, subregion: p.subregion,
  country: p.country, type: p.type, tier: p.tier,
  allocated: p.allocated, spent: p.spent, pending: p.pending,
  status: p.status, note: p.note || '',
  contact_name: p.contactName, contact_email: p.contactEmail,
  account_manager: p.accountManager,
});

const dbToRequest = (r, items) => ({
  id: r.id, partner: r.partner_name,
  submitted: r.submitted, status: r.status,
  assignedTo: r.assigned_to || '',
  poNumber: r.po_number || '', note: r.note || '',
  partnerNotified: r.partner_notified || false,
  notifiedAt: r.notified_at || '',
  bpGeneratedAt: r.bp_generated_at || '',
  signedDoc: r.signed_doc || null,
  items: (items || []).map(dbToItem),
});

const dbToItem = (it) => ({
  id: it.id, tactic: it.tactic, title: it.title,
  productGroup: it.product_group,
  targetSolutions: typeof it.target_solutions === 'string' ? JSON.parse(it.target_solutions) : (it.target_solutions || []),
  amount: Number(it.amount || 0),
  mdfRequest: Number(it.mdf_request || 0),
  localCurrency: it.local_currency || 'EUR',
  fyHalf: it.fy_half, fyQuarter: it.fy_quarter,
  month: it.month, period: it.period,
  where: it.where_field, targetAudience: it.target_audience,
  targetAttendees: it.target_attendees, objective: it.objective,
  itemStatus: it.item_status || 'request_submitted',
  assignedTo: it.assigned_to || '',
  allocadiaId: it.allocadia_id || '',
  campaignId: it.campaign_id || '',
  cancelReason: it.cancel_reason, postponedTo: it.postponed_to,
  acknowledged: it.acknowledged || false,
});

const itemToDB = (it, requestId) => ({
  id: it.id || (`${requestId}-${Date.now()}`),
  request_id: requestId,
  tactic: it.tactic || '',
  title: it.title || it.objective || '',
  product_group: it.productGroup || it.product_group || '',
  target_solutions: JSON.stringify(Array.isArray(it.targetSolutions) ? it.targetSolutions : [it.targetSolutions || '']),
  amount: Number(it.amount || 0),
  mdf_request: Number(it.mdfRequest || it.mdf_request || Math.round((it.amount || 0) * 0.5)),
  local_currency: it.localCurrency || it.local_currency || 'EUR',
  fy_half: it.fyHalf || it.fy_half || '',
  fy_quarter: it.fyQuarter || it.fy_quarter || '',
  month: it.month || '',
  period: it.period || '',
  where_field: it.where || it.where_field || '',
  target_audience: it.targetAudience || it.target_audience || '',
  target_attendees: it.targetAttendees || it.target_attendees || '',
  objective: it.objective || '',
  item_status: it.itemStatus || it.item_status || 'request_submitted',
  assigned_to: it.assignedTo || it.assigned_to || '',
  allocadia_id: it.allocadiaId || it.allocadia_id || '',
  campaign_id: it.campaignId || it.campaign_id || '',
  cancel_reason: it.cancelReason || null,
  postponed_to: it.postponedTo || null,
  acknowledged: it.acknowledged || false,
});

const dbToClaim = (r) => ({
  id: r.id, reqId: r.req_id, itemId: r.item_id,
  partner: r.partner, activity: r.activity,
  claimAmount: Number(r.claim_amount || 0),
  vatPct: Number(r.vat_pct || 0),
  totalValue: Number(r.total_value || 0),
  currency: r.currency || 'EUR',
  submitted: r.submitted,
  status: r.status,
  files: typeof r.files === 'string' ? JSON.parse(r.files) : (r.files || {}),
  notes: r.notes || '',
  fyHalf: r.fy_half, fyQuarter: r.fy_quarter, month: r.month,
  reviewNotesMarketing: r.review_notes_marketing || '',
  reviewNotesFinance: r.review_notes_finance || '',
  statusHistory: typeof r.status_history === 'string' ? JSON.parse(r.status_history) : (r.status_history || []),
});

const claimToDB = (c) => ({
  id: c.id, req_id: c.reqId, item_id: c.itemId,
  partner: c.partner, activity: c.activity,
  claim_amount: c.claimAmount, vat_pct: c.vatPct, total_value: c.totalValue,
  currency: c.currency, submitted: c.submitted, status: c.status,
  files: JSON.stringify(c.files || {}),
  notes: c.notes || '',
  fy_half: c.fyHalf, fy_quarter: c.fyQuarter, month: c.month,
  review_notes_marketing: c.reviewNotesMarketing || '',
  review_notes_finance: c.reviewNotesFinance || '',
  status_history: JSON.stringify(c.statusHistory || []),
});


const DARK_THEME = {
  bg: '#07090f',
  surface: '#0d1117',
  card: '#111926',
  border: '#1a2640',
  accent: '#3b82f6',
  accentGlow: 'rgba(59,130,246,0.12)',
  cyan: '#06b6d4',
  success: '#10b981',
  warning: '#f59e0b',
  danger: '#ef4444',
  purple: '#8b5cf6',
  teal: '#14b8a6',
  text: '#e2e8f0',
  muted: '#4b6080',
  faint: '#1e2d45',
  navBg: '#0d1117',
  navText: '#e2e8f0',
};
const LIGHT_THEME = {
  bg: '#f0f4ff',
  surface: '#ffffff',
  card: '#ffffff',
  border: '#c8d4f0',
  accent: '#1a6aff',
  accentGlow: 'rgba(26,106,255,0.10)',
  cyan: '#0055bb',
  success: '#007a3d',
  warning: '#b85c00',
  danger: '#cc0000',
  purple: '#5533cc',
  teal: '#006688',
  text: '#0a0a1a',
  muted: '#4455aa',
  faint: '#eef1fb',
  navBg: '#00008b',
  navText: '#ffffff',
};
let C = { ...DARK_THEME };

const OT_LOGO =
  'data:image/png;base64,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';

const SAMPLE_PARTNERS = [
  {
    id: 'P001',
    name: 'TechVision Ltd',
    tier: 'Platinum',
    country: 'UK',
    region: 'Europe',
    subregion: 'UK&I',
    allocated: 50000,
    spent: 32000,
    pending: 8000,
    status: 'Active',
    note: '',
    contactName: 'James Smith',
    contactEmail: 'j.smith@techvision.com',
    type: 'Reseller',
  },
  {
    id: 'P002',
    name: 'CloudSys GmbH',
    tier: 'Gold',
    country: 'Germany',
    region: 'Europe',
    subregion: 'DACH',
    allocated: 30000,
    spent: 18000,
    pending: 5000,
    status: 'Active',
    note: '',
    contactName: 'Klaus Weber',
    contactEmail: 'k.weber@cloudsys.de',
    type: 'Reseller',
  },
  {
    id: 'P003',
    name: 'Nordic IT AB',
    tier: 'Silver',
    country: 'Sweden',
    region: 'Europe',
    subregion: 'Nordics',
    allocated: 15000,
    spent: 9000,
    pending: 0,
    status: 'Active',
    note: '',
    contactName: 'Anna Lindqvist',
    contactEmail: 'a.lindqvist@nordicit.se',
    type: 'Distributor',
  },
  {
    id: 'P004',
    name: 'Benelux Solutions BV',
    tier: 'Gold',
    country: 'Netherlands',
    region: 'Europe',
    subregion: 'Benelux',
    allocated: 25000,
    spent: 10000,
    pending: 3000,
    status: 'Active',
    note: '',
    contactName: 'Erik van der Berg',
    contactEmail: 'e.berg@beneluxsol.nl',
    type: 'GSI',
  },
  {
    id: 'P005',
    name: 'Iberia Tech SL',
    tier: 'Silver',
    country: 'Spain',
    region: 'Europe',
    subregion: 'Iberia',
    allocated: 12000,
    spent: 7500,
    pending: 0,
    status: 'Active',
    note: '',
    contactName: 'Carlos Mendez',
    contactEmail: 'c.mendez@iberiatech.es',
    type: 'Reseller',
  },
  {
    id: 'P006',
    name: 'France Digitale SARL',
    tier: 'Platinum',
    country: 'France',
    region: 'Europe',
    subregion: 'France',
    allocated: 45000,
    spent: 28000,
    pending: 10000,
    status: 'Active',
    note: '',
    contactName: 'Sophie Martin',
    contactEmail: 's.martin@francedigitale.fr',
    type: 'Reseller',
  },
  {
    id: 'P007',
    name: 'Italia Cloud SpA',
    tier: 'Gold',
    country: 'Italy',
    region: 'Europe',
    subregion: 'Italy',
    allocated: 20000,
    spent: 12000,
    pending: 4000,
    status: 'Active',
    note: '',
    contactName: 'Marco Rossi',
    contactEmail: 'm.rossi@italiacloud.it',
    type: 'Reseller',
  },
  {
    id: 'P008',
    name: 'CEE Systems sro',
    tier: 'Silver',
    country: 'Czech Republic',
    region: 'Europe',
    subregion: 'CEE',
    allocated: 10000,
    spent: 4000,
    pending: 0,
    status: 'Active',
    note: '',
    contactName: 'Petr Novak',
    contactEmail: 'p.novak@ceesystems.cz',
    type: 'ISVP',
  },
  {
    id: 'P009',
    name: 'American Tech Corp',
    tier: 'Platinum',
    country: 'USA',
    region: 'US',
    subregion: 'US',
    allocated: 80000,
    spent: 55000,
    pending: 15000,
    status: 'Active',
    note: '',
    contactName: 'Michael Johnson',
    contactEmail: 'm.johnson@amtech.com',
    type: 'Reseller',
  },
  {
    id: 'P010',
    name: 'Canada Cloud Inc',
    tier: 'Gold',
    country: 'Canada',
    region: 'US',
    subregion: 'Canada',
    allocated: 35000,
    spent: 20000,
    pending: 5000,
    status: 'Active',
    note: '',
    contactName: 'Sarah Thompson',
    contactEmail: 's.thompson@canadacloud.ca',
    type: 'Reseller',
  },
  {
    id: 'P011',
    name: 'LatAm Digital SA',
    tier: 'Silver',
    country: 'Brazil',
    region: 'US',
    subregion: 'LATAM',
    allocated: 18000,
    spent: 8000,
    pending: 2000,
    status: 'Active',
    note: '',
    contactName: 'Ana Souza',
    contactEmail: 'a.souza@latamdigital.com.br',
    type: 'Distributor',
  },
  {
    id: 'P012',
    name: 'Gulf Tech LLC',
    tier: 'Gold',
    country: 'UAE',
    region: 'International',
    subregion: 'META',
    allocated: 28000,
    spent: 15000,
    pending: 6000,
    status: 'Active',
    note: '',
    contactName: 'Ahmed Al-Rashid',
    contactEmail: 'a.rashid@gulftech.ae',
    type: 'Reseller',
  },
  {
    id: 'P013',
    name: 'AsiaPac Solutions',
    tier: 'Platinum',
    country: 'Singapore',
    region: 'International',
    subregion: 'APAC',
    allocated: 40000,
    spent: 25000,
    pending: 8000,
    status: 'Active',
    note: '',
    contactName: 'Wei Chen',
    contactEmail: 'w.chen@asiapac.sg',
    type: 'GSI',
  },
  {
    id: 'P014',
    name: 'ANZ Systems Pty',
    tier: 'Gold',
    country: 'Australia',
    region: 'International',
    subregion: 'ANZ',
    allocated: 22000,
    spent: 14000,
    pending: 3000,
    status: 'Active',
    note: '',
    contactName: 'James Wilson',
    contactEmail: 'j.wilson@anzsystems.com.au',
    type: 'Reseller',
  },
];

const MONTHS = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

const SAMPLE_REQUESTS = [
  {
    id: 'REQ-001',
    partner: 'TechVision Ltd',
    submitted: '2026-01-15',
    status: 'request_submitted',
    poNumber: '',
    note: '',
    items: [
      {
        id: 'REQ-001-A',
        tactic: 'Virtual Event / Webinar',
        title: 'Webinar Cloud Security Q1',
        productGroup: 'CyberSecurity',
        targetSolutions: ['Security'],
        amount: 4000,
        localCurrency: 'EUR',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q1',
        period: 'Jan-Mar 2026',
        where: 'Online',
        targetAudience: 'IT Decision Makers',
        targetAttendees: '200',
        objective: 'Lead generation',
      },
      {
        id: 'REQ-001-B',
        tactic: 'Digital Advertising',
        title: 'LinkedIn Campaign Q1',
        productGroup: 'Content (ECS)',
        targetSolutions: ['Cloud Infrastructure'],
        amount: 2500,
        localCurrency: 'EUR',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q1',
        period: 'Jan-Mar 2026',
        where: 'LinkedIn',
        targetAudience: 'C-Suite',
        targetAttendees: '',
        objective: 'Brand awareness',
      },
    ],
  },
  {
    id: 'REQ-002',
    partner: 'CloudSys GmbH',
    submitted: '2026-01-20',
    status: 'sent_for_signature',
    poNumber: 'PO-2026-0012',
    note: 'Reviewed and BP sent',
    items: [
      {
        id: 'REQ-002-A',
        tactic: 'Trade Show / Exhibition',
        title: 'CeBIT Frankfurt',
        productGroup: 'Portfolio',
        targetSolutions: ['Cloud Infrastructure', 'Security'],
        amount: 8000,
        localCurrency: 'EUR',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q2',
        period: 'Apr-Jun 2026',
        where: 'Frankfurt',
        targetAudience: 'Enterprise IT',
        targetAttendees: '500',
        objective: 'Pipeline generation',
      },
      {
        id: 'REQ-002-B',
        tactic: 'In-Person Event',
        title: 'DACH Partner Day',
        productGroup: 'Experience (DX)',
        targetSolutions: ['AI / ML'],
        amount: 3500,
        localCurrency: 'EUR',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q1',
        period: 'Jan-Mar 2026',
        where: 'Munich',
        targetAudience: 'Partners',
        targetAttendees: '80',
        objective: 'Enablement',
      },
    ],
  },
  {
    id: 'REQ-003',
    partner: 'American Tech Corp',
    submitted: '2026-02-01',
    status: 'signed',
    poNumber: 'PO-2026-0008',
    note: 'Signed and returned',
    items: [
      {
        id: 'REQ-003-A',
        tactic: 'In-Person Event',
        title: 'East Coast Partner Summit',
        productGroup: 'Portfolio',
        targetSolutions: ['Cloud Infrastructure'],
        amount: 15000,
        localCurrency: 'USD',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q1',
        period: 'Jan-Mar 2026',
        where: 'New York',
        targetAudience: 'Enterprise',
        targetAttendees: '300',
        objective: 'Pipeline',
      },
    ],
  },
  {
    id: 'REQ-004',
    partner: 'Nordic IT AB',
    submitted: '2026-02-10',
    status: 'rejected',
    poNumber: '',
    note: 'Outside MDF eligible activities',
    items: [
      {
        id: 'REQ-004-A',
        tactic: 'Print Advertising',
        title: 'Magazine Ad Campaign',
        productGroup: 'Content (ECS)',
        targetSolutions: ['Other'],
        amount: 5000,
        localCurrency: 'SEK',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q2',
        period: 'Apr-Jun 2026',
        where: 'Sweden',
        targetAudience: 'SMB',
        targetAttendees: '',
        objective: 'Brand',
      },
    ],
  },
  {
    id: 'REQ-005',
    partner: 'France Digitale SARL',
    submitted: '2026-02-15',
    status: 'request_submitted',
    poNumber: '',
    note: '',
    items: [
      {
        id: 'REQ-005-A',
        tactic: 'Virtual Event / Webinar',
        title: 'Webinar IA et Securite',
        productGroup: 'CyberSecurity',
        targetSolutions: ['Security', 'AI / ML'],
        amount: 3000,
        localCurrency: 'EUR',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q2',
        period: 'Apr-Jun 2026',
        where: 'Online',
        targetAudience: 'IT Managers',
        targetAttendees: '150',
        objective: 'Leads',
      },
      {
        id: 'REQ-005-B',
        tactic: 'Content Syndication',
        title: 'Gartner Report Syndication',
        productGroup: 'Content (ECS)',
        targetSolutions: ['Cloud Infrastructure'],
        amount: 4500,
        localCurrency: 'EUR',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q2',
        period: 'Apr-Jun 2026',
        where: 'Online',
        targetAudience: 'C-Level',
        targetAttendees: '',
        objective: 'Awareness',
      },
      {
        id: 'REQ-005-C',
        tactic: 'Digital Advertising',
        title: 'Google Ads Campaign',
        productGroup: 'Experience (DX)',
        targetSolutions: ['Cloud Infrastructure'],
        amount: 2000,
        localCurrency: 'EUR',
        fyHalf: 'FY26 H1',
        fyQuarter: 'Q2',
        period: 'Apr-Jun 2026',
        where: 'Online',
        targetAudience: 'SMB',
        targetAttendees: '',
        objective: 'Lead gen',
      },
    ],
  },
];

const SAMPLE_HISTORY = [
  {
    id: 1,
    ts: '2025-03-08 09:14',
    user: 'Marco R.',
    action: 'Approveta richiesta MDF-002',
    entity: 'MDF-002',
    type: 'approve',
  },
  {
    id: 2,
    ts: '2025-03-07 16:32',
    user: 'Sara B.',
    action: 'Updated spent budget: CloudSys SpA +6,000',
    entity: 'CloudSys SpA',
    type: 'edit',
  },
  {
    id: 3,
    ts: '2025-03-06 11:05',
    user: 'Marco R.',
    action: 'New request MDF-008 created',
    entity: 'MDF-008',
    type: 'create',
  },
  {
    id: 4,
    ts: '2025-03-05 14:20',
    user: 'Luca T.',
    action: 'Notes added: NexGen Solutions',
    entity: 'NexGen Solutions',
    type: 'note',
  },
  {
    id: 5,
    ts: '2025-03-04 10:00',
    user: 'Sara B.',
    action: 'Rejectta richiesta MDF-006',
    entity: 'MDF-006',
    type: 'reject',
  },
];

function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  var s = String(val)
    .replace(/[^0-9.,]/g, '')
    .trim();
  if (!s) return 0;
  var lc = s.lastIndexOf(','),
    ld = s.lastIndexOf('.');
  if (lc > ld) {
    var ac = s.slice(lc + 1);
    return ac.length === 3
      ? parseFloat(s.replace(/,/g, '')) || 0
      : parseFloat(s.replace(/\./g, '').replace(',', '.')) || 0;
  }
  if (ld > lc) {
    var ad = s.slice(ld + 1);
    return ad.length === 3 && lc === -1
      ? parseFloat(s.replace(/\./g, '')) || 0
      : parseFloat(s.replace(/,/g, '')) || 0;
  }
  return parseFloat(s) || 0;
}
const fmtEUR = (n) => 'EUR ' + Number(n || 0).toLocaleString('it-IT');
const fmtUSD = (n, r) =>
  'USD ' + Math.round(Number(n || 0) * (r || 1)).toLocaleString('en-US');
const CURRENCY_SYMBOLS = {
  USD: 'USD ',
  EUR: '\u20ac',
  GBP: '\u00a3',
  CHF: 'CHF ',
  SEK: 'kr ',
  NOK: 'kr ',
  DKK: 'kr ',
  PLN: 'z\u0142 ',
  CZK: 'K\u010d ',
  AED: 'AED ',
  SGD: 'S$ ',
  AUD: 'A$ ',
};
const fmtLC = (amount, currency) => {
  const sym = CURRENCY_SYMBOLS[currency] || currency + ' ';
  return sym + Number(amount || 0).toLocaleString('en-US');
};

// -- WORKFLOW STATUSES ---------------------------------------------------------
const WORKFLOW_STEPS = [
  {
    id: 'request_submitted',
    label: 'Submitted',
    short: 'Submitted',
    color: '#f59e0b',
    icon: '1',
  },
  {
    id: 'approved',
    label: 'In Review',
    short: 'In Review',
    color: '#f59e0b',
    icon: '2',
  },
  {
    id: 'sent_for_signature',
    label: 'Sent for Signature',
    short: 'Sent for Sign.',
    color: '#8b5cf6',
    icon: '3',
  },
  {
    id: 'signed',
    label: 'Signed',
    short: 'Signed',
    color: '#06b6d4',
    icon: '4',
  },
  {
    id: 'po_raised',
    label: 'Approved',
    short: 'Approved',
    color: '#10b981',
    icon: '5',
  },
  {
    id: 'rejected',
    label: 'Rejected',
    short: 'Rejected',
    color: '#ef4444',
    icon: 'x',
  },
  {
    id: 'cancelled_by_partner',
    label: 'Cancelled by Partner',
    short: 'Cancelled',
    color: '#ef4444',
    icon: 'x',
  },
  {
    id: 'postponed',
    label: 'Postponed',
    short: 'Postponed',
    color: '#f59e0b',
    icon: '~',
  },
];
const STATUS_COLOR = Object.fromEntries(
  WORKFLOW_STEPS.map((s) => [s.id, s.color])
);
const STATUS_LABEL = Object.fromEntries(
  WORKFLOW_STEPS.map((s) => [s.id, s.label])
);
const STATUS_SHORT = Object.fromEntries(
  WORKFLOW_STEPS.map((s) => [s.id, s.short])
);
const getStatusColor = (s) => STATUS_COLOR[s] || '#4b6080';
const getStatusLabel = (s) => STATUS_LABEL[s] || s;
const getStatusShort = (s) => STATUS_SHORT[s] || s;
const tierColor = { Platinum: '#e2e8f0', Gold: '#f59e0b', Silver: '#94a3b8' };
const histColor = {
  approve: C.success,
  reject: C.danger,
  edit: C.accent,
  create: C.purple,
  note: C.warning,
};
const claimStatusColor = {
  draft: C.muted,
  submitted: C.accent,
  marketing_review: C.warning,
  finance_review: C.purple,
  approved: C.success,
  rejected: C.danger,
  paid: C.teal,
};
const claimStatusLabel = {
  draft: 'Draft',
  submitted: 'Submitted',
  marketing_review: 'Mktg Review',
  finance_review: 'Finance Review',
  approved: 'Approved',
  rejected: 'Rejected',
  paid: 'Paid',
};

const SAMPLE_CLAIMS = [
  {
    id: 'CLM-001',
    reqId: 'REQ-001',
    itemId: 'REQ-001-A',
    partner: 'TechVision Ltd',
    activity: 'Cloud Security Webinar Q1',
    claimAmount: 1950,
    vatPct: 20,
    totalValue: 2340,
    currency: 'EUR',
    submitted: '2026-10-15',
    status: 'submitted',
    files: {
      partnerInvoice: 'TechVision_Invoice_Oct2026.pdf',
      thirdParty: null,
      inHouse: null,
      merchandise: null,
      additional: [],
    },
    notes:
      'Webinar delivered July 28. 87 registrations, 62 attendees. 14 qualified leads.',
    fyHalf: 'FY26 H1',
    fyQuarter: 'Q1',
    month: 'July',
    reviewNotesMarketing: '',
    reviewNotesFinance: '',
    statusHistory: [
      { status: 'submitted', by: 'TechVision Ltd', at: '2026-10-15 09:00', note: 'Claim submitted via portal' },
    ],
  },
  {
    id: 'CLM-002',
    reqId: 'REQ-002',
    itemId: 'REQ-002-A',
    partner: 'CloudSys GmbH',
    activity: 'CeBIT Frankfurt 2026',
    claimAmount: 4200,
    vatPct: 19,
    totalValue: 4998,
    currency: 'EUR',
    submitted: '2026-10-20',
    status: 'marketing_review',
    files: {
      partnerInvoice: 'CloudSys_Invoice_2026.pdf',
      thirdParty: 'Venue_Invoice_Frankfurt.pdf',
      inHouse: null,
      merchandise: 'Merch_receipts.pdf',
      additional: ['Event_photos.zip'],
    },
    notes:
      'Full day event in Frankfurt. 34 partners attended. Venue + catering + materials.',
    fyHalf: 'FY26 H1',
    fyQuarter: 'Q2',
    month: 'October',
    reviewNotesMarketing: 'Checking venue invoice details.',
    reviewNotesFinance: '',
    statusHistory: [
      { status: 'submitted', by: 'CloudSys GmbH', at: '2026-10-20 11:30', note: 'Claim submitted via portal' },
      { status: 'marketing_review', by: 'Decio A.', at: '2026-10-21 09:15', note: 'Checking venue invoice details.' },
    ],
  },
  {
    id: 'CLM-003',
    reqId: 'REQ-003',
    itemId: 'REQ-003-A',
    partner: 'American Tech Corp',
    activity: 'East Coast Partner Summit',
    claimAmount: 8500,
    vatPct: 0,
    totalValue: 8500,
    currency: 'USD',
    submitted: '2026-10-28',
    status: 'finance_review',
    files: {
      partnerInvoice: 'AmericanTech_Invoice.pdf',
      thirdParty: 'Summit_Organiser_Invoice.pdf',
      inHouse: null,
      merchandise: null,
      additional: ['Attendance_report.pdf', 'Lead_list.xlsx'],
    },
    notes:
      'Partner summit in New York. 320 attendees, 28 qualified leads. Strong brand exposure.',
    fyHalf: 'FY26 H1',
    fyQuarter: 'Q1',
    month: 'September',
    reviewNotesMarketing: 'Approved by marketing - strong ROI.',
    reviewNotesFinance: 'Checking documentation.',
    statusHistory: [
      { status: 'submitted', by: 'American Tech Corp', at: '2026-10-28 14:00', note: 'Claim submitted via portal' },
      { status: 'marketing_review', by: 'Decio A.', at: '2026-10-29 10:00', note: 'Started marketing review' },
      { status: 'finance_review', by: 'Decio A.', at: '2026-10-30 11:45', note: 'Approved by marketing - strong ROI.' },
    ],
  },
  {
    id: 'CLM-004',
    reqId: 'REQ-004',
    itemId: 'REQ-004-B',
    partner: 'France Digitale',
    activity: 'Gartner Report Syndication',
    claimAmount: 2100,
    vatPct: 20,
    totalValue: 2520,
    currency: 'EUR',
    submitted: '2026-01-20',
    status: 'approved',
    files: {
      partnerInvoice: 'FranceDigitale_Invoice.pdf',
      thirdParty: null,
      inHouse: null,
      merchandise: null,
      additional: [],
    },
    notes: 'Gartner report syndication campaign. 180 downloads, 22 MQL.',
    fyHalf: 'FY26 H1',
    fyQuarter: 'Q2',
    month: 'December',
    reviewNotesMarketing: 'Approved.',
    reviewNotesFinance: 'VAT verified. Ready to pay.',
    statusHistory: [
      { status: 'submitted', by: 'France Digitale', at: '2026-01-20 08:00', note: 'Claim submitted via portal' },
      { status: 'marketing_review', by: 'Kaila', at: '2026-01-21 09:00', note: 'Started review' },
      { status: 'finance_review', by: 'Kaila', at: '2026-01-22 14:30', note: 'Approved.' },
      { status: 'approved', by: 'Finance', at: '2026-01-23 10:00', note: 'VAT verified. Ready to pay.' },
    ],
  },
  {
    id: 'CLM-005',
    reqId: 'REQ-001',
    itemId: 'REQ-001-B',
    partner: 'TechVision Ltd',
    activity: 'LinkedIn Campaign Q1',
    claimAmount: 1200,
    vatPct: 20,
    totalValue: 1440,
    currency: 'EUR',
    submitted: '2026-10-29',
    status: 'on_hold',
    files: {
      partnerInvoice: 'TechVision_LinkedIn_Invoice.pdf',
      thirdParty: null,
      inHouse: null,
      merchandise: null,
      additional: [],
    },
    notes: 'LinkedIn sponsored content campaign. 4,200 impressions, 87 clicks.',
    fyHalf: 'FY26 H1',
    fyQuarter: 'Q1',
    month: 'August',
    reviewNotesMarketing:
      'On hold - partner invoice missing campaign dates. Please resubmit.',
    reviewNotesFinance: '',
    statusHistory: [
      { status: 'submitted', by: 'TechVision Ltd', at: '2026-10-29 16:00', note: 'Claim submitted via portal' },
      { status: 'marketing_review', by: 'Decio A.', at: '2026-10-30 09:30', note: 'Started review' },
      { status: 'on_hold', by: 'Decio A.', at: '2026-10-30 10:15', note: 'On hold - partner invoice missing campaign dates. Please resubmit.' },
    ],
  },
  {
    id: 'CLM-006',
    reqId: 'REQ-003',
    itemId: 'REQ-003-A',
    partner: 'American Tech Corp',
    activity: 'East Coast Partner Summit',
    claimAmount: 3800,
    vatPct: 0,
    totalValue: 3800,
    currency: 'USD',
    submitted: '2026-04-10',
    status: 'paid',
    files: {
      partnerInvoice: 'AmericanTech_Invoice_Q4.pdf',
      thirdParty: null,
      inHouse: null,
      merchandise: null,
      additional: [],
    },
    notes: 'Q4 partner summit expenses.',
    fyHalf: 'FY25 H2',
    fyQuarter: 'Q4',
    month: 'April',
    reviewNotesMarketing: 'Approved.',
    reviewNotesFinance: 'Paid on 2026-04-28. Bank transfer confirmed.',
    statusHistory: [
      { status: 'submitted', by: 'American Tech Corp', at: '2026-04-10 09:00', note: 'Claim submitted via portal' },
      { status: 'marketing_review', by: 'Umair', at: '2026-04-11 10:00', note: 'Started review' },
      { status: 'finance_review', by: 'Umair', at: '2026-04-12 11:00', note: 'Approved.' },
      { status: 'approved', by: 'Finance', at: '2026-04-15 09:00', note: 'All docs verified.' },
      { status: 'paid', by: 'Finance', at: '2026-04-28 14:00', note: 'Paid on 2026-04-28. Bank transfer confirmed.' },
    ],
  },
];
// --- Persistence helpers ---
const STORAGE_KEY = 'mdf_manager_data';
const loadPersistedData = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch { return null; }
};
const persistedData = loadPersistedData();


const StatusBadge = ({ status }) => {
  const cfg = {
    active: { c: C.success, l: 'Active' },
    inactive: { c: C.muted, l: 'Inactive' },
    'at risk': { c: C.warning, l: 'At Risk' },
    request_submitted: { c: C.warning, l: 'Request Submitted' },
    approved: { c: '#10b981', l: 'Approved' },
    sent_for_signature: { c: C.purple, l: 'Sent for Signature' },
    signed: { c: C.cyan || '#06b6d4', l: 'Signed' },
    po_raised: { c: C.accent, l: 'PO Raised' },
    approved_and_signed: { c: C.success, l: 'Approved and Signed' },
    rejected: { c: C.danger, l: 'Rejected' },
    cancelled_by_partner: { c: C.danger, l: 'Cancelled by Partner' },
    postponed: { c: C.warning, l: 'Postponed' },
    claim_submitted: { c: C.cyan || '#06b6d4', l: 'Claim Submitted' },
    paid: { c: C.teal || '#14b8a6', l: 'Paid' },
    on_hold: { c: C.danger, l: 'On Hold' },
    marketing_review: { c: C.cyan || '#06b6d4', l: 'Mktg Review' },
    finance_review: { c: C.purple, l: 'Finance Review' },
  };
  const { c, l } = cfg[status] || { c: C.muted, l: status || 'Unknown' };
  return (
    <span
      style={{
        background: c + '18',
        color: c,
        border: `1px solid ${c}35`,
        borderRadius: 20,
        padding: '2px 10px',
        fontSize: 10,
        fontWeight: 800,
        letterSpacing: '0.06em',
        textTransform: 'uppercase',
      }}
    >
      {l}
    </span>
  );
};

const Bar = ({ value, max, color = C.accent, height = 5 }) => {
  const pct = Math.min((value / (max || 1)) * 100, 100);
  const col = pct > 90 ? C.danger : pct > 75 ? C.warning : color;
  return (
    <div
      style={{
        background: C.faint,
        borderRadius: 4,
        height,
        overflow: 'hidden',
      }}
    >
      <div
        style={{
          width: `${pct}%`,
          height: '100%',
          background: col,
          borderRadius: 4,
          transition: 'width 0.8s ease',
        }}
      />
    </div>
  );
};

const CurrencyToggle = ({ currency, onChange, rate, loading }) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
    {!loading && rate && (
      <span style={{ fontSize: 11, color: C.muted, fontFamily: 'monospace' }}>
        1 EUR = {rate.toFixed(4)} USD
      </span>
    )}
    {loading && <span style={{ fontSize: 11, color: C.muted }}>...</span>}
    <div
      style={{
        display: 'flex',
        background: C.faint,
        borderRadius: 8,
        padding: 3,
        border: `1px solid ${C.border}`,
      }}
    >
      {['EUR', 'USD'].map((cur) => (
        <button
          key={cur}
          onClick={() => onChange(cur)}
          style={{
            background: currency === cur ? C.accent : 'transparent',
            color: currency === cur ? '#fff' : C.muted,
            border: 'none',
            borderRadius: 6,
            padding: '5px 14px',
            fontWeight: 700,
            fontSize: 12,
            cursor: 'pointer',
            fontFamily: 'monospace',
          }}
        >
          {cur}
        </button>
      ))}
    </div>
  </div>
);

const SparkBar = ({ data, colors, labels, height = 120 }) => {
  const max = Math.max(...data.map((d) => d.reduce((a, b) => a + b, 0)), 1);
  const w = 100 / data.length;
  return (
    <svg
      viewBox={'0 0 100 ' + height}
      style={{ width: '100%', height }}
      preserveAspectRatio="none"
    >
      {data.map((bars, i) => {
        let y = height;
        return bars.map((val, j) => {
          const h = (val / max) * height * 0.9;
          y -= h;
          return (
            <rect
              key={j}
              x={i * w + w * 0.1}
              y={y}
              width={w * 0.8}
              height={h}
              fill={colors[j] || C.accent}
              rx="2"
            />
          );
        });
      })}
    </svg>
  );
};

const LineChart = ({ data, color = C.accent, height = 80, fill = false }) => {
  if (!data || data.length < 2) return null;
  const max = Math.max(...data, 1),
    min = 0;
  const pts = data
    .map(
      (v, i) =>
        `${(i / (data.length - 1)) * 100},${
          100 - ((v - min) / (max - min)) * 90
        }`
    )
    .join(' ');
  const area = `0,100 ${pts} 100,100`;
  return (
    <svg
      viewBox="0 0 100 100"
      style={{ width: '100%', height }}
      preserveAspectRatio="none"
    >
      {fill && <polygon points={area} fill={color + '22'} />}
      <polyline
        points={pts}
        fill="none"
        stroke={color}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      {data.map((v, i) => (
        <circle
          key={i}
          cx={(i / (data.length - 1)) * 100}
          cy={100 - ((v - min) / (max - min)) * 90}
          r="2.5"
          fill={color}
        />
      ))}
    </svg>
  );
};

const DonutChart = ({ slices, size = 120 }) => {
  const total = slices.reduce((s, x) => s + x.value, 0) || 1;
  let angle = -90;
  const r = 40,
    cx = 50,
    cy = 50,
    stroke = 18;
  const toXY = (a, radius) => ({
    x: cx + radius * Math.cos((a * Math.PI) / 180),
    y: cy + radius * Math.sin((a * Math.PI) / 180),
  });
  return (
    <svg viewBox="0 0 100 100" style={{ width: size, height: size }}>
      <circle
        cx={cx}
        cy={cy}
        r={r}
        fill="none"
        stroke={C.faint}
        strokeWidth={stroke}
      />
      {slices.length === 1 ? (
        <circle cx={cx} cy={cy} r={r} fill={slices[0].color} opacity="0.85" />
      ) : slices.map((s, i) => {
        const sweep = (s.value / total) * 360;
        if (sweep === 0) return null;
        const start = angle,
          end = angle + sweep;
        angle = end;
        const large = sweep > 180 ? 1 : 0;
        const s1 = toXY(start, r),
          e1 = toXY(end, r);
        return (
          <path
            key={i}
            d={`M ${cx} ${cy} L ${s1.x} ${s1.y} A ${r} ${r} 0 ${large} 1 ${e1.x} ${e1.y} Z`}
            fill={s.color}
            opacity="0.85"
          />
        );
      })}
      <circle cx={cx} cy={cy} r={r - stroke / 2} fill={C.card} />
    </svg>
  );
};

const TACTICS = [
  'Customer Assessment',
  'Assets Development',
  'Customer Events',
  'Digital Advertising',
  'Digital Content',
  'Electronic Direct Mail',
  'Enablement Activity',
  'Event Sponsorship',
  'MDF Champion',
  'Partner Sales Incentive',
  'Print Advertising',
  'Print Brochure',
  'Print Direct Mail',
  'Social Media',
  'Search Engine Marketing',
  'Telemarketing',
  'Solution for Demo/Dev/Eval',
  'Sponsoring OT Event',
  'E2E Campaign',
  'Content Syndication',
  'SPIFF',
];
const PRODUCT_GROUPS = [
  'Content (ECS)',
  'Experience (DX)',
  'Experience / Content',
  'OSM',
  'Portfolio',
  'CyberSecurity',
  'Cross BU',
];
const PARTNER_TYPES = ['Reseller', 'Distributor', 'GSI', 'ISVP'];
const SOLUTIONS = [
  'Cloud Infrastructure',
  'Security',
  'Networking',
  'Data & Analytics',
  'AI / ML',
  'Collaboration',
  'DevOps',
  'ERP / CRM',
  'Other',
];

const TEAM_MEMBERS = ['Decio A.', 'Kaila', 'Umair'];

const REGIONS = {
  Europe: [
    'UK&I',
    'Nordics',
    'Benelux',
    'France',
    'Iberia',
    'Italy',
    'DACH',
    'CEE',
  ],
  US: ['US', 'Canada', 'LATAM'],
  International: ['META', 'APAC', 'ANZ'],
};
const ALL_MACROS = Object.keys(REGIONS);
const ALL_SUBREGIONS = Object.values(REGIONS).flat();
const getSubregions = (macro) => REGIONS[macro] || ALL_SUBREGIONS;
const getMacro = (subregion) =>
  Object.entries(REGIONS).find(([, subs]) => subs.includes(subregion))?.[0] ||
  '';

const FormField = ({ label, req, error, children }) => (
  <div>
    <div
      style={{ fontSize: 12, color: C.muted, fontWeight: 700, marginBottom: 6 }}
    >
      {label}
      {req && <span style={{ color: C.danger }}> *</span>}
    </div>
    {children}
    {error && (
      <div style={{ fontSize: 11, color: C.danger, marginTop: 4 }}>{error}</div>
    )}
  </div>
);
const inpStyle = (err) => ({
  width: '100%',
  background: C.surface,
  border: `1px solid ${err ? C.danger : C.border}`,
  color: C.text,
  borderRadius: 8,
  padding: '10px 12px',
  fontSize: 13,
  outline: 'none',
  fontFamily: 'inherit',
});

const LocalTextarea = ({ value, onCommit, placeholder, height = 60 }) => {
  const [local, setLocal] = useState(value || '');
  useEffect(() => {
    setLocal(value || '');
  }, [value]);
  return (
    <textarea
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={() => onCommit(local)}
      placeholder={placeholder}
      style={{
        width: '100%',
        background: C.surface,
        border: `1px solid ${C.border}`,
        color: C.text,
        borderRadius: 8,
        padding: 10,
        fontSize: 12,
        resize: 'none',
        height,
        marginBottom: 10,
        fontFamily: 'inherit',
        outline: 'none',
      }}
    />
  );
};

const ApprovelNotes = ({ note, setNotes }) => {
  const [local, setLocal] = useState(note || '');
  useEffect(() => {
    setLocal(note || '');
  }, [note]);
  return (
    <textarea
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={() => setNotes(local)}
      placeholder="Approvel note..."
      style={{
        width: '100%',
        background: C.bg,
        border: `1px solid ${C.border}`,
        color: C.text,
        borderRadius: 10,
        padding: 10,
        fontSize: 12,
        resize: 'none',
        height: 60,
        marginBottom: 10,
        fontFamily: 'inherit',
        outline: 'none',
      }}
    />
  );
};

const NewRequestModal = ({
  partners,
  onAdd,
  onClose,
  currentUser,
  portalPartner,
  isPartnerView,
}) => {
  const F = FormField;
  const inp = inpStyle;

  const FY_HALVES = [
    'FY26 H1',
    'FY26 H2',
    'FY27 H1',
    'FY27 H2',
    'FY28 H1',
    'FY28 H2',
  ];
  const QUARTERS_BY_HALF = { H1: ['Q1', 'Q2'], H2: ['Q3', 'Q4'] };
  const MONTHS_BY_QUARTER = {
    Q1: ['July', 'August', 'September'],
    Q2: ['October', 'November', 'December'],
    Q3: ['January', 'February', 'March'],
    Q4: ['April', 'May', 'June'],
  };

  const emptyItem = () => ({
    id: Date.now() + Math.random(),
    fyHalf: '',
    fyQuarter: '',
    month: '',
    productGroup: '',
    tactic: '',
    location: '',
    targetAudience: '',
    targetSolutions: '',
    objective: '',
    totalCost: '',
    currency: 'EUR',
    mdfRequest: '',
  });

  const [partnerInput, setPartnerInput] = useState(
    isPartnerView ? portalPartner?.name || '' : ''
  );
  const [partnerType, setPartnerType] = useState(
    isPartnerView ? portalPartner?.type || '' : ''
  );
  const [partnerTier, setPartnerTier] = useState(
    isPartnerView ? portalPartner?.tier || '' : ''
  );
  const [partnerContact, setPartnerContact] = useState(
    isPartnerView ? portalPartner?.contactName || '' : ''
  );
  const [partnerEmail, setPartnerEmail] = useState(
    isPartnerView ? portalPartner?.contactEmail || '' : ''
  );
  const [partnerManager, setPartnerManager] = useState(
    isPartnerView ? portalPartner?.accountManager || '' : ''
  );
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [items, setItems] = useState([emptyItem()]);
  const [errors, setErrors] = useState({});

  const suggestions =
    partnerInput.length > 0
      ? partners
          .filter((p) =>
            p.name.toLowerCase().includes(partnerInput.toLowerCase())
          )
          .slice(0, 6)
      : [];

  const setItem = (idx, k, v) =>
    setItems((p) =>
      p.map((it, i) => {
        if (i !== idx) return it;
        const updated = { ...it, [k]: v };

        if (k === 'fyHalf') {
          updated.fyQuarter = '';
          updated.month = '';
        }
        if (k === 'fyQuarter') {
          updated.month = '';
        }

        if (k === 'totalCost') {
          const num = parseNum(String(v));
          if (num > 0) updated.mdfRequest = String(Math.round(num * 0.5));
        }
        return updated;
      })
    );

  const getHalfKey = (fyHalf) =>
    fyHalf.includes('H1') ? 'H1' : fyHalf.includes('H2') ? 'H2' : '';
  const getQuarters = (fyHalf) => QUARTERS_BY_HALF[getHalfKey(fyHalf)] || [];
  const getMonths = (fyQuarter) => MONTHS_BY_QUARTER[fyQuarter] || [];

  const validate = () => {
    const e = {};
    if (!partnerInput.trim()) e.partner = 'Partner name required';
    if (!partnerType) e.partnerType = 'Required';
    if (!partnerTier) e.partnerTier = 'Required';
    if (!partnerContact) e.partnerContact = 'Required';
    if (!partnerEmail) e.partnerEmail = 'Required';
    if (!partnerManager) e.partnerManager = 'Required';
    items.forEach((it, i) => {
      const pf = (f, msg) => {
        if (!it[f] || !String(it[f]).trim()) e[`${f}_${i}`] = msg;
      };
      pf('fyHalf', 'Required');
      pf('fyQuarter', 'Required');
      pf('month', 'Required');
      pf('productGroup', 'Required');
      pf('tactic', 'Required');
      pf('location', 'Required');
      pf('targetAudience', 'Required');
      pf('targetSolutions', 'Required');
      pf('objective', 'Required');
      pf('totalCost', 'Required');
      pf('mdfRequest', 'Required');
      if (it.totalCost && isNaN(parseNum(String(it.totalCost))))
        e[`totalCost_${i}`] = 'Invalid amount';
      if (!e[`totalCost_${i}`] && parseNum(String(it.totalCost || 0)) <= 0)
        e[`totalCost_${i}`] = 'Must be greater than 0';
      if (it.mdfRequest && isNaN(parseNum(String(it.mdfRequest))))
        e[`mdfRequest_${i}`] = 'Invalid amount';
      if (parseNum(String(it.mdfRequest)) > parseNum(String(it.totalCost)))
        e[`mdfRequest_${i}`] = 'MDF cannot exceed total cost';
    });
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const submit = () => {
    if (!validate()) return;
    const reqId = 'REQ-' + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2,5).toUpperCase();
    onAdd({
      id: reqId,
      partner: partnerInput.trim(),
      partnerType: partnerType,
      partnerTier: partnerTier,
      partnerContact: partnerContact,
      partnerEmail: partnerEmail,
      partnerManager: partnerManager,
      submitted: new Date().toISOString().slice(0, 10),
      status: 'request_submitted',
      poNumber: '',
      note: '',
      items: items.map((it, i) => ({
        id: `${reqId}-${String.fromCharCode(65 + i)}`,
        fyHalf: it.fyHalf,
        fyQuarter: it.fyQuarter,
        month: it.month,
        period: `${it.month} (${it.fyQuarter})`,
        productGroup: it.productGroup,
        tactic: it.tactic,
        where: it.location,
        targetAudience: it.targetAudience,
        targetSolutions: [it.targetSolutions],
        objective: it.objective,
        amount: parseNum(String(it.totalCost)),
        mdfRequest: parseNum(String(it.mdfRequest)),
        localCurrency: it.currency,
        title: it.objective.slice(0, 50),
      })),
    });
    onClose();
  };

  const err = (key) =>
    errors[key] ? { border: `1.5px solid ${C.danger}` } : {};
  const ErrMsg = ({ k }) =>
    errors[k] ? (
      <div style={{ fontSize: 10, color: C.danger, marginTop: 3 }}>
        ^ {errors[k]}
      </div>
    ) : null;

  return (
    <div
      data-modal="true"
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.88)',
        zIndex: 300,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <div
        style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 20,
          padding: 32,
          width: 680,
          maxHeight: '92vh',
          overflowY: 'auto',
          boxSizing: 'border-box',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: 24,
          }}
        >
          <div>
            <div
              style={{
                fontFamily: "'Syne',sans-serif",
                fontWeight: 800,
                fontSize: 22,
              }}
            >
              New MDF Request
            </div>
            <div style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>
              All fields are mandatory
            </div>
          </div>
          <button
            onClick={onClose}
            style={{
              background: 'transparent',
              border: 'none',
              color: C.muted,
              fontSize: 22,
              cursor: 'pointer',
              lineHeight: 1,
            }}
          >
            x
          </button>
        </div>

        {/* Partner */}
        <div
          style={{
            background: C.faint,
            borderRadius: 12,
            padding: '14px 16px',
            marginBottom: 20,
          }}
        >
          <div
            style={{
              fontSize: 10,
              fontWeight: 800,
              color: C.accent,
              letterSpacing: '0.12em',
              marginBottom: 12,
            }}
          >
            PARTNER
          </div>
          {/* Partner Name with autocomplete */}
          {isPartnerView ? (
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 12 }}>
              {partnerInput}
            </div>
          ) : (
            <div style={{ position: 'relative', marginBottom: 12 }}>
              <label
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: C.muted,
                  display: 'block',
                  marginBottom: 4,
                }}
              >
                Partner Name *
              </label>
              <input
                value={partnerInput}
                onChange={(e) => {
                  setPartnerInput(e.target.value);
                  setShowSuggestions(true);

                  const match = partners.find(
                    (p) => p.name.toLowerCase() === e.target.value.toLowerCase()
                  );
                  if (match) {
                    setPartnerType(match.type || '');
                    setPartnerTier(match.tier || '');
                    setPartnerContact(match.contactName || '');
                    setPartnerEmail(match.contactEmail || '');
                    setPartnerManager(match.accountManager || '');
                  }
                }}
                onFocus={() => setShowSuggestions(true)}
                onBlur={() => setTimeout(() => setShowSuggestions(false), 150)}
                placeholder="Type partner name or add new..."
                style={{
                  ...inp(errors.partner),
                  ...err('partner'),
                  width: '100%',
                }}
              />
              <ErrMsg k="partner" />
              {showSuggestions && suggestions.length > 0 && (
                <div
                  style={{
                    position: 'absolute',
                    top: '100%',
                    left: 0,
                    right: 0,
                    background: C.surface,
                    border: `1px solid ${C.border}`,
                    borderRadius: 10,
                    zIndex: 10,
                    boxShadow: '0 8px 24px rgba(0,0,0,0.3)',
                    marginTop: 4,
                    overflow: 'hidden',
                  }}
                >
                  {suggestions.map((p) => (
                    <div
                      key={p.id}
                      onClick={() => {
                        setPartnerInput(p.name);
                        setPartnerType(p.type || '');
                        setPartnerTier(p.tier || '');
                        setPartnerContact(p.contactName || '');
                        setPartnerEmail(p.contactEmail || '');
                        setPartnerManager(p.accountManager || '');
                        setShowSuggestions(false);
                      }}
                      style={{
                        padding: '10px 14px',
                        cursor: 'pointer',
                        fontSize: 13,
                        borderBottom: `1px solid ${C.border}20`,
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                      }}
                    >
                      <span style={{ fontWeight: 600 }}>{p.name}</span>
                      <span style={{ fontSize: 11, color: C.muted }}>
                        {p.tier} . {p.type} . {p.subregion}
                      </span>
                    </div>
                  ))}
                  {partnerInput &&
                    !partners.find(
                      (p) => p.name.toLowerCase() === partnerInput.toLowerCase()
                    ) && (
                      <div
                        style={{
                          padding: '10px 14px',
                          fontSize: 13,
                          color: C.accent,
                          fontWeight: 600,
                          borderTop: `1px solid ${C.border}`,
                        }}
                      >
                        + Add "{partnerInput}" as new partner
                      </div>
                    )}
                </div>
              )}
            </div>
          )}
          {/* Partner details - 3 col grid */}
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: '1fr 1fr 1fr',
              gap: 12,
              marginBottom: 12,
            }}
          >
            <F label="Partner Type" req error={errors.partnerType}>
              <select
                value={partnerType}
                onChange={(e) => setPartnerType(e.target.value)}
                style={{ ...inp(errors.partnerType), width: '100%' }}
              >
                <option value="">Select...</option>
                {PARTNER_TYPES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </F>
            <F label="Partner Level" req error={errors.partnerTier}>
              <select
                value={partnerTier}
                onChange={(e) => setPartnerTier(e.target.value)}
                style={{ ...inp(errors.partnerTier), width: '100%' }}
              >
                <option value="">Select...</option>
                {['Platinum', 'Gold', 'Silver'].map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </F>
            <F label="OT Partner Manager" req error={errors.partnerManager}>
              <input
                value={partnerManager}
                onChange={(e) => setPartnerManager(e.target.value)}
                placeholder="e.g. Jane Smith"
                style={{ ...inp(errors.partnerManager), width: '100%' }}
              />
            </F>
          </div>
          <div
            style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}
          >
            <F label="Partner Contact Name" req error={errors.partnerContact}>
              <input
                value={partnerContact}
                onChange={(e) => setPartnerContact(e.target.value)}
                placeholder="e.g. Marco Rossi"
                style={{ ...inp(errors.partnerContact), width: '100%' }}
              />
            </F>
            <F label="Partner Contact Email" req error={errors.partnerEmail}>
              <input
                value={partnerEmail}
                onChange={(e) => setPartnerEmail(e.target.value)}
                placeholder="e.g. m.rossi@partner.com"
                type="email"
                style={{ ...inp(errors.partnerEmail), width: '100%' }}
              />
            </F>
          </div>
        </div>

        {/* Activities */}
        <div
          style={{
            fontSize: 10,
            fontWeight: 800,
            color: C.accent,
            letterSpacing: '0.12em',
            marginBottom: 10,
          }}
        >
          ACTIVITIES ({items.length})
        </div>

        {items.map((item, idx) => (
          <div
            key={item.id}
            style={{
              background: C.faint,
              border: `1px solid ${C.border}`,
              borderRadius: 12,
              padding: 16,
              marginBottom: 14,
            }}
          >
            {/* Activity header */}
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginBottom: 14,
              }}
            >
              <div style={{ fontSize: 12, fontWeight: 800, color: C.accent }}>
                Activity {idx + 1}
              </div>
              {items.length > 1 && (
                <button
                  onClick={() => setItems((p) => p.filter((_, i) => i !== idx))}
                  style={{
                    background: 'transparent',
                    border: 'none',
                    color: C.muted,
                    fontSize: 12,
                    cursor: 'pointer',
                  }}
                >
                  x Remove
                </button>
              )}
            </div>

            {/* Row 1: FY Half + Quarter + Month */}
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr 1fr',
                gap: 12,
                marginBottom: 12,
              }}
            >
              <div>
                <F label="FY Half" req error={errors[`fyHalf_${idx}`]}>
                  <select
                    value={item.fyHalf}
                    onChange={(e) => setItem(idx, 'fyHalf', e.target.value)}
                    style={{ ...inp(errors[`fyHalf_${idx}`]), width: '100%' }}
                  >
                    <option value="">Select...</option>
                    {FY_HALVES.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </F>
              </div>
              <div>
                <F label="FY Quarter" req error={errors[`fyQuarter_${idx}`]}>
                  <select
                    value={item.fyQuarter}
                    onChange={(e) => setItem(idx, 'fyQuarter', e.target.value)}
                    style={{
                      ...inp(errors[`fyQuarter_${idx}`]),
                      width: '100%',
                    }}
                    disabled={!item.fyHalf}
                  >
                    <option value="">Select...</option>
                    {getQuarters(item.fyHalf).map((q) => (
                      <option key={q} value={q}>
                        {q}
                      </option>
                    ))}
                  </select>
                </F>
              </div>
              <div>
                <F label="Month" req error={errors[`month_${idx}`]}>
                  <select
                    value={item.month}
                    onChange={(e) => setItem(idx, 'month', e.target.value)}
                    style={{ ...inp(errors[`month_${idx}`]), width: '100%' }}
                    disabled={!item.fyQuarter}
                  >
                    <option value="">Select...</option>
                    {getMonths(item.fyQuarter).map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </F>
              </div>
            </div>

            {/* Row 2: Product Group + Tactic */}
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 12,
                marginBottom: 12,
              }}
            >
              <F
                label="Product Group (BU Focus)"
                req
                error={errors[`productGroup_${idx}`]}
              >
                <select
                  value={item.productGroup}
                  onChange={(e) => setItem(idx, 'productGroup', e.target.value)}
                  style={{
                    ...inp(errors[`productGroup_${idx}`]),
                    width: '100%',
                  }}
                >
                  <option value="">Select...</option>
                  {PRODUCT_GROUPS.map((g) => (
                    <option key={g} value={g}>
                      {g}
                    </option>
                  ))}
                </select>
              </F>
              <F label="Marketing Tactic" req error={errors[`tactic_${idx}`]}>
                <select
                  value={item.tactic}
                  onChange={(e) => setItem(idx, 'tactic', e.target.value)}
                  style={{ ...inp(errors[`tactic_${idx}`]), width: '100%' }}
                >
                  <option value="">Select...</option>
                  {TACTICS.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              </F>
            </div>

            {/* Row 3: Location + Target Audience */}
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 12,
                marginBottom: 12,
              }}
            >
              <F label="Location" req error={errors[`location_${idx}`]}>
                <input
                  value={item.location}
                  onChange={(e) => setItem(idx, 'location', e.target.value)}
                  placeholder="e.g. Milan, Online, London"
                  style={{ ...inp(errors[`location_${idx}`]), width: '100%' }}
                />
              </F>
              <F
                label="Target Audience"
                req
                error={errors[`targetAudience_${idx}`]}
              >
                <input
                  value={item.targetAudience}
                  onChange={(e) =>
                    setItem(idx, 'targetAudience', e.target.value)
                  }
                  placeholder="e.g. IT Decision Makers, C-Suite"
                  style={{
                    ...inp(errors[`targetAudience_${idx}`]),
                    width: '100%',
                  }}
                />
              </F>
            </div>

            {/* Row 4: Target Solutions (free text) + Objective */}
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 12,
                marginBottom: 12,
              }}
            >
              <F
                label="Target Solutions"
                req
                error={errors[`targetSolutions_${idx}`]}
              >
                <input
                  value={item.targetSolutions}
                  onChange={(e) =>
                    setItem(idx, 'targetSolutions', e.target.value)
                  }
                  placeholder="e.g. Cloud, Security, AI/ML"
                  style={{
                    ...inp(errors[`targetSolutions_${idx}`]),
                    width: '100%',
                  }}
                />
              </F>
              <F label="Objective" req error={errors[`objective_${idx}`]}>
                <input
                  value={item.objective}
                  onChange={(e) => setItem(idx, 'objective', e.target.value)}
                  placeholder="e.g. Lead generation, Pipeline"
                  style={{ ...inp(errors[`objective_${idx}`]), width: '100%' }}
                />
              </F>
            </div>

            {/* Row 5: Total Cost + Currency + MDF Request */}
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '2fr 1fr 2fr',
                gap: 12,
              }}
            >
              <F label="Total Cost" req error={errors[`totalCost_${idx}`]}>
                <input
                  value={item.totalCost}
                  onChange={(e) => setItem(idx, 'totalCost', e.target.value)}
                  placeholder="0"
                  style={{ ...inp(errors[`totalCost_${idx}`]), width: '100%' }}
                />
              </F>
              <F label="Currency">
                <select
                  value={item.currency}
                  onChange={(e) => setItem(idx, 'currency', e.target.value)}
                  style={{ ...inp(false), width: '100%' }}
                >
                  {[
                    'EUR',
                    'USD',
                    'GBP',
                    'CHF',
                    'SEK',
                    'NOK',
                    'DKK',
                    'PLN',
                    'CZK',
                    'AED',
                    'SGD',
                    'AUD',
                  ].map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </F>
              <F label="MDF Request" req error={errors[`mdfRequest_${idx}`]}>
                <input
                  value={item.mdfRequest}
                  onChange={(e) => setItem(idx, 'mdfRequest', e.target.value)}
                  placeholder="0"
                  style={{ ...inp(errors[`mdfRequest_${idx}`]), width: '100%' }}
                />
              </F>
            </div>

            {/* Show MDF % if both filled */}
            {item.totalCost &&
              item.mdfRequest &&
              parseNum(String(item.totalCost)) > 0 && (
                <div
                  style={{
                    fontSize: 11,
                    color: C.muted,
                    marginTop: 6,
                    fontFamily: 'monospace',
                  }}
                >
                  MDF ={' '}
                  {Math.round(
                    (parseNum(String(item.mdfRequest)) /
                      parseNum(String(item.totalCost))) *
                      100
                  )}
                  % of total cost
                </div>
              )}
          </div>
        ))}

        {/* Add activity */}
        <button
          onClick={() => setItems((p) => [...p, emptyItem()])}
          style={{
            width: '100%',
            background: 'transparent',
            border: `2px dashed ${C.border}`,
            color: C.accent,
            borderRadius: 12,
            padding: 11,
            fontWeight: 700,
            fontSize: 13,
            cursor: 'pointer',
            marginBottom: 20,
          }}
        >
          + Add Another Activity
        </button>

        {/* Actions */}
        <div style={{ display: 'flex', gap: 12 }}>
          <button
            onClick={onClose}
            style={{
              flex: 1,
              background: 'transparent',
              color: C.muted,
              border: `1px solid ${C.border}`,
              borderRadius: 10,
              padding: 13,
              fontWeight: 600,
              cursor: 'pointer',
            }}
          >
            Cancel
          </button>
          <button
            onClick={submit}
            style={{
              flex: 2,
              background: C.accent,
              color: '#fff',
              border: 'none',
              borderRadius: 10,
              padding: 13,
              fontWeight: 800,
              fontSize: 14,
              cursor: 'pointer',
            }}
          >
            Submit Request
          </button>
        </div>
      </div>
    </div>
  );
};

const ImportModal = ({ onImport, onClose }) => {
  const [step, setStep] = useState('upload');
  const [wb, setWb] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [si, setSi] = useState(0);
  const [headers, setHeaders] = useState([]);
  const [preview, setPreview] = useState([]);
  const [mapping, setMapping] = useState({
    name: '',
    subregion: '',
    type: '',
    accountManager: '',
    // spent is auto-calculated, not imported
    tier: '',
    country: '',
    region: '',
    allocated: '',
    spent: '',
  });
  const fileRef = useRef();
  const loadSheet = (workbook, idx) => {
    const ws = workbook.Sheets[workbook.SheetNames[idx]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
    setHeaders(rows[0] || []);
    setPreview(rows.slice(0, 7));
  };
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const workbook = XLSX.read(ev.target.result, { type: 'binary' });
        setWb(workbook);
        setSheets(workbook.SheetNames);
        loadSheet(workbook, 0);
        setStep('preview');
      } catch {
        alert('Errore file');
      }
    };
    reader.readAsBinaryString(file);
  };
  const doImport = () => {
    const ws = wb.Sheets[wb.SheetNames[si]];
    const rowsR = XLSX.utils.sheet_to_json(ws, { raw: true, defval: '' });
    const rowsS = XLSX.utils.sheet_to_json(ws, { raw: false, defval: '' });
    const imported = rowsR.map((rowR, i) => {
      const rowS = rowsS[i] || {};
      const getAlloc = () => {
        const r = rowR[mapping.allocated],
          s = rowS[mapping.allocated];
        return typeof r === 'number' && !isNaN(r) ? r : parseNum(s);
      };
      return {
        id: 'p' + Date.now().toString(36) + i.toString(36) + Math.random().toString(36).slice(2,5),
        name: String(
          rowS[mapping.name] || rowR[mapping.name] || `Partner ${i + 1}`
        ).trim(),
        region: String(rowS[mapping.region] || '').trim(),
        subregion: String(rowS[mapping.subregion] || '').trim(),
        country: String(rowS[mapping.country] || '').trim(),
        type: String(rowS[mapping.type] || 'Reseller').trim(),
        tier: String(rowS[mapping.tier] || 'Silver').trim(),
        accountManager: String(rowS[mapping.accountManager] || '').trim(),
        allocated: getAlloc(),
        spent: 0,  // auto-calculated from approved claims
        pending: 0,
        status: 'active',
        note: '',
      };
    });
    onImport(imported);
    onClose();
  };
  const si2 = { upload: 0, preview: 1, mapping: 2 }[step];
  const inp = (err) => ({
    width: '100%',
    background: C.surface,
    border: `1px solid ${err ? C.danger : C.border}`,
    color: C.text,
    borderRadius: 8,
    padding: '10px 12px',
    fontSize: 13,
    outline: 'none',
  });
  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.85)',
        zIndex: 300,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <div
        style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 20,
          padding: 32,
          width: 660,
          maxHeight: '88vh',
          overflowY: 'auto',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            marginBottom: 24,
          }}
        >
          <div>
            <div style={{ fontWeight: 800, fontSize: 20, marginBottom: 2 }}>
              Import from Excel
            </div>
            <div style={{ color: C.muted, fontSize: 13 }}>
              Supporta $15,000 . EUR 15.000 . 15000.00
            </div>
          </div>
          <button
            onClick={onClose}
            style={{ background: 'transparent', color: C.muted, fontSize: 22 }}
          >
            x
          </button>
        </div>
        <div
          style={{ display: 'flex', alignItems: 'center', marginBottom: 28 }}
        >
          {['Upload', 'Preview', 'Map Columns'].map((s, i) => (
            <div
              key={s}
              style={{
                display: 'flex',
                alignItems: 'center',
                flex: i < 2 ? '1' : '0',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <div
                  style={{
                    width: 26,
                    height: 26,
                    borderRadius: '50%',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    background: i <= si2 ? C.accent : C.faint,
                    color: i <= si2 ? '#fff' : C.muted,
                    fontSize: 12,
                    fontWeight: 800,
                  }}
                >
                  {i + 1}
                </div>
                <span
                  style={{
                    fontSize: 13,
                    fontWeight: i === si2 ? 700 : 400,
                    color: i <= si2 ? C.text : C.muted,
                  }}
                >
                  {s}
                </span>
              </div>
              {i < 2 && (
                <div
                  style={{
                    flex: 1,
                    height: 1,
                    background: i < si2 ? C.accent : C.border,
                    margin: '0 16px',
                  }}
                />
              )}
            </div>
          ))}
        </div>
        {step === 'upload' && (
          <div
            onClick={() => fileRef.current.click()}
            style={{
              border: `2px dashed ${C.border}`,
              borderRadius: 16,
              padding: '52px 32px',
              textAlign: 'center',
              cursor: 'pointer',
            }}
          >
            <div style={{ fontSize: 44, marginBottom: 12 }}></div>
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 6 }}>
              Click to upload
            </div>
            <div style={{ color: C.muted, fontSize: 13 }}>
              Supports .xlsx, .xls, .csv
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFile}
              style={{ display: 'none' }}
            />
          </div>
        )}
        {step === 'preview' && (
          <>
            {sheets.length > 1 && (
              <div style={{ marginBottom: 14 }}>
                <div
                  style={{
                    fontSize: 12,
                    color: C.muted,
                    marginBottom: 8,
                    fontWeight: 600,
                  }}
                >
                  Sheet
                </div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  {sheets.map((s, i) => (
                    <button
                      key={s}
                      onClick={() => {
                        setSi(i);
                        loadSheet(wb, i);
                      }}
                      style={{
                        background: i === si ? C.accent : C.faint,
                        color: i === si ? '#fff' : C.muted,
                        border: `1px solid ${i === si ? C.accent : C.border}`,
                        borderRadius: 8,
                        padding: '6px 14px',
                        fontSize: 12,
                        fontWeight: 600,
                        cursor: 'pointer',
                      }}
                    >
                      {s}
                    </button>
                  ))}
                </div>
              </div>
            )}
            <div
              style={{
                overflowX: 'auto',
                borderRadius: 10,
                border: `1px solid ${C.border}`,
                marginBottom: 18,
              }}
            >
              <table
                style={{
                  width: '100%',
                  borderCollapse: 'collapse',
                  fontSize: 12,
                }}
              >
                <thead>
                  <tr style={{ background: C.faint }}>
                    {headers.map((h, i) => (
                      <th
                        key={i}
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          color: C.muted,
                          fontWeight: 700,
                          borderBottom: `1px solid ${C.border}`,
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {String(h)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.slice(1).map((row, ri) => (
                    <tr
                      key={ri}
                      style={{ borderBottom: `1px solid ${C.border}20` }}
                    >
                      {headers.map((_, ci) => (
                        <td
                          key={ci}
                          style={{ padding: '8px 14px', color: C.text }}
                        >
                          {row[ci] !== undefined ? String(row[ci]) : ''}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <button
                onClick={() => setStep('upload')}
                style={{
                  flex: 1,
                  background: C.faint,
                  color: C.muted,
                  border: `1px solid ${C.border}`,
                  borderRadius: 10,
                  padding: 12,
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                Back
              </button>
              <button
                onClick={() => setStep('mapping')}
                style={{
                  flex: 2,
                  background: C.accent,
                  color: '#fff',
                  border: 'none',
                  borderRadius: 10,
                  padding: 12,
                  fontWeight: 700,
                  cursor: 'pointer',
                }}
              >
                Map Columns
              </button>
            </div>
          </>
        )}
        {step === 'mapping' && (
          <>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 14,
                marginBottom: 24,
              }}
            >
              {[
                { key: 'name', label: 'Partner Name *', req: true },
                { key: 'region', label: 'Region' },
                { key: 'subregion', label: 'Sub-Region' },
                { key: 'country', label: 'Country' },
                { key: 'type', label: 'Type (Reseller, GSI...)' },
                { key: 'tier', label: 'Level (Platinum, Gold...)' },
                { key: 'accountManager', label: 'Account Manager' },
                { key: 'allocated', label: 'Budget Allocated *', req: true },
        
              ].map((f) => (
                <div key={f.key}>
                  <div
                    style={{
                      fontSize: 12,
                      color: C.muted,
                      marginBottom: 6,
                      fontWeight: 600,
                    }}
                  >
                    {f.label}
                  </div>
                  <select
                    value={mapping[f.key]}
                    onChange={(e) =>
                      setMapping((p) => ({ ...p, [f.key]: e.target.value }))
                    }
                    style={inp(f.req && !mapping[f.key])}
                  >
                    <option value="">-- select column --</option>
                    {headers.map((h, i) => (
                      <option key={i} value={String(h)}>
                        {String(h)}
                      </option>
                    ))}
                  </select>
                  {(f.key === 'allocated' || f.key === 'spent') &&
                    mapping[f.key] &&
                    preview[1] && (
                      <div
                        style={{ fontSize: 10, color: C.success, marginTop: 4 }}
                      >
                        {'Preview: ' +
                          parseNum(
                            String(
                              preview[1][headers.indexOf(mapping[f.key])] || 0
                            )
                          ).toLocaleString()}
                      </div>
                    )}
                </div>
              ))}
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <button
                onClick={() => setStep('preview')}
                style={{
                  flex: 1,
                  background: C.faint,
                  color: C.muted,
                  border: `1px solid ${C.border}`,
                  borderRadius: 10,
                  padding: 12,
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                Back
              </button>
              <button
                onClick={doImport}
                disabled={!mapping.name || !mapping.allocated}
                style={{
                  flex: 2,
                  background: C.success,
                  color: '#000',
                  border: 'none',
                  borderRadius: 10,
                  padding: 12,
                  fontWeight: 700,
                  cursor: 'pointer',
                  opacity: !mapping.name || !mapping.allocated ? 0.5 : 1,
                }}
              >
                Import {Math.max(0, preview.length - 1)} Partners
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const doExport = (
  tab,
  partners,
  requests,
  filteredPartners,
  filteredRequests,
  filteredRows,
  rate,
  toUSD
) => {
  const wb = XLSX.utils.book_new();
  const ts = new Date().toISOString().slice(0, 10);

  if (tab === 'dashboard' || tab === 'partners') {
    const pData = (filteredPartners || partners).map((p) => ({
      Region: p.region,
      'Sub-Region': p.subregion || '',
      Country: p.country || '',
      Partner: p.name,
      Type: p.type || '',
      Level: p.tier || '',
      'Allocated (USD)': Math.round((p.allocated || 0) * rate),
      'Spent (USD)': Math.round((p.spent || 0) * rate),
      'Pending (USD)': Math.round((p.pending || 0) * rate),
      'Available (USD)': Math.round(
        ((p.allocated || 0) - (p.spent || 0) - (p.pending || 0)) * rate
      ),
      'Util %': p.allocated
        ? Math.round(((p.spent + p.pending) / p.allocated) * 100) + '%'
        : '0%',
    }));
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(pData),
      'Partners'
    );
  }

  if (tab === 'requests') {
    const rData = (filteredRequests || requests).flatMap((r) =>
      (r.items || []).map((it) => ({
        'Request ID': r.id,
        Partner: r.partner,
        Status: r.status,
        'PO Number': r.poNumber || '',
        Submitted: r.submitted || '',
        'FY Half': it.fyHalf || '',
        Quarter: it.fyQuarter || '',
        Month: it.month || it.period || '',
        'Product Group': it.productGroup || '',
        Tactic: it.tactic || '',
        Location: it.where || '',
        'Target Audience': it.targetAudience || '',
        Objective: it.objective || '',
        Currency: it.localCurrency || 'EUR',
        'Total Cost (LC)': it.amount || 0,
        'MDF Request (LC)': it.mdfRequest || Math.round((it.amount || 0) * 0.5),
        'MDF Request (USD)': toUSD
          ? toUSD(
              it.mdfRequest || Math.round((it.amount || 0) * 0.5),
              it.localCurrency || 'EUR'
            )
          : 0,
      }))
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(rData),
      'Requests'
    );
  }

  if (tab === 'analytics') {
    const aData = (filteredRows || []).map((row) => ({
      Region: row.macro,
      'Sub-Region': row.subregion,
      Country: row.country,
      Partner: row.partnerName,
      'Contact Name': row.contactName,
      'Contact Email': row.contactEmail,
      'OT Partner Mgr': row.accountManager,
      'FY Half': row.fyHalf,
      Quarter: row.fyQuarter,
      Month: row.period,
      'BU / Product Group': row.productGroup,
      Tactic: row.tactic,
      Activity: row.title,
      'Allocadia ID': row.allocadiaId || '',
      'Campaign ID': row.campaignId || '',
      'Partner Notified': row.partnerNotified ? 'Yes' : 'No',
      Currency: row.localCurrency,
      'Total Cost (LC)': row.totalCost,
      'MDF Request (LC)': row.mdfRequest,
      'MDF Request (USD)': toUSD ? toUSD(row.mdfRequest, row.localCurrency) : 0,
      'Request ID': row.reqId,
      Status: row.reqStatus,
      'PO Number': row.poNumber || '',
    }));
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(aData),
      'MDF Overview'
    );
  }

  if (tab === 'claims') {
    const cData = (filteredRequests || requests).map((r) => ({
      'Request ID': r.id,
      Partner: r.partner,
      Status: r.status,
      'PO Number': r.poNumber || '',
      Submitted: r.submitted || '',
      Activities: (r.items || []).length,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cData), 'Claims');
  }

  const summary = [
    {
      'Exported From': tab,
      Date: ts,
      Rows: wb.SheetNames.length > 0 ? 'See sheets' : 0,
    },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'Info');

  if (wb.SheetNames.length === 1) {
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        partners.map((p) => ({ Partner: p.name, Allocated: p.allocated }))
      ),
      'Partners'
    );
  }

  XLSX.writeFile(wb, `MDF_${tab}_${ts}.xlsx`);
};

const doPPTExport = async (partners, requests, rate, currency, fmtA) => {
  const btn = document.querySelector('[data-ppt-btn]');
  if (btn) {
    btn.textContent = 'Generating...';
    btn.disabled = true;
  }

  try {
    const totalAlloc = partners.reduce((s, p) => s + p.allocated, 0);
    const totalSpent = partners.reduce((s, p) => s + p.spent, 0);
    const totalPend = partners.reduce((s, p) => s + p.pending, 0);
    const utilPct = totalAlloc
      ? Math.round(((totalSpent + totalPend) / totalAlloc) * 100)
      : 0;

    const byRegion = ['Europe', 'US', 'International']
      .map((m) => {
        const mp = partners.filter((p) => p.region === m);
        return {
          name: m,
          count: mp.length,
          alloc: mp.reduce((s, p) => s + p.allocated, 0),
          spent: mp.reduce((s, p) => s + p.spent, 0),
        };
      })
      .filter((m) => m.count > 0);

    const statusCounts = [
      'request_submitted',
      'sent_for_signature',
      'signed',
      'rejected',
    ].map((k) => ({
      key: k,
      label: k.replace(/_/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase()),
      count: requests.filter((r) => r.status === k).length,
      amt: requests
        .filter((r) => r.status === k)
        .reduce(
          (s, r) => (r.items || []).reduce((ss, i) => ss + i.amount, s),
          0
        ),
    }));

    const topPartners = [...partners]
      .sort((a, b) => b.allocated - a.allocated)
      .slice(0, 8);

    const allItems = requests.flatMap((r) =>
      (r.items || []).map((it) => ({
        ...it,
        partner: r.partner,
        status: r.status,
      }))
    );
    const buMap = {};
    allItems.forEach((it) => {
      const bu = it.productGroup || 'Other';
      if (!buMap[bu]) buMap[bu] = { count: 0, amt: 0 };
      buMap[bu].count++;
      buMap[bu].amt += it.amount || 0;
    });
    const buList = Object.entries(buMap)
      .sort((a, b) => b[1].amt - a[1].amt)
      .slice(0, 8);

    const fmt = (n) => {
      const v = Number(n || 0);
      return v >= 1e6
        ? (v / 1e6).toFixed(1) + 'M'
        : v >= 1e3
        ? (v / 1e3).toFixed(0) + 'K'
        : v.toLocaleString();
    };

    const prompt = `You are a Node.js script generator. Generate a complete Node.js script using pptxgenjs to create an MDF Manager PowerPoint report.

The script must:
1. Use: const pptxgen = require('pptxgenjs');
2. Create a professional dark-themed presentation with these slides:
   - Slide 1: Title slide (dark navy #1E3A5F bg), "MDF Manager" large title, subtitle "Market Development Fund Report", date ${new Date().toLocaleDateString(
     'en-US',
     { month: 'long', day: 'numeric', year: 'numeric' }
   )}, stats: ${partners.length} Partners, ${requests.length} Requests, ${
      allItems.length
    } Activities
   - Slide 2: Budget KPIs (light bg), 4 cards: Allocated ${fmt(
     totalAlloc
   )}, Spent ${fmt(totalSpent)}, Pending ${fmt(totalPend)}, Available ${fmt(
      totalAlloc - totalSpent - totalPend
    )}, utilization bar ${utilPct}%
   - Slide 3: Region breakdown - ${byRegion
     .map(
       (r) =>
         `${r.name}: ${r.count} partners, ${fmt(r.alloc)} allocated, ${fmt(
           r.spent
         )} spent`
     )
     .join('; ')}
   - Slide 4: Request statuses - ${statusCounts
     .map((s) => `${s.label}: ${s.count} (${fmt(s.amt)})`)
     .join('; ')}
   - Slide 5: Top partners table - ${topPartners
     .map(
       (p) =>
         `${p.name}|${p.region}|${p.type || '-'}|${p.tier}|${fmt(
           p.allocated
         )}|${fmt(p.spent)}`
     )
     .join('; ')}
   - Slide 6: BU breakdown - ${buList
     .map(([n, d]) => `${n}: ${d.count} activities, ${fmt(d.amt)}`)
     .join('; ')}
3. Use colors: navy #1E3A5F, blue #2563EB, green #10B981, amber #F59E0B, purple #8B5CF6
4. End with: prs.writeFile({fileName:'MDF_Report.pptx'}).then(()=>console.log('DONE'));
5. Return ONLY the complete Node.js script, no explanation, no markdown fences.`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 4000,
        messages: [{ role: 'user', content: prompt }],
      }),
    });

    if (response.status === 401 || response.status === 403) {
      throw new Error('API access not available in this environment. PPT export requires the StackBlitz platform with API injection enabled.');
    }

    const data = await response.json();
    if (data.error) throw new Error(data.error.message || 'API error');
    const script = data.content?.find((b) => b.type === 'text')?.text || '';

    if (!script) throw new Error('No script generated');

    const blob = new Blob([script], { type: 'text/javascript' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `MDF_PPT_Generator_${new Date()
      .toISOString()
      .slice(0, 10)}.js`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    alert(
      'PPT generator script downloaded!\n\nTo create the PPTX:\n1. Make sure Node.js is installed\n2. Run: npm install pptxgenjs\n3. Run: node MDF_PPT_Generator_*.js\n\nThis will create MDF_Report.pptx in the same folder.'
    );
  } catch (err) {
    console.error('PPT export error:', err);
    alert('PPT export failed: ' + err.message);
  } finally {
    if (btn) {
      btn.textContent = 'PPT';
      btn.disabled = false;
    }
  }
};

// Mirrors testMDF.docx structure: Partners Info + Activities table + Totals + Signatures
const generateMDFBusinessPlan = (
  partner,
  activities,
  poNumber = '',
  companyName = 'OT'
) => {
  const fmt = (n, cur) => {
    const num = Number(n || 0).toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
    return cur ? `${cur} ${num}` : num;
  };

  // Totals per currency (can't sum mixed currencies)
  const currencyTotals = {};
  activities.forEach((a) => {
    const cur = a.localCurrency || 'EUR';
    if (!currencyTotals[cur]) currencyTotals[cur] = { cost: 0, mdf: 0 };
    currencyTotals[cur].cost += Number(a.amount || 0);
    currencyTotals[cur].mdf += Number(
      a.mdfRequest || Math.round((a.amount || 0) * 0.5)
    );
  });
  const totalCostStr =
    Object.entries(currencyTotals)
      .map(([c, v]) => fmt(v.cost, c))
      .join(' + ') || '0';
  const totalMDFStr =
    Object.entries(currencyTotals)
      .map(([c, v]) => fmt(v.mdf, c))
      .join(' + ') || '0';

  const hdr = (cols) =>
    cols
      .map(
        ([text, w]) =>
          `<th style="background:#1F3864;color:#fff;padding:8px;font-size:11px;border:1px solid #888;width:${w}px;text-align:left">${text}</th>`
      )
      .join('');
  const td = (text, right = false, bold = false) =>
    `<td style="padding:7px 8px;border:1px solid #bbb;font-size:11px;text-align:${
      right ? 'right' : 'left'
    };${bold ? 'font-weight:700;background:#D6E4F0;' : 'background:#fff;'}">${
      text || ''
    }</td>`;

  const actRows = activities
    .map(
      (a) => `<tr>
    ${td(a.fyHalf || '')}
    ${td(a.fyQuarter || '')}
    ${td(a.productGroup || '')}
    ${td(a.tactic || a.category || '')}
    ${td(a.title || a.activity || '')}
    ${td(a.allocadiaId || '')}
    ${td(a.campaignId || '')}
    ${td(fmt(a.amount, a.localCurrency || 'EUR'), true)}
    ${td(
      fmt(
        a.mdfRequest || Math.round((a.amount || 0) * 0.5),
        a.localCurrency || 'EUR'
      ),
      true
    )}
  </tr>`
    )
    .join('');

  const emptyRows = Array(Math.max(0, 4 - activities.length))
    .fill(
      `<tr>${['', '', '', '', '', '', '', ''].map(() => td('')).join('')}</tr>`
    )
    .join('');

  const logoSrc = typeof OT_LOGO !== 'undefined' ? OT_LOGO : '';

  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  body { font-family: Calibri, Arial, sans-serif; font-size: 12px; margin: 30px; color: #000; }
  h1 { color: #1F3864; font-size: 18px; margin-bottom: 4px; }
  h2 { color: #1F3864; font-size: 14px; margin: 16px 0 6px; border-bottom: 2px solid #1F3864; padding-bottom: 4px; }
  p  { font-size: 11px; line-height: 1.5; margin: 4px 0 10px; color: #222; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 10px; }
  .confidential { color: #888; font-style: italic; font-size: 10px; text-align: right; }
  .sig-table td { padding: 18px 10px; border: 1px solid #bbb; font-size: 11px; min-height: 40px; }
  .sig-hdr { background: #1F3864; color: #fff; font-weight: 700; padding: 8px 10px; font-size: 12px; }
</style>
</head><body>

<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
  ${
    logoSrc
      ? `<img src="${logoSrc}" style="height:32px;object-fit:contain" alt="OT"/>`
      : `<span style="font-weight:800;font-size:18px;color:#1F3864">${companyName}</span>`
  }
  <span class="confidential">CONFIDENTIAL</span>
</div>

<h1>Section I - MDF Business Plan</h1>
<h2>Overview</h2>
<p>Subject to the terms and conditions hereto attached, this MDF Business Plan summarizes the goals, activities and commitments agreed between <strong>${companyName}</strong> and <strong>${
    partner.name || partner.partnerName || ''
  }</strong>. This document is confidential between ${companyName} and ${
    partner.name || partner.partnerName || ''
  } and must not be disclosed to any other parties.</p>
<p>Notwithstanding the use of the terms "partner" and "partnership" in the MDF Business Plan, neither ${companyName} nor the partner is entitled to act, incur any liability, make any representation or enter any commitment or contractual arrangement for and on behalf of the other party, without the prior written approval of the other party.</p>

<h2>Partners Information</h2>
<table>
  <thead><tr>
    ${hdr([
      ['Partner Name', 140],
      ['Primary Contact Name', 140],
      ['Primary Contact Email', 170],
      ['Partner Type', 100],
      ['Partner Level', 100],
      ['OT Account Manager', 140],
    ])}
  </tr></thead>
  <tbody><tr>
    ${td(partner.name || partner.partnerName || '')}
    ${td(partner.contactName || partner.partnerContact || '')}
    ${td(partner.contactEmail || partner.partnerEmail || '')}
    ${td(partner.type || partner.partnerType || '')}
    ${td(partner.tier || partner.partnerTier || '')}
    ${td(partner.accountManager || partner.poNumber || '')}
  </tr></tbody>
</table>

<h2>Approved MDF Activities</h2>
<table>
  <thead><tr>
    ${hdr([
      ['Fiscal Year Half', 70],
      ['FY Quarter', 60],
      ['Product Group (BU Focus)', 100],
      ['Marketing Tactic', 100],
      ['Activity Description', 160],
      ['Allocadia ID', 75],
      ['Campaign ID', 75],
      ['Total Cost (LCY)', 110],
      ['MDF Request (LCY)', 100],
    ])}
  </tr></thead>
  <tbody>
    ${actRows}
    ${emptyRows}
    <tr>
      <td colspan="7" style="padding:8px;border:1px solid #bbb;font-weight:700;background:#D6E4F0;font-size:11px">TOTAL Cost of ALL Marketing Activities</td>
      <td style="padding:8px;border:1px solid #bbb;font-weight:700;background:#D6E4F0;font-size:11px;text-align:right">${totalCostStr}</td>
      <td style="padding:8px;border:1px solid #bbb;font-weight:700;background:#D6E4F0;font-size:11px;text-align:right">${totalMDFStr}</td>
    </tr>
  </tbody>
</table>
${
  poNumber
    ? `<p style="font-size:11px;color:#555"><strong>PO Number:</strong> ${poNumber}</p>`
    : ''
}

<h2 style="margin-top:24px">Signatures</h2>
<table class="sig-table" style="margin-top:12px">
  <thead><tr>
    <td class="sig-hdr">${companyName} Representative</td>
    <td class="sig-hdr">${
      partner.name || partner.partnerName || ''
    } Representative</td>
  </tr></thead>
  <tbody>
    <tr>
      <td>Name: _______________________________<br/><br/>Title: _______________________________<br/><br/>Date: ________________________________<br/><br/>Signature: ___________________________</td>
      <td>Name: _______________________________<br/><br/>Title: _______________________________<br/><br/>Date: ________________________________<br/><br/>Signature: ___________________________</td>
    </tr>
  </tbody>
</table>
</body></html>`;
  const blob = new Blob([html], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `MDF_BusinessPlan_${(
    partner.name ||
    partner.partnerName ||
    'Partner'
  ).replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

const NavIcon = ({ id, size = 18, color }) => {
  const s = {
    width: size,
    height: size,
    viewBox: '0 0 24 24',
    fill: 'none',
    stroke: color,
    strokeWidth: 2,
    strokeLinecap: 'round',
    strokeLinejoin: 'round',
    flexShrink: 0,
  };
  if (id === 'dashboard')
    return (
      <svg {...s}>
        <rect x="3" y="3" width="7" height="7" rx="1" />
        <rect x="14" y="3" width="7" height="7" rx="1" />
        <rect x="3" y="14" width="7" height="7" rx="1" />
        <rect x="14" y="14" width="7" height="7" rx="1" />
      </svg>
    );
  if (id === 'analytics')
    return (
      <svg {...s}>
        <path d="M3 3v18h18" />
        <path d="M7 16l4-4 4 4 4-7" />
      </svg>
    );
  if (id === 'partners')
    return (
      <svg {...s}>
        <circle cx="9" cy="7" r="4" />
        <path d="M3 21v-2a4 4 0 0 1 4-4h4a4 4 0 0 1 4 4v2" />
        <path d="M16 3.13a4 4 0 0 1 0 7.75" />
        <path d="M21 21v-2a4 4 0 0 0-3-3.85" />
      </svg>
    );
  if (id === 'requests')
    return (
      <svg {...s}>
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
        <polyline points="14,2 14,8 20,8" />
        <line x1="16" y1="13" x2="8" y2="13" />
        <line x1="16" y1="17" x2="8" y2="17" />
      </svg>
    );
  if (id === 'claims')
    return (
      <svg {...s}>
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
        <polyline points="22,4 12,14.01 9,11.01" />
      </svg>
    );
  if (id === 'history')
    return (
      <svg {...s}>
        <circle cx="12" cy="12" r="10" />
        <polyline points="12,6 12,12 16,14" />
      </svg>
    );
  return (
    <svg {...s}>
      <circle cx="12" cy="12" r="3" />
    </svg>
  );
};

const PlanModal = ({
  modal,
  partners,
  currentUser,
  onClose,
  generateFn,
  onUpdateStatus,
}) => {
  const [activities, setActivities] = useState(modal.activities);
  const [poNumber, setPoNumber] = useState('');
  const [allocadiaIds, setAllocadiaIds] = useState({});
  const [campaignIds, setCampaignIds] = useState({});
  const setAllocadia = (id, val) =>
    setAllocadiaIds((p) => ({ ...p, [id]: val }));
  const setCampaignId = (id, val) =>
    setCampaignIds((p) => ({ ...p, [id]: val }));

  const toggle = (id) =>
    setActivities((prev) =>
      prev.map((a) => (a.id === id ? { ...a, selected: !a.selected } : a))
    );
  const selected = activities.filter((a) => a.selected);

  const handleGenerate = () => {
    const partner = modal.partner;
    const partnerInfo = {
      name: partner.name,
      contactName: partner.contactName || '',
      contactEmail: partner.contactEmail || '',
      type: partner.type || 'Partner',
      tier: partner.tier || '',
      accountManager: currentUser,
      poNumber: poNumber,
    };
    const acts = selected.map((r) => ({
      fyHalf: r.fyHalf || 'FY26 H1',
      fyQuarter: r.fyQuarter || 'Q1',
      productGroup:
        r.productGroup ||
        (r.targetSolutions || []).join(', ') ||
        r.category ||
        '',
      tactic: r.tactic || r.category || '',
      title: r.title || r.activity || '',
      amount: r.amount,
      mdfRequest: r.mdfRequest || Math.round((r.amount || 0) * 0.5),
      localCurrency: r.localCurrency || 'EUR',
      allocadiaId: allocadiaIds[r.id] || r.allocadiaId || '',
      campaignId: campaignIds[r.id] || r.campaignId || '',
    }));
    generateFn(partnerInfo, acts, poNumber);
    onUpdateStatus(modal.reqId, poNumber, allocadiaIds, campaignIds);
    onClose();
  };

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.85)',
        zIndex: 300,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <div
        style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 20,
          padding: 32,
          width: 660,
          maxHeight: '88vh',
          overflowY: 'auto',
          boxSizing: 'border-box',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            fontFamily: "'Syne',sans-serif",
            fontWeight: 800,
            fontSize: 20,
            marginBottom: 4,
          }}
        >
          Generate MDF Business Plan
        </div>
        <div style={{ color: C.muted, fontSize: 13, marginBottom: 20 }}>
          Select activities to include in the document
        </div>
        <div
          style={{
            display: 'flex',
            flexDirection: 'column',
            gap: 8,
            marginBottom: 16,
          }}
        >
          {activities.map((a) => (
            <div
              key={a.id}
              onClick={() => toggle(a.id)}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 14,
                padding: '12px 16px',
                borderRadius: 12,
                border: `1px solid ${a.selected ? C.accent + '50' : C.border}`,
                background: a.selected ? C.accentGlow : C.faint,
                cursor: 'pointer',
                transition: 'all 0.15s',
              }}
            >
              <div
                style={{
                  width: 20,
                  height: 20,
                  borderRadius: 4,
                  border: `2px solid ${a.selected ? C.accent : C.muted}`,
                  background: a.selected ? C.accent : 'transparent',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  flexShrink: 0,
                }}
              >
                {a.selected && (
                  <span
                    style={{ color: '#fff', fontSize: 12, fontWeight: 800 }}
                  >
                    v
                  </span>
                )}
              </div>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 2 }}>
                  {a.title || a.activity}
                </div>
                <div
                  style={{
                    display: 'flex',
                    gap: 8,
                    fontSize: 11,
                    color: C.muted,
                    flexWrap: 'wrap',
                  }}
                >
                  <span>{a.tactic || a.category}</span>
                  <span>.</span>
                  <span>{a.fyQuarter}</span>
                  <span>.</span>
                  <span style={{ fontFamily: 'monospace', color: C.accent }}>
                    {a.localCurrency || 'EUR'}{' '}
                    {Number(a.amount || 0).toLocaleString()} . MDF:{' '}
                    {a.localCurrency || 'EUR'}{' '}
                    {Number(
                      a.mdfRequest || Math.round((a.amount || 0) * 0.5)
                    ).toLocaleString()}
                  </span>
                </div>
                <div
                  style={{
                    display: 'grid',
                    gridTemplateColumns: '1fr 1fr',
                    gap: 8,
                    marginTop: 8,
                  }}
                  onClick={(e) => e.stopPropagation()}
                >
                  <div>
                    <label
                      style={{
                        fontSize: 10,
                        fontWeight: 700,
                        color: C.muted,
                        display: 'block',
                        marginBottom: 3,
                        letterSpacing: '0.06em',
                      }}
                    >
                      ALLOCADIA ID
                    </label>
                    <input
                      value={allocadiaIds[a.id] || a.allocadiaId || ''}
                      onChange={(e) => setAllocadia(a.id, e.target.value)}
                      placeholder="e.g. ALO-2026-001"
                      style={{
                        width: '100%',
                        background: C.card,
                        border: `1px solid ${C.accent}40`,
                        color: C.text,
                        borderRadius: 6,
                        padding: '5px 10px',
                        fontSize: 12,
                        fontFamily: 'monospace',
                        outline: 'none',
                        boxSizing: 'border-box',
                      }}
                    />
                  </div>
                  <div>
                    <label
                      style={{
                        fontSize: 10,
                        fontWeight: 700,
                        color: C.muted,
                        display: 'block',
                        marginBottom: 3,
                        letterSpacing: '0.06em',
                      }}
                    >
                      MDF CAMPAIGN ID
                    </label>
                    <input
                      value={campaignIds[a.id] || a.campaignId || ''}
                      onChange={(e) => setCampaignId(a.id, e.target.value)}
                      placeholder="e.g. CMP-2026-001"
                      style={{
                        width: '100%',
                        background: C.card,
                        border: `1px solid ${C.accent}40`,
                        color: C.text,
                        borderRadius: 6,
                        padding: '5px 10px',
                        fontSize: 12,
                        fontFamily: 'monospace',
                        outline: 'none',
                        boxSizing: 'border-box',
                      }}
                    />
                  </div>
                </div>
              </div>
              <div
                style={{
                  background: C.success + '18',
                  color: C.success,
                  border: `1px solid ${C.success}30`,
                  borderRadius: 20,
                  padding: '2px 10px',
                  fontSize: 10,
                  fontWeight: 800,
                }}
              >
                APPROVED
              </div>
            </div>
          ))}
        </div>
        <div
          style={{
            background: C.faint,
            border: `1px solid ${C.border}`,
            borderRadius: 10,
            padding: '12px 16px',
            marginBottom: 16,
          }}
        >
          <div
            style={{
              display: 'flex',
              justifyContent: 'space-between',
              marginBottom: 6,
            }}
          >
            <span style={{ fontSize: 12, color: C.muted }}>
              Total Cost (in local currencies)
            </span>
            <span
              style={{
                fontFamily: 'monospace',
                fontWeight: 700,
                fontSize: 13,
                color: C.accent,
              }}
            >
              {selected
                .map(
                  (a) =>
                    `${a.localCurrency || 'EUR'} ${Number(
                      a.amount || 0
                    ).toLocaleString()}`
                )
                .join(' + ') || '-'}
            </span>
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between' }}>
            <span style={{ fontSize: 12, color: C.muted }}>
              MDF Request - 50% (in local currencies)
            </span>
            <span
              style={{
                fontFamily: 'monospace',
                fontWeight: 700,
                fontSize: 13,
                color: C.success,
              }}
            >
              {selected
                .map(
                  (a) =>
                    `${a.localCurrency || 'EUR'} ${Number(
                      Math.round((a.amount || 0) * 0.5)
                    ).toLocaleString()}`
                )
                .join(' + ') || '-'}
            </span>
          </div>
        </div>

        <div style={{ display: 'flex', gap: 10 }}>
          <button
            onClick={onClose}
            style={{
              flex: 1,
              background: 'transparent',
              color: C.muted,
              border: `1px solid ${C.border}`,
              borderRadius: 10,
              padding: 13,
              fontWeight: 600,
              cursor: 'pointer',
            }}
          >
            Cancel
          </button>
          <button
            onClick={handleGenerate}
            disabled={selected.length === 0}
            style={{
              flex: 2,
              background: selected.length === 0 ? C.faint : C.accent,
              color: selected.length === 0 ? C.muted : '#fff',
              border: 'none',
              borderRadius: 10,
              padding: 13,
              fontWeight: 700,
              fontSize: 14,
              cursor: selected.length === 0 ? 'not-allowed' : 'pointer',
              opacity: selected.length === 0 ? 0.6 : 1,
            }}
          >
            Download Business Plan ({selected.length}{' '}
            {selected.length === 1 ? 'activity' : 'activities'})
          </button>
        </div>
      </div>
    </div>
  );
};

const RequestReviewModal = ({
  request,
  partners,
  requests,
  setRequests,
  selectedItems,
  setSelectedItems,
  addHistory,
  toast,
  openPlanModal,
  setPlanModal,
  onClose,
  fmtA,
  currency,
}) => {
  const FY_HALVES = [
    'FY26 H1',
    'FY26 H2',
    'FY27 H1',
    'FY27 H2',
    'FY28 H1',
    'FY28 H2',
  ];
  const QUARTERS_BY_HALF = { H1: ['Q1', 'Q2'], H2: ['Q3', 'Q4'] };
  const MONTHS_BY_QUARTER = {
    Q1: ['July', 'August', 'September'],
    Q2: ['October', 'November', 'December'],
    Q3: ['January', 'February', 'March'],
    Q4: ['April', 'May', 'June'],
  };
  const getHalfKey = (fyHalf) => (fyHalf?.includes('H1') ? 'H1' : 'H2');
  const getQuarters = (fyHalf) => QUARTERS_BY_HALF[getHalfKey(fyHalf)] || [];
  const getMonths = (fyQ) => MONTHS_BY_QUARTER[fyQ] || [];

  const [r, setR] = useState({
    ...request,
    items: (request.items || []).map((i) => ({ ...i })),
  });
  const [note, setNote] = useState(request.note || '');
  const [hasEdits, setHasEdits] = useState(false);
  const [signedDoc, setSignedDoc] = useState(request.signedDoc || null);

  React.useEffect(() => {
    const updated = requests.find((x) => x.id === r.id);
    if (updated) {
      const statusChanged = updated.status !== r.status;
      const bpChanged = updated.bpGeneratedAt !== r.bpGeneratedAt;
      if (statusChanged || bpChanged) {
        setR((prev) => ({
          ...prev,
          status: updated.status,
          bpGeneratedAt: updated.bpGeneratedAt,
          items: updated.items || prev.items,
        }));
        if (statusChanged) setHasEdits(false);
      }
    }
  }, [requests]);

  const setItem = (idx, k, v) => {
    setHasEdits(true);
    setR((prev) => {
      const newItems = prev.items.map((it, i) => {
        if (i !== idx) return it;
        const updated = { ...it, [k]: v };
        if (k === 'fyHalf') {
          updated.fyQuarter = '';
          updated.month = '';
        }
        if (k === 'fyQuarter') {
          updated.month = '';
        }
        if (k === 'totalCost') {
          const num = parseNum(String(v));
          if (num > 0) updated.mdfRequest = Math.round(num * 0.5);
          updated.amount = parseNum(String(v));
        }
        if (k === 'mdfRequest') updated.mdfRequest = parseNum(String(v));
        return updated;
      });
      // When CMM is assigned to any item, apply to ALL items + request level
      if (k === 'assignedTo' && v) {
        const allAssigned = newItems.map(it => ({ ...it, assignedTo: v }));
        return { ...prev, assignedTo: v, items: allAssigned };
      }
      return { ...prev, items: newItems };
    });
  };

  const saveChanges = () => {
    // Validate: approved items must have Campaign ID
    const missingCampaignId = r.items.filter(it =>
      it.itemStatus === 'approved' && !it.campaignId
    );
    if (missingCampaignId.length > 0) {
      toast(`Campaign ID required for: ${missingCampaignId.map(it => it.title || it.id).join(', ')}`, C.danger);
      return;
    }
    const updatedItems = r.items.map((it) => ({
      ...it,
      month:
        it.month || (it.period || '').replace(/\s*\(.*\)/, '').trim() || '',
      itemStatus: it.itemStatus || r.status,
    }));
    setRequests((prev) =>
      prev.map((x) => {
        if (x.id !== r.id) return x;
        // Preserve parent status if it has advanced beyond local copy (e.g. after BP generation)
        // Auto-advance to 'approved' if all active items are approved
        // Use r.status directly - user explicitly set it via status buttons
        // Only auto-advance from request_submitted->approved if all items are approved
        const activeItems = (r.items || []).filter(it => !['cancelled_by_partner','postponed','rejected'].includes(it.itemStatus));
        const allApproved = activeItems.length > 0 && activeItems.every(it => it.itemStatus === 'approved');
        const autoStatus = allApproved && r.status === 'request_submitted' ? 'approved' : r.status;
        // Never revert to a previous status - always use the most advanced
        const statusOrder = ['request_submitted','approved','sent_for_signature','signed','po_raised'];
        const rStatusIdx = statusOrder.indexOf(autoStatus);
        const xStatusIdx = statusOrder.indexOf(x.status);
        const safeStatus = (x.status === 'rejected') ? x.status
          : (autoStatus === 'rejected') ? 'rejected'
          : rStatusIdx >= xStatusIdx ? autoStatus : x.status;
        return {
          ...r,
          status: safeStatus,
          note,
          items: updatedItems,
          partnerContact: r.partnerContact,
          partnerEmail: r.partnerEmail,
          partnerManager: r.partnerManager,
          partnerType: r.partnerType,
          partnerTier: r.partnerTier,
        };
      })
    );
    addHistory(`Request ${r.id} updated`, r.id, 'edit');
    toast('Changes saved!');
    setHasEdits(false);
  };

  // Update status locally - user must click Save Changes to persist
  const setStatus = (newStatus) => {
    setR(prev => ({ ...prev, status: newStatus }));
    setHasEdits(true);
  };

  const approveAndGenerate = () => {
    const partnerRecord = partners.find((p) => p.name === r.partner) || {};
    const partner = {
      ...partnerRecord,
      name: r.partner,
      contactName: r.partnerContact || partnerRecord.contactName || '',
      contactEmail: r.partnerEmail || partnerRecord.contactEmail || '',
      accountManager: r.partnerManager || partnerRecord.accountManager || '',
      type: r.partnerType || partnerRecord.type || '',
      tier: r.partnerTier || partnerRecord.tier || '',
    };
    const chosen = r.items.filter((i) => i.itemStatus === 'approved');
    if (chosen.length === 0) {
      toast('Approve at least one activity first', '#f59e0b');
      return;
    }
    setPlanModal({
      partner,
      activities: chosen.map((i) => ({ ...i, selected: true })),
      reqId: r.id,
    });
  };

  const rejectReq = () => {
    setStatus('rejected');
    onClose();
  };

  const markSigned = () => {
    setStatus('signed');
    onClose();
  };

  const approvedItemIds = r.items
    .filter((i) => i.itemStatus === 'approved')
    .map((i) => i.id);
  const allItemIds = approvedItemIds;
  const allSelected =
    allItemIds.length > 0 && allItemIds.every((id) => selectedItems[id]);

  const selItemIds = Object.keys(selectedItems).filter((k) => selectedItems[k]);
  const statusColor = (s) =>
    s === 'signed'
      ? C.success
      : s === 'rejected'
      ? C.danger
      : s === 'sent_for_signature'
      ? C.purple
      : C.warning;
  const statusLabel = (s) =>
    s === 'request_submitted'
      ? 'Request Submitted'
      : s === 'sent_for_signature'
      ? 'Sent for Signature'
      : s === 'signed'
      ? 'Approved & Signed'
      : s === 'rejected'
      ? 'Rejected'
      : s;
  const inp = inpStyle;
  const F = FormField;

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.88)',
        zIndex: 300,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <div
        style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 20,
          width: 740,
          maxHeight: '92vh',
          overflowY: 'auto',
          boxSizing: 'border-box',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            padding: '20px 28px 16px',
            borderBottom: `1px solid ${C.border}`,
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            position: 'sticky',
            top: 0,
            background: C.card,
            zIndex: 1,
          }}
        >
          <div>
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 10,
                marginBottom: 4,
              }}
            >
              <span
                style={{
                  fontFamily: 'monospace',
                  fontSize: 12,
                  color: C.muted,
                }}
              >
                {r.id}
              </span>
              <span
                style={{
                  background: statusColor(r.status) + '20',
                  color: statusColor(r.status),
                  borderRadius: 6,
                  padding: '2px 10px',
                  fontSize: 11,
                  fontWeight: 700,
                }}
              >
                {statusLabel(r.status)}
              </span>
              {hasEdits && (
                <span
                  style={{
                    background: C.warning + '20',
                    color: C.warning,
                    borderRadius: 6,
                    padding: '2px 8px',
                    fontSize: 10,
                    fontWeight: 700,
                  }}
                >
                  Unsaved changes
                </span>
              )}
            </div>
            <div
              style={{
                fontFamily: "'Syne',sans-serif",
                fontWeight: 800,
                fontSize: 20,
              }}
            >
              {r.partner}
            </div>
            <div
              style={{
                fontSize: 12,
                color: C.muted,
                marginTop: 2,
                display: 'flex',
                alignItems: 'center',
                gap: 10,
              }}
            >
              <span>
                Submitted: {r.submitted} . {r.items.length}{' '}
                {r.items.length === 1 ? 'activity' : 'activities'}
              </span>
              {r.partnerNotified &&
                (r.status === 'signed' || r.status === 'po_raised') && (
                  <span
                    title={
                      r.notifiedAt ? 'Signed BP emailed on ' + r.notifiedAt : ''
                    }
                    style={{
                      background: C.success + '18',
                      color: C.success,
                      border: `1px solid ${C.success}30`,
                      borderRadius: 6,
                      padding: '2px 8px',
                      fontSize: 10,
                      fontWeight: 700,
                    }}
                  >
                    Signed BP Sent to Partner
                  </span>
                )}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button
                onClick={saveChanges}
                style={{
                  background: hasEdits ? C.success : C.faint,
                  color: hasEdits ? '#000' : C.muted,
                  border: hasEdits ? 'none' : `1px solid ${C.border}`,
                  borderRadius: 10,
                  padding: '8px 16px',
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: 'pointer',
                }}
              >
                Save Changes
              </button>
            <button
              onClick={() => {
                if (hasEdits && !window.confirm('You have unsaved changes. Close without saving?')) return;
                onClose();
              }}
              style={{
                background: 'transparent',
                border: `1px solid ${C.border}`,
                color: C.muted,
                borderRadius: 10,
                padding: '8px 14px',
                fontSize: 13,
                cursor: 'pointer',
              }}
            >
              Close
            </button>
          </div>
        </div>
        <div style={{ padding: '20px 28px' }}>
          {/* Partner info - editable */}
          <div
            style={{
              background: C.faint,
              borderRadius: 10,
              padding: '14px 16px',
              marginBottom: 20,
            }}
          >
            <div
              style={{
                fontSize: 10,
                fontWeight: 800,
                color: C.accent,
                letterSpacing: '0.1em',
                marginBottom: 12,
              }}
            >
              PARTNER INFORMATION
            </div>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr 1fr',
                gap: 12,
                marginBottom: 12,
              }}
            >
              {[
                ['Partner Name', 'partnerName', r.partner],
                [
                  'Partner Type',
                  'partnerType',
                  r.partnerType ||
                    partners.find((p) => p.name === r.partner)?.type ||
                    '',
                ],
                [
                  'Partner Level',
                  'partnerTier',
                  r.partnerTier ||
                    partners.find((p) => p.name === r.partner)?.tier ||
                    '',
                ],
              ].map(([label, field, val]) => (
                <div key={field}>
                  <div
                    style={{
                      fontSize: 9,
                      color: C.muted,
                      fontWeight: 700,
                      letterSpacing: '0.08em',
                      marginBottom: 4,
                    }}
                  >
                    {label.toUpperCase()}
                  </div>
                  <input
                    value={r[field] !== undefined ? r[field] : val}
                    onChange={(e) => {
                      setHasEdits(true);
                      setR((prev) => ({ ...prev, [field]: e.target.value }));
                    }}
                    style={{
                      width: '100%',
                      background: C.card,
                      border: `1px solid ${C.border}`,
                      color: C.text,
                      borderRadius: 6,
                      padding: '5px 10px',
                      fontSize: 12,
                      fontFamily: 'inherit',
                    }}
                  />
                </div>
              ))}
            </div>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr 1fr',
                gap: 12,
              }}
            >
              {[
                [
                  'Contact Name',
                  'partnerContact',
                  r.partnerContact ||
                    partners.find((p) => p.name === r.partner)?.contactName ||
                    '',
                ],
                [
                  'Contact Email',
                  'partnerEmail',
                  r.partnerEmail ||
                    partners.find((p) => p.name === r.partner)?.contactEmail ||
                    '',
                ],
                [
                  'OT Partner Mgr',
                  'partnerManager',
                  r.partnerManager ||
                    partners.find((p) => p.name === r.partner)
                      ?.accountManager ||
                    '',
                ],
              ].map(([label, field, val]) => (
                <div key={field}>
                  <div
                    style={{
                      fontSize: 9,
                      color: C.muted,
                      fontWeight: 700,
                      letterSpacing: '0.08em',
                      marginBottom: 4,
                    }}
                  >
                    {label.toUpperCase()}
                  </div>
                  <input
                    value={r[field] !== undefined ? r[field] : val}
                    onChange={(e) => {
                      setHasEdits(true);
                      setR((prev) => ({ ...prev, [field]: e.target.value }));
                    }}
                    style={{
                      width: '100%',
                      background: C.card,
                      border: `1px solid ${C.border}`,
                      color: C.text,
                      borderRadius: 6,
                      padding: '5px 10px',
                      fontSize: 12,
                      fontFamily: 'inherit',
                    }}
                  />
                </div>
              ))}
            </div>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 12,
                marginTop: 12,
              }}
            >
              <div>
                <div
                  style={{
                    fontSize: 9,
                    color: C.muted,
                    fontWeight: 700,
                    letterSpacing: '0.08em',
                    marginBottom: 4,
                  }}
                >
                  SUBMITTED
                </div>
                <div style={{ fontSize: 12, color: C.muted, padding: '6px 0' }}>
                  {r.submitted}
                </div>
              </div>
            </div>
          </div>
          <div
            style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              marginBottom: 10,
            }}
          >
            <div
              style={{
                fontSize: 11,
                fontWeight: 800,
                color: C.accent,
                letterSpacing: '0.1em',
              }}
            >
              ACTIVITIES ({r.items.length})
            </div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              {r.items.some((it) => it.itemStatus === 'approved') ? (
                <label
                  style={{
                    fontSize: 11,
                    color: C.muted,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: 6,
                  }}
                >
                  <div
                    onClick={() =>
                      setSelectedItems(
                        allSelected
                          ? {}
                          : Object.fromEntries(
                              approvedItemIds.map((id) => [id, true])
                            )
                      )
                    }
                    style={{
                      width: 16,
                      height: 16,
                      borderRadius: 3,
                      border: `2px solid ${allSelected ? C.success : C.muted}`,
                      background: allSelected ? C.success : 'transparent',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      cursor: 'pointer',
                      flexShrink: 0,
                    }}
                  >
                    {allSelected && (
                      <span
                        style={{ color: '#000', fontSize: 10, fontWeight: 900 }}
                      >
                        v
                      </span>
                    )}
                  </div>
                  Select all approved for BP ({approvedItemIds.length})
                </label>
              ) : (
                <span style={{ fontSize: 11, color: C.muted }}>
                  Approve activities, then select for BP
                </span>
              )}
            </div>
          </div>
          {r.items.map((item, idx) => (
            <div
              key={item.id}
              style={{
                background: C.faint,
                border: `1px solid ${C.border}`,
                borderRadius: 12,
                padding: 16,
                marginBottom: 14,
              }}
            >
              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  marginBottom: 12,
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  {/* Checkbox - only for approved items to include in BP */}
                  {item.itemStatus === 'approved' && (
                    <div
                      onClick={() =>
                        setSelectedItems((p) => ({
                          ...p,
                          [item.id]: !p[item.id],
                        }))
                      }
                      style={{
                        width: 18,
                        height: 18,
                        borderRadius: 4,
                        border: `2px solid ${
                          selectedItems[item.id] ? C.success : C.muted
                        }`,
                        background: selectedItems[item.id]
                          ? C.success
                          : 'transparent',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        cursor: 'pointer',
                        flexShrink: 0,
                      }}
                    >
                      {selectedItems[item.id] && (
                        <span
                          style={{
                            color: '#000',
                            fontSize: 11,
                            fontWeight: 900,
                          }}
                        >
                          v
                        </span>
                      )}
                    </div>
                  )}
                  <span
                    style={{ fontSize: 12, fontWeight: 800, color: C.accent }}
                  >
                    Activity {idx + 1}
                  </span>
                  <span
                    style={{
                      fontSize: 10,
                      background: C.card,
                      borderRadius: 6,
                      padding: '2px 8px',
                      color: C.muted,
                    }}
                  >
                    {item.tactic || '-'}
                  </span>
                </div>
                {/* Per-item status + Approve/Reject */}
                <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                  {(() => {
                    const iStatus = item.itemStatus || 'request_submitted';
                    if (iStatus === 'approved')
                      return (
                        <div
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: 6,
                          }}
                        >
                          <span
                            style={{
                              background: C.success + '20',
                              color: C.success,
                              border: `1px solid ${C.success}40`,
                              borderRadius: 8,
                              padding: '4px 10px',
                              fontSize: 11,
                              fontWeight: 700,
                            }}
                          >
                            Approved
                          </span>
                          <button
                            onClick={() => {
                              setHasEdits(true);
                              setR((prev) => ({
                                ...prev,
                                items: prev.items.map((it) =>
                                  it.id === item.id
                                    ? { ...it, itemStatus: 'request_submitted' }
                                    : it
                                ),
                              }));
                            }}
                            style={{
                              background: 'transparent',
                              border: `1px solid ${C.border}`,
                              color: C.muted,
                              borderRadius: 6,
                              padding: '4px 8px',
                              fontSize: 10,
                              cursor: 'pointer',
                            }}
                          >
                            Undo
                          </button>
                        </div>
                      );
                    if (iStatus === 'rejected')
                      return (
                        <div
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: 6,
                          }}
                        >
                          <span
                            style={{
                              background: C.danger + '20',
                              color: C.danger,
                              border: `1px solid ${C.danger}40`,
                              borderRadius: 8,
                              padding: '4px 10px',
                              fontSize: 11,
                              fontWeight: 700,
                            }}
                          >
                            Rejected
                          </span>
                          <button
                            onClick={() => {
                              setHasEdits(true);
                              setR((prev) => ({
                                ...prev,
                                items: prev.items.map((it) =>
                                  it.id === item.id
                                    ? { ...it, itemStatus: 'request_submitted' }
                                    : it
                                ),
                              }));
                            }}
                            style={{
                              background: 'transparent',
                              border: `1px solid ${C.border}`,
                              color: C.muted,
                              borderRadius: 6,
                              padding: '4px 8px',
                              fontSize: 10,
                              cursor: 'pointer',
                            }}
                          >
                            Undo
                          </button>
                        </div>
                      );
                    if (iStatus === 'cancelled_by_partner')
                      return (
                        <div
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: 6,
                            flexWrap: 'wrap',
                          }}
                        >
                          <span
                            style={{
                              background: C.danger + '20',
                              color: C.danger,
                              border: `1px solid ${C.danger}40`,
                              borderRadius: 8,
                              padding: '4px 10px',
                              fontSize: 11,
                              fontWeight: 700,
                            }}
                          >
                            Cancelled by Partner
                          </span>
                          {item.cancelReason && (
                            <span
                              style={{
                                fontSize: 10,
                                color: C.muted,
                                fontStyle: 'italic',
                              }}
                            >
                              "{item.cancelReason}"
                            </span>
                          )}

                        </div>
                      );
                    if (iStatus === 'postponed')
                      return (
                        <div
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: 6,
                            flexWrap: 'wrap',
                          }}
                        >
                          <span
                            style={{
                              background: C.warning + '20',
                              color: C.warning,
                              border: `1px solid ${C.warning}40`,
                              borderRadius: 8,
                              padding: '4px 10px',
                              fontSize: 11,
                              fontWeight: 700,
                            }}
                          >
                            Postponed
                            {item.postponedTo ? ` to ${item.postponedTo}` : ''}
                          </span>
                          {item.cancelReason && (
                            <span
                              style={{
                                fontSize: 10,
                                color: C.muted,
                                fontStyle: 'italic',
                              }}
                            >
                              "{item.cancelReason}"
                            </span>
                          )}
                          <button
                            onClick={() => {
                              setHasEdits(true);
                              setR((prev) => ({
                                ...prev,
                                items: prev.items.map((it) =>
                                  it.id === item.id
                                    ? { ...it, itemStatus: 'approved', acknowledged: true }
                                    : it
                                ),
                              }));
                            }}
                            style={{
                              background: 'transparent',
                              border: `1px solid ${C.border}`,
                              color: C.muted,
                              borderRadius: 6,
                              padding: '4px 8px',
                              fontSize: 10,
                              cursor: 'pointer',
                            }}
                          >
                            Restore
                          </button>
                        </div>
                      );
                    return (
                      <div
                        style={{
                          display: 'flex',
                          gap: 6,
                          flexWrap: 'wrap',
                          alignItems: 'center',
                        }}
                      >
                        <select
                          value={item.assignedTo || ''}
                          onChange={(e) =>
                            setItem(idx, 'assignedTo', e.target.value)
                          }
                          style={{
                            background: C.faint,
                            border: `1px solid ${C.border}`,
                            color: C.text,
                            borderRadius: 6,
                            padding: '4px 8px',
                            fontSize: 11,
                            cursor: 'pointer',
                          }}
                        >
                          <option value="">Assign CMM...</option>
                          {TEAM_MEMBERS.map((m) => (
                            <option key={m} value={m}>
                              {m}
                            </option>
                          ))}
                        </select>
                        <button
                          disabled={item.itemStatus === 'approved'}
                          onClick={() => {
                            if (item.itemStatus === 'approved') return; // BREAK-006: no double-approve
                            setHasEdits(true);
                            setR((prev) => ({
                              ...prev,
                              assignedTo: prev.assignedTo || currentUser,
                              items: prev.items.map((it) =>
                                it.id === item.id
                                  ? {
                                      ...it,
                                      itemStatus: 'approved',
                                      assignedTo: it.assignedTo || prev.assignedTo || currentUser,
                                    }
                                  : { ...it, assignedTo: it.assignedTo || prev.assignedTo || currentUser }
                              ),
                            }));
                          }}
                          style={{
                            background: item.itemStatus === 'approved' ? C.faint : C.success,
                            color: item.itemStatus === 'approved' ? C.muted : '#000',
                            border: 'none',
                            borderRadius: 8,
                            padding: '5px 12px',
                            fontSize: 11,
                            fontWeight: 700,
                            cursor: item.itemStatus === 'approved' ? 'default' : 'pointer',
                          }}
                        >
                          {item.itemStatus === 'approved' ? '✓ Approved' : 'Approve'}
                        </button>
                        <button
                          onClick={() => {
                            setHasEdits(true);
                            setR((prev) => ({
                              ...prev,
                              items: prev.items.map((it) =>
                                it.id === item.id
                                  ? { ...it, itemStatus: 'rejected' }
                                  : it
                              ),
                            }));
                          }}
                          style={{
                            background: 'transparent',
                            color: C.danger,
                            border: `1px solid ${C.danger}`,
                            borderRadius: 8,
                            padding: '5px 12px',
                            fontSize: 11,
                            fontWeight: 700,
                            cursor: 'pointer',
                          }}
                        >
                          Reject
                        </button>
                      </div>
                    );
                  })()}
                </div>
              </div>
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr 1fr',
                  gap: 10,
                  marginBottom: 10,
                }}
              >
                <F label="FY Half">
                  <select
                    value={item.fyHalf || ''}
                    onChange={(e) => setItem(idx, 'fyHalf', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  >
                    <option value="">Select...</option>
                    {FY_HALVES.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </F>
                <F label="Quarter">
                  <select
                    value={item.fyQuarter || ''}
                    onChange={(e) => setItem(idx, 'fyQuarter', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                    disabled={!item.fyHalf}
                  >
                    <option value="">Select...</option>
                    {getQuarters(item.fyHalf).map((q) => (
                      <option key={q} value={q}>
                        {q}
                      </option>
                    ))}
                  </select>
                </F>
                <F label="Month">
                  <select
                    value={item.month || ''}
                    onChange={(e) => setItem(idx, 'month', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                    disabled={!item.fyQuarter}
                  >
                    <option value="">Select...</option>
                    {getMonths(item.fyQuarter).map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </F>
              </div>
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr',
                  gap: 10,
                  marginBottom: 10,
                }}
              >
                <F label="Product Group">
                  <select
                    value={item.productGroup || ''}
                    onChange={(e) =>
                      setItem(idx, 'productGroup', e.target.value)
                    }
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  >
                    <option value="">Select...</option>
                    {PRODUCT_GROUPS.map((g) => (
                      <option key={g} value={g}>
                        {g}
                      </option>
                    ))}
                  </select>
                </F>
                <F label="Tactic">
                  <select
                    value={item.tactic || ''}
                    onChange={(e) => setItem(idx, 'tactic', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  >
                    <option value="">Select...</option>
                    {TACTICS.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </F>
              </div>
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr',
                  gap: 10,
                  marginBottom: 10,
                }}
              >
                <F label="Location">
                  <input
                    value={item.where || ''}
                    onChange={(e) => setItem(idx, 'where', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  />
                </F>
                <F label="Target Audience">
                  <input
                    value={item.targetAudience || ''}
                    onChange={(e) =>
                      setItem(idx, 'targetAudience', e.target.value)
                    }
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  />
                </F>
              </div>
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr',
                  gap: 10,
                  marginBottom: 10,
                }}
              >
                <F label="Solutions">
                  <input
                    value={
                      Array.isArray(item.targetSolutions)
                        ? item.targetSolutions.join(', ')
                        : item.targetSolutions || ''
                    }
                    onChange={(e) =>
                      setItem(idx, 'targetSolutions', e.target.value)
                    }
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  />
                </F>
                <F label="Objective">
                  <input
                    value={item.objective || ''}
                    onChange={(e) => setItem(idx, 'objective', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  />
                </F>
              </div>
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: '2fr 1fr 2fr',
                  gap: 10,
                  marginBottom: 10,
                }}
              >
                <F label="Total Cost">
                  <input
                    value={item.amount || ''}
                    onChange={(e) => setItem(idx, 'totalCost', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  />
                </F>
                <F label="Currency">
                  <select
                    value={item.localCurrency || 'EUR'}
                    onChange={(e) =>
                      setItem(idx, 'localCurrency', e.target.value)
                    }
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  >
                    {[
                      'EUR',
                      'USD',
                      'GBP',
                      'CHF',
                      'SEK',
                      'NOK',
                      'DKK',
                      'PLN',
                      'AED',
                      'SGD',
                      'AUD',
                    ].map((c) => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </F>
                <F label="MDF Request">
                  <input
                    value={
                      item.mdfRequest || Math.round((item.amount || 0) * 0.5)
                    }
                    onChange={(e) => setItem(idx, 'mdfRequest', e.target.value)}
                    style={{ ...inp(false), width: '100%', fontSize: 12 }}
                  />
                </F>
              </div>
              {(item.amount || 0) > 0 && (
                <div
                  style={{
                    fontSize: 11,
                    color: C.muted,
                    marginTop: 2,
                    fontFamily: 'monospace',
                  }}
                >
                  MDF ={' '}
                  {Math.round(
                    ((item.mdfRequest || item.amount * 0.5) / item.amount) * 100
                  )}
                  % of total cost
                </div>
              )}
              {/* Allocadia ID - editable when approved */}
              {item.itemStatus === 'approved' && (
                <div
                  style={{
                    marginTop: 10,
                    paddingTop: 10,
                    borderTop: `1px solid ${C.border}`,
                    display: 'flex',
                    alignItems: 'center',
                    gap: 10,
                  }}
                >
                  <span
                    style={{
                      fontSize: 10,
                      color: C.muted,
                      fontWeight: 700,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    ALLOCADIA ID
                  </span>
                  <input
                    value={item.allocadiaId || ''}
                    onChange={(e) => {
                      setHasEdits(true);
                      setR((prev) => ({
                        ...prev,
                        items: prev.items.map((it) =>
                          it.id === item.id
                            ? { ...it, allocadiaId: e.target.value }
                            : it
                        ),
                      }));
                    }}
                    placeholder="e.g. ALO-2026-001"
                    style={{
                      flex: 1,
                      background: C.card,
                      border: `1px solid ${C.accent}40`,
                      color: C.text,
                      borderRadius: 6,
                      padding: '5px 10px',
                      fontSize: 12,
                      fontFamily: 'monospace',
                      outline: 'none',
                    }}
                  />
                </div>
              )}
              {/* Campaign ID - mandatory when approved */}
              {item.itemStatus === 'approved' && (
                <div style={{ marginTop: 8, paddingTop: 8, borderTop: `1px solid ${C.border}`, display: 'flex', alignItems: 'center', gap: 10 }}>
                  <span style={{ fontSize: 10, color: !item.campaignId ? C.danger : C.muted, fontWeight: 700, whiteSpace: 'nowrap' }}>
                    CAMPAIGN ID *
                  </span>
                  <input
                    value={item.campaignId || ''}
                    onChange={(e) => {
                      setHasEdits(true);
                      setR((prev) => ({
                        ...prev,
                        items: prev.items.map((it) =>
                          it.id === item.id ? { ...it, campaignId: e.target.value } : it
                        ),
                      }));
                    }}
                    placeholder="Required for pipeline tracking"
                    style={{
                      flex: 1, background: C.card,
                      border: `1px solid ${!item.campaignId ? C.danger : C.accent + '40'}`,
                      color: C.text, borderRadius: 6, padding: '5px 10px',
                      fontSize: 12, fontFamily: 'monospace', outline: 'none',
                    }}
                  />
                  {!item.campaignId && (
                    <span style={{ fontSize: 9, color: C.danger, whiteSpace: 'nowrap', fontWeight: 700 }}>Required</span>
                  )}
                </div>
              )}
              {item.itemStatus !== 'approved' && item.allocadiaId && (
                <div style={{ marginTop: 8, fontSize: 10, fontFamily: 'monospace', color: C.muted }}>
                  Allocadia: {item.allocadiaId}
                </div>
              )}
              {item.itemStatus !== 'approved' && item.campaignId && (
                <div style={{ marginTop: 4, fontSize: 10, fontFamily: 'monospace', color: C.muted }}>
                  Campaign ID: {item.campaignId}
                </div>
              )}
            </div>
          ))}
          <div style={{ marginBottom: 20 }}>
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: C.text,
                marginBottom: 6,
              }}
            >
              Internal Note
            </div>
            <LocalTextarea
              value={note}
              onCommit={(v) => {
                setNote(v);
                setHasEdits(true);
              }}
              placeholder="Add review notes..."
            />
          </div>
          {/* -- WORKFLOW ACTIONS -- */}
          <div style={{ marginBottom: 20 }}>
            <div
              style={{
                fontSize: 10,
                fontWeight: 800,
                color: C.accent,
                letterSpacing: '0.1em',
                marginBottom: 12,
              }}
            >
              WORKFLOW ACTIONS
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {/* Generate BP v only when activities are approved and status is still request_submitted */}
              {r.items.some((it) => it.itemStatus === 'approved') &&
                !['rejected','cancelled_by_partner'].includes(r.status) &&
                r.status === 'request_submitted' && (
                  <div>
                    <button
                      onClick={approveAndGenerate}
                      style={{
                        width: '100%',
                        background: C.success,
                        color: '#000',
                        border: 'none',
                        borderRadius: 10,
                        padding: '11px 16px',
                        fontWeight: 800,
                        fontSize: 13,
                        cursor: 'pointer',
                        textAlign: 'left',
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                      }}
                    >
                      <span>Generate Business Plan</span>
                      <span style={{ fontSize: 11, opacity: 0.8 }}>
                        {r.items.filter((i) => i.itemStatus === 'approved').length}{' '}
                        {r.items.filter((i) => i.itemStatus === 'approved').length === 1 ? 'activity' : 'activities'}
                      </span>
                    </button>
                    {r.bpGeneratedAt && (
                      <div style={{ fontSize: 10, color: C.muted, marginTop: 5, textAlign: 'center' }}>
                        BP previously generated: {r.bpGeneratedAt}
                      </div>
                    )}
                  </div>
                )}

              {/* APPROVED but no BP yet → show Generate BP button */}
              {r.status === 'approved' && !r.bpGeneratedAt && (
                <div style={{ background: C.warning + '15', border: `1px solid ${C.warning}40`, borderRadius: 10, padding: '14px 16px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}>
                  <div>
                    <div style={{ fontSize: 13, color: C.warning, fontWeight: 700, marginBottom: 2 }}>Activities approved — Business Plan not yet generated</div>
                    <div style={{ fontSize: 11, color: C.muted }}>Generate and download the BP before sending for signature</div>
                  </div>
                  <button
                    onClick={approveAndGenerate}
                    style={{ background: C.success, color: '#000', border: 'none', borderRadius: 10, padding: '10px 18px', fontWeight: 800, fontSize: 13, cursor: 'pointer', whiteSpace: 'nowrap', flexShrink: 0 }}
                  >
                    Generate BP →
                  </button>
                </div>
              )}

              {/* APPROVED + BP generated → confirm sent for signature */}
              {r.status === 'approved' && r.bpGeneratedAt && (
                <div
                  style={{
                    background: C.faint,
                    borderRadius: 10,
                    padding: '12px 16px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: 12,
                  }}
                >
                  <div style={{ flex: 1 }}>
                    <div
                      style={{
                        fontSize: 12,
                        color: C.success,
                        fontWeight: 700,
                        marginBottom: 2,
                      }}
                    >
                      Business Plan downloaded
                    </div>
                    <div style={{ fontSize: 11, color: C.muted }}>
                      Send the BP via your signature tool, then confirm below
                    </div>
                  </div>
                  <button
                    onClick={() => setStatus('sent_for_signature')}
                    style={{
                      background: '#8b5cf6',
                      color: '#fff',
                      border: 'none',
                      borderRadius: 8,
                      padding: '7px 14px',
                      fontSize: 11,
                      fontWeight: 700,
                      cursor: 'pointer',
                      whiteSpace: 'nowrap',
                      flexShrink: 0,
                    }}
                  >
                    Confirm Sent for Signature
                  </button>
                </div>
              )}

              {/* SENT FOR SIGNATURE v just confirm it's been sent */}
              {r.status === 'sent_for_signature' && (
                <div
                  style={{
                    background: C.faint,
                    borderRadius: 10,
                    padding: '12px 16px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                  }}
                >
                  <div
                    style={{ fontSize: 12, color: '#8b5cf6', fontWeight: 600 }}
                  >
                    BP sent for signature via your tool
                  </div>
                  <button
                    onClick={() => setStatus('sent_for_signature')}
                    style={{
                      background: 'transparent',
                      border: `1px solid #8b5cf6`,
                      color: '#8b5cf6',
                      borderRadius: 8,
                      padding: '5px 12px',
                      fontSize: 11,
                      fontWeight: 700,
                      cursor: 'pointer',
                    }}
                  >
                    Confirm Sent
                  </button>
                </div>
              )}

              {/* SIGNED v Step 1: Upload signed BP. Step 2: Notify partner */}
              {r.status === 'signed' && (
                <div
                  style={{
                    background: C.faint,
                    borderRadius: 10,
                    padding: '14px 16px',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: 10,
                  }}
                >
                  <div
                    style={{
                      fontSize: 10,
                      fontWeight: 800,
                      color: '#06b6d4',
                      letterSpacing: '0.08em',
                    }}
                  >
                    SIGNED v COMPLETE THE STEPS BELOW
                  </div>

                  {/* Step 1: Upload signed doc */}
                  <div>
                    <div
                      style={{
                        fontSize: 10,
                        color: C.muted,
                        fontWeight: 700,
                        marginBottom: 6,
                      }}
                    >
                      1. UPLOAD SIGNED BUSINESS PLAN
                    </div>
                    <label style={{ cursor: 'pointer', display: 'block' }}>
                      <input
                        type="file"
                        accept=".pdf,.doc,.docx"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (!file) return;
                          const reader = new FileReader();
                          reader.onload = (ev) => {
                            const doc = {
                              name: file.name,
                              size: file.size,
                              dataUrl: ev.target.result,
                            };
                            setSignedDoc(doc);
                            setRequests((prev) =>
                              prev.map((x) =>
                                x.id === r.id ? { ...x, signedDoc: doc } : x
                              )
                            );
                            setR((prev) => ({ ...prev, signedDoc: doc }));
                            toast('Signed document uploaded.');
                          };
                          reader.readAsDataURL(file);
                        }}
                        style={{ display: 'none' }}
                      />
                      {signedDoc || r.signedDoc ? (
                        <div
                          style={{
                            background: C.card,
                            border: `1px solid #06b6d440`,
                            borderRadius: 8,
                            padding: '8px 12px',
                            display: 'flex',
                            justifyContent: 'space-between',
                            alignItems: 'center',
                          }}
                        >
                          <div
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              gap: 8,
                            }}
                          >
                            <div
                              style={{
                                width: 28,
                                height: 28,
                                borderRadius: 6,
                                background: '#06b6d420',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                fontSize: 9,
                                fontWeight: 700,
                                color: '#06b6d4',
                                flexShrink: 0,
                              }}
                            >
                              PDF
                            </div>
                            <div>
                              <div
                                style={{
                                  fontSize: 11,
                                  fontWeight: 600,
                                  color: '#06b6d4',
                                }}
                              >
                                {(signedDoc || r.signedDoc).name}
                              </div>
                              <div style={{ fontSize: 9, color: C.muted }}>
                                {(
                                  (signedDoc || r.signedDoc).size / 1024
                                ).toFixed(0)}{' '}
                                KB
                              </div>
                            </div>
                          </div>
                          <button
                            onClick={(e) => {
                              e.preventDefault();
                              setSignedDoc(null);
                              setR((prev) => ({ ...prev, signedDoc: null }));
                              setRequests((prev) =>
                                prev.map((x) =>
                                  x.id === r.id ? { ...x, signedDoc: null } : x
                                )
                              );
                            }}
                            style={{
                              background: 'transparent',
                              border: 'none',
                              color: C.muted,
                              cursor: 'pointer',
                              fontSize: 14,
                            }}
                          >
                            x
                          </button>
                        </div>
                      ) : (
                        <div
                          style={{
                            background: C.card,
                            border: `2px dashed ${C.border}`,
                            borderRadius: 8,
                            padding: '10px 14px',
                            textAlign: 'center',
                          }}
                        >
                          <div style={{ fontSize: 12, color: C.muted }}>
                            Click to upload signed Business Plan (PDF or Word)
                          </div>
                        </div>
                      )}
                    </label>
                  </div>

                  {/* Step 2: Send email to partner */}
                  <div>
                    <div
                      style={{
                        fontSize: 10,
                        color: C.muted,
                        fontWeight: 700,
                        marginBottom: 6,
                      }}
                    >
                      2. NOTIFY PARTNER BY EMAIL
                    </div>
                    <button
                      onClick={() => {
                        const partnerRecord =
                          partners.find((p) => p.name === r.partner) || {};
                        const toEmail =
                          r.partnerEmail || partnerRecord.contactEmail || '';
                        const toName =
                          r.partnerContact ||
                          partnerRecord.contactName ||
                          r.partner;
                        const acts = (r.items || []).filter(
                          (i) => i.itemStatus === 'approved'
                        );
                        const actLines = acts
                          .map(
                            (a) =>
                              `  - ${a.title || a.tactic} (${a.fyHalf} ${
                                a.fyQuarter
                              } ${a.month})`
                          )
                          .join('\n');
                        const subject = encodeURIComponent(
                          `MDF Business Plan Signed - ${r.id} | OT`
                        );
                        const emailBody = [
                          `Dear ${toName},`,
                          ``,
                          `I am pleased to confirm that your MDF Business Plan has been signed by both parties.`,
                          ``,
                          `Request ID: ${r.id}`,
                          ``,
                          `Approved Activities:`,
                          actLines,
                          ``,
                          `You may now proceed with the execution of the approved activities.`,
                          `Once completed, please submit your claim through the Partner Portal,`,
                          `uploading all relevant invoices and supporting documentation.`,
                          ``,
                          signedDoc || r.signedDoc
                            ? `The signed Business Plan document is attached to this email for your records.`
                            : ``,
                          ``,
                          `Should you have any questions, please do not hesitate to reach out.`,
                          ``,
                          `Best regards,`,
                          `Channel Marketing Team | OT`,
                        ]
                          .filter((l) => l !== undefined)
                          .join('\n');
                        window.open(
                          `mailto:${toEmail}?subject=${subject}&body=${encodeURIComponent(
                            emailBody
                          )}`,
                          '_blank'
                        );
                        setRequests((prev) =>
                          prev.map((x) =>
                            x.id === r.id
                              ? {
                                  ...x,
                                  partnerNotified: true,
                                  notifiedAt: new Date().toLocaleString(
                                    'en-GB',
                                    {
                                      day: '2-digit',
                                      month: 'short',
                                      year: 'numeric',
                                      hour: '2-digit',
                                      minute: '2-digit',
                                    }
                                  ),
                                }
                              : x
                          )
                        );
                        setR((prev) => ({ ...prev, partnerNotified: true }));
                        addHistory(
                          `Signed BP emailed to partner - ${r.partner}`,
                          r.id,
                          'approve'
                        );
                        toast(
                          'Email client opened' +
                            (signedDoc || r.signedDoc
                              ? ' - attach signed doc before sending.'
                              : '.')
                        );
                      }}
                      style={{
                        width: '100%',
                        background: '#06b6d4',
                        color: '#000',
                        border: 'none',
                        borderRadius: 8,
                        padding: '9px 14px',
                        fontWeight: 700,
                        fontSize: 12,
                        cursor: 'pointer',
                      }}
                    >
                      Open Notification Email
                    </button>
                    {(signedDoc || r.signedDoc) && (
                      <div
                        style={{
                          fontSize: 10,
                          color: C.muted,
                          marginTop: 4,
                          textAlign: 'center',
                        }}
                      >
                        Remember to attach the signed document before sending
                      </div>
                    )}
                  </div>

                  {/* Step 3: Confirm PO raised */}
                  <div>
                    <div
                      style={{
                        fontSize: 10,
                        color: C.muted,
                        fontWeight: 700,
                        marginBottom: 6,
                      }}
                    >
                      3. CONFIRM PO RAISED
                    </div>
                    {!r.partnerNotified && (
                      <div style={{ fontSize: 11, color: C.warning, marginBottom: 8, padding: '6px 10px', background: C.warning + '15', borderRadius: 6 }}>
                        Complete step 2 (notify partner) before raising PO
                      </div>
                    )}
                    <button
                      disabled={!r.partnerNotified}
                      onClick={() => setStatus('po_raised')}
                      style={{
                        width: '100%',
                        background: r.partnerNotified ? C.accent : C.faint,
                        color: r.partnerNotified ? '#fff' : C.muted,
                        border: 'none',
                        borderRadius: 8,
                        padding: '9px 14px',
                        fontWeight: 700,
                        fontSize: 12,
                        cursor: r.partnerNotified ? 'pointer' : 'not-allowed',
                        opacity: r.partnerNotified ? 1 : 0.6,
                      }}
                    >
                      Confirm PO Raised
                    </button>
                  </div>
                </div>
              )}

              {/* PO RAISED v Enter PO number when ready */}
              {r.status === 'po_raised' && (
                <div
                  style={{
                    background: C.faint,
                    borderRadius: 10,
                    padding: '14px 16px',
                  }}
                >
                  <div
                    style={{
                      fontSize: 10,
                      fontWeight: 800,
                      color: C.accent,
                      letterSpacing: '0.08em',
                      marginBottom: 10,
                    }}
                  >
                    ENTER PO NUMBER WHEN READY
                  </div>
                  {(signedDoc || r.signedDoc) && (
                    <div
                      style={{
                        fontSize: 11,
                        color: C.success,
                        marginBottom: 10,
                        display: 'flex',
                        alignItems: 'center',
                        gap: 6,
                      }}
                    >
                      <span>v</span> Signed document:{' '}
                      {(signedDoc || r.signedDoc).name}
                    </div>
                  )}
                  <div style={{ display: 'flex', gap: 8 }}>
                    <input
                      value={r.poNumber || ''}
                      onChange={(e) => {
                        setHasEdits(true);
                        setR((prev) => ({ ...prev, poNumber: e.target.value }));
                      }}
                      placeholder="e.g. PO-2026-0042"
                      style={{
                        flex: 1,
                        background: C.card,
                        border: `1px solid ${C.border}`,
                        color: C.text,
                        borderRadius: 8,
                        padding: '8px 12px',
                        fontSize: 13,
                        fontFamily: 'monospace',
                        outline: 'none',
                      }}
                    />
                    <button
                      onClick={() => {
                        if (!r.poNumber) {
                          toast('Enter a PO Number first', '#f59e0b');
                          return;
                        }
                        setRequests((prev) =>
                          prev.map((x) =>
                            x.id === r.id ? { ...x, poNumber: r.poNumber } : x
                          )
                        );
                        addHistory(
                          `PO ${r.poNumber} saved for ${r.id}`,
                          r.id,
                          'approve'
                        );
                        toast(`PO ${r.poNumber} saved. Process complete!`);
                        setHasEdits(false);
                      }}
                      style={{
                        background: C.success,
                        color: '#000',
                        border: 'none',
                        borderRadius: 8,
                        padding: '8px 16px',
                        fontWeight: 700,
                        fontSize: 13,
                        cursor: 'pointer',
                        whiteSpace: 'nowrap',
                      }}
                    >
                      Save PO
                    </button>
                  </div>
                  {r.poNumber && (
                    <div
                      style={{
                        fontSize: 11,
                        color: C.success,
                        marginTop: 8,
                        fontFamily: 'monospace',
                        fontWeight: 600,
                      }}
                    >
                      v PO {r.poNumber} v Process complete
                    </div>
                  )}
                </div>
              )}
              <div
                style={{ borderTop: `1px solid ${C.border}`, margin: '4px 0' }}
              />
              <div>
                <div
                  style={{
                    fontSize: 10,
                    color: C.muted,
                    fontWeight: 700,
                    letterSpacing: '0.06em',
                    marginBottom: 8,
                  }}
                >
                  OVERRIDE STATUS (admin)
                </div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  {WORKFLOW_STEPS.map((step) => {
                    const isCurrent = r.status === step.id;
                    return (
                      <button
                        key={step.id}
                        onClick={() => {
                          if (isCurrent) return;
                          setRequests((prev) =>
                            prev.map((x) =>
                              x.id === r.id ? { ...x, status: step.id } : x
                            )
                          );
                          setR((prev) => ({ ...prev, status: step.id }));
                          setStatus(step.id);
                          addHistory(
                            `Request ${r.id} -> ${step.label}`,
                            r.id,
                            step.id === 'rejected' ? 'reject' : 'approve'
                          );
                          toast(`Status: ${step.label}`);
                        }}
                        style={{
                          background: isCurrent
                            ? step.color + '30'
                            : 'transparent',
                          color: isCurrent ? step.color : C.muted,
                          border: `1px solid ${
                            isCurrent ? step.color : C.border
                          }`,
                          borderRadius: 20,
                          padding: '4px 12px',
                          fontSize: 11,
                          fontWeight: 700,
                          cursor: isCurrent ? 'default' : 'pointer',
                          transition: 'all 0.15s',
                        }}
                      >
                        {isCurrent ? '* ' : ''}
                        {step.short}
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

const AnalyticsTab = ({
  partners,
  requests,
  claims = [],
  pipelineData = {},
  currency,
  fmtA,
  fmtB,
  toUSD,
  onExport,
  onSaveOverride,
  onImportPipeline,
  onSavePipeline,
  isDark = true,
}) => {
  const [anMacro, setAnMacro] = useState('All');
  const [anSubregion, setAnSubregion] = useState('All');
  const [anTier, setAnTier] = useState('All');
  const [anType, setAnType] = useState('All');
  const [anStatus, setAnStatus] = useState('All');
  const [anPAM, setAnPAM] = useState('All');
  const [anFyHalf, setAnFyHalf] = useState('All');
  const [anFyQ, setAnFyQ] = useState('All');
  const [anBU, setAnBU] = useState('All');
  const [anTactic, setAnTactic] = useState('All');
  const [anExec, setAnExec] = useState('All');
  const [anClaim, setAnClaim] = useState('All');
  const [anPartner, setAnPartner] = useState('All');
  const [anSearch, setAnSearch] = useState('');
  const [sortCol, setSortCol] = useState('macro');
  const [sortDir, setSortDir] = useState('asc');
  const [editCell, setEditCell] = useState(null);
  const [editVal, setEditVal] = useState('');
  const [overrides, setOverrides] = useState({});

  const EXEC_STATUSES = ['', 'Executed', 'Canceled'];
  const CLAIM_STATUSES = ['', 'Approved', 'On Hold', 'Canceled', 'Expired'];

  const ALL_FILTER_OPTS = [
    { key: 'region', label: 'Region' },
    { key: 'subregion', label: 'Sub-Region' },
    { key: 'tier', label: 'Partner Level' },
    { key: 'type', label: 'Partner Type' },
    { key: 'status', label: 'Request Status' },
    { key: 'fyHalf', label: 'Fiscal Year' },
    { key: 'fyQ', label: 'FY Quarter' },
    { key: 'bu', label: 'BU / Product Focus' },
    { key: 'tactic', label: 'Tactic' },
    { key: 'exec', label: 'Activity Status' },
    { key: 'claim', label: 'Claim Status' },
    { key: 'partner', label: 'Partner Name' },
    { key: 'pam', label: 'PAM' },
  ];
  const [activeFilters, setActiveFilters] = useState([
    'region',
    'subregion',
    'type',
    'status',
  ]);
  const [showFilterPicker, setShowFilterPicker] = useState(false);
  const toggleFilter = (key) =>
    setActiveFilters((p) =>
      p.includes(key) ? p.filter((k) => k !== key) : [...p, key]
    );

  const getOverride = (rowKey, field, def) =>
    overrides[`${rowKey}_${field}`] ?? def;
  const setOverride = (rowKey, field, val) =>
    setOverrides((p) => ({ ...p, [`${rowKey}_${field}`]: val }));

  const startEdit = (rowKey, field, current) => {
    setEditCell({ rowKey, field });
    setEditVal(String(current));
  };
  const commitEdit = (rowKey, field) => {
    if (editCell?.rowKey === rowKey && editCell?.field === field) {
      setOverride(rowKey, field, editVal);
      setEditCell(null);
      // Write allocadiaId/campaignId/poNumber back to requests state
      if (['allocadiaId','campaignId','poNumber'].includes(field) && onSaveOverride) {
        onSaveOverride(rowKey, field, editVal);
      }
      // Write pipeline back to App pipelineData (persisted)
      if (field === 'pipelineGenerated' && onSavePipeline) {
        const row = rows.find(r => r.rowKey === rowKey);
        const campaignIdOverride = overrides[`${rowKey}_campaignId`] || row?.campaignId;
        const key = campaignIdOverride || rowKey;
        onSavePipeline(key, Number(editVal) || 0);
      }
    }
  };

  const clearAll = () => {
    setAnMacro('All');
    setAnSubregion('All');
    setAnTier('All');
    setAnType('All');
    setAnStatus('All');
    setAnFyHalf('All');
    setAnFyQ('All');
    setAnBU('All');
    setAnTactic('All');
    setAnExec('All');
    setAnClaim('All');
    setAnPartner('All');
      setAnPAM('All');
    setAnSearch('');
  };
  const hasActiveFilter =
    anMacro !== 'All' ||
    anSubregion !== 'All' ||
    anTier !== 'All' ||
    anType !== 'All' ||
    anStatus !== 'All' ||
    anFyHalf !== 'All' ||
    anFyQ !== 'All' ||
    anBU !== 'All' ||
    anTactic !== 'All' ||
    anExec !== 'All' ||
    anClaim !== 'All' ||
    anPartner !== 'All' ||
    anPAM !== 'All' ||
    anSearch;

  const rows = useMemo(() => {
    const flat = [];
    requests.forEach((r) => {
      const p = partners.find((x) => x.name === r.partner) || {};
      (r.items || []).forEach((item, idx) => {
        const rowKey = item.id || `${r.id}-${idx}`;
        flat.push({
          rowKey,
          macro: p.region || '-',
          subregion: p.subregion || '-',
          country: p.country || '-',
          partnerName: r.partner || '-',
          contactName: p.contactName || r.partnerContact || '-',
          contactEmail: p.contactEmail || r.partnerEmail || '-',
          accountManager: p.accountManager || r.partnerManager || '-',
          fyHalf: item.fyHalf || '-',
          fyQuarter: item.fyQuarter || '-',
          period:
            item.month ||
            (item.period || '').replace(/\s*\(.*\)/, '').trim() ||
            '-',
          productGroup: item.productGroup || '-',
          tactic: item.tactic || '-',
          title: item.title || '-',
          allocadiaId: item.allocadiaId || '',
          campaignId: item.campaignId || '',
          partnerNotified: r.partnerNotified || false,
          totalCost: item.amount || 0,
          mdfRequest: item.mdfRequest || Math.round((item.amount || 0) * 0.5),
          localCurrency: item.localCurrency || 'EUR',
          reqStatus: (() => {
            // Use request-level status when it has advanced past item approval
            const advanced = ['sent_for_signature','signed','po_raised','rejected'];
            if (advanced.includes(r.status)) return r.status;
            const s = item.itemStatus || r.status || '-';
            if (s === 'cancelled_by_partner') return 'cancelled_by_partner';
            return s;
          })(),
          poNumber: r.poNumber || '-',
          assignedTo: item.assignedTo || '-',
          reqId: r.id || '-',
          execStatus: '',
          claimStatus: (() => {
            const actSt = item.itemStatus || r.status || '';
            if (actSt === 'cancelled_by_partner') return 'cancelled';
            if (!['signed', 'po_raised'].includes(actSt)) return '';
            // Match by itemId first (most reliable), then fall back to title
            const c = claims.find(
              (cl) =>
                cl.reqId === r.id &&
                (cl.itemId === item.id || cl.activity === item.title)
            );
            return c ? c.status : '';
          })(),
          claimAmount: (() => {
            const actSt = item.itemStatus || r.status || '';
            if (
              !['signed', 'po_raised', 'cancelled_by_partner'].includes(actSt)
            )
              return 0;
            const c = claims.find(
              (cl) =>
                cl.reqId === r.id &&
                (cl.itemId === item.id || cl.activity === item.title)
            );
            return c ? c.claimAmount || 0 : 0;
          })(),
          pipelineGenerated: pipelineData[item.campaignId] || pipelineData[item.id] || 0,
          tier: p.tier || '-',
          partnerType: p.type || '-',
        });
      });
    });
    return flat;
  }, [partners, requests, claims]);

  const filtered = useMemo(
    () =>
      rows.filter((row) => {
        if (anMacro !== 'All' && (Array.isArray(anMacro) ? (anMacro.length > 0 && !anMacro.includes(row.macro)) : row.macro !== anMacro)) return false;
        if (anSubregion !== 'All' && (Array.isArray(anSubregion) ? (anSubregion.length > 0 && !anSubregion.includes(row.subregion)) : row.subregion !== anSubregion)) return false;
        if (anTier !== 'All' && (Array.isArray(anTier) ? (anTier.length > 0 && !anTier.includes(row.tier)) : row.tier !== anTier)) return false;
        if (anType !== 'All' && (Array.isArray(anType) ? (anType.length > 0 && !anType.includes(row.partnerType)) : row.partnerType !== anType)) return false;
        if (anStatus !== 'All') {
          const claimStatusMap = {
            claim_submitted: 'submitted',
            claim_mktg: 'marketing_review',
            claim_finance: 'finance_review',
            claim_approved: 'approved',
            claim_paid: 'paid',
            claim_hold: 'on_hold',
          };
          if (Array.isArray(anStatus)) {
            const reqStatuses = anStatus.filter((s) => !s.startsWith('claim_'));
            const claimStatuses = anStatus
              .filter((s) => s.startsWith('claim_'))
              .map((s) => claimStatusMap[s]);
            const matchReq =
              reqStatuses.length === 0 || reqStatuses.includes(row.reqStatus);
            const matchClaim =
              claimStatuses.length === 0 ||
              claimStatuses.includes(row.claimStatus);
            if (reqStatuses.length > 0 && claimStatuses.length > 0) {
              if (!matchReq && !matchClaim) return false;
            } else if (reqStatuses.length > 0 && !matchReq) return false;
            else if (claimStatuses.length > 0 && !matchClaim) return false;
          } else {
            if (row.reqStatus !== anStatus) return false;
          }
        }
        if (anFyHalf !== 'All' && (Array.isArray(anFyHalf) ? (anFyHalf.length > 0 && !anFyHalf.includes(row.fyHalf)) : row.fyHalf !== anFyHalf)) return false;
        if (anFyQ !== 'All' && (Array.isArray(anFyQ) ? (anFyQ.length > 0 && !anFyQ.includes(row.fyQuarter)) : row.fyQuarter !== anFyQ)) return false;
        if (anBU !== 'All' && (Array.isArray(anBU) ? (anBU.length > 0 && !anBU.includes(row.productGroup)) : row.productGroup !== anBU)) return false;
        if (anTactic !== 'All' && (Array.isArray(anTactic) ? (anTactic.length > 0 && !anTactic.includes(row.tactic)) : row.tactic !== anTactic)) return false;
        if (
          anExec !== 'All' &&
          (() => {
            const v = getOverride(row.rowKey, 'execStatus', row.execStatus);
            return Array.isArray(anExec) ? (anExec.length > 0 && !anExec.includes(v)) : v !== anExec;
          })()
        )
          return false;
        if (anClaim !== 'All') {
          const cs = row.claimStatus || '';
          if (Array.isArray(anClaim)) {
            if (anClaim.length > 0 && !anClaim.includes(cs)) return false;
          } else if (cs !== anClaim) return false;
        }
        if (anPartner !== 'All' && (Array.isArray(anPartner) ? (anPartner.length > 0 && !anPartner.includes(row.partnerName)) : row.partnerName !== anPartner)) return false;
        if (anPAM !== 'All' && (Array.isArray(anPAM) ? (anPAM.length > 0 && !anPAM.includes(row.accountManager)) : row.accountManager !== anPAM)) return false;
        if (anSearch) {
          const q = anSearch.toLowerCase();
          if (
            ![
              row.partnerName,
              row.title,
              row.tactic,
              row.productGroup,
              row.subregion,
              row.reqId,
              row.macro,
              row.country,
            ]
              .join(' ')
              .toLowerCase()
              .includes(q)
          )
            return false;
        }
        return true;
      }),
    [
      rows,
      anMacro,
      anSubregion,
      anTier,
      anType,
      anStatus,
      anFyHalf,
      anFyQ,
      anBU,
      anTactic,
      anExec,
      anClaim,
      anPartner,
      anPAM,
      anSearch,
      overrides,
    ]
  );

  const sorted = useMemo(
    () =>
      [...filtered].sort((a, b) => {
        let av = getOverride(a.rowKey, sortCol, a[sortCol]) ?? '';
        let bv = getOverride(b.rowKey, sortCol, b[sortCol]) ?? '';
        if (typeof av === 'number' && typeof bv === 'number')
          return sortDir === 'asc' ? av - bv : bv - av;
        return sortDir === 'asc'
          ? String(av).localeCompare(String(bv))
          : String(bv).localeCompare(String(av));
      }),
    [filtered, sortCol, sortDir, overrides]
  );

  const totalCost = filtered.reduce((s, r) => s + r.totalCost, 0);
  const totalMDF = filtered.reduce((s, r) => s + r.mdfRequest, 0);
  const totalClaim = filtered.reduce(
    (s, r) =>
      s + (Number(getOverride(r.rowKey, 'claimAmount', r.claimAmount)) || 0),
    0
  );
  const totalPipeline = filtered.reduce(
    (s, r) =>
      s +
      (Number(
        getOverride(r.rowKey, 'pipelineGenerated', r.pipelineGenerated)
      ) || 0),
    0
  );
  const uniquePartners = [...new Set(filtered.map((r) => r.partnerName))]
    .length;
  const uniqueReqs = [...new Set(filtered.map((r) => r.reqId))].length;

  const statusLabel = (s) =>
    s === 'request_submitted'
      ? 'Submitted'
      : s === 'approved'
      ? 'Approved'
      : s === 'sent_for_signature'
      ? 'Sent for Sign.'
      : s === 'signed'
      ? 'Signed'
      : s === 'po_raised'
      ? 'PO Raised'
      : s === 'rejected'
      ? 'Rejected'
      : s === 'cancelled_by_partner'
      ? 'Cancelled by Partner'
      : s === 'postponed'
      ? 'Postponed'
      : s || '-';
  const statusColor = (s) =>
    s === 'signed'
      ? C.success
      : s === 'rejected'
      ? C.danger
      : s === 'sent_for_signature'
      ? C.purple
      : C.warning;

  const EditableNum = ({ rowKey, field, def }) => {
    const val = getOverride(rowKey, field, def);
    const isEditing = editCell?.rowKey === rowKey && editCell?.field === field;
    return isEditing ? (
      <input
        autoFocus
        data-editing="true"
        value={editVal}
        onChange={(e) => setEditVal(e.target.value)}
        onBlur={() => commitEdit(rowKey, field)}
        onKeyDown={(e) => e.key === 'Enter' && commitEdit(rowKey, field)}
        style={{
          width: 90,
          background: C.faint,
          border: `1px solid ${C.accent}`,
          color: C.text,
          borderRadius: 4,
          padding: '2px 6px',
          fontSize: 12,
          fontFamily: 'monospace',
          textAlign: 'right',
        }}
      />
    ) : (
      <span
        onClick={() => startEdit(rowKey, field, val)}
        style={{
          cursor: 'pointer',
          borderBottom: `1px dashed ${C.border}`,
          fontFamily: 'monospace',
          fontSize: 12,
        }}
        title="Click to edit"
      >
        {Number(val || 0) > 0
          ? `USD ${Number(val).toLocaleString('en-US')}`
          : '-'}
      </span>
    );
  };

  const EditableText = ({ rowKey, field, def }) => {
    const val = getOverride(rowKey, field, def) || '';
    const isEditing = editCell?.rowKey === rowKey && editCell?.field === field;
    return isEditing ? (
      <input
        autoFocus
        value={editVal}
        onChange={(e) => setEditVal(e.target.value)}
        onBlur={() => commitEdit(rowKey, field)}
        onKeyDown={(e) => e.key === 'Enter' && commitEdit(rowKey, field)}
        style={{
          width: 100,
          background: C.faint,
          border: `1px solid ${C.accent}`,
          color: C.text,
          borderRadius: 4,
          padding: '2px 6px',
          fontSize: 12,
          fontFamily: 'monospace',
        }}
      />
    ) : (
      <span
        onClick={() => startEdit(rowKey, field, val)}
        style={{
          cursor: 'pointer',
          borderBottom: `1px dashed ${C.border}`,
          fontSize: 12,
          fontFamily: 'monospace',
          color: val ? C.text : C.muted,
        }}
        title="Click to edit"
      >
        {val || '--'}
      </span>
    );
  };

  const EditableSelect = ({ rowKey, field, options, def }) => {
    const val = getOverride(rowKey, field, def) || '';
    return (
      <select
        value={val}
        onChange={(e) => setOverride(rowKey, field, e.target.value)}
        style={{
          background: 'transparent',
          border: `1px solid ${C.border}`,
          color: val ? C.text : C.muted,
          borderRadius: 6,
          padding: '2px 6px',
          fontSize: 11,
          cursor: 'pointer',
          maxWidth: 110,
        }}
      >
        {options.map((o) => (
          <option key={o} value={o}>
            {o || '--'}
          </option>
        ))}
      </select>
    );
  };

  const TH = ({ col, label, right, width }) => (
    <th
      onClick={() => {
        setSortCol(col);
        setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
      }}
      style={{
        padding: '9px 10px',
        textAlign: right ? 'right' : 'left',
        fontSize: 9,
        color: sortCol === col ? C.accent : C.muted,
        fontWeight: 700,
        letterSpacing: '0.07em',
        textTransform: 'uppercase',
        cursor: 'pointer',
        whiteSpace: 'nowrap',
        userSelect: 'none',
        background: sortCol === col ? C.faint : 'transparent',
        borderBottom: `2px solid ${C.border}`,
        position: 'sticky',
        top: 0,
        zIndex: 1,
        minWidth: width || 80,
      }}
    >
      {label}
      {sortCol === col ? (sortDir === 'asc' ? ' ^' : ' v') : ''}
    </th>
  );

  return (
    <div
      style={{
        color: C.text,
        display: 'flex',
        flexDirection: 'column',
        height: '100%',
        overflow: 'hidden',
        padding: 28,
        boxSizing: 'border-box',
      }}
    >
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'flex-start',
          marginBottom: 10,
          flexShrink: 0,
        }}
      >
        <div>
          <div
            style={{
              fontFamily: "'Syne',sans-serif",
              fontSize: 26,
              fontWeight: 800,
              marginBottom: 4,
            }}
          >
            MDF Overview
          </div>
          <div style={{ color: C.muted, fontSize: 13 }}>
            {filtered.length} activities . {uniquePartners} partners .{' '}
            {uniqueReqs} requests
          </div>
        </div>
        <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
          <span
            style={{
              fontFamily: "'Syne',sans-serif",
              fontWeight: 800,
              fontSize: 16,
              color: C.accent,
              letterSpacing: '0.05em',
            }}
          >
            OT
          </span>
          <div
            style={{
              width: 1,
              height: 20,
              background: C.border,
              flexShrink: 0,
            }}
          />
          {onImportPipeline && (
            <button
              onClick={onImportPipeline}
              style={{
                background: '#10b981',
                color: '#000',
                border: 'none',
                borderRadius: 10,
                padding: '7px 14px',
                fontWeight: 700,
                fontSize: 12,
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: 6,
              }}
            >
              📊 Import Pipeline
            </button>
          )}
          {onSavePipeline && (() => {
            // Check if there are any pipeline overrides to save
            const pipelineOverrides = Object.entries(overrides)
              .filter(([k]) => k.endsWith('_pipelineGenerated'));
            return pipelineOverrides.length > 0 ? (
              <button
                onClick={() => {
                  pipelineOverrides.forEach(([k, val]) => {
                    const rowKey = k.replace('_pipelineGenerated', '');
                    const row = rows.find(r => r.rowKey === rowKey);
                    // Use campaignId override if entered manually, else item campaignId, else rowKey
                    const campaignIdOverride = overrides[`${rowKey}_campaignId`] || row?.campaignId;
                    const key = campaignIdOverride || rowKey;
                    onSavePipeline(key, Number(val) || 0);
                  });
                  setOverrides(prev => {
                    const next = {...prev};
                    Object.keys(next).filter(k => k.endsWith('_pipelineGenerated')).forEach(k => delete next[k]);
                    return next;
                  });
                }}
                style={{
                  background: '#10b981',
                  color: '#000',
                  border: 'none',
                  borderRadius: 10,
                  padding: '7px 14px',
                  fontWeight: 800,
                  fontSize: 12,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: 6,
                  animation: 'pulse 1.5s infinite',
                }}
              >
                💾 Save Pipeline ({pipelineOverrides.length})
              </button>
            ) : null;
          })()}
          <button
            onClick={() => onExport && onExport(sorted)}
            style={{
              background: C.faint,
              color: C.text,
              border: `1px solid ${C.border}`,
              borderRadius: 10,
              padding: '7px 14px',
              fontWeight: 600,
              fontSize: 12,
              cursor: 'pointer',
            }}
          >
            Excel
          </button>
        </div>
      </div>
      <div
        style={{
          background: C.surface,
          border: `1px solid ${C.border}`,
          borderRadius: 12,
          padding: '10px 14px',
          marginBottom: 8,
          flexShrink: 0,
        }}
      >
        <div
          style={{
            display: 'flex',
            gap: 8,
            flexWrap: 'wrap',
            alignItems: 'center',
          }}
        >
          <button
            onClick={() => setShowFilterPicker((p) => !p)}
            style={{
              background: showFilterPicker ? C.accent : C.faint,
              color: showFilterPicker ? '#fff' : C.muted,
              border: `1px solid ${showFilterPicker ? C.accent : C.border}`,
              borderRadius: 8,
              padding: '5px 12px',
              fontSize: 12,
              fontWeight: 600,
              cursor: 'pointer',
              flexShrink: 0,
            }}
          >
            {showFilterPicker ? 'x' : '+'} Filters{' '}
            {activeFilters.length > 0 && `(${activeFilters.length})`}
          </button>
          <div
            style={{
              width: 1,
              height: 20,
              background: C.border,
              flexShrink: 0,
            }}
          />
          <input
            value={anSearch}
            onChange={(e) => setAnSearch(e.target.value)}
            placeholder="Search..."
            style={{
              background: C.faint,
              border: `1px solid ${C.border}`,
              color: C.text,
              borderRadius: 8,
              padding: '5px 10px',
              fontSize: 12,
              width: 160,
              fontFamily: 'inherit',
            }}
          />
          {activeFilters.includes('region') && (
            <MultiSelect
              value={
                anMacro === 'All'
                  ? []
                  : Array.isArray(anMacro)
                  ? anMacro
                  : [anMacro]
              }
              onChange={(v) => {
                setAnMacro(v.length === 0 ? 'All' : v);
                setAnSubregion('All');
              }}
              placeholder="All Regions"
              options={ALL_MACROS}
            />
          )}
          {activeFilters.includes('subregion') && (
            <MultiSelect
              value={
                anSubregion === 'All'
                  ? []
                  : Array.isArray(anSubregion)
                  ? anSubregion
                  : [anSubregion]
              }
              onChange={(v) => setAnSubregion(v.length === 0 ? 'All' : v)}
              placeholder="All Sub-Regions"
              options={
                !anMacro ||
                anMacro === 'All' ||
                (Array.isArray(anMacro) && anMacro.length === 0)
                  ? ALL_SUBREGIONS
                  : Array.isArray(anMacro)
                  ? anMacro.flatMap((m) => getSubregions(m))
                  : getSubregions(anMacro)
              }
            />
          )}
          {activeFilters.includes('tier') && (
            <MultiSelect
              value={
                anTier === 'All'
                  ? []
                  : Array.isArray(anTier)
                  ? anTier
                  : [anTier]
              }
              onChange={(v) => setAnTier(v.length === 0 ? 'All' : v)}
              placeholder="All Levels"
              options={['Platinum', 'Gold', 'Silver']}
            />
          )}
          {activeFilters.includes('type') && (
            <MultiSelect
              value={
                anType === 'All'
                  ? []
                  : Array.isArray(anType)
                  ? anType
                  : [anType]
              }
              onChange={(v) => setAnType(v.length === 0 ? 'All' : v)}
              placeholder="All Types"
              options={PARTNER_TYPES}
            />
          )}
          {activeFilters.includes('status') && (
            <MultiSelect
              value={
                anStatus === 'All'
                  ? []
                  : Array.isArray(anStatus)
                  ? anStatus
                  : [anStatus]
              }
              onChange={(v) => setAnStatus(v.length === 0 ? 'All' : v)}
              placeholder="All Statuses"
              options={[
                { value: 'request_submitted', label: 'Submitted' },
                { value: 'approved', label: 'Approved' },
                { value: 'sent_for_signature', label: 'Sent for Signature' },
                { value: 'signed', label: 'Signed' },
                { value: 'po_raised', label: 'PO Raised' },
                { value: 'rejected', label: 'Rejected' },
                {
                  value: 'cancelled_by_partner',
                  label: 'Cancelled by Partner',
                },
                { value: 'postponed', label: 'Postponed' },
              ]}
            />
          )}
          {activeFilters.includes('fyHalf') && (
            <MultiSelect
              value={
                anFyHalf === 'All'
                  ? []
                  : Array.isArray(anFyHalf)
                  ? anFyHalf
                  : [anFyHalf]
              }
              onChange={(v) => setAnFyHalf(v.length === 0 ? 'All' : v)}
              placeholder="All FY"
              options={[...new Set(rows.map((r) => r.fyHalf).filter(Boolean))]}
            />
          )}
          {activeFilters.includes('fyQ') && (
            <MultiSelect
              value={
                anFyQ === 'All' ? [] : Array.isArray(anFyQ) ? anFyQ : [anFyQ]
              }
              onChange={(v) => setAnFyQ(v.length === 0 ? 'All' : v)}
              placeholder="All Quarters"
              options={['Q1', 'Q2', 'Q3', 'Q4']}
            />
          )}
          {activeFilters.includes('bu') && (
            <MultiSelect
              value={anBU === 'All' ? [] : Array.isArray(anBU) ? anBU : [anBU]}
              onChange={(v) => setAnBU(v.length === 0 ? 'All' : v)}
              placeholder="All BU"
              options={PRODUCT_GROUPS}
            />
          )}
          {activeFilters.includes('tactic') && (
            <MultiSelect
              value={
                anTactic === 'All'
                  ? []
                  : Array.isArray(anTactic)
                  ? anTactic
                  : [anTactic]
              }
              onChange={(v) => setAnTactic(v.length === 0 ? 'All' : v)}
              placeholder="All Tactics"
              options={TACTICS}
            />
          )}
          {activeFilters.includes('exec') && (
            <MultiSelect
              value={
                anExec === 'All'
                  ? []
                  : Array.isArray(anExec)
                  ? anExec
                  : [anExec]
              }
              onChange={(v) => setAnExec(v.length === 0 ? 'All' : v)}
              placeholder="All Activity Status"
              options={['Executed', 'Canceled']}
            />
          )}
          {activeFilters.includes('claim') && (
            <MultiSelect
              value={
                anClaim === 'All'
                  ? []
                  : Array.isArray(anClaim)
                  ? anClaim
                  : [anClaim]
              }
              onChange={(v) => setAnClaim(v.length === 0 ? 'All' : v)}
              placeholder="All Claim Status"
              options={[
                { value: 'submitted', label: 'Submitted' },
                { value: 'marketing_review', label: 'Mktg Review' },
                { value: 'finance_review', label: 'Finance Review' },
                { value: 'approved', label: 'Approved' },
                { value: 'paid', label: 'Paid' },
                { value: 'on_hold', label: 'On Hold' },
                { value: 'cancelled', label: 'Cancelled' },
              ]}
            />
          )}
          {activeFilters.includes('partner') && (
            <MultiSelect
              value={
                anPartner === 'All'
                  ? []
                  : Array.isArray(anPartner)
                  ? anPartner
                  : [anPartner]
              }
              onChange={(v) => setAnPartner(v.length === 0 ? 'All' : v)}
              placeholder="All Partners"
              options={[...new Set(rows.map((r) => r.partnerName))].sort()}
            />
          )}
          {activeFilters.includes('pam') && (
            <MultiSelect
              value={anPAM === 'All' ? [] : Array.isArray(anPAM) ? anPAM : [anPAM]}
              onChange={(v) => setAnPAM(v.length === 0 ? 'All' : v)}
              placeholder="All PAMs"
              options={[...new Set(rows.map((r) => r.accountManager).filter(Boolean))].sort()}
            />
          )}
          {hasActiveFilter && (
            <button
              onClick={clearAll}
              style={{
                background: 'transparent',
                border: `1px solid ${C.border}`,
                color: C.muted,
                borderRadius: 8,
                padding: '4px 10px',
                fontSize: 12,
                cursor: 'pointer',
              }}
            >
              Clear all
            </button>
          )}
        </div>
        {showFilterPicker && (
          <div
            style={{
              marginTop: 12,
              paddingTop: 12,
              borderTop: `1px solid ${C.border}`,
              display: 'flex',
              flexWrap: 'wrap',
              gap: 8,
            }}
          >
            <span
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: C.muted,
                width: '100%',
                marginBottom: 2,
              }}
            >
              Choose which filters to show:
            </span>
            {ALL_FILTER_OPTS.map((f) => {
              const active = activeFilters.includes(f.key);
              return (
                <button
                  key={f.key}
                  onClick={() => toggleFilter(f.key)}
                  style={{
                    background: active ? C.accent : C.faint,
                    color: active ? '#fff' : C.muted,
                    border: `1px solid ${active ? C.accent : C.border}`,
                    borderRadius: 20,
                    padding: '4px 14px',
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: 'pointer',
                  }}
                >
                  {active ? 'x ' : ''}
                  {f.label}
                </button>
              );
            })}
          </div>
        )}
      </div>
      <div
        style={{
          display: 'flex',
          gap: 8,
          marginBottom: 8,
          flexWrap: 'wrap',
          flexShrink: 0,
        }}
      >
        {[
          { label: 'Activities', val: filtered.length, color: C.accent },
          { label: 'Partners', val: uniquePartners, color: C.purple },
          { label: 'Requests', val: uniqueReqs, color: C.warning },
          {
            label: 'MDF Req (USD)',
            val: `USD ${filtered
              .reduce(
                (s, r) =>
                  s + (toUSD ? toUSD(r.mdfRequest, r.localCurrency) : 0),
                0
              )
              .toLocaleString('en-US')}`,
            color: '#f59e0b',
          },
          { label: 'Claim Apprvd', val: fmtA(totalClaim), color: C.success },
          {
            label: 'Pipeline',
            val: fmtA(totalPipeline),
            color: C.cyan || C.accent,
          },
        ].map((k) => (
          <div
            key={k.label}
            style={{
              background: C.surface,
              border: `1px solid ${C.border}`,
              borderRadius: 8,
              padding: '6px 12px',
              display: 'flex',
              alignItems: 'center',
              gap: 8,
            }}
          >
            <span
              style={{
                fontSize: 9,
                color: C.muted,
                fontWeight: 700,
                letterSpacing: '0.05em',
                textTransform: 'uppercase',
                whiteSpace: 'nowrap',
              }}
            >
              {k.label}
            </span>
            <span
              style={{
                fontFamily: 'monospace',
                fontSize: 13,
                fontWeight: 800,
                color: k.color,
                whiteSpace: 'nowrap',
              }}
            >
              {k.val}
            </span>
          </div>
        ))}
      </div>
      <div
        style={{
          background: C.surface,
          border: `1px solid ${C.border}`,
          borderRadius: 14,
          overflow: 'hidden',
          flex: 1,
          display: 'flex',
          flexDirection: 'column',
          minHeight: 0,
        }}
      >
        <div
          style={{
            overflowX: 'auto',
            overflowY: 'auto',
            flex: 1,
            minHeight: 0,
          }}
        >
          <table
            style={{
              width: '100%',
              borderCollapse: 'collapse',
              fontSize: 12,
              minWidth: 2200,
            }}
          >
            <thead>
              <tr>
                <TH col="macro" label="Region" width={90} />
                <TH col="subregion" label="Sub-Region" width={90} />
                <TH col="country" label="Country" width={80} />
                <TH col="partnerName" label="Partner" width={140} />
                <TH col="contactName" label="Contact Name" width={120} />
                <TH col="contactEmail" label="Contact Email" width={160} />
                <TH col="accountManager" label="OT Partner Mgr" width={110} />
                <TH col="fyHalf" label="Fiscal Year" width={80} />
                <TH col="fyQuarter" label="FY Quarter" width={80} />
                <TH col="period" label="Month" width={90} />
                <TH col="productGroup" label="BU / Product Focus" width={130} />
                <TH col="tactic" label="Marketing Tactic" width={150} />
                <TH col="title" label="Activity Description" width={180} />
                <TH col="allocadiaId" label="Allocadia ID" width={100} />
                <TH col="campaignId" label="Campaign ID" width={100} />
                <TH col="partnerNotified" label="BP Sent" width={80} />
                <TH col="totalCost" label="Total Cost (LC)" right width={110} />
                <TH
                  col="mdfRequest"
                  label="MDF Request (LC)"
                  right
                  width={110}
                />
                <TH col="localCurrency" label="Currency" width={70} />
                <TH
                  col="mdfRequestUSD"
                  label="MDF Request (USD)"
                  right
                  width={120}
                />
                <TH col="reqStatus" label="Status" width={110} />
                <TH col="poNumber" label="PO Number" width={100} />
                <TH col="reqId" label="Request ID" width={90} />
                <TH col="execStatus" label="Activity Status" width={100} />
                <TH col="claimStatus" label="Claim Status" width={100} />
                <TH col="claimAmount" label="Claim Amount" right width={100} />
                <TH
                  col="pipelineGenerated"
                  label="Pipeline Generated"
                  right
                  width={120}
                />
              </tr>
            </thead>
            <tbody>
              {sorted.map((row, i) => (
                <tr
                  key={row.rowKey}
                  style={{
                    borderBottom: `1px solid ${C.border}15`,
                    background: i % 2 === 0 ? C.faint : C.surface,
                  }}
                >
                  <td
                    style={{
                      padding: '7px 10px',
                      fontWeight: 600,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.macro}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      color: C.muted,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.subregion}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      color: C.muted,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.country}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      fontWeight: 600,
                      whiteSpace: 'nowrap',
                      maxWidth: 150,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                    }}
                    title={row.partnerName}
                  >
                    {row.partnerName}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    {row.contactName}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      color: C.muted,
                      fontSize: 11,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.contactEmail}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    {row.accountManager}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    {row.fyHalf}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    {row.fyQuarter}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      color: C.muted,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.period}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    {row.productGroup}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      color: C.muted,
                      whiteSpace: 'nowrap',
                      maxWidth: 150,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                    }}
                    title={row.tactic}
                  >
                    {row.tactic}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      maxWidth: 200,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}
                    title={row.title}
                  >
                    {row.title}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    <EditableText
                      rowKey={row.rowKey}
                      field="allocadiaId"
                      def={row.allocadiaId}
                    />
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    <EditableText
                      rowKey={row.rowKey}
                      field="campaignId"
                      def={row.campaignId}
                    />
                  </td>
                  <td style={{ padding: '7px 10px', textAlign: 'center' }}>
                    {row.partnerNotified ? (
                      <span
                        style={{
                          background: C.success + '18',
                          color: C.success,
                          border: `1px solid ${C.success}30`,
                          borderRadius: 6,
                          padding: '2px 8px',
                          fontSize: 10,
                          fontWeight: 700,
                        }}
                      >
                        Yes
                      </span>
                    ) : (
                      <span style={{ color: C.muted, fontSize: 11 }}>-</span>
                    )}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 700,
                      color: C.accent,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.totalCost.toLocaleString('en-US')}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 700,
                      color: C.success,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.mdfRequest.toLocaleString('en-US')}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      color: C.muted,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.localCurrency}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 700,
                      color: '#f59e0b',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {toUSD
                      ? `USD ${toUSD(
                          row.mdfRequest,
                          row.localCurrency
                        ).toLocaleString('en-US')}`
                      : '-'}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    <span
                      style={{
                        background: statusColor(row.reqStatus) + '20',
                        color: statusColor(row.reqStatus),
                        borderRadius: 6,
                        padding: '2px 8px',
                        fontSize: 10,
                        fontWeight: 700,
                      }}
                    >
                      {statusLabel(row.reqStatus)}
                    </span>
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    <EditableText
                      rowKey={row.rowKey}
                      field="poNumber"
                      def={row.poNumber}
                    />
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      fontFamily: 'monospace',
                      color: C.accent,
                      fontSize: 11,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {row.reqId}
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    <EditableSelect
                      rowKey={row.rowKey}
                      field="execStatus"
                      options={EXEC_STATUSES}
                      def={row.execStatus}
                    />
                  </td>
                  <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                    {(() => {
                      const cs = row.claimStatus;
                      if (!cs)
                        return (
                          <span style={{ color: C.muted, fontSize: 11 }}>
                            -
                          </span>
                        );
                      const col =
                        cs === 'submitted'
                          ? C.warning
                          : cs === 'marketing_review'
                          ? C.cyan || '#06b6d4'
                          : cs === 'finance_review'
                          ? C.purple
                          : cs === 'approved'
                          ? C.success
                          : cs === 'paid'
                          ? C.teal || '#14b8a6'
                          : cs === 'on_hold'
                          ? C.danger
                          : cs === 'cancelled'
                          ? C.danger
                          : C.muted;
                      const lbl =
                        cs === 'submitted'
                          ? 'Submitted'
                          : cs === 'marketing_review'
                          ? 'Mktg Review'
                          : cs === 'finance_review'
                          ? 'Finance Rev'
                          : cs === 'approved'
                          ? 'Approved'
                          : cs === 'paid'
                          ? 'Paid'
                          : cs === 'on_hold'
                          ? 'On Hold'
                          : cs === 'cancelled'
                          ? 'Cancelled'
                          : cs;
                      return (
                        <span
                          style={{
                            background: col + '18',
                            color: col,
                            border: `1px solid ${col}30`,
                            borderRadius: 6,
                            padding: '2px 8px',
                            fontSize: 10,
                            fontWeight: 700,
                            whiteSpace: 'nowrap',
                          }}
                        >
                          {lbl}
                        </span>
                      );
                    })()}
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      textAlign: 'right',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    <EditableNum
                      rowKey={row.rowKey}
                      field="claimAmount"
                      def={row.claimAmount}
                    />
                  </td>
                  <td
                    style={{
                      padding: '7px 10px',
                      textAlign: 'right',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    <EditableNum
                      rowKey={row.rowKey}
                      field="pipelineGenerated"
                      def={row.pipelineGenerated}
                    />
                  </td>
                </tr>
              ))}
              {sorted.length === 0 && (
                <tr>
                  <td
                    colSpan={25}
                    style={{ padding: 40, textAlign: 'center', color: C.muted }}
                  >
                    No activities match the current filters
                  </td>
                </tr>
              )}
              {sorted.length > 0 && (
                <tr
                  style={{
                    borderTop: `2px solid ${C.border}`,
                    background: C.faint,
                    position: 'sticky',
                    bottom: 0,
                  }}
                >
                  <td
                    colSpan={17}
                    style={{
                      padding: '9px 10px',
                      fontWeight: 800,
                      fontSize: 12,
                    }}
                  >
                    TOTAL ({filtered.length} activities, {uniquePartners}{' '}
                    partners)
                  </td>
                  <td
                    style={{
                      padding: '9px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 800,
                      color: C.accent,
                    }}
                  >
                    Mixed LCY
                  </td>
                  <td
                    style={{
                      padding: '9px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 800,
                      color: C.success,
                    }}
                  >
                    Mixed LCY
                  </td>
                  <td
                    style={{
                      padding: '9px 10px',
                      color: C.muted,
                      whiteSpace: 'nowrap',
                    }}
                  />
                  <td
                    style={{
                      padding: '9px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 800,
                      color: '#f59e0b',
                    }}
                  >
                    USD{' '}
                    {toUSD
                      ? filtered
                          .reduce(
                            (s, r) => s + toUSD(r.mdfRequest, r.localCurrency),
                            0
                          )
                          .toLocaleString('en-US')
                      : '-'}
                  </td>
                  <td colSpan={4} />
                  <td
                    style={{
                      padding: '9px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 800,
                      color: C.warning,
                    }}
                  >{`USD ${totalClaim.toLocaleString('en-US')}`}</td>
                  <td
                    style={{
                      padding: '9px 10px',
                      textAlign: 'right',
                      fontFamily: 'monospace',
                      fontWeight: 800,
                      color: C.cyan || C.accent,
                    }}
                  >{`USD ${totalPipeline.toLocaleString('en-US')}`}</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// -- MULTI-SELECT FILTER --------------------------------------------------
const MultiSelect = ({
  value = [],
  onChange,
  options = [],
  placeholder = '',
}) => {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef(null);
  React.useEffect(() => {
    const h = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', h);
    return () => document.removeEventListener('mousedown', h);
  }, []);
  const toggle = (val) => {
    if (val === '__all__') {
      onChange([]);
      return;
    }
    const next = value.includes(val)
      ? value.filter((v) => v !== val)
      : [...value, val];
    onChange(next);
  };
  const label =
    value.length === 0
      ? placeholder
      : value.length === 1
      ? String(value[0])
      : `${value.length} selected`;
  const active = value.length > 0;
  return (
    <div ref={ref} style={{ position: 'relative', minWidth: 120 }}>
      <button
        onClick={() => setOpen((o) => !o)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setOpen((o) => !o); }
          if (e.key === 'Escape') setOpen(false);
        }}
        aria-haspopup="listbox"
        aria-expanded={open}
        style={{
          background: C.faint,
          border: `1px solid ${active ? C.accent : C.border}`,
          color: active ? C.accent : C.text,
          borderRadius: 10,
          padding: '7px 12px',
          fontSize: 13,
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          whiteSpace: 'nowrap',
          width: '100%',
          minWidth: 120,
        }}
      >
        <span
          style={{
            flex: 1,
            textAlign: 'left',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
          }}
        >
          {label}
        </span>
        {active && (
          <span
            onClick={(e) => {
              e.stopPropagation();
              onChange([]);
            }}
            style={{
              color: C.muted,
              fontSize: 11,
              lineHeight: 1,
              flexShrink: 0,
            }}
          >
            x
          </span>
        )}
        <span style={{ color: C.muted, fontSize: 9, flexShrink: 0 }}>
          {open ? '-' : '-'}
        </span>
      </button>
      {open && (
        <div
          role="listbox"
          aria-multiselectable="true"
          style={{
            position: 'absolute',
            top: 'calc(100% + 4px)',
            left: 0,
            minWidth: '100%',
            background: C.surface,
            border: `1px solid ${C.border}`,
            borderRadius: 10,
            zIndex: 500,
            boxShadow: '0 8px 24px rgba(0,0,0,0.4)',
            overflow: 'hidden',
            maxHeight: 280,
            overflowY: 'auto',
          }}
        >
          <div
            onClick={() => toggle('__all__')}
            style={{
              padding: '8px 14px',
              cursor: 'pointer',
              fontSize: 12,
              color: value.length === 0 ? C.accent : C.text,
              fontWeight: value.length === 0 ? 700 : 400,
              borderBottom: `1px solid ${C.border}20`,
            }}
          >
            All (clear)
          </div>
          {options.map((opt) => {
            const val = typeof opt === 'object' ? opt.value : opt;
            const lbl = typeof opt === 'object' ? opt.label : String(opt);
            const sel = value.includes(val);
            return (
              <div
                key={val}
                onClick={() => toggle(val)}
                style={{
                  padding: '8px 14px',
                  cursor: 'pointer',
                  fontSize: 12,
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  color: sel ? C.accent : C.text,
                  background: sel ? C.accentGlow : 'transparent',
                }}
              >
                <div
                  style={{
                    width: 14,
                    height: 14,
                    borderRadius: 3,
                    border: `2px solid ${sel ? C.accent : C.muted}`,
                    background: sel ? C.accent : 'transparent',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    flexShrink: 0,
                  }}
                >
                  {sel && (
                    <span
                      style={{
                        color: '#fff',
                        fontSize: 9,
                        fontWeight: 900,
                        lineHeight: 1,
                      }}
                    >
                      v
                    </span>
                  )}
                </div>
                {lbl}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

const ClaimReviewModal = ({ claim, onClose, onMove, onHold, onNotify }) => {
  const [holdNote, setHoldNote] = useState('');
  const [showHold, setShowHold] = useState(false);

  const STATUS_COLOR = {
    submitted: C.warning,
    marketing_review: C.cyan || '#06b6d4',
    finance_review: C.purple,
    approved: C.success,
    paid: C.teal || '#14b8a6',
    on_hold: C.danger,
  };
  const STATUS_LABEL = {
    submitted: 'Submitted',
    marketing_review: 'Marketing Review',
    finance_review: 'Finance Review',
    approved: 'Approved',
    paid: 'Paid',
    on_hold: 'On Hold',
  };
  const sCol = STATUS_COLOR[claim.status] || C.muted;

  const Row = ({ label, value, mono, accent }) => (
    <div
      style={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'flex-start',
        padding: '8px 0',
        borderBottom: `1px solid ${C.border}20`,
      }}
    >
      <span
        style={{
          fontSize: 11,
          color: C.muted,
          fontWeight: 600,
          minWidth: 150,
          flexShrink: 0,
        }}
      >
        {label}
      </span>
      <span
        style={{
          fontSize: 12,
          fontWeight: accent ? 700 : 500,
          color: accent ? C.accent : C.text,
          fontFamily: mono ? 'monospace' : 'inherit',
          textAlign: 'right',
        }}
      >
        {value || '-'}
      </span>
    </div>
  );

  const vatAmt = Math.round(
    ((claim.claimAmount || 0) * (claim.vatPct || 0)) / 100
  );
  const totalDocs =
    [
      claim.files?.partnerInvoice,
      claim.files?.thirdParty,
      claim.files?.inHouse,
      claim.files?.merchandise,
    ].filter(Boolean).length + (claim.files?.additional?.length || 0);

  const steps = [
    'submitted',
    'marketing_review',
    'finance_review',
    'approved',
    'paid',
  ];
  const stepLabels = [
    'Submitted',
    'Mktg Review',
    'Finance',
    'Approved',
    'Paid',
  ];
  const stepColors = [
    C.warning,
    C.cyan || '#06b6d4',
    C.purple,
    C.success,
    C.teal || '#14b8a6',
  ];
  const currentIdx = steps.indexOf(claim.status);

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.85)',
        zIndex: 400,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: 16,
      }}
    >
      <div
        style={{
          background: C.surface,
          border: `1px solid ${C.border}`,
          borderRadius: 20,
          width: 'min(720px,95vw)',
          maxHeight: '92vh',
          overflowY: 'auto',
          boxSizing: 'border-box',
        }}
      >
        {/* Sticky header */}
        <div
          style={{
            padding: '20px 28px 16px',
            borderBottom: `1px solid ${C.border}`,
            position: 'sticky',
            top: 0,
            background: C.surface,
            zIndex: 1,
          }}
        >
          <div
            style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'flex-start',
            }}
          >
            <div>
              <div
                style={{
                  fontFamily: "'Syne',sans-serif",
                  fontWeight: 800,
                  fontSize: 22,
                  marginBottom: 6,
                }}
              >
                Claim Review
              </div>
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 10,
                  flexWrap: 'wrap',
                }}
              >
                <span
                  style={{
                    fontFamily: 'monospace',
                    fontSize: 13,
                    color: C.accent,
                    fontWeight: 700,
                  }}
                >
                  {claim.id}
                </span>
                <span
                  style={{
                    background: sCol + '18',
                    color: sCol,
                    border: `1px solid ${sCol}30`,
                    borderRadius: 20,
                    padding: '3px 12px',
                    fontSize: 11,
                    fontWeight: 700,
                  }}
                >
                  {STATUS_LABEL[claim.status] || claim.status}
                </span>
                <span style={{ fontSize: 12, color: C.muted }}>
                  {claim.partner}
                </span>
              </div>
            </div>
            <button
              onClick={onClose}
              style={{
                background: 'transparent',
                border: 'none',
                color: C.muted,
                fontSize: 20,
                cursor: 'pointer',
                padding: 4,
              }}
            >
              x
            </button>
          </div>
          {/* Progress */}
          {claim.status !== 'on_hold' && (
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                marginTop: 14,
                gap: 0,
              }}
            >
              {steps.map((s, i) => {
                const done = currentIdx > i,
                  active = currentIdx === i;
                const col = stepColors[i];
                return (
                  <div
                    key={s}
                    style={{ display: 'flex', alignItems: 'center', flex: 1 }}
                  >
                    <div
                      style={{
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        flex: 1,
                      }}
                    >
                      <div
                        style={{
                          width: 22,
                          height: 22,
                          borderRadius: '50%',
                          background: done || active ? col : C.faint,
                          border: `2px solid ${
                            done || active ? col : C.border
                          }`,
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          boxShadow: active ? `0 0 0 3px ${col}30` : 'none',
                          flexShrink: 0,
                        }}
                      >
                        {done ? (
                          <span
                            style={{
                              color: '#fff',
                              fontSize: 10,
                              fontWeight: 900,
                            }}
                          >
                            v
                          </span>
                        ) : (
                          <span
                            style={{
                              color: active ? '#fff' : C.muted,
                              fontSize: 9,
                              fontWeight: 700,
                            }}
                          >
                            {i + 1}
                          </span>
                        )}
                      </div>
                      <span
                        style={{
                          fontSize: 9,
                          color: active ? col : C.muted,
                          fontWeight: active ? 700 : 400,
                          marginTop: 4,
                          textAlign: 'center',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {stepLabels[i]}
                      </span>
                    </div>
                    {i < steps.length - 1 && (
                      <div
                        style={{
                          height: 2,
                          flex: 1,
                          background: done ? col : C.border,
                          marginBottom: 16,
                        }}
                      />
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div
          style={{
            padding: '20px 28px',
            display: 'grid',
            gridTemplateColumns: '1fr 1fr',
            gap: 24,
          }}
        >
          {/* LEFT: Activity + Financials */}
          <div>
            <div
              style={{
                fontSize: 10,
                fontWeight: 800,
                color: C.accent,
                letterSpacing: '0.1em',
                marginBottom: 10,
              }}
            >
              ACTIVITY
            </div>
            <div
              style={{
                background: C.faint,
                borderRadius: 10,
                padding: '12px 14px',
                marginBottom: 18,
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 6 }}>
                {claim.activity}
              </div>
              <Row label="Partner" value={claim.partner} />
              <Row
                label="FY / Quarter"
                value={`${claim.fyHalf || ''} ${claim.fyQuarter || ''} ${
                  claim.month || ''
                }`}
              />
              <Row label="Submitted" value={claim.submitted} />
              <Row label="Request ID" value={claim.reqId} />
            </div>

            <div
              style={{
                fontSize: 10,
                fontWeight: 800,
                color: C.accent,
                letterSpacing: '0.1em',
                marginBottom: 10,
              }}
            >
              FINANCIALS
            </div>
            <div
              style={{
                background: C.faint,
                borderRadius: 10,
                padding: '12px 14px',
                marginBottom: 18,
              }}
            >
              <Row
                label="Net Claim Amount"
                value={`${claim.currency} ${Number(
                  claim.claimAmount || 0
                ).toLocaleString()}`}
                mono
                accent
              />
              {claim.vatPct > 0 && (
                <Row
                  label={`VAT (${claim.vatPct}%)`}
                  value={`${claim.currency} ${vatAmt.toLocaleString()}`}
                  mono
                />
              )}
              <Row
                label="Total (incl. VAT)"
                value={`${claim.currency} ${Number(
                  claim.totalValue || claim.claimAmount || 0
                ).toLocaleString()}`}
                mono
                accent
              />
              <Row label="Currency" value={claim.currency} />
            </div>

            {claim.notes && (
              <>
                <div
                  style={{
                    fontSize: 10,
                    fontWeight: 800,
                    color: C.accent,
                    letterSpacing: '0.1em',
                    marginBottom: 10,
                  }}
                >
                  PARTNER NOTES
                </div>
                <div
                  style={{
                    background: C.faint,
                    borderRadius: 10,
                    padding: '12px 14px',
                    fontSize: 12,
                    color: C.muted,
                    fontStyle: 'italic',
                    lineHeight: 1.6,
                    marginBottom: 18,
                  }}
                >
                  "{claim.notes}"
                </div>
              </>
            )}

            {claim.reviewNotesMarketing && (
              <div
                style={{
                  background: C.warning + '10',
                  border: `1px solid ${C.warning}30`,
                  borderRadius: 10,
                  padding: '10px 14px',
                  fontSize: 12,
                  color: C.warning,
                  marginBottom: 10,
                }}
              >
                <div style={{ fontWeight: 700, marginBottom: 4 }}>
                  Marketing Review Note:
                </div>
                {claim.reviewNotesMarketing}
              </div>
            )}

            {/* Status audit trail */}
            {(claim.statusHistory || []).length > 0 && (
              <div>
                <div style={{ fontSize: 10, fontWeight: 800, color: C.accent, letterSpacing: '0.1em', marginBottom: 10 }}>
                  STATUS HISTORY
                </div>
                <div style={{ position: 'relative' }}>
                  <div style={{ position: 'absolute', left: 7, top: 0, bottom: 0, width: 1, background: C.border }} />
                  {(claim.statusHistory || []).map((h, i) => {
                    const col = h.status === 'submitted' ? C.warning
                      : h.status === 'marketing_review' ? (C.cyan || '#06b6d4')
                      : h.status === 'finance_review' ? C.purple
                      : h.status === 'approved' ? C.success
                      : h.status === 'paid' ? (C.teal || '#14b8a6')
                      : h.status === 'on_hold' ? C.danger
                      : C.muted;
                    const lbl = h.status === 'submitted' ? 'Submitted'
                      : h.status === 'marketing_review' ? 'Mktg Review'
                      : h.status === 'finance_review' ? 'Finance Review'
                      : h.status === 'approved' ? 'Approved'
                      : h.status === 'paid' ? 'Paid'
                      : h.status === 'on_hold' ? 'On Hold'
                      : h.status;
                    return (
                      <div key={i} style={{ display: 'flex', gap: 10, marginBottom: 10, position: 'relative' }}>
                        <div style={{ width: 15, height: 15, borderRadius: '50%', background: col, border: `2px solid ${col}`, flexShrink: 0, marginTop: 1, zIndex: 1 }} />
                        <div style={{ flex: 1 }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 2 }}>
                            <span style={{ fontSize: 11, fontWeight: 700, color: col }}>{lbl}</span>
                            <span style={{ fontSize: 10, color: C.muted }}>· {h.by}</span>
                          </div>
                          <div style={{ fontSize: 10, color: C.muted, fontFamily: 'monospace' }}>{h.at}</div>
                          {h.note && h.note !== lbl && (
                            <div style={{ fontSize: 11, color: C.muted, marginTop: 2, fontStyle: 'italic' }}>"{h.note}"</div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          {/* RIGHT: Documents + Actions */}
          <div>
            <div
              style={{
                fontSize: 10,
                fontWeight: 800,
                color: C.accent,
                letterSpacing: '0.1em',
                marginBottom: 10,
              }}
            >
              DOCUMENTS ({totalDocs})
            </div>
            <div
              style={{
                background: C.faint,
                borderRadius: 10,
                padding: '12px 14px',
                marginBottom: 18,
              }}
            >
              {[
                [
                  'Partner Invoice (required)',
                  claim.files?.partnerInvoice,
                  true,
                ],
                ['Third Party Supplier', claim.files?.thirdParty, false],
                ['Partner In-House Activity', claim.files?.inHouse, false],
                ['Merchandise Receipt', claim.files?.merchandise, false],
              ].map(([lbl, name, req]) => (
                <div
                  key={lbl}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 10,
                    padding: '8px 0',
                    borderBottom: `1px solid ${C.border}20`,
                  }}
                >
                  {name ? (
                    <>
                      <div
                        style={{
                          width: 28,
                          height: 28,
                          borderRadius: 6,
                          background: C.success + '20',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          fontSize: 9,
                          fontWeight: 700,
                          color: C.success,
                          flexShrink: 0,
                        }}
                      >
                        PDF
                      </div>
                      <div>
                        <div
                          style={{
                            fontSize: 11,
                            fontWeight: 600,
                            color: C.success,
                          }}
                        >
                          {name}
                        </div>
                        <div style={{ fontSize: 9, color: C.muted }}>{lbl}</div>
                      </div>
                    </>
                  ) : (
                    <>
                      <div
                        style={{
                          width: 28,
                          height: 28,
                          borderRadius: 6,
                          background: C.faint,
                          border: `1px solid ${C.border}`,
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          fontSize: 9,
                          color: C.muted,
                          flexShrink: 0,
                        }}
                      >
                        -
                      </div>
                      <div style={{ fontSize: 11, color: C.muted }}>
                        {lbl}
                        {req ? (
                          <span style={{ color: C.danger }}> *</span>
                        ) : (
                          ''
                        )}{' '}
                        - not uploaded
                      </div>
                    </>
                  )}
                </div>
              ))}
              {(claim.files?.additional || []).map((f, i) => (
                <div
                  key={i}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 10,
                    padding: '8px 0',
                    borderBottom: `1px solid ${C.border}20`,
                  }}
                >
                  <div
                    style={{
                      width: 28,
                      height: 28,
                      borderRadius: 6,
                      background: C.accent + '15',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      fontSize: 9,
                      fontWeight: 700,
                      color: C.accent,
                      flexShrink: 0,
                    }}
                  >
                    DOC
                  </div>
                  <div style={{ fontSize: 11, color: C.muted }}>{f}</div>
                </div>
              ))}
              {totalDocs === 0 && (
                <div style={{ fontSize: 12, color: C.muted }}>
                  No documents uploaded
                </div>
              )}
            </div>

            {/* Workflow actions */}
            <div
              style={{
                fontSize: 10,
                fontWeight: 800,
                color: C.accent,
                letterSpacing: '0.1em',
                marginBottom: 10,
              }}
            >
              ACTIONS
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {claim.status === 'submitted' && (
                <button
                  onClick={() => {
                    onMove(claim.id, 'marketing_review');
                    onClose();
                  }}
                  style={{
                    background: C.accent,
                    color: '#fff',
                    border: 'none',
                    borderRadius: 10,
                    padding: '10px 16px',
                    fontWeight: 700,
                    fontSize: 13,
                    cursor: 'pointer',
                  }}
                >
                  Start Marketing Review
                </button>
              )}
              {claim.status === 'marketing_review' && (
                <button
                  onClick={() => {
                    onMove(claim.id, 'finance_review');
                    onClose();
                  }}
                  style={{
                    background: C.success,
                    color: '#fff',
                    border: 'none',
                    borderRadius: 10,
                    padding: '10px 16px',
                    fontWeight: 700,
                    fontSize: 13,
                    cursor: 'pointer',
                  }}
                >
                  Approve - Send to Finance
                </button>
              )}
              {claim.status === 'finance_review' && (
                <button
                  onClick={() => {
                    onMove(claim.id, 'approved');
                    onNotify(claim);
                    onClose();
                  }}
                  style={{
                    background: C.success,
                    color: '#fff',
                    border: 'none',
                    borderRadius: 10,
                    padding: '10px 16px',
                    fontWeight: 700,
                    fontSize: 13,
                    cursor: 'pointer',
                  }}
                >
                  Approve & Notify Partner
                </button>
              )}
              {claim.status === 'approved' && (
                <>
                  <button
                    onClick={() => {
                      onNotify(claim);
                    }}
                    style={{
                      background: 'transparent',
                      border: `1px solid ${C.success}`,
                      color: C.success,
                      borderRadius: 10,
                      padding: '10px 16px',
                      fontWeight: 700,
                      fontSize: 13,
                      cursor: 'pointer',
                    }}
                  >
                    Resend Notification Email
                  </button>
                  <button
                    onClick={() => {
                      onMove(claim.id, 'paid');
                      onClose();
                    }}
                    style={{
                      background: C.teal || '#14b8a6',
                      color: '#fff',
                      border: 'none',
                      borderRadius: 10,
                      padding: '10px 16px',
                      fontWeight: 700,
                      fontSize: 13,
                      cursor: 'pointer',
                    }}
                  >
                    Mark as Paid
                  </button>
                </>
              )}
              {!['on_hold', 'approved', 'paid'].includes(claim.status) &&
                (showHold ? (
                  <div>
                    <textarea
                      value={holdNote}
                      onChange={(e) => setHoldNote(e.target.value)}
                      rows={3}
                      placeholder="Explain what is missing or incorrect (this will be shared with the partner)..."
                      style={{
                        width: '100%',
                        background: C.faint,
                        border: `1px solid ${C.danger}`,
                        color: C.text,
                        borderRadius: 8,
                        padding: '8px 12px',
                        fontSize: 12,
                        outline: 'none',
                        resize: 'vertical',
                        boxSizing: 'border-box',
                        marginBottom: 8,
                      }}
                    />
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button
                        onClick={() => {
                          setShowHold(false);
                          setHoldNote('');
                        }}
                        style={{
                          flex: 1,
                          background: 'transparent',
                          color: C.muted,
                          border: `1px solid ${C.border}`,
                          borderRadius: 8,
                          padding: '8px',
                          fontSize: 12,
                          cursor: 'pointer',
                        }}
                      >
                        Cancel
                      </button>
                      <button
                        onClick={() => {
                          if (!holdNote.trim()) return;
                          onHold(claim.id, holdNote);
                          onClose();
                        }}
                        style={{
                          flex: 2,
                          background: C.danger,
                          color: '#fff',
                          border: 'none',
                          borderRadius: 8,
                          padding: '8px',
                          fontWeight: 700,
                          fontSize: 12,
                          cursor: 'pointer',
                        }}
                      >
                        Confirm - Put on Hold
                      </button>
                    </div>
                  </div>
                ) : (
                  <button
                    onClick={() => setShowHold(true)}
                    style={{
                      background: 'transparent',
                      color: C.danger,
                      border: `1px solid ${C.danger}`,
                      borderRadius: 10,
                      padding: '10px 16px',
                      fontWeight: 700,
                      fontSize: 13,
                      cursor: 'pointer',
                    }}
                  >
                    Put on Hold
                  </button>
                ))}
              <button
                onClick={onClose}
                style={{
                  background: 'transparent',
                  color: C.muted,
                  border: `1px solid ${C.border}`,
                  borderRadius: 10,
                  padding: '9px 16px',
                  fontWeight: 600,
                  fontSize: 12,
                  cursor: 'pointer',
                }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

const ClaimsTab = ({ claims, setClaims, partners, requests = [], addHistory, toast, toUSD, fmtA }) => {
  const STATUS_COLOR = {
    submitted: C.warning,
    marketing_review: C.cyan || '#06b6d4',
    finance_review: C.purple,
    approved: C.success,
    paid: C.teal || '#14b8a6',
    on_hold: C.danger,
  };
  const STATUS_LABEL = {
    submitted: 'Submitted',
    marketing_review: 'Mktg Review',
    finance_review: 'Finance Review',
    approved: 'Approved',
    paid: 'Paid',
    on_hold: 'On Hold',
  };

  const [reviewClaim, setReviewClaim] = useState(null);
  const [filterPartner, setFilterPartner] = useState([]);
  const [filterStatus, setFilterStatus] = useState([]);
  const [filterQuarter, setFilterQuarter] = useState([]);
  const [searchClaim, setSearchClaim] = useState('');
  const [sortDir, setSortDir] = useState('desc');
  const [sortCol, setSortCol] = useState('submitted');

  const moveClaim = (id, newStatus, note = '') => {
    const now = new Date().toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    setClaims(prev =>
      prev.map(c => {
        if (c.id !== id) return c;
        const histEntry = { status: newStatus, by: currentUser, at: now, note: note || STATUS_LABEL[newStatus] || newStatus };
        addHistory(`Claim ${id} → ${STATUS_LABEL[newStatus] || newStatus}`, id, 'approve');
        return { ...c, status: newStatus, statusHistory: [...(c.statusHistory || []), histEntry] };
      })
    );
    toast(`Claim updated: ${STATUS_LABEL[newStatus]}`);
  };

  const holdClaim = (id, note) => {
    const now = new Date().toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    setClaims(prev => {
      const next = prev.map(c => {
        if (c.id !== id) return c;
        const histEntry = { status: 'on_hold', by: currentUser, at: now, note };
        return { ...c, status: 'on_hold', reviewNotesMarketing: note, statusHistory: [...(c.statusHistory || []), histEntry] };
      });
      // BUG-001 FIX: persist hold status to Supabase immediately
      const updated = next.find(c => c.id === id);
      if (updated) dbSaveClaim(updated).catch(e => console.warn('[MDF] holdClaim save error:', e.message));
      return next;
    });
    addHistory(`Claim ${id} put on hold`, id, 'reject');
    toast('Claim put on hold. Partner notified.');
  };

  const notifyApproved = (c) => {
    const partnerRecord = partners.find(p => p.name === c.partner) || {};
    const toEmail = partnerRecord.contactEmail || '';
    const toName = partnerRecord.contactName || c.partner;
    const subject = encodeURIComponent(`MDF Claim Approved - ${c.id} | OT`);
    const body = encodeURIComponent(
      `Dear ${toName},\n\nWe are pleased to confirm that your MDF claim has been approved.\n\nClaim ID: ${c.id}\nActivity: ${c.activity}\nApproved Amount: ${c.currency} ${Number(c.claimAmount || 0).toLocaleString()}\n\nPlease now send your invoice to the OT Finance team for payment processing.\n\nBest regards,\nChannel Marketing Team | OT`
    );
    window.open(`mailto:${toEmail}?subject=${subject}&body=${body}`, '_blank');
    toast('Email opened - notify partner to send invoice.');
  };

  const getRegion = (partnerName) => {
    const p = partners.find(x => x.name === partnerName);
    return p ? p.region || '-' : '-';
  };

  const getSubregion = (partnerName) => {
    const p = partners.find(x => x.name === partnerName);
    return p ? p.subregion || '-' : '-';
  };

  const toUSDsafe = (amount, currency) => {
    if (toUSD) return toUSD(amount, currency);
    const fallback = { EUR: 1.09, GBP: 1.27, CHF: 1.12, SEK: 0.092, AED: 0.272, SGD: 0.74, AUD: 0.65 };
    if (!currency || currency === 'USD') return Number(amount || 0);
    return Math.round(Number(amount || 0) * (fallback[currency] || 1));
  };

  const partnerOptions = [...new Set(claims.map(c => c.partner))].sort();

  const filtered = claims
    .filter(c => {
      if (filterPartner.length > 0 && !filterPartner.includes(c.partner)) return false;
      if (filterStatus.length > 0 && !filterStatus.includes(c.status)) return false;
      if (filterQuarter.length > 0 && !filterQuarter.includes(c.fyQuarter)) return false;
      if (searchClaim) {
        const q = searchClaim.toLowerCase();
        if (![c.id, c.partner, c.activity, c.reqId].join(' ').toLowerCase().includes(q)) return false;
      }
      return true;
    })
    .sort((a, b) => {
      let av = a[sortCol] ?? '';
      let bv = b[sortCol] ?? '';
      if (sortCol === 'claimAmount' || sortCol === 'totalValue') {
        av = Number(av) || 0; bv = Number(bv) || 0;
        return sortDir === 'desc' ? bv - av : av - bv;
      }
      return sortDir === 'desc'
        ? String(bv).localeCompare(String(av))
        : String(av).localeCompare(String(bv));
    });

  const hasActiveFilter = filterPartner.length > 0 || filterStatus.length > 0 || filterQuarter.length > 0 || searchClaim.length > 0;
  const kpiSource = hasActiveFilter ? filtered : claims;
  const kpiData = [
    {
      label: 'New / Submitted',
      color: C.warning,
      sub: 'awaiting review',
      claims: kpiSource.filter(c => c.status === 'submitted'),
    },
    {
      label: 'In Review',
      color: C.cyan || '#06b6d4',
      sub: 'marketing + finance',
      claims: kpiSource.filter(c => ['marketing_review', 'finance_review'].includes(c.status)),
    },
    {
      label: 'Approved',
      color: C.success,
      sub: 'partner to invoice',
      claims: kpiSource.filter(c => c.status === 'approved'),
    },
    {
      label: 'On Hold',
      color: C.danger,
      sub: 'more info needed',
      claims: kpiSource.filter(c => c.status === 'on_hold'),
    },
  ];

  const TH = ({ col, label, right, width }) => (
    <th
      onClick={() => { setSortCol(col); setSortDir(d => col === sortCol ? (d === 'asc' ? 'desc' : 'asc') : 'desc'); }}
      style={{
        padding: '9px 12px',
        textAlign: right ? 'right' : 'left',
        fontSize: 10,
        color: sortCol === col ? C.accent : C.muted,
        fontWeight: 700,
        letterSpacing: '0.07em',
        textTransform: 'uppercase',
        cursor: 'pointer',
        whiteSpace: 'nowrap',
        userSelect: 'none',
        background: sortCol === col ? C.faint : 'transparent',
        borderBottom: `2px solid ${C.border}`,
        position: 'sticky',
        top: 0,
        zIndex: 1,
        minWidth: width || 80,
      }}
    >
      {label}{sortCol === col ? (sortDir === 'asc' ? ' ^' : ' v') : ''}
    </th>
  );

  const totalNetUSD = filtered.reduce((s, c) => s + toUSDsafe(c.claimAmount || 0, c.currency), 0);
  const totalTotalUSD = filtered.reduce((s, c) => s + toUSDsafe(c.totalValue || c.claimAmount || 0, c.currency), 0);

  return (
    <>
      <div style={{ animation: 'slideIn 0.3s ease', padding: 28, overflowY: 'auto', height: '100%', boxSizing: 'border-box' }}>

        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 20 }}>
          <div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontSize: 26, fontWeight: 800, marginBottom: 4 }}>
              MDF Claims
            </div>
            <div style={{ color: C.muted, fontSize: 13 }}>
              {claims.length} total &nbsp;.&nbsp;
              {claims.filter(c => c.status === 'submitted').length} new &nbsp;.&nbsp;
              {claims.filter(c => ['marketing_review', 'finance_review'].includes(c.status)).length} in review &nbsp;.&nbsp;
              {claims.filter(c => c.status === 'on_hold').length} on hold
            </div>
          </div>
        </div>

        {/* KPI cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 12, marginBottom: 20 }}>
          {kpiData.map(k => {
            const usdTotal = k.claims.reduce((s, c) => s + toUSDsafe(c.claimAmount || 0, c.currency), 0);
            return (
              <div key={k.label} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: '14px 16px', position: 'relative', overflow: 'hidden' }}>
                <div style={{ position: 'absolute', top: 0, right: 0, width: 40, height: 40, background: `radial-gradient(circle at top right,${k.color}18,transparent)`, borderRadius: '0 12px 0 40px' }} />
                <div style={{ fontSize: 10, color: C.muted, fontWeight: 700, letterSpacing: '0.08em', marginBottom: 6, textTransform: 'uppercase' }}>{k.label}</div>
                <div style={{ fontSize: 28, fontWeight: 800, color: k.color, fontFamily: 'monospace', marginBottom: 2 }}>{k.claims.length}</div>
                <div style={{ fontSize: 11, color: k.color, opacity: 0.85, fontFamily: 'monospace', fontWeight: 600, marginBottom: 2 }}>
                  USD {usdTotal.toLocaleString('en-US')}
                </div>
                <div style={{ fontSize: 10, color: C.muted }}>{k.sub}</div>
              </div>
            );
          })}
        </div>

        {/* Filter bar */}
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center', padding: '10px 14px', background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, marginBottom: 14 }}>
          <input
            value={searchClaim}
            onChange={e => setSearchClaim(e.target.value)}
            placeholder="Search ID, partner, activity..."
            style={{ background: C.faint, border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: '6px 12px', fontSize: 12, outline: 'none', minWidth: 200, fontFamily: 'inherit' }}
          />
          <MultiSelect value={filterPartner} onChange={setFilterPartner} placeholder="All Partners" options={partnerOptions} />
          <MultiSelect
            value={filterStatus}
            onChange={setFilterStatus}
            placeholder="All Statuses"
            options={[
              { value: 'submitted', label: 'Submitted' },
              { value: 'marketing_review', label: 'Mktg Review' },
              { value: 'finance_review', label: 'Finance Review' },
              { value: 'approved', label: 'Approved' },
              { value: 'paid', label: 'Paid' },
              { value: 'on_hold', label: 'On Hold' },
            ]}
          />
          <MultiSelect value={filterQuarter} onChange={setFilterQuarter} placeholder="All Quarters" options={['Q1', 'Q2', 'Q3', 'Q4']} />
          <button
            onClick={() => { setSortCol('submitted'); setSortDir(d => d === 'desc' ? 'asc' : 'desc'); }}
            style={{ background: C.faint, border: `1px solid ${C.border}`, color: C.muted, borderRadius: 8, padding: '6px 12px', fontSize: 12, cursor: 'pointer', whiteSpace: 'nowrap' }}
          >
            Date {sortDir === 'desc' ? '↓ Newest' : '↑ Oldest'}
          </button>
          {(filterPartner.length > 0 || filterStatus.length > 0 || filterQuarter.length > 0 || searchClaim) && (
            <button
              onClick={() => { setFilterPartner([]); setFilterStatus([]); setFilterQuarter([]); setSearchClaim(''); }}
              style={{ background: 'transparent', border: `1px solid ${C.border}`, color: C.muted, borderRadius: 8, padding: '6px 12px', fontSize: 12, cursor: 'pointer' }}
            >
              Clear
            </button>
          )}
          <span style={{ marginLeft: 'auto', fontSize: 11, color: C.muted }}>{filtered.length} of {claims.length} claims</span>
        </div>

        {/* Table */}
        {claims.length === 0 ? (
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: 60, textAlign: 'center', color: C.muted }}>
            <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 8 }}>No Claims Yet</div>
            <div style={{ fontSize: 13 }}>Claims appear here when partners submit approved activities for reimbursement.</div>
          </div>
        ) : (
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, overflow: 'hidden' }}>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, minWidth: 1400 }}>
                <thead>
                  <tr style={{ background: C.faint }}>
                    <TH col="id" label="ID" width={90} />
                    <TH col="region" label="Region" width={90} />
                    <TH col="partner" label="Partner" width={140} />
                    <TH col="cmmRef" label="CMM Ref" width={90} />
                    <TH col="activity" label="Activity" width={180} />
                    <TH col="fyQuarter" label="Quarter" width={75} />
                    <TH col="submitted" label="Created" width={95} />
                    <TH col="claimAmount" label="Net Claim (LCY)" right width={120} />
                    <TH col="currency" label="CCY" width={55} />
                    <TH col="claimUSD" label="Net Claim (USD)" right width={120} />
                    <TH col="vatPct" label="VAT %" right width={60} />
                    <TH col="totalValue" label="Total (LCY)" right width={110} />
                    <TH col="totalUSD" label="Total (USD)" right width={110} />
                    <TH col="status" label="Status" width={110} />
                    <TH col="docs" label="Docs" width={55} />
                    <th style={{ padding: '9px 12px', borderBottom: `2px solid ${C.border}`, background: 'transparent', width: 70 }} />
                  </tr>
                </thead>
                <tbody>
                  {filtered.length > 0 && (() => {
                    // Split into 3 sections
                    const claimsNeedsAttention = filtered.filter(c => c.status === 'submitted');
                    const claimsInProgress = filtered.filter(c => ['marketing_review', 'finance_review'].includes(c.status));
                    const claimsCompleted = filtered.filter(c => ['approved', 'paid', 'on_hold'].includes(c.status));

                    const renderClaimRow = (c, i) => {
                      const sCol = STATUS_COLOR[c.status] || C.muted;
                      const sLbl = STATUS_LABEL[c.status] || c.status;
                      const region = getRegion(c.partner);
                      const netUSD = toUSDsafe(c.claimAmount || 0, c.currency);
                      const totalUSD = toUSDsafe(c.totalValue || c.claimAmount || 0, c.currency);
                      const totalDocs = [c.files?.partnerInvoice, c.files?.thirdParty, c.files?.inHouse, c.files?.merchandise].filter(Boolean).length + (c.files?.additional?.length || 0);
                      const isNew = c.status === 'submitted';
                      const isInReview = ['marketing_review','finance_review'].includes(c.status);
                      const isOnHold = c.status === 'on_hold';
                      const rowBg = isNew ? `${C.warning}08` : isInReview ? `${C.cyan||'#06b6d4'}08` : isOnHold ? `${C.danger}08` : i % 2 === 0 ? C.faint : C.surface;
                      const rowBorderLeft = isNew ? `3px solid ${C.warning}` : isInReview ? `3px solid ${C.cyan||'#06b6d4'}` : isOnHold ? `3px solid ${C.danger}` : '3px solid transparent';
                      const rowHoverBg = isNew ? `${C.warning}18` : isInReview ? `${C.cyan||'#06b6d4'}18` : isOnHold ? `${C.danger}18` : C.accentGlow;
                      return (
                        <tr
                          key={c.id}
                          onClick={() => setReviewClaim(c)}
                          style={{ borderBottom: `1px solid ${C.border}20`, background: rowBg, borderLeft: rowBorderLeft, cursor: 'pointer', transition: 'background 0.1s' }}
                          onMouseEnter={e => e.currentTarget.style.background = rowHoverBg}
                          onMouseLeave={e => e.currentTarget.style.background = rowBg}
                        >
                          <td style={{ padding: '9px 12px', fontFamily: 'monospace', fontSize: 11, color: C.accent, fontWeight: 700, whiteSpace: 'nowrap' }}>{c.id}</td>
                          <td style={{ padding: '9px 12px', color: C.muted, fontSize: 11, whiteSpace: 'nowrap' }}>{region}</td>
                          <td style={{ padding: '9px 12px', fontWeight: 600, whiteSpace: 'nowrap', maxWidth: 150, overflow: 'hidden', textOverflow: 'ellipsis' }}>{c.partner}</td>
                          <td style={{ padding: '9px 12px', whiteSpace: 'nowrap' }}>
                            {(() => {
                              const req = requests.find(r => r.id === c.reqId);
                              const cmm = req?.assignedTo || '-';
                              const col = cmm === 'Decio A.' ? C.accent : cmm === 'Kaila' ? C.success : cmm === 'Umair' ? C.purple : C.muted;
                              return cmm !== '-'
                                ? <span style={{ background: col+'18', color: col, border: `1px solid ${col}30`, borderRadius: 20, padding: '2px 9px', fontSize: 10, fontWeight: 700 }}>{cmm}</span>
                                : <span style={{ color: C.muted, fontSize: 11 }}>-</span>;
                            })()}
                          </td>
                          <td style={{ padding: '9px 12px', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={c.activity}>{c.activity}</td>
                          <td style={{ padding: '9px 12px', color: C.muted, whiteSpace: 'nowrap' }}>{c.fyHalf ? `${c.fyHalf} ` : ''}{c.fyQuarter}</td>
                          <td style={{ padding: '9px 12px', color: C.muted, whiteSpace: 'nowrap', fontFamily: 'monospace', fontSize: 11 }}>{c.submitted}</td>
                          <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 700, color: C.accent, whiteSpace: 'nowrap' }}>
                            {Number(c.claimAmount || 0).toLocaleString('en-US')}
                          </td>
                          <td style={{ padding: '9px 12px', color: C.muted, fontSize: 11, whiteSpace: 'nowrap' }}>{c.currency}</td>
                          <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 700, color: '#f59e0b', whiteSpace: 'nowrap' }}>
                            USD {netUSD.toLocaleString('en-US')}
                          </td>
                          <td style={{ padding: '9px 12px', textAlign: 'right', color: C.muted, whiteSpace: 'nowrap' }}>{c.vatPct || 0}%</td>
                          <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 600, whiteSpace: 'nowrap' }}>
                            {Number(c.totalValue || c.claimAmount || 0).toLocaleString('en-US')}
                          </td>
                          <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 700, color: C.success, whiteSpace: 'nowrap' }}>
                            USD {totalUSD.toLocaleString('en-US')}
                          </td>
                          <td style={{ padding: '9px 12px', whiteSpace: 'nowrap' }}>
                            <span style={{ background: sCol + '18', color: sCol, border: `1px solid ${sCol}30`, borderRadius: 20, padding: '3px 10px', fontSize: 10, fontWeight: 700 }}>
                              {sLbl}
                            </span>
                          </td>
                          <td style={{ padding: '9px 12px', textAlign: 'center', color: totalDocs > 0 ? C.success : C.muted, fontWeight: 700 }}>
                            {totalDocs > 0 ? totalDocs : '-'}
                          </td>
                          <td style={{ padding: '9px 12px', textAlign: 'center' }}>
                            <span style={{ fontSize: 11, color: C.accent, fontWeight: 600, whiteSpace: 'nowrap' }}>Review &rarr;</span>
                          </td>
                        </tr>
                      );
                    };

                    return (
                      <>
                        {/* === NEEDS ATTENTION === */}
                        {claimsNeedsAttention.length > 0 && (
                          <tr style={{ background: C.warning + '10' }}>
                            <td colSpan={16} style={{ padding: '8px 14px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                <div style={{ width: 8, height: 8, borderRadius: '50%', background: C.warning, flexShrink: 0 }} />
                                <span style={{ fontSize: 11, fontWeight: 800, color: C.warning, letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                                  Needs Attention — {claimsNeedsAttention.length} new claim{claimsNeedsAttention.length !== 1 ? 's' : ''}
                                </span>
                                <span style={{ fontSize: 10, color: C.muted }}>awaiting review</span>
                              </div>
                            </td>
                          </tr>
                        )}
                        {claimsNeedsAttention.map((c, i) => renderClaimRow(c, i))}

                        {/* === IN PROGRESS === */}
                        {claimsInProgress.length > 0 && (
                          <tr style={{ background: `${C.cyan||'#06b6d4'}10`, borderTop: claimsNeedsAttention.length > 0 ? `2px solid ${C.border}` : 'none' }}>
                            <td colSpan={16} style={{ padding: '8px 14px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                <div style={{ width: 8, height: 8, borderRadius: '50%', background: C.cyan || '#06b6d4', flexShrink: 0 }} />
                                <span style={{ fontSize: 11, fontWeight: 800, color: C.cyan || '#06b6d4', letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                                  In Progress — {claimsInProgress.length} claim{claimsInProgress.length !== 1 ? 's' : ''}
                                </span>
                                <span style={{ fontSize: 10, color: C.muted }}>
                                  {claimsInProgress.filter(c => c.status === 'marketing_review').length > 0 && `${claimsInProgress.filter(c => c.status === 'marketing_review').length} mktg review`}
                                  {claimsInProgress.filter(c => c.status === 'marketing_review').length > 0 && claimsInProgress.filter(c => c.status === 'finance_review').length > 0 && ' · '}
                                  {claimsInProgress.filter(c => c.status === 'finance_review').length > 0 && `${claimsInProgress.filter(c => c.status === 'finance_review').length} finance review`}
                                </span>
                              </div>
                            </td>
                          </tr>
                        )}
                        {claimsInProgress.map((c, i) => renderClaimRow(c, i + 100))}

                        {/* === COMPLETED === */}
                        {claimsCompleted.length > 0 && (
                          <tr style={{ background: C.success + '10', borderTop: (claimsNeedsAttention.length > 0 || claimsInProgress.length > 0) ? `2px solid ${C.border}` : 'none' }}>
                            <td colSpan={16} style={{ padding: '8px 14px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                <div style={{ width: 8, height: 8, borderRadius: '50%', background: C.success, flexShrink: 0 }} />
                                <span style={{ fontSize: 11, fontWeight: 800, color: C.success, letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                                  Completed — {claimsCompleted.length} claim{claimsCompleted.length !== 1 ? 's' : ''}
                                </span>
                                <span style={{ fontSize: 10, color: C.muted }}>
                                  {claimsCompleted.filter(c => c.status === 'approved').length > 0 && `${claimsCompleted.filter(c => c.status === 'approved').length} approved`}
                                  {claimsCompleted.filter(c => c.status === 'approved').length > 0 && claimsCompleted.filter(c => c.status === 'paid').length > 0 && ' · '}
                                  {claimsCompleted.filter(c => c.status === 'paid').length > 0 && `${claimsCompleted.filter(c => c.status === 'paid').length} paid`}
                                  {claimsCompleted.filter(c => c.status === 'on_hold').length > 0 && ` · ${claimsCompleted.filter(c => c.status === 'on_hold').length} on hold`}
                                </span>
                              </div>
                            </td>
                          </tr>
                        )}
                        {claimsCompleted.map((c, i) => renderClaimRow(c, i + 200))}
                      </>
                    );
                  })()}
                  {filtered.length === 0 && (
                    <tr>
                      <td colSpan={16} style={{ padding: 40, textAlign: 'center', color: C.muted }}>
                        No claims match the current filters
                      </td>
                    </tr>
                  )}
                </tbody>
                <tfoot>
                  <tr style={{ background: C.faint, borderTop: `2px solid ${C.border}`, position: 'sticky', bottom: 0 }}>
                    <td colSpan={6} style={{ padding: '9px 12px', fontWeight: 800, fontSize: 12 }}>
                      TOTAL ({filtered.length} claims)
                    </td>
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 800, color: C.accent }}>
                      —
                    </td>
                    <td />
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 800, color: '#f59e0b' }}>
                      USD {totalNetUSD.toLocaleString('en-US')}
                    </td>
                    <td />
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 800 }}>—</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 800, color: C.success }}>
                      USD {totalTotalUSD.toLocaleString('en-US')}
                    </td>
                    <td colSpan={3} />
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* Claim Review Modal */}
      {reviewClaim && (
        <ClaimReviewModal
          claim={reviewClaim}
          onClose={() => setReviewClaim(null)}
          onMove={(id, status) => { moveClaim(id, status); setReviewClaim(prev => prev ? { ...prev, status } : null); }}
          onHold={(id, note) => { holdClaim(id, note); setReviewClaim(prev => prev ? { ...prev, status: 'on_hold', reviewNotesMarketing: note } : null); }}
          onNotify={notifyApproved}
        />
      )}
    </>
  );
};



// Partner form field components (defined outside modal to prevent remount on re-render)
const PartnerLabel = ({ children, req }) => (
  <div style={{ fontSize: 10, fontWeight: 700, color: C.muted, letterSpacing: '0.08em', marginBottom: 5, textTransform: 'uppercase' }}>
    {children}{req && <span style={{ color: C.danger }}> *</span>}
  </div>
);

const PartnerField = ({ label, req, children }) => (
  <div>
    <PartnerLabel req={req}>{label}</PartnerLabel>
    {children}
  </div>
);

// ============================================================
// EDIT / ADD PARTNER MODAL
// ============================================================
const EditPartnerModal = React.memo(({ partner, onSave, onClose }) => {
  const isNew = !partner?.id;
  const [form, setForm] = React.useState({
    name: partner?.name || '',
    region: partner?.region || '',
    subregion: partner?.subregion || '',
    country: partner?.country || '',
    type: partner?.type || 'Reseller',
    tier: partner?.tier || 'Gold',
    allocated: partner?.allocated || 0,
    pending: partner?.pending || 0,
    status: partner?.status || 'Active',
    contactName: partner?.contactName || '',
    contactEmail: partner?.contactEmail || '',
    accountManager: partner?.accountManager || '',
    note: partner?.note || '',
  });
  const [errors, setErrors] = React.useState({});

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const validate = () => {
    const e = {};
    if (!form.name.trim()) e.name = 'Required';
    if (!form.region.trim()) e.region = 'Required';
    if (!form.allocated || isNaN(Number(form.allocated))) e.allocated = 'Required';
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const save = () => {
    if (!validate()) return;
    onSave({
      ...partner,
      id: partner?.id || 'p' + Date.now(),
      name: form.name.trim(),
      region: form.region.trim(),
      subregion: form.subregion.trim(),
      country: form.country.trim(),
      type: form.type,
      tier: form.tier,
      allocated: Number(form.allocated) || 0,
      spent: partner?.spent || 0,  // auto-calculated from approved claims
      pending: Number(form.pending) || 0,
      status: form.status,
      contactName: form.contactName.trim(),
      contactEmail: form.contactEmail.trim(),
      accountManager: form.accountManager.trim(),
      note: form.note.trim(),
    });
  };

  // Stable base style - avoids React DOM patching on every render
  const _inpBase = { background: C.faint, color: C.text, borderRadius: 8, padding: '8px 12px', fontSize: 13, width: '100%', outline: 'none', boxSizing: 'border-box' };
  const inp = (err) => ({ ..._inpBase, border: `1px solid ${err ? C.danger : C.border}` });
  const sel = (err) => ({ ...inp(err), cursor: 'pointer' });



  const grid2 = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 };

  return (
    <div data-modal="true" style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.7)', zIndex: 500, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
      <div style={{ background: C.surface, borderRadius: 20, width: '100%', maxWidth: 620, maxHeight: '90vh', overflowY: 'auto', border: `1px solid ${C.border}` }}>
        {/* Header */}
        <div style={{ padding: '20px 24px', borderBottom: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', position: 'sticky', top: 0, background: C.surface, zIndex: 1 }}>
          <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 18 }}>
            {isNew ? 'Add New Partner' : `Edit — ${partner.name}`}
          </div>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: C.muted, fontSize: 22, cursor: 'pointer', lineHeight: 1 }}>×</button>
        </div>

        {/* Body */}
        <div style={{ padding: 24, display: 'flex', flexDirection: 'column', gap: 16 }}>

          {/* Partner Name */}
          <PartnerField label="Partner Name" req>
            <input value={form.name} onChange={e => set('name', e.target.value)}
              placeholder="e.g. TechVision Ltd" style={inp(errors.name)} />
            {errors.name && <div style={{ fontSize: 10, color: C.danger, marginTop: 3 }}>{errors.name}</div>}
          </PartnerField>

          {/* Geography */}
          <div style={grid2}>
            <PartnerField label="Region" req>
              <select value={form.region} onChange={e => set('region', e.target.value)} style={sel(errors.region)}>
                <option value="">-- select --</option>
                {['Europe', 'US', 'International'].map(r => <option key={r}>{r}</option>)}
              </select>
              {errors.region && <div style={{ fontSize: 10, color: C.danger, marginTop: 3 }}>{errors.region}</div>}
            </PartnerField>
            <PartnerField label="Sub-Region">
              <input value={form.subregion} onChange={e => set('subregion', e.target.value)}
                placeholder="e.g. UK&I, DACH, Nordics" style={inp()} />
            </PartnerField>
          </div>

          <PartnerField label="Country">
            <input value={form.country} onChange={e => set('country', e.target.value)}
              placeholder="e.g. Italy" style={inp()} />
          </PartnerField>

          {/* Classification */}
          <div style={grid2}>
            <PartnerField label="Type">
              <select value={form.type} onChange={e => set('type', e.target.value)} style={sel()}>
                {['Reseller', 'Distributor', 'GSI', 'ISVP', 'MSP', 'Other'].map(t => <option key={t}>{t}</option>)}
              </select>
            </PartnerField>
            <PartnerField label="Level">
              <select value={form.tier} onChange={e => set('tier', e.target.value)} style={sel()}>
                {['Platinum', 'Gold', 'Silver', 'Bronze'].map(t => <option key={t}>{t}</option>)}
              </select>
            </PartnerField>
          </div>

          {/* Budget */}
          <div style={grid2}>
            <PartnerField label="Budget Allocated (USD)" req>
              <input type="number" value={form.allocated} onChange={e => set('allocated', e.target.value)}
                placeholder="0" style={inp(errors.allocated)} />
              {errors.allocated && <div style={{ fontSize: 10, color: C.danger, marginTop: 3 }}>{errors.allocated}</div>}
            </PartnerField>

          </div>

          {/* Contact */}
          <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 14 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: C.accent, letterSpacing: '0.08em', marginBottom: 12, textTransform: 'uppercase' }}>Contact & Team</div>
            <div style={grid2}>
              <PartnerField label="Contact Name">
                <input value={form.contactName} onChange={e => set('contactName', e.target.value)}
                  placeholder="e.g. John Smith" style={inp()} />
              </PartnerField>
              <PartnerField label="Contact Email">
                <input value={form.contactEmail} onChange={e => set('contactEmail', e.target.value)}
                  placeholder="john@partner.com" style={inp()} />
              </PartnerField>
            </div>
            <div style={{ marginTop: 14 }}>
              <PartnerField label="Partner Account Manager (PAM)">
                <input value={form.accountManager} onChange={e => set('accountManager', e.target.value)}
                  placeholder="Full name" style={inp()} />
              </PartnerField>
            </div>
          </div>

          {/* Status & Note */}
          <div style={grid2}>
            <PartnerField label="Status">
              <select value={form.status} onChange={e => set('status', e.target.value)} style={sel()}>
                {['Active', 'Inactive', 'Prospect'].map(s => <option key={s}>{s}</option>)}
              </select>
            </PartnerField>
            <PartnerField label="Note">
              <input value={form.note} onChange={e => set('note', e.target.value)}
                placeholder="Optional note" style={inp()} />
            </PartnerField>
          </div>
        </div>

        {/* Footer */}
        <div style={{ padding: '16px 24px', borderTop: `1px solid ${C.border}`, display: 'flex', justifyContent: 'flex-end', gap: 10, position: 'sticky', bottom: 0, background: C.surface }}>
          <button onClick={onClose} style={{ background: 'transparent', border: `1px solid ${C.border}`, color: C.muted, borderRadius: 10, padding: '9px 20px', fontSize: 13, cursor: 'pointer' }}>
            Cancel
          </button>
          <button onClick={save} style={{ background: C.accent, color: '#fff', border: 'none', borderRadius: 10, padding: '9px 24px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
            {isNew ? 'Add Partner' : 'Save Changes'}
          </button>
        </div>
      </div>
    </div>
  );
});


// ============================================================
// PIPELINE IMPORT MODAL
// ============================================================
const PipelineImportModal = ({ onImport, onClose, C }) => {
  const [step, setStep] = React.useState('upload'); // upload | preview | done
  const [rows, setRows] = React.useState([]);
  const [error, setError] = React.useState('');
  const fileRef = React.useRef(null);

  const parseCSV = (text) => {
    const lines = text.trim().split('\n').filter(l => l.trim());
    if (lines.length < 2) { setError('File must have a header row and at least one data row'); return; }
    const headers = lines[0].split(',').map(h => h.trim().toLowerCase().replace(/[^a-z0-9_]/g,'_'));
    const campaignCol = headers.findIndex(h => h.includes('campaign'));
    const pipelineCol = headers.findIndex(h => h.includes('pipeline') || h.includes('amount') || h.includes('revenue') || h.includes('value'));
    if (campaignCol === -1) { setError('Could not find a Campaign ID column. Expected header containing "campaign"'); return; }
    if (pipelineCol === -1) { setError('Could not find a Pipeline column. Expected header containing "pipeline", "amount", "revenue", or "value"'); return; }
    const parsed = lines.slice(1).map(line => {
      const cols = line.split(',').map(c => c.trim().replace(/^"|"$/g,''));
      const pipeline = parseFloat(cols[pipelineCol]) || 0;
      return { campaignId: cols[campaignCol] || '', pipeline: Math.max(0, pipeline) }; // BREAK-010: clamp negatives to 0
    }).filter(r => r.campaignId);
    // Warn if any negative values were clamped
    const negCount = lines.slice(1).filter(line => {
      const cols = line.split(',').map(c => c.trim().replace(/^"|"$/g,''));
      return parseFloat(cols[pipelineCol]) < 0;
    }).length;
    if (negCount > 0) setError(`${negCount} negative value(s) were set to 0. Check your source data.`);
    if (parsed.length === 0) { setError('No valid rows found'); return; }
    setRows(parsed);
    setStep('preview');
    setError('');
  };

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => parseCSV(ev.target.result);
    reader.readAsText(file);
  };

  const handleImport = () => {
    const data = {};
    rows.forEach(r => { data[r.campaignId] = (data[r.campaignId] || 0) + r.pipeline; });
    onImport(data, rows.length);
    setStep('done');
  };

  const inp = { background: C.faint, border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: '8px 12px', fontSize: 13, width: '100%', outline: 'none', boxSizing: 'border-box' };

  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.7)', zIndex:600, display:'flex', alignItems:'center', justifyContent:'center', padding:20 }}>
      <div style={{ background:C.surface, borderRadius:20, width:'100%', maxWidth:560, border:`1px solid ${C.border}` }}>
        {/* Header */}
        <div style={{ padding:'18px 24px', borderBottom:`1px solid ${C.border}`, display:'flex', justifyContent:'space-between', alignItems:'center' }}>
          <div style={{ fontFamily:"'Syne',sans-serif", fontWeight:800, fontSize:18 }}>Import Pipeline Data</div>
          <button onClick={onClose} style={{ background:'transparent', border:'none', color:C.muted, fontSize:22, cursor:'pointer' }}>×</button>
        </div>

        <div style={{ padding:24 }}>
          {step === 'upload' && (
            <div>
              <p style={{ fontSize:13, color:C.muted, marginBottom:16 }}>
                Upload a CSV file with two columns: <strong style={{color:C.text}}>Campaign ID</strong> and <strong style={{color:C.text}}>Pipeline Amount</strong>. 
                The Campaign ID must match the Campaign IDs entered in approved activities.
              </p>
              <div style={{ background:C.faint, border:`1px dashed ${C.border}`, borderRadius:12, padding:32, textAlign:'center', cursor:'pointer', marginBottom:16 }}
                onClick={() => fileRef.current?.click()}>
                <div style={{ fontSize:28, marginBottom:8 }}>📊</div>
                <div style={{ fontSize:14, fontWeight:700, color:C.text, marginBottom:4 }}>Click to upload CSV</div>
                <div style={{ fontSize:12, color:C.muted }}>campaign_id, pipeline_amount</div>
              </div>
              <input ref={fileRef} type="file" accept=".csv,.txt" onChange={handleFile} style={{ display:'none' }} />
              {error && <div style={{ fontSize:12, color:C.danger, padding:'8px 12px', background:C.danger+'15', borderRadius:8 }}>{error}</div>}
              <div style={{ marginTop:16, background:C.faint, borderRadius:10, padding:'12px 16px', fontSize:11, color:C.muted }}>
                <div style={{ fontWeight:700, marginBottom:4, color:C.text }}>Expected CSV format:</div>
                <code style={{ fontFamily:'monospace' }}>campaign_id,pipeline_amount<br/>ALO-2026-001,120000<br/>ALO-2026-002,85000</code>
              </div>
            </div>
          )}

          {step === 'preview' && (
            <div>
              <div style={{ fontSize:13, color:C.muted, marginBottom:12 }}>
                Found <strong style={{color:C.text}}>{rows.length} records</strong>. This will update pipeline values for matching Campaign IDs only — no other data will be changed.
              </div>
              <div style={{ maxHeight:280, overflowY:'auto', border:`1px solid ${C.border}`, borderRadius:10 }}>
                <table style={{ width:'100%', borderCollapse:'collapse' }}>
                  <thead>
                    <tr style={{ background:C.faint }}>
                      <th style={{ padding:'8px 12px', textAlign:'left', fontSize:11, fontWeight:700, color:C.muted }}>Campaign ID</th>
                      <th style={{ padding:'8px 12px', textAlign:'right', fontSize:11, fontWeight:700, color:C.muted }}>Pipeline Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r, i) => (
                      <tr key={i} style={{ borderTop:`1px solid ${C.border}20`, background: i%2===0 ? C.faint : C.surface }}>
                        <td style={{ padding:'7px 12px', fontFamily:'monospace', fontSize:12, color:C.accent }}>{r.campaignId}</td>
                        <td style={{ padding:'7px 12px', textAlign:'right', fontFamily:'monospace', fontSize:12, fontWeight:700, color:'#10b981' }}>
                          {Number(r.pipeline).toLocaleString('en-US')}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ marginTop:12, padding:'10px 14px', background:'#10b98115', border:'1px solid #10b98130', borderRadius:8, fontSize:12, color:'#10b981', fontWeight:600 }}>
                ✓ Safe import — only pipeline values are updated. Partner allocations, requests, and claims are not affected.
              </div>
            </div>
          )}

          {step === 'done' && (
            <div style={{ textAlign:'center', padding:'20px 0' }}>
              <div style={{ fontSize:40, marginBottom:12 }}>✅</div>
              <div style={{ fontSize:16, fontWeight:700, color:C.text, marginBottom:8 }}>Pipeline data imported!</div>
              <div style={{ fontSize:13, color:C.muted }}>{rows.length} Campaign IDs updated</div>
            </div>
          )}
        </div>

        <div style={{ padding:'16px 24px', borderTop:`1px solid ${C.border}`, display:'flex', justifyContent:'flex-end', gap:10 }}>
          <button onClick={onClose} style={{ background:'transparent', border:`1px solid ${C.border}`, color:C.muted, borderRadius:10, padding:'9px 20px', fontSize:13, cursor:'pointer' }}>
            {step === 'done' ? 'Close' : 'Cancel'}
          </button>
          {step === 'preview' && (
            <button onClick={handleImport} style={{ background:'#10b981', color:'#000', border:'none', borderRadius:10, padding:'9px 24px', fontSize:13, fontWeight:700, cursor:'pointer' }}>
              Import {rows.length} Records
            </button>
          )}
        </div>
      </div>
    </div>
  );
};


// ── Dashboard partner performance row (defined outside App to prevent remount) ──
const DashPartnerRow = ({ p }) => {
  const ROI_TARGET = 30;
  const pct = Math.min(Math.round((p.ratio / ROI_TARGET) * 100), 150);
  const barColor = p.ratio >= ROI_TARGET ? '#10b981' : p.ratio >= ROI_TARGET * 0.7 ? '#f59e0b' : '#ef4444';
  const tierCol = p.tier === 'Platinum' ? '#f59e0b' : p.tier === 'Gold' ? C.accent : C.muted;
  return (
    <div style={{ display:'flex', alignItems:'center', gap:12, padding:'10px 0', borderBottom:`1px solid ${C.border}20` }}>
      <div style={{ flex:1, minWidth:0 }}>
        <div style={{ display:'flex', alignItems:'center', gap:6, marginBottom:4 }}>
          <span style={{ fontSize:12, fontWeight:700, color:C.text }}>{p.name}</span>
          {p.tier && <span style={{ fontSize:9, fontWeight:700, color:tierCol, background:tierCol+'20', borderRadius:4, padding:'1px 5px' }}>{p.tier}</span>}
        </div>
        <div style={{ background:C.faint, borderRadius:4, height:6, overflow:'hidden' }}>
          <div style={{ width:`${Math.min(pct,100)}%`, height:'100%', background:barColor, borderRadius:4 }} />
        </div>
      </div>
      <div style={{ textAlign:'right', flexShrink:0 }}>
        <div style={{ fontSize:13, fontWeight:800, color:barColor, fontFamily:'monospace' }}>
          {p.hasPipeline ? `1:${Math.round(p.ratio)}` : '—'}
        </div>
        <div style={{ fontSize:10, color:C.muted }}>
          {p.hasPipeline ? `USD ${Math.round(p.totalPipeline).toLocaleString()}` : 'No pipeline'}
        </div>
      </div>
    </div>
  );
};


export default function App() {
  // ── SSO: replace this with the identity provider's display name ──
  // e.g. from Okta: currentUser = ssoSession.user.displayName
  // e.g. from Azure AD: currentUser = msalAccount.name
  const currentUser = 'Decio A.';
  const [loggedIn, setLoggedIn] = useState(false);
  const [pipelineData, setPipelineData] = useState(() => {
    try { const d = loadPersistedData(); return d?.pipelineData || {}; } catch { return {}; }
  }); // { campaignId: pipelineAmount } - persisted to localStorage
  const [showPipelineImport, setShowPipelineImport] = useState(false);
  const [dbLoaded, setDbLoaded] = useState(false);
  const [dbError, setDbError] = useState(null);
  const [portalUser, setPortalUser] = useState(null); // { partner, email }
  const [loginPwd, setLoginPwd] = useState('');
  const [loginErr, setLoginErr] = useState(false);
  const [tab, setTab] = useState('dashboard');
  const [isDark, setIsDark] = useState(false);
  Object.assign(C, isDark ? DARK_THEME : LIGHT_THEME);
  const [currency, setCurrency] = useState('USD');
  const [rate, setRate] = useState(1.09);
  const [allRates, setAllRates] = useState({});
  const [rateLoading, setRateLoading] = useState(false);
  const [rateTs, setRateTs] = useState(null);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [partnerFY, setPartnerFY] = useState('All');
  const [selReq, setSelReq] = useState(null);
  const [selectedItems, setSelectedItems] = useState({});
  const [showNewReq, setShowNewReq] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [planModal, setPlanModal] = useState(null);
  const [portalPartner, setPortalPartner] = useState(null);
  const [isPartnerView] = useState(false); // Partner portal mode - disabled in internal build
  const [notif, setNotif] = useState(null);
  const [search, setSearch] = useState('');
  const [filterMacro, setFilterMacro] = useState([]);
  const [filterRegion, setFilterRegion] = useState([]);
  const [filterTier, setFilterTier] = useState([]);
  const [filterType, setFilterType] = useState([]);
  const [filterPAM, setFilterPAM] = useState([]);
  const [filterStatus, setFilterStatus] = useState([]);
  const [filterPartnerReqStatus, setFilterPartnerReqStatus] = useState([]);
  const [filterReqStatus, setFilterReqStatus] = useState([]);
  const [filterReqPartnerType, setFilterReqPartnerType] = useState([]);
  const [filterReqFY, setFilterReqFY] = useState([]);
  const [filterReqQ, setFilterReqQ] = useState([]);
  const [filterReqMonth, setFilterReqMonth] = useState([]);
  const [filterReqPartner, setFilterReqPartner] = useState([]);
  const [filterCMM, setFilterCMM] = useState([]);
  const [combinedBPSel, setCombinedBPSel] = useState({}); // {itemId: {item, request}}
  const [showCombinedBP, setShowCombinedBP] = useState(false);
  const [sortReqDir, setSortReqDir] = useState('desc'); // newest first
  const [editPartner, setEditPartner] = useState(null);
  const [dashMacro, setDashMacro] = useState([]);
  const [dashSubregion, setDashSubregion] = useState([]);
  const [dashTier, setDashTier] = useState([]);
  const [dashPartner, setDashPartner] = useState([]);
  const [dashPartnerType, setDashPartnerType] = useState([]);
  const [dashPAM, setDashPAM] = useState([]);

  const [partners, setPartners] = useState([
    {
      id: 'p1',
      name: 'TechVision Ltd',
      region: 'Europe',
      subregion: 'UK&I',
      country: 'UK',
      type: 'Reseller',
      tier: 'Platinum',
      allocated: 50000,
      spent: 32000,
      pending: 8000,
      status: 'Active',
      note: '',
      contactName: 'John Smith',
      contactEmail: 'j.smith@techvision.com',
      accountManager: 'Sarah Jones',
    },
    {
      id: 'p2',
      name: 'CloudSys GmbH',
      region: 'Europe',
      subregion: 'DACH',
      country: 'Germany',
      type: 'Reseller',
      tier: 'Gold',
      allocated: 30000,
      spent: 18000,
      pending: 5000,
      status: 'Active',
      note: '',
      contactName: 'Hans Mueller',
      contactEmail: 'h.mueller@cloudsys.de',
      accountManager: 'Sarah Jones',
    },
    {
      id: 'p3',
      name: 'Nordic IT AB',
      region: 'Europe',
      subregion: 'Nordics',
      country: 'Sweden',
      type: 'Distributor',
      tier: 'Silver',
      allocated: 15000,
      spent: 9000,
      pending: 0,
      status: 'Active',
      note: '',
      contactName: 'Anna Svensson',
      contactEmail: 'a.svensson@nordicit.se',
      accountManager: 'Mike Brown',
    },
    {
      id: 'p4',
      name: 'France Digitale',
      region: 'Europe',
      subregion: 'France',
      country: 'France',
      type: 'Reseller',
      tier: 'Platinum',
      allocated: 45000,
      spent: 28000,
      pending: 10000,
      status: 'Active',
      note: '',
      contactName: 'Marie Dupont',
      contactEmail: 'm.dupont@france-dig.fr',
      accountManager: 'Sarah Jones',
    },
    {
      id: 'p5',
      name: 'Italia Cloud SpA',
      region: 'Europe',
      subregion: 'Italy',
      country: 'Italy',
      type: 'Reseller',
      tier: 'Gold',
      allocated: 20000,
      spent: 12000,
      pending: 4000,
      status: 'Active',
      note: '',
      contactName: 'Marco Rossi',
      contactEmail: 'm.rossi@italiacloud.it',
      accountManager: 'Mike Brown',
    },
    {
      id: 'p6',
      name: 'American Tech Corp',
      region: 'US',
      subregion: 'US',
      country: 'USA',
      type: 'Reseller',
      tier: 'Platinum',
      allocated: 80000,
      spent: 55000,
      pending: 15000,
      status: 'Active',
      note: '',
      contactName: 'Bob Johnson',
      contactEmail: 'b.johnson@amtech.com',
      accountManager: 'Tom Davis',
    },
    {
      id: 'p7',
      name: 'Canada Cloud Inc',
      region: 'US',
      subregion: 'Canada',
      country: 'Canada',
      type: 'Reseller',
      tier: 'Gold',
      allocated: 35000,
      spent: 20000,
      pending: 5000,
      status: 'Active',
      note: '',
      contactName: 'Lisa Chen',
      contactEmail: 'l.chen@canadacloud.ca',
      accountManager: 'Tom Davis',
    },
    {
      id: 'p8',
      name: 'Gulf Tech LLC',
      region: 'International',
      subregion: 'META',
      country: 'UAE',
      type: 'Reseller',
      tier: 'Gold',
      allocated: 28000,
      spent: 15000,
      pending: 6000,
      status: 'Active',
      note: '',
      contactName: 'Ahmed Al-Rashid',
      contactEmail: 'a.rashid@gulftech.ae',
      accountManager: 'Mike Brown',
    },
    {
      id: 'p9',
      name: 'AsiaPac Solutions',
      region: 'International',
      subregion: 'APAC',
      country: 'Singapore',
      type: 'GSI',
      tier: 'Platinum',
      allocated: 40000,
      spent: 25000,
      pending: 8000,
      status: 'Active',
      note: '',
      contactName: 'Wei Zhang',
      contactEmail: 'w.zhang@asiapac.sg',
      accountManager: 'Tom Davis',
    },
  ]);

  const [requests, setRequests] = useState([
    {
      id: 'REQ-001',
      partner: 'TechVision Ltd',
      submitted: '2026-01-15',
      status: 'request_submitted',
      assignedTo: '',
      poNumber: '',
      note: '',
      items: [
        {
          id: 'REQ-001-A',
          assignedTo: 'Decio A.',
          fyHalf: 'FY26 H1',
          fyQuarter: 'Q1',
          month: 'July',
          period: 'July (Q1)',
          productGroup: 'CyberSecurity',
          tactic: 'Virtual Event / Webinar',
          title: 'Cloud Security Webinar Q1',
          where: 'Online',
          targetAudience: 'IT Decision Makers',
          targetSolutions: 'Cloud, Security',
          objective: 'Lead generation',
          amount: 4000,
          mdfRequest: 2000,
          localCurrency: 'EUR',
          allocadiaId: '',
          itemStatus: 'request_submitted',
        },
        {
          id: 'REQ-001-B',
          assignedTo: 'Decio A.',
          fyHalf: 'FY26 H1',
          fyQuarter: 'Q1',
          month: 'August',
          period: 'August (Q1)',
          productGroup: 'Content (ECS)',
          tactic: 'Digital Advertising',
          title: 'LinkedIn Campaign Q1',
          where: 'Online',
          targetAudience: 'C-Suite',
          targetSolutions: 'Content Management',
          objective: 'Brand awareness',
          amount: 2500,
          mdfRequest: 1250,
          localCurrency: 'EUR',
          allocadiaId: '',
          itemStatus: 'cancelled_by_partner',
          cancelReason: 'Campaign budget reallocated to Q2 event',
          acknowledged: false,
        },
      ],
    },
    {
      id: 'REQ-002',
      partner: 'CloudSys GmbH',
      submitted: '2026-01-20',
      status: 'sent_for_signature',
      assignedTo: 'Kaila',
      poNumber: 'PO-2026-001',
      note: 'Approved in review.',
      items: [
        {
          id: 'REQ-002-A',
          assignedTo: 'Kaila',
          fyHalf: 'FY26 H1',
          fyQuarter: 'Q2',
          month: 'October',
          period: 'October (Q2)',
          productGroup: 'Portfolio',
          tactic: 'Trade Show / Exhibition',
          title: 'CeBIT Frankfurt 2026',
          where: 'Frankfurt, Germany',
          targetAudience: 'IT Managers',
          targetSolutions: 'Portfolio Solutions',
          objective: 'Pipeline generation',
          amount: 8000,
          mdfRequest: 4000,
          localCurrency: 'EUR',
          allocadiaId: 'ALO-2026-001',
        },
      ],
    },
    {
      id: 'REQ-003',
      partner: 'American Tech Corp',
      submitted: '2026-02-01',
      status: 'signed',
      assignedTo: 'Umair',
      poNumber: 'PO-2026-002',
      note: '',
      items: [
        {
          id: 'REQ-003-A',
          assignedTo: 'Umair',
          fyHalf: 'FY26 H1',
          fyQuarter: 'Q1',
          month: 'September',
          period: 'September (Q1)',
          productGroup: 'Portfolio',
          tactic: 'In-Person Event',
          title: 'East Coast Partner Summit',
          where: 'New York, USA',
          targetAudience: 'Enterprise Customers',
          targetSolutions: 'Full Portfolio',
          objective: 'Partner enablement',
          amount: 15000,
          mdfRequest: 7500,
          localCurrency: 'USD',
          allocadiaId: 'ALO-2026-002',
        },
      ],
    },
    {
      id: 'REQ-004',
      partner: 'France Digitale',
      submitted: '2026-02-10',
      status: 'request_submitted',
      assignedTo: '',
      poNumber: '',
      note: '',
      items: [
        {
          id: 'REQ-004-A',
          assignedTo: 'Decio A.',
          fyHalf: 'FY26 H1',
          fyQuarter: 'Q2',
          month: 'November',
          period: 'November (Q2)',
          productGroup: 'CyberSecurity',
          tactic: 'Virtual Event / Webinar',
          title: 'IA et Securite Webinar',
          where: 'Online',
          targetAudience: 'CISO, IT Directors',
          targetSolutions: 'Security Suite',
          objective: 'Lead generation',
          amount: 3000,
          mdfRequest: 1500,
          localCurrency: 'EUR',
          allocadiaId: '',
          itemStatus: 'postponed',
          cancelReason: 'Speaker unavailable - moving to Q3',
          acknowledged: false,
        },
        {
          id: 'REQ-004-B',
          assignedTo: 'Kaila',
          fyHalf: 'FY26 H1',
          fyQuarter: 'Q2',
          month: 'December',
          period: 'December (Q2)',
          productGroup: 'Content (ECS)',
          tactic: 'Content Syndication',
          title: 'Gartner Report Syndication',
          where: 'Online',
          targetAudience: 'C-Suite France',
          targetSolutions: 'Content Solutions',
          objective: 'Brand awareness',
          amount: 4500,
          mdfRequest: 2250,
          localCurrency: 'EUR',
          allocadiaId: '',
        },
      ],
    },
  ]);

  const [claims, setClaims] = useState([]);

  // Auto-calculate spent per partner from approved/paid claims
  const getPartnerSpent = React.useCallback((partnerName) => {
    return claims
      .filter(c => c.partner === partnerName && ['approved','paid'].includes(c.status))
      .reduce((s, c) => {
        // Use claimAmount directly in local currency - stored as EUR equivalent
        return s + Number(c.claimAmount || 0);
      }, 0);
  }, [claims]);

  // Enrich partners with auto-calculated spent
  const enrichedPartners = React.useMemo(() =>
    partners.map(p => ({
      ...p,
      spent: getPartnerSpent(p.name),
      // Sum pipeline for all this partner's activities
      // Key lookup matches AnalyticsTab: campaignId first, then item.id (rowKey)
      pipeline: requests
        .filter(r => r.partner === p.name)
        .flatMap(r => r.items || [])
        .reduce((s, it) => {
          return s + (pipelineData[it.campaignId] || pipelineData[it.id] || 0);
        }, 0),
    }))
  , [partners, claims, requests, pipelineData, getPartnerSpent]);
  const [history, setHistory] = useState([]);

  const addHistory = useCallback(
    (action, entity, type) => {
      setHistory((prev) =>
        [
          {
            id: Date.now().toString(36) + Math.random().toString(36).slice(2,5),
            action,
            entity,
            type,
            user: portalUser ? portalUser.partner.name : currentUser,
            ts: new Date().toLocaleString('en-US'),
          },
          ...prev,
        ].slice(0, 200)
      );
      dbAddHistory(action, entity, type).catch(() => {});
    },
    [currentUser, portalUser]
  );

  const toastTimerRef = useRef(null);
  const toast = useCallback((msg, color = C.success) => {
    setNotif({ msg, color });
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    toastTimerRef.current = setTimeout(() => setNotif(null), 3000);
  }, []);

  const toUSD = useCallback(
    (amount, currency) => {
      if (!currency || currency === 'USD') return Number(amount || 0);
      if (Object.keys(allRates).length === 0) {
        const fallback = {
          EUR: 1.09,
          GBP: 1.27,
          CHF: 1.12,
          SEK: 0.092,
          NOK: 0.091,
          DKK: 0.146,
          PLN: 0.25,
          CZK: 0.044,
          AED: 0.272,
          SGD: 0.74,
          AUD: 0.65,
        };
        return Math.round(Number(amount || 0) * (fallback[currency] || 1));
      }
      const r = allRates[currency];
      return r ? Math.round(Number(amount || 0) / r) : Number(amount || 0);
    },
    [allRates]
  );

  useEffect(() => {
    const fetch2 = async () => {
      setRateLoading(true);
      try {
        const res = await fetch(
          'https://api.exchangerate-api.com/v4/latest/USD'
        );
        const data = await res.json();
        if (data?.rates) {
          const r = data.rates;
          setAllRates(r);
          setRate(1 / r['EUR'] || 1.09);
          setRateTs(
            new Date().toLocaleTimeString('en-US', {
              hour: '2-digit',
              minute: '2-digit',
            })
          );
        }
      } catch {
        // BUG-011 FIX: notify user that FX rates are stale/unavailable
        setRateTs('⚠ Rates unavailable — using fallback (1 USD = 1.09 EUR)');
      } finally {
        setRateLoading(false);
      }
    };
    fetch2();
    const iv = setInterval(fetch2, 3600000);
    return () => clearInterval(iv);
  }, []);

  // ============================================================
  // SUPABASE: Initial data load
  // ============================================================
  useEffect(() => {
    const loadAll = async () => {
      try {
        // PERF FIX: load all tables in parallel instead of sequentially
        // Previously 5 sequential awaits (~750-1000ms). Now 1 parallel await (~150-200ms).
        console.log('[MDF] Loading from Supabase (parallel)...');
        const [pRows, rRows, iRows, cRows] = await Promise.all([
          supa.from('partners').select(),
          supa.from('requests').select(),
          supa.from('request_items').select(),
          supa.from('claims').select(),
        ]);
        console.log('[MDF] Loaded:', pRows.length, 'partners,', rRows.length, 'requests,', cRows.length, 'claims');

        const partners = pRows.map(dbToPartner);
        const reqs = rRows.map(r => dbToRequest(r, iRows.filter(it => it.request_id === r.id)));
        const claims = cRows.map(dbToClaim);

        setPartners(partners);
        setRequests(reqs);
        setClaims(claims);

        // History is non-critical — load async after main data is shown
        supa.from('history').select('order=ts.desc&limit=100')
          .then(hRows => setHistory(hRows.map(h => ({
            id: h.id, action: h.action, entity: h.entity,
            type: h.type, user: h.user_name,
            ts: new Date(h.ts).toLocaleString('en-US'),
          })))).catch(() => {});

        setDbLoaded(true);
        console.log('[MDF] All data loaded successfully');
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ partners, requests: reqs, claims, pipelineData }));
      } catch (err) {
        console.error('[MDF] DB load error:', err.message);
        setDbError(err.message);
        // Fall back to localStorage
        const cached = loadPersistedData();
        if (cached?.partners) setPartners(cached.partners);
        if (cached?.requests) setRequests(cached.requests);
        if (cached?.claims) setClaims(cached.claims);
        setDbLoaded(true);
      }
    };
    // Show cached data immediately while Supabase loads (instant perceived performance)
    const cached = loadPersistedData();
    if (cached?.partners?.length) {
      setPartners(cached.partners);
      setRequests(cached.requests || []);
      setClaims(cached.claims || []);
      setDbLoaded(true); // show UI right away
    }
    loadAll(); // then overwrite with fresh data from DB
  }, []);

  // ============================================================
  // SUPABASE: Polling for real-time sync (every 5 seconds)
  // ============================================================
  useEffect(() => {
    if (!dbLoaded) return;
    const poll = async () => {
      // Don't poll while user has a modal open or is editing a cell - prevents focus loss
      if (document.querySelector('[data-modal="true"]')) return;
      if (document.querySelector('[data-editing="true"]')) return;
      try {
        const [rRows, iRows, cRows] = await Promise.all([
          supa.from('requests').select(),
          supa.from('request_items').select(),
          supa.from('claims').select(),
        ]);
        setRequests(rRows.map(r => dbToRequest(r, iRows.filter(it => it.request_id === r.id))));
        setClaims(cRows.map(dbToClaim));
        // Clear DB error if poll succeeds after a failure
        if (dbError) setDbError(null);
      } catch (e) {
        // BUG-009 FIX: surface DB errors to the user
        console.warn('[MDF] Poll error:', e.message);
        setDbError('Connection issue — changes may not be saving. Retrying...');
      }
    };
    const interval = setInterval(poll, 5000);
    return () => clearInterval(interval);
  }, [dbLoaded]);

    // ============================================================
  // DB WRITE HELPERS (called alongside setState)
  // ============================================================
  const dbSaveRequest = async (req) => {
    try {
      console.log('[MDF] Saving request:', req.id);
      await supa.upsert('requests', {
        id: req.id, partner_name: req.partner,
        submitted: req.submitted, status: req.status,
        assigned_to: req.assignedTo || '',
        po_number: req.poNumber || '', note: req.note || '',
        partner_notified: req.partnerNotified || false,
        notified_at: req.notifiedAt || null,
        bp_generated_at: req.bpGeneratedAt || null,
        signed_doc: req.signedDoc ? JSON.stringify(req.signedDoc) : null,
        updated_at: new Date().toISOString(),
      });
      // Upsert items
      if (req.items?.length) {
        await supa.upsert('request_items', req.items.map(it => itemToDB(it, req.id)));
      }
    } catch (err) { console.warn('dbSaveRequest error:', err.message); }
  };

  const dbSaveClaim = async (claim) => {
    try {
      await supa.upsert('claims', claimToDB(claim));
    } catch (err) { console.warn('dbSaveClaim error:', err.message); }
  };

  const dbSavePipeline = async (campaignId, amount) => {
    try {
      await supa.upsert('pipeline', { campaign_id: campaignId, pipeline_amount: amount, updated_at: new Date().toISOString() });
    } catch (err) { console.warn('dbSavePipeline error:', err.message); }
  };

  const dbSavePartner = async (partner) => {
    try {
      await supa.upsert('partners', partnerToDB(partner));
    } catch (err) { console.warn('dbSavePartner error:', err.message); }
  };

  const dbAddHistory = async (action, entity, type) => {
    try {
      await supa.insert('history', {
        action, entity, type,
        user_name: portalUser ? portalUser.partner.name : currentUser,
      });
    } catch (err) { console.warn('dbAddHistory error:', err.message); }
  };

  // Persist to localStorage as offline fallback (includes pipelineData)
  useEffect(() => {
    if (!dbLoaded) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ partners, requests, claims, pipelineData }));
    } catch {}
  }, [partners, requests, claims, pipelineData, dbLoaded]);

  const fmtA = (n) => (currency === 'EUR' ? fmtEUR(n) : fmtUSD(n, rate));
  const fmtB = (n) => (currency === 'EUR' ? fmtUSD(n, rate) : fmtEUR(n));

  const parseNum = (s) => {
    const n = parseFloat(String(s).replace(/[^0-9.-]/g, ''));
    return isNaN(n) ? 0 : n;
  };

  const updatePartnersSpent = (id, val) => {
    const num = parseNum(String(val));
    setPartners((p) => p.map((x) => (x.id === id ? { ...x, spent: num } : x)));
    setEditPartner(null);
  };
  const updatePartnersNotes = (id, val) =>
    setPartners((p) => p.map((x) => (x.id === id ? { ...x, note: val } : x)));
  const commitPartnersNotes = (id) => {};
  const openPlanModal = (partnerName) => {
    const p = partners.find((x) => x.name === partnerName);
    if (!p) return;
    const acts = requests
      .filter((r) => r.partner === partnerName && r.status === 'signed')
      .flatMap((r) => (r.items || []).map((i) => ({ ...i, selected: true })));
    if (!acts.length) {
      toast('No approved activities for this partner', '#f59e0b');
      return;
    }
    setPlanModal({ partner: p, activities: acts, reqId: null });
  };

  // Badge = requests with no CMM assigned + requests with unacknowledged partner changes
  const pendingCnt = requests.filter(
    (r) => !r.assignedTo || (r.items || []).some(it =>
      (['cancelled_by_partner','postponed'].includes(it.itemStatus) && !it.acknowledged)
    )
  ).length;
  const cancelledByPartnerCnt = requests.reduce(
    (s, r) =>
      (r.items || []).filter((it) => it.itemStatus === 'cancelled_by_partner' && !it.acknowledged)
        .length + s,
    0
  );
  // Unassigned new requests count
  const unassignedReqCnt = requests.filter(r =>
    !r.assignedTo && ['request_submitted','approved'].includes(r.status)
  ).length;
  const postponedByPartnerCnt = requests.reduce(
    (s, r) =>
      (r.items || []).filter((it) => it.itemStatus === 'postponed' && !it.acknowledged).length + s,
    0
  );
  const claimPendingCnt = claims.filter(
    (c) => c.status === 'submitted'
  ).length;

  const dashPartners = useMemo(
    () =>
      enrichedPartners.filter((p) => {
        if (
          Array.isArray(dashMacro)
            ? dashMacro.length > 0 && !dashMacro.includes(p.region)
            : dashMacro !== 'All' && p.region !== dashMacro
        )
          return false;
        if (
          Array.isArray(dashSubregion)
            ? dashSubregion.length > 0 && !dashSubregion.includes(p.subregion)
            : dashSubregion !== 'All' && p.subregion !== dashSubregion
        )
          return false;
        if (
          Array.isArray(dashTier)
            ? dashTier.length > 0 && !dashTier.includes(p.tier)
            : dashTier !== 'All' && p.tier !== dashTier
        )
          return false;
        if (
          Array.isArray(dashPartner)
            ? dashPartner.length > 0 && !dashPartner.includes(p.name)
            : dashPartner !== 'All' && p.name !== dashPartner
        )
          return false;
        if (
          Array.isArray(dashPartnerType)
            ? dashPartnerType.length > 0 && !dashPartnerType.includes(p.type)
            : dashPartnerType &&
              dashPartnerType !== 'All' &&
              p.type !== dashPartnerType
        )
          return false;
        if (
          Array.isArray(dashPAM)
            ? dashPAM.length > 0 && !dashPAM.includes(p.accountManager)
            : dashPAM !== 'All' && p.accountManager !== dashPAM
        )
          return false;
        return true;
      }),
    [partners, dashMacro, dashSubregion, dashTier, dashPartner, dashPartnerType, dashPAM]
  );

  const dashPartnerNames = useMemo(
    () =>
      partners
        .filter(
          (p) =>
            (dashMacro === 'All' || p.region === dashMacro) &&
            (dashSubregion === 'All' || p.subregion === dashSubregion) &&
            (dashTier === 'All' || p.tier === dashTier)
        )
        .map((p) => p.name),
    [partners, dashMacro, dashSubregion, dashTier]
  );

  const totalAlloc = dashPartners.reduce((s, p) => s + p.allocated, 0);
  // Pending = MDF from signed requests (committed, waiting for PO)
  const totalPending = (() => {
    const partnerNames = new Set(dashPartners.map(p => p.name));
    return requests
      .filter(r => partnerNames.has(r.partner) && r.status === 'signed')
      .flatMap(r => r.items || [])
      .reduce((s, it) => s + toUSD(it.mdfRequest || Math.round((it.amount||0)*0.5), it.localCurrency||'EUR'), 0);
  })();
  // Allocated = MDF from PO raised requests only (real budget commitment)
  const totalAllocated = (() => {
    const partnerNames = new Set(dashPartners.map(p => p.name));
    return requests
      .filter(r => partnerNames.has(r.partner) && r.status === 'po_raised')
      .flatMap(r => r.items || [])
      .reduce((s, it) => s + toUSD(it.mdfRequest || Math.round((it.amount||0)*0.5), it.localCurrency||'EUR'), 0);
  })();
  // Spent/Used = actually reimbursed via approved or paid claims
  const totalClaimed = (() => {
    const partnerNames = new Set(dashPartners.map(p => p.name));
    return claims
      .filter(c => partnerNames.has(c.partner) && ['approved','paid'].includes(c.status))
      .reduce((s, c) => s + toUSD(c.claimAmount||0, c.currency||'EUR'), 0);
  })();
  const totalSpent = dashPartners.reduce((s, p) => s + p.spent, 0);
  const totalPend = dashPartners.reduce((s, p) => s + p.pending, 0);
  const totalAvail = totalAlloc - totalAllocated;

  const dashPartnerNameSet = new Set(dashPartners.map((p) => p.name));
  const dashClaims = claims.filter((c) => dashPartnerNameSet.has(c.partner));
  const totalClaimsSubmitted = dashClaims
    .filter((c) => c.status === 'submitted')
    .reduce((s, c) => s + (c.claimAmount || 0), 0);
  const totalClaimsInReview = dashClaims
    .filter((c) => ['marketing_review','finance_review'].includes(c.status))
    .reduce((s, c) => s + (c.claimAmount || 0), 0);
  const totalClaimsPaid = dashClaims
    .filter((c) => c.status === 'paid')
    .reduce((s, c) => s + (c.claimAmount || 0), 0);
  const totalClaimsApproved = dashClaims
    .filter((c) => c.status === 'approved' || c.status === 'paid')
    .reduce((s, c) => s + (c.claimAmount || 0), 0);
  const claimsCnt = dashClaims.length;
  const claimsPaidCnt = dashClaims.filter((c) => c.status === 'paid').length;

  const filteredP = useMemo(
    () =>
      enrichedPartners.filter((p) => {
        const matchSearch =
          !search || p.name.toLowerCase().includes(search.toLowerCase());
        const matchMacro =
          filterMacro.length === 0 || filterMacro.includes(p.region);
        const matchRegion =
          filterRegion.length === 0 || filterRegion.includes(p.subregion);
        const matchTier =
          filterTier.length === 0 || filterTier.includes(p.tier);
        const matchStatus =
          filterStatus.length === 0 || filterStatus.includes(p.status);
        const matchType =
          filterType.length === 0 || filterType.includes(p.type);
        const matchPAM =
          filterPAM.length === 0 || filterPAM.includes(p.accountManager);
        const matchPortal = !isPartnerView || p.name === portalPartner?.name;
        const matchReqStatus =
          filterPartnerReqStatus.length === 0 ||
          requests.some(
            (r) =>
              r.partner === p.name && filterPartnerReqStatus.includes(r.status)
          );
        return (
          matchSearch &&
          matchMacro &&
          matchRegion &&
          matchTier &&
          matchStatus &&
          matchType &&
          matchPAM &&
          matchPortal &&
          matchReqStatus
        );
      }),
    [
      partners,
      search,
      filterMacro,
      filterRegion,
      filterTier,
      filterStatus,
      filterType,
      filterPartnerReqStatus,
      isPartnerView,
      portalPartner,
      requests,
    ]
  );

  const filteredR = useMemo(
    () =>
      requests.filter((r) => {
        const p = partners.find((x) => x.name === r.partner) || {};
        if (
          filterReqStatus.length > 0 &&
          !filterReqStatus.some((s) => {
            if (s === 'bp_sent') return r.partnerNotified;
            // Request-level statuses: only check r.status
            const reqLevelStatuses = ['request_submitted','approved','sent_for_signature','signed','po_raised','rejected'];
            if (reqLevelStatuses.includes(s)) return r.status === s;
            // Item-level statuses (cancelled_by_partner, postponed): check items
            return (r.items || []).some((it) => it.itemStatus === s);
          })
        )
          return false;
        if (
          filterReqPartnerType.length > 0 &&
          !filterReqPartnerType.includes(p.type)
        )
          return false;
        if (
          filterReqPartner.length > 0 &&
          !filterReqPartner.includes(r.partner)
        )
          return false;
        if (
          filterCMM.length > 0 &&
          !(r.items || []).some((it) => filterCMM.includes(it.assignedTo))
        )
          return false;
        if (
          filterReqFY.length > 0 &&
          !(r.items || []).some((it) => filterReqFY.includes(it.fyHalf))
        )
          return false;
        if (
          filterReqQ.length > 0 &&
          !(r.items || []).some((it) => filterReqQ.includes(it.fyQuarter))
        )
          return false;
        if (
          filterReqMonth.length > 0 &&
          !(r.items || []).some((it) => filterReqMonth.includes(it.month))
        )
          return false;
        if (
          search &&
          ![r.id, r.partner, r.poNumber]
            .join(' ')
            .toLowerCase()
            .includes(search.toLowerCase())
        )
          return false;
        return true;
      }),
    [
      requests,
      filterReqStatus,
      filterReqPartnerType,
      filterReqPartner,
      filterReqFY,
      filterReqQ,
      filterReqMonth,
      filterCMM,
      search,
      partners,
    ]
  );

  const sortedR = useMemo(
    () =>
      [...filteredR].sort((a, b) =>
        sortReqDir === 'desc'
          ? (b.submitted || '').localeCompare(a.submitted || '')
          : (a.submitted || '').localeCompare(b.submitted || '')
      ),
    [filteredR, sortReqDir]
  );

  const ALL_SUBREGIONS = Object.values(REGIONS).flat();
  const tierColor = { Platinum: '#f59e0b', Gold: C.accent, Silver: C.muted };

  const navTabs = [
    { id: 'dashboard', icon: 'dashboard', label: 'Dashboard' },
    { id: 'partners', icon: 'partners', label: 'Partners' },
    { id: 'analytics', icon: 'analytics', label: 'MDF Overview' },
    {
      id: 'requests',
      icon: 'requests',
      label: 'Requests',
      badge: unassignedReqCnt || 0,
      alertBadge: cancelledByPartnerCnt + postponedByPartnerCnt || 0,
    },
    { id: 'claims', icon: 'claims', label: 'Claims', badge: claimPendingCnt },
    { id: 'history', icon: 'history', label: 'History' },
  ];

  // DB loading screen
  if (!dbLoaded) return (
    <div style={{ minHeight: '100vh', background: '#00008b', display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: 16 }}>
      <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 28, color: '#60aaff' }}>MDF Manager</div>
      {dbError ? (
        <div style={{ textAlign: 'center' }}>
          <div style={{ color: '#ef4444', fontSize: 13, marginBottom: 8 }}>Database connection failed — using local data</div>
          <div style={{ color: 'rgba(255,255,255,0.5)', fontSize: 11 }}>{dbError}</div>
        </div>
      ) : (
        <div style={{ color: 'rgba(255,255,255,0.6)', fontSize: 13 }}>Connecting to database...</div>
      )}
    </div>
  );

    if (!loggedIn)
    return (
      <div
        style={{
          minHeight: '100vh',
          background: C.bg,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: 20,
        }}
      >
        <div
          style={{
            background: C.surface,
            border: `1px solid ${C.border}`,
            borderRadius: 20,
            padding: 40,
            width: 380,
            boxSizing: 'border-box',
            textAlign: 'center',
          }}
        >
          <div
            style={{
              fontFamily: "'Syne',sans-serif",
              fontWeight: 800,
              fontSize: 28,
              color: C.accent,
              marginBottom: 4,
            }}
          >
            MDF Manager
          </div>
          <div style={{ fontSize: 13, color: C.muted, marginBottom: 32 }}>
            Channel Marketing - Internal Tool
          </div>
          <div style={{ textAlign: 'left', marginBottom: 12 }}>
            <label
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: C.muted,
                letterSpacing: '0.06em',
                display: 'block',
                marginBottom: 6,
              }}
            >
              PASSWORD
            </label>
            <input
              type="password"
              value={loginPwd}
              onChange={(e) => {
                setLoginPwd(e.target.value);
                setLoginErr(false);
              }}
              onKeyDown={(e) =>
                e.key === 'Enter' &&
                (loginPwd === 'mdf2026' ? setLoggedIn(true) : setLoginErr(true))
              }
              placeholder="Enter access password"
              style={{
                width: '100%',
                background: C.faint,
                border: `1px solid ${loginErr ? C.danger : C.border}`,
                color: C.text,
                borderRadius: 10,
                padding: '10px 14px',
                fontSize: 14,
                outline: 'none',
                boxSizing: 'border-box',
              }}
            />
            {loginErr && (
              <div style={{ fontSize: 11, color: C.danger, marginTop: 4 }}>
                Incorrect password
              </div>
            )}
            <div style={{ fontSize: 10, color: C.muted, marginTop: 6 }}>
              Demo password: mdf2026
            </div>
          </div>
          <button
            onClick={() =>
              loginPwd === 'mdf2026' ? setLoggedIn(true) : setLoginErr(true)
            }
            style={{
              width: '100%',
              background: C.accent,
              color: '#fff',
              border: 'none',
              borderRadius: 10,
              padding: 13,
              fontWeight: 800,
              fontSize: 15,
              cursor: 'pointer',
              marginTop: 8,
            }}
          >
            Sign In
          </button>
        </div>
      </div>
    );

  return (
    <div
      key={isDark ? 'dk' : 'lt'}
      style={{
        display: 'flex',
        height: '100vh',
        overflow: 'hidden',
        background: C.bg,
        color: C.text,
        fontFamily: 'system-ui,sans-serif',
        fontSize: 14,
        position: 'relative',
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Mono:ital,wght@0,400;0,500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        input,select,textarea{font-family:inherit}
        button{font-family:inherit;cursor:pointer}
        @keyframes slideIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        select option{background:${C.surface};color:${C.text}}
        body{color:${C.text};background:${C.bg}}
        /* Force all inputs/selects to inherit theme */
        input,select,textarea{background:${C.faint};color:${C.text};border-color:${C.border}}
      `}</style>

      {/* Notification */}
      {notif && (
        <div
          style={{
            position: 'fixed',
            top: 20,
            right: 20,
            zIndex: 999,
            background: notif.color || C.success,
            color: '#000',
            borderRadius: 12,
            padding: '12px 20px',
            fontWeight: 700,
            fontSize: 13,
            boxShadow: '0 4px 20px rgba(0,0,0,0.3)',
            animation: 'slideIn 0.3s ease',
          }}
        >
          {notif.msg}
        </div>
      )}

      {/* Sidebar */}
      <div
        style={{
          width: sidebarCollapsed ? 52 : 215,
          background: isDark ? C.surface : C.navBg,
          borderRight: `1px solid ${isDark ? C.border : '#00006a'}`,
          display: 'flex',
          flexDirection: 'column',
          position: 'fixed',
          top: 0,
          bottom: 0,
          transition: 'width 0.2s ease',
          zIndex: 100,
          overflow: 'hidden',
        }}
      >
        <div
          style={{
            padding: sidebarCollapsed ? '16px 8px' : '22px 20px',
            borderBottom: `1px solid ${C.border}`,
            position: 'relative',
            minHeight: 80,
          }}
        >
          <button
            onClick={() => setSidebarCollapsed((p) => !p)}
            title={sidebarCollapsed ? 'Expand menu' : 'Collapse menu'}
            style={{
              position: 'absolute',
              top: 12,
              right: sidebarCollapsed ? 10 : 12,
              width: 28,
              height: 28,
              borderRadius: 6,
              background: 'transparent',
              border: `1px solid ${C.border}`,
              color: C.muted,
              fontSize: 14,
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 2,
              flexDirection: 'column',
              gap: 3,
              padding: 4,
            }}
          >
            <div
              style={{
                width: 14,
                height: 2,
                background: C.muted,
                borderRadius: 1,
              }}
            />
            <div
              style={{
                width: 14,
                height: 2,
                background: C.muted,
                borderRadius: 1,
              }}
            />
            <div
              style={{
                width: 14,
                height: 2,
                background: C.muted,
                borderRadius: 1,
              }}
            />
          </button>
          {!sidebarCollapsed && (
            <>
              <span
                style={{
                  fontFamily: "'Syne',sans-serif",
                  fontWeight: 800,
                  fontSize: 16,
                  color: isDark ? C.accent : '#60aaff',
                  letterSpacing: '0.05em',
                }}
              >
                OT
              </span>
              <div
                style={{
                  fontFamily: "'Syne',sans-serif",
                  fontSize: 18,
                  fontWeight: 800,
                  lineHeight: 1.2,
                  color: isDark ? C.text : '#ffffff',
                }}
              >
                MDF{' '}
                <span style={{ color: isDark ? C.accent : '#60aaff' }}>
                  Manager
                </span>
              </div>
              <div
                style={{
                  marginTop: 8,
                  background: isDark ? C.accentGlow : 'rgba(255,255,255,0.15)',
                  border: `1px solid ${
                    isDark ? C.accent + '30' : 'rgba(255,255,255,0.3)'
                  }`,
                  borderRadius: 8,
                  padding: '5px 10px',
                  fontSize: 11,
                  color: isDark ? C.accent : '#ffffff',
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                }}
              >
                <span style={{ opacity: 0.6, marginRight: 5 }}>👤</span>{currentUser}
              </div>
              <button
                onClick={() => setIsDark((d) => !d)}
                style={{
                  marginTop: 8,
                  width: '100%',
                  background: isDark ? 'transparent' : 'rgba(255,255,255,0.15)',
                  border: `1px solid ${
                    isDark ? C.border : 'rgba(255,255,255,0.4)'
                  }`,
                  color: isDark ? C.muted : 'rgba(255,255,255,0.9)',
                  borderRadius: 8,
                  padding: '5px 10px',
                  fontSize: 11,
                  cursor: 'pointer',
                  fontWeight: 600,
                }}
              >
                {isDark ? 'Light Mode' : 'Dark Mode'}
              </button>
            </>
          )}
          {sidebarCollapsed && (
            <div
              style={{
                marginTop: 34,
                display: 'flex',
                justifyContent: 'center',
              }}
            >
              <span
                style={{
                  fontFamily: "'Syne',sans-serif",
                  fontWeight: 800,
                  fontSize: 14,
                  color: C.accent,
                  letterSpacing: '0.05em',
                }}
              >
                OT
              </span>
            </div>
          )}
        </div>
        <nav
          style={{
            padding: sidebarCollapsed ? '10px 6px' : '14px 10px',
            flex: 1,
            overflow: 'hidden',
          }}
        >
          {navTabs.map((t) => (
            <button
              key={t.id}
              onClick={() => {
                setTab(t.id);
                setSelReq(null);
                setPlanModal(null);
                setShowCombinedBP(false);
                setShowNewReq(false);
                setEditPartner(null);
              }}
              title={sidebarCollapsed ? t.label : ''}
              style={{
                width: '100%',
                display: 'flex',
                alignItems: 'center',
                gap: 10,
                padding: sidebarCollapsed ? '10px 0' : '10px 12px',
                justifyContent: sidebarCollapsed ? 'center' : 'flex-start',
                borderRadius: 10,
                marginBottom: 3,
                background:
                  tab === t.id
                    ? isDark
                      ? C.accentGlow
                      : 'rgba(255,255,255,0.18)'
                    : 'transparent',
                color:
                  tab === t.id
                    ? isDark
                      ? C.accent
                      : '#ffffff'
                    : isDark
                    ? C.muted
                    : 'rgba(255,255,255,0.65)',
                fontWeight: tab === t.id ? 700 : 500,
                fontSize: 13,
                border: `1px solid ${
                  tab === t.id ? C.accent + '30' : 'transparent'
                }`,
                transition: 'all 0.15s',
                overflow: 'hidden',
                position: 'relative',
              }}
            >
              <NavIcon
                id={t.icon}
                size={sidebarCollapsed ? 20 : 16}
                color={tab === t.id ? C.accent : C.muted}
              />
              {!sidebarCollapsed && (
                <>
                  <span
                    style={{ whiteSpace: 'nowrap', flex: 1, textAlign: 'left' }}
                  >
                    {t.label}
                  </span>
                  {t.badge > 0 && (
                    <span
                      style={{
                        background: isDark ? C.warning : '#ffd700',
                        color: '#000',
                        borderRadius: 20,
                        padding: '1px 7px',
                        fontSize: 10,
                        fontWeight: 800,
                      }}
                    >
                      {t.badge}
                    </span>
                  )}
                  {t.alertBadge > 0 && (
                    <span
                      style={{
                        background: C.danger,
                        color: '#fff',
                        borderRadius: 20,
                        padding: '1px 7px',
                        fontSize: 10,
                        fontWeight: 800,
                      }}
                      title="Partner cancellations/postponements"
                    >
                      {t.alertBadge}
                    </span>
                  )}
                </>
              )}
              {sidebarCollapsed && (t.badge > 0 || t.alertBadge > 0) && (
                <div
                  style={{
                    position: 'absolute',
                    top: 6,
                    right: 6,
                    width: 7,
                    height: 7,
                    borderRadius: '50%',
                    background: t.alertBadge > 0 ? C.danger : C.warning,
                  }}
                />
              )}
            </button>
          ))}
        </nav>
      </div>

      {/* Main content */}
      <div
        style={{
          marginLeft: sidebarCollapsed ? 52 : 215,
          flex: 1,
          transition: 'margin-left 0.2s ease',
          display: 'flex',
          flexDirection: 'column',
          height: '100vh',
          overflow: 'hidden',
          boxSizing: 'border-box',
        }}
      >
        {/* FX rate warning banner */}
        {Object.keys(allRates).length === 0 && !rateLoading && (
          <div style={{ background: C.warning + '18', borderBottom: `1px solid ${C.warning}30`, padding: '6px 20px', fontSize: 12, color: C.warning, fontWeight: 600, display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
            <span>&#9888;</span>
            <span>Exchange rates unavailable — USD values are estimated. Check your connection.</span>
          </div>
        )}
        {rateTs && Object.keys(allRates).length > 0 && (
          <div style={{ background: C.faint, borderBottom: `1px solid ${C.border}20`, padding: '3px 20px', fontSize: 10, color: C.muted, flexShrink: 0 }}>
            Rates updated: {rateTs}
          </div>
        )}

        {/* Dashboard */}
        {tab === 'dashboard' && (
          <div
            style={{
              animation: 'slideIn 0.3s ease',
              padding: 28,
              overflowY: 'auto',
              height: '100%',
            }}
          >
            <div style={{ marginBottom: 20 }}>
              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'flex-start',
                  marginBottom: 14,
                }}
              >
                <div>
                  <div
                    style={{
                      fontFamily: "'Syne',sans-serif",
                      fontSize: 26,
                      fontWeight: 800,
                      marginBottom: 4,
                    }}
                  >
                    MDF Dashboard
                  </div>
                  <div style={{ color: C.muted, fontSize: 13 }}>
                    {dashPartners.length} of {partners.length} partners
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
                  <span
                    style={{
                      fontFamily: "'Syne',sans-serif",
                      fontWeight: 800,
                      fontSize: 16,
                      color: C.accent,
                      letterSpacing: '0.05em',
                    }}
                  >
                    OT
                  </span>
                  <div style={{ width: 1, height: 20, background: C.border }} />
                  <button
                    onClick={() =>
                      doExport(
                        'dashboard',
                        partners,
                        requests,
                        dashPartners,
                        null,
                        null,
                        rate,
                        toUSD
                      )
                    }
                    style={{
                      background: C.faint,
                      color: C.text,
                      border: `1px solid ${C.border}`,
                      borderRadius: 10,
                      padding: '7px 14px',
                      fontWeight: 600,
                      fontSize: 12,
                      cursor: 'pointer',
                    }}
                  >
                    Excel
                  </button>
                  <button
                    onClick={() =>
                      doPPTExport(partners, requests, rate, currency, fmtA)
                    }
                    data-ppt-btn
                    style={{
                      background: C.accent,
                      color: '#fff',
                      border: 'none',
                      borderRadius: 10,
                      padding: '7px 14px',
                      fontWeight: 700,
                      fontSize: 12,
                      cursor: 'pointer',
                    }}
                  >
                    PPT
                  </button>
                </div>
              </div>
              <div
                style={{
                  display: 'flex',
                  gap: 8,
                  flexWrap: 'wrap',
                  padding: '12px 16px',
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 12,
                  alignItems: 'center',
                }}
              >
                <span
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: C.muted,
                    marginRight: 4,
                  }}
                >
                  FILTER:
                </span>
                <MultiSelect
                  value={dashMacro}
                  onChange={(v) => {
                    setDashMacro(v);
                    setDashSubregion([]);
                    setDashPartner([]);
                  }}
                  placeholder="All Regions"
                  options={ALL_MACROS}
                />
                <MultiSelect
                  value={dashSubregion}
                  onChange={(v) => {
                    setDashSubregion(v);
                    setDashPartner([]);
                  }}
                  placeholder="All Sub-Regions"
                  options={
                    Array.isArray(dashMacro) && dashMacro.length > 0
                      ? dashMacro.flatMap((m) => getSubregions(m))
                      : ALL_SUBREGIONS
                  }
                />
                <MultiSelect
                  value={dashTier}
                  onChange={(v) => {
                    setDashTier(v);
                    setDashPartner([]);
                  }}
                  placeholder="All Levels"
                  options={['Platinum', 'Gold', 'Silver']}
                />
                <MultiSelect
                  value={dashPartnerType}
                  onChange={setDashPartnerType}
                  placeholder="All Types"
                  options={PARTNER_TYPES}
                />
                <MultiSelect
                  value={dashPAM === 'All' ? [] : Array.isArray(dashPAM) ? dashPAM : [dashPAM]}
                  onChange={(v) => setDashPAM(v.length === 0 ? 'All' : v)}
                  placeholder="All PAMs"
                  options={[...new Set(partners.map(p => p.accountManager).filter(Boolean))].sort()}
                />
                <MultiSelect
                  value={dashPartner}
                  onChange={setDashPartner}
                  placeholder="All Partners"
                  options={dashPartnerNames.map((n) => ({
                    value: n,
                    label: n,
                  }))}
                />
                {(dashMacro.length > 0 ||
                  dashSubregion.length > 0 ||
                  dashTier.length > 0 ||
                  dashPartner.length > 0 ||
                  dashPartnerType.length > 0) && (
                  <button
                    onClick={() => {
                      setDashMacro([]);
                      setDashSubregion([]);
                      setDashTier([]);
                      setDashPartner([]);
                      setDashPartnerType([]);
                    }}
                    style={{
                      background: 'transparent',
                      border: `1px solid ${C.border}`,
                      color: C.muted,
                      borderRadius: 8,
                      padding: '5px 12px',
                      fontSize: 12,
                      cursor: 'pointer',
                    }}
                  >
                    Clear
                  </button>
                )}
              </div>
            </div>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(3,1fr)',
                gap: 14,
                marginBottom: 14,
              }}
            >
              {[
                {
                  label: 'Available',
                  val: `USD ${Math.round(totalAlloc * rate).toLocaleString('en-US')}`,
                  color: C.accent,
                  sub: `Total budget · ${dashPartners.length} partners`,
                },
                {
                  label: 'Pending',
                  val: `USD ${Math.round(totalPending * rate).toLocaleString('en-US')}`,
                  color: C.warning,
                  sub: `Signed · awaiting PO · ${Math.round((totalPending / (totalAlloc || 1)) * 100)}% of budget`,
                },
                {
                  label: 'Allocated',
                  val: `USD ${Math.round(totalAllocated * rate).toLocaleString('en-US')}`,
                  color: C.purple,
                  sub: `PO raised · ${Math.round((totalAllocated / (totalAlloc || 1)) * 100)}% of budget`,
                },
                {
                  label: 'Spent / Used',
                  val: `USD ${Math.round(totalClaimed * rate).toLocaleString('en-US')}`,
                  color: C.success,
                  sub: `Approved & paid claims · ${Math.round((totalClaimed / (totalAlloc || 1)) * 100)}% of budget`,
                },
              ].map((k) => (
                <div
                  key={k.label}
                  style={{
                    background: C.surface,
                    border: `1px solid ${C.border}`,
                    borderRadius: 14,
                    padding: '16px 20px',
                    position: 'relative',
                    overflow: 'hidden',
                  }}
                >
                  <div
                    style={{
                      position: 'absolute',
                      top: 0,
                      right: 0,
                      width: 50,
                      height: 50,
                      background: `radial-gradient(circle at top right,${k.color}15,transparent)`,
                      borderRadius: '0 14px 0 50px',
                    }}
                  />
                  <div
                    style={{
                      fontSize: 10,
                      color: C.muted,
                      fontWeight: 700,
                      textTransform: 'uppercase',
                      letterSpacing: '0.08em',
                      marginBottom: 8,
                    }}
                  >
                    {k.label}
                  </div>
                  <div
                    style={{
                      fontFamily: "'DM Mono',monospace",
                      fontSize: 20,
                      fontWeight: 800,
                      color: k.color,
                      marginBottom: 2,
                    }}
                  >
                    {k.val}
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: k.color,
                      marginTop: 4,
                      fontWeight: 600,
                      opacity: 0.85,
                    }}
                  >
                    {k.sub}
                  </div>
                </div>
              ))}
            </div>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(3,1fr)',
                gap: 14,
                marginBottom: 20,
              }}
            >
              {[
                {
                  label: 'Pending Approval',
                  val: `USD ${Math.round(totalPend * rate).toLocaleString(
                    'en-US'
                  )}`,
                  color: C.warning,
                  sub: `${pendingCnt} requests awaiting review`,
                },
                {
                  label: 'Claims Submitted',
                  val: `USD ${Math.round(
                    totalClaimsSubmitted * rate
                  ).toLocaleString('en-US')}`,
                  color: C.cyan || '#06b6d4',
                  sub: `${dashClaims.filter(c => c.status === 'submitted').length} new · ${dashClaims.filter(c => ['marketing_review','finance_review'].includes(c.status)).length} in review`,
                },
                {
                  label: 'Claims Approved / Paid',
                  val: `USD ${Math.round(
                    totalClaimsApproved * rate
                  ).toLocaleString('en-US')}`,
                  color: C.teal || '#14b8a6',
                  sub: `${claimsPaidCnt} paid . ${
                    dashClaims.filter((c) => c.status === 'approved').length
                  } approved`,
                },
              ].map((k) => (
                <div
                  key={k.label}
                  style={{
                    background: C.surface,
                    border: `1px solid ${C.border}`,
                    borderRadius: 14,
                    padding: '16px 20px',
                    position: 'relative',
                    overflow: 'hidden',
                  }}
                >
                  <div
                    style={{
                      position: 'absolute',
                      top: 0,
                      right: 0,
                      width: 50,
                      height: 50,
                      background: `radial-gradient(circle at top right,${k.color}15,transparent)`,
                      borderRadius: '0 14px 0 50px',
                    }}
                  />
                  <div
                    style={{
                      fontSize: 10,
                      color: C.muted,
                      fontWeight: 700,
                      textTransform: 'uppercase',
                      letterSpacing: '0.08em',
                      marginBottom: 8,
                    }}
                  >
                    {k.label}
                  </div>
                  <div
                    style={{
                      fontFamily: "'DM Mono',monospace",
                      fontSize: 20,
                      fontWeight: 800,
                      color: k.color,
                      marginBottom: 2,
                    }}
                  >
                    {k.val}
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: k.color,
                      marginTop: 4,
                      fontWeight: 600,
                      opacity: 0.85,
                    }}
                  >
                    {k.sub}
                  </div>
                </div>
              ))}
            </div>

            {/* ROI / Pipeline KPI boxes */}
            {(() => {
              const ROI_TARGET = 30;
              // Sum pipeline for all campaigns belonging to dashPartners
              const partnerNames = new Set(dashPartners.map(p => p.name));
              const allItems = requests
                .filter(r => partnerNames.has(r.partner))
                .flatMap(r => r.items || []);
              const totalPipeline = allItems.reduce((s, it) =>
                s + (pipelineData[it.campaignId] || pipelineData[it.id] || 0), 0);
              // Use totalAllocated (signed/PO raised) as the MDF invested base
              // Falls back to totalAlloc (partner budget) if nothing signed yet
              // ROI base = PO raised (real commitment). Falls back to total budget if no PO yet.
              const totalMDFInvested = totalAllocated > 0 ? totalAllocated : (totalPending > 0 ? totalPending : totalAlloc);
              const targetPipeline = totalMDFInvested * ROI_TARGET;
              const actualRatio = totalMDFInvested > 0 ? totalPipeline / totalMDFInvested : 0;
              const roiColor = actualRatio >= ROI_TARGET ? '#10b981' : actualRatio >= ROI_TARGET * 0.7 ? C.warning : C.danger;
              const pipelineUSD = Math.round(totalPipeline);
              const targetUSD = Math.round(targetPipeline);
              return (
                <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:14, marginBottom:14 }}>
                  <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:14, padding:'16px 20px' }}>
                    <div style={{ fontSize:10, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'0.08em', marginBottom:8 }}>Pipeline Target (1:30)</div>
                    <div style={{ fontFamily:'monospace', fontSize:22, fontWeight:800, color:C.accent }}>
                      USD {targetUSD.toLocaleString('en-US')}
                    </div>
                    <div style={{ fontSize:11, color:C.muted, marginTop:4 }}>
                      {totalMDFInvested > 0 ? `Based on USD ${Math.round(totalMDFInvested).toLocaleString('en-US')} MDF ${totalAllocated > 0 ? 'allocated (PO raised)' : totalPending > 0 ? 'pending (signed)' : 'budgeted'}` : 'No MDF invested yet'}
                    </div>
                  </div>
                  <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:14, padding:'16px 20px' }}>
                    <div style={{ fontSize:10, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'0.08em', marginBottom:8 }}>Pipeline Generated</div>
                    <div style={{ fontFamily:'monospace', fontSize:22, fontWeight:800, color:roiColor }}>
                      {totalPipeline > 0 ? `USD ${pipelineUSD.toLocaleString('en-US')}` : '— No data yet'}
                    </div>
                    <div style={{ fontSize:11, color:C.muted, marginTop:4 }}>
                      {totalPipeline > 0 ? `${allItems.filter(it => pipelineData[it.campaignId]).length} campaigns reporting` : 'Import pipeline via MDF Overview tab'}
                    </div>
                  </div>
                  <div style={{ background:C.surface, border:`2px solid ${roiColor}40`, borderRadius:14, padding:'16px 20px' }}>
                    <div style={{ fontSize:10, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'0.08em', marginBottom:8 }}>ROI Ratio</div>
                    <div style={{ fontFamily:'monospace', fontSize:22, fontWeight:800, color:roiColor }}>
                      {totalPipeline > 0 ? `1 : ${Math.round(actualRatio)}` : '—'}
                    </div>
                    <div style={{ fontSize:11, marginTop:4, fontWeight:600, color:roiColor }}>
                      {totalPipeline > 0
                        ? actualRatio >= ROI_TARGET ? '✓ Above 1:30 target' : `${((actualRatio/ROI_TARGET)*100).toFixed(0)}% of 1:30 target`
                        : 'Target: 1:30'}
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* Pie charts */}
            <div
              key={`pie-${dashMacro}-${dashSubregion}-${dashTier}-${dashPartner}-${dashPartnerType}`}
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 16,
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 14,
                  padding: 22,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2 }}>
                  Budget by Macro Region
                </div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 16 }}>
                  Allocated vs Spent
                </div>
                {(() => {
                  const macroColors = {
                    Europe: C.accent,
                    US: C.success,
                    International: C.purple,
                  };
                  const macroData = ALL_MACROS.map((m) => {
                    const mp = dashPartners.filter((p) => p.region === m);
                    const conv = (n) => currency === 'USD' ? Math.round(n * rate) : n;
                    return {
                      name: m,
                      allocated: conv(mp.reduce((s, p) => s + p.allocated, 0)),
                      spent: conv(mp.reduce((s, p) => s + p.spent, 0)),
                      color: macroColors[m] || C.muted,
                    };
                  }).filter((m) => m.allocated > 0);
                  const total =
                    macroData.reduce((s, m) => s + m.allocated, 0) || 1;
                  let angle = -90;
                  const R = 42,
                    CX = 60,
                    CY = 60,
                    size = 120;
                  const slices = macroData.map((m) => {
                    const pct = m.allocated / total;
                    const sweep = pct * 360;
                    const startA = angle;
                    angle += sweep;
                    const toXY = (a, r) => ({
                      x: CX + r * Math.cos((a * Math.PI) / 180),
                      y: CY + r * Math.sin((a * Math.PI) / 180),
                    });
                    const s2 = toXY(startA, R);
                    const e = toXY(startA + sweep, R);
                    const large = sweep > 180 ? 1 : 0;
                    const midA = startA + sweep / 2;
                    const lp = toXY(midA, R * 0.65);
                    return {
                      ...m,
                      pct: Math.round(pct * 100),
                      sweep,
                      s2,
                      e,
                      large,
                      lp,
                    };
                  });
                  return (
                    <div
                      style={{ display: 'flex', alignItems: 'center', gap: 20 }}
                    >
                      <svg
                        width={size}
                        height={size}
                        viewBox={`0 0 ${size} ${size}`}
                        style={{ flexShrink: 0 }}
                      >
                        {slices.length === 1 ? (
                          <circle cx={CX} cy={CY} r={R} fill={slices[0].color} />
                        ) : slices.map(
                          (s, i) =>
                            s.sweep > 1 && (
                              <path
                                key={i}
                                d={`M ${CX} ${CY} L ${s.s2.x} ${s.s2.y} A ${R} ${R} 0 ${s.large} 1 ${s.e.x} ${s.e.y} Z`}
                                fill={s.color}
                              />
                            )
                        )}
                        {slices.map(
                          (s, i) =>
                            s.pct >= 8 && (
                              <text
                                key={'l' + i}
                                x={s.lp.x}
                                y={s.lp.y + 4}
                                textAnchor="middle"
                                fontSize="9"
                                fontWeight="700"
                                fill="#fff"
                                fontFamily="monospace"
                              >
                                {s.pct}%
                              </text>
                            )
                        )}
                        <circle cx={CX} cy={CY} r={22} fill={C.card} />
                        <text
                          x={CX}
                          y={CY + 5}
                          textAnchor="middle"
                          fontSize="7"
                          fill={C.muted}
                          fontFamily="Arial"
                        >
                          {fmtA(total)}
                        </text>
                      </svg>
                      <div style={{ flex: 1 }}>
                        {macroData.map((m) => (
                          <div key={m.name} style={{ marginBottom: 12 }}>
                            <div
                              style={{
                                display: 'flex',
                                justifyContent: 'space-between',
                                alignItems: 'center',
                                marginBottom: 3,
                              }}
                            >
                              <div
                                style={{
                                  display: 'flex',
                                  alignItems: 'center',
                                  gap: 8,
                                }}
                              >
                                <div
                                  style={{
                                    width: 10,
                                    height: 10,
                                    borderRadius: 2,
                                    background: m.color,
                                  }}
                                />
                                <span style={{ fontSize: 13, fontWeight: 700 }}>
                                  {m.name}
                                </span>
                              </div>
                              <span
                                style={{
                                  fontSize: 12,
                                  fontFamily: 'monospace',
                                  color: m.color,
                                  fontWeight: 700,
                                }}
                              >
                                {Math.round((m.allocated / total) * 100)}%
                              </span>
                            </div>
                            <div
                              style={{
                                fontSize: 11,
                                color: C.muted,
                                marginBottom: 3,
                              }}
                            >
                              {fmtA(m.spent)} spent of {fmtA(m.allocated)}
                            </div>
                            <div
                              style={{
                                background: C.faint,
                                borderRadius: 3,
                                height: 5,
                                overflow: 'hidden',
                              }}
                            >
                              <div
                                style={{
                                  width: `${
                                    m.allocated
                                      ? Math.round(
                                          (m.spent / m.allocated) * 100
                                        )
                                      : 0
                                  }%`,
                                  height: '100%',
                                  background: m.color,
                                  borderRadius: 3,
                                }}
                              />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}
              </div>
              <div
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 14,
                  padding: 22,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2 }}>
                  Budget by Business Unit
                </div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 16 }}>
                  MDF requested by Product Group
                </div>
                {(() => {
                  const buColors = [
                    C.accent,
                    '#f59e0b',
                    C.success,
                    C.purple,
                    '#ec4899',
                    C.cyan || '#06b6d4',
                    '#f97316',
                    '#14b8a6',
                  ];
                  const dashReqs = requests.filter((r) =>
                    dashPartners.some((p) => p.name === r.partner)
                  );
                  const buMap = {};
                  dashReqs.forEach((r) =>
                    (r.items || []).forEach((item) => {
                      const bu = item.productGroup || 'Other';
                      if (!buMap[bu])
                        buMap[bu] = { name: bu, amount: 0, count: 0 };
                      buMap[bu].amount += item.amount || 0;
                      buMap[bu].count++;
                    })
                  );
                  const buData = Object.values(buMap)
                    .sort((a, b) => b.amount - a.amount)
                    .slice(0, 8)
                    .map((b, i) => ({
                      ...b,
                      color: buColors[i % buColors.length],
                    }));
                  if (!buData.length)
                    return (
                      <div
                        style={{
                          color: C.muted,
                          fontSize: 13,
                          padding: '20px 0',
                          textAlign: 'center',
                        }}
                      >
                        No request data yet
                      </div>
                    );
                  const total = buData.reduce((s, b) => s + b.amount, 0) || 1;
                  let angle2 = -90;
                  const R = 42,
                    CX = 60,
                    CY = 60,
                    size = 120;
                  const slices2 = buData.map((b) => {
                    const pct = b.amount / total;
                    const sweep = pct * 360;
                    const startA = angle2;
                    angle2 += sweep;
                    const toXY = (a, r) => ({
                      x: CX + r * Math.cos((a * Math.PI) / 180),
                      y: CY + r * Math.sin((a * Math.PI) / 180),
                    });
                    const s2 = toXY(startA, R);
                    const e = toXY(startA + sweep, R);
                    const large = sweep > 180 ? 1 : 0;
                    const midA = startA + sweep / 2;
                    const lp = toXY(midA, R * 0.65);
                    return {
                      ...b,
                      pct: Math.round(pct * 100),
                      sweep,
                      s2,
                      e,
                      large,
                      lp,
                    };
                  });
                  return (
                    <div
                      style={{ display: 'flex', alignItems: 'center', gap: 20 }}
                    >
                      <svg
                        width={size}
                        height={size}
                        viewBox={`0 0 ${size} ${size}`}
                        style={{ flexShrink: 0 }}
                      >
                        {slices2.map(
                          (s, i) =>
                            s.sweep > 1 && (
                              <path
                                key={i}
                                d={`M ${CX} ${CY} L ${s.s2.x} ${s.s2.y} A ${R} ${R} 0 ${s.large} 1 ${s.e.x} ${s.e.y} Z`}
                                fill={s.color}
                              />
                            )
                        )}
                        {slices2.map(
                          (s, i) =>
                            s.pct >= 8 && (
                              <text
                                key={'l' + i}
                                x={s.lp.x}
                                y={s.lp.y + 4}
                                textAnchor="middle"
                                fontSize="9"
                                fontWeight="700"
                                fill="#fff"
                                fontFamily="monospace"
                              >
                                {s.pct}%
                              </text>
                            )
                        )}
                        <circle cx={CX} cy={CY} r={22} fill={C.card} />
                        <text
                          x={CX}
                          y={CY + 5}
                          textAnchor="middle"
                          fontSize="7"
                          fill={C.muted}
                          fontFamily="Arial"
                        >
                          {buData.length} BUs
                        </text>
                      </svg>
                      <div
                        style={{ flex: 1, overflowY: 'auto', maxHeight: 160 }}
                      >
                        {slices2.map((b) => (
                          <div key={b.name} style={{ marginBottom: 8 }}>
                            <div
                              style={{
                                display: 'flex',
                                justifyContent: 'space-between',
                                alignItems: 'center',
                                marginBottom: 2,
                              }}
                            >
                              <div
                                style={{
                                  display: 'flex',
                                  alignItems: 'center',
                                  gap: 6,
                                  flex: 1,
                                  minWidth: 0,
                                }}
                              >
                                <div
                                  style={{
                                    width: 8,
                                    height: 8,
                                    borderRadius: 2,
                                    background: b.color,
                                    flexShrink: 0,
                                  }}
                                />
                                <span
                                  style={{
                                    fontSize: 12,
                                    fontWeight: 600,
                                    overflow: 'hidden',
                                    textOverflow: 'ellipsis',
                                    whiteSpace: 'nowrap',
                                  }}
                                >
                                  {b.name}
                                </span>
                              </div>
                              <span
                                style={{
                                  fontSize: 12,
                                  fontFamily: 'monospace',
                                  color: b.color,
                                  fontWeight: 700,
                                  marginLeft: 8,
                                  flexShrink: 0,
                                }}
                              >
                                {b.pct}%
                              </span>
                            </div>
                            <div
                              style={{
                                background: C.faint,
                                borderRadius: 3,
                                height: 4,
                                overflow: 'hidden',
                              }}
                            >
                              <div
                                style={{
                                  width: `${b.pct}%`,
                                  height: '100%',
                                  background: b.color,
                                  borderRadius: 3,
                                }}
                              />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}
              </div>
            </div>
            {/* Charts Row 1 */}
            <div
              key={`charts1-${dashMacro}-${dashSubregion}-${dashTier}-${dashPartnerType}`}
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 16,
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 14,
                  padding: 22,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>
                  Spend by Sub-Region
                </div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 16 }}>
                  Allocated vs Spent
                </div>
                {(() => {
                  const regions = [
                    ...new Set(
                      dashPartners.map((p) => p.subregion || p.region)
                    ),
                  ].slice(0, 8);
                  const maxVal = Math.max(
                    ...regions.map((r) =>
                      dashPartners
                        .filter((p) => (p.subregion || p.region) === r)
                        .reduce((s, p) => s + p.allocated, 0)
                    ),
                    1
                  );
                  return regions.map((reg) => {
                    const rp = dashPartners.filter(
                      (p) => (p.subregion || p.region) === reg
                    );
                    const alloc = rp.reduce((s, p) => s + p.allocated, 0);
                    const spent = rp.reduce((s, p) => s + p.spent, 0);
                    const pct = Math.round((alloc / maxVal) * 100);
                    const spentPct = Math.round((spent / alloc) * 100) || 0;
                    return (
                      <div key={reg} style={{ marginBottom: 10 }}>
                        <div
                          style={{
                            display: 'flex',
                            justifyContent: 'space-between',
                            marginBottom: 3,
                          }}
                        >
                          <span style={{ fontSize: 12, fontWeight: 600 }}>
                            {reg}
                          </span>
                          <span
                            style={{
                              fontSize: 11,
                              fontFamily: 'monospace',
                              color: C.muted,
                            }}
                          >
                            {fmtA(spent)} / {fmtA(alloc)}
                          </span>
                        </div>
                        <div
                          style={{
                            background: C.faint,
                            borderRadius: 4,
                            height: 8,
                            overflow: 'hidden',
                          }}
                        >
                          <div
                            style={{
                              width: `${pct}%`,
                              height: '100%',
                              background: C.accent + '40',
                              borderRadius: 4,
                              position: 'relative',
                            }}
                          >
                            <div
                              style={{
                                width: `${spentPct}%`,
                                height: '100%',
                                background: C.accent,
                                borderRadius: 4,
                              }}
                            />
                          </div>
                        </div>
                      </div>
                    );
                  });
                })()}
              </div>
              <div
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 14,
                  padding: 22,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>
                  Spend by Partner Type
                </div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 16 }}>
                  Budget allocated per type
                </div>
                {(() => {
                  const typeColors = {
                    Reseller: C.accent,
                    Distributor: C.success,
                    GSI: C.purple,
                    ISVP: C.warning,
                  };
                  const total =
                    dashPartners.reduce((s, p) => s + p.allocated, 0) || 1;
                  return PARTNER_TYPES.map((type) => {
                    const tp = dashPartners.filter((p) => p.type === type);
                    if (!tp.length) return null;
                    const alloc = tp.reduce((s, p) => s + p.allocated, 0);
                    const spent = tp.reduce((s, p) => s + p.spent, 0);
                    const pct = Math.round((alloc / total) * 100);
                    const col = typeColors[type] || C.accent;
                    return (
                      <div key={type} style={{ marginBottom: 14 }}>
                        <div
                          style={{
                            display: 'flex',
                            justifyContent: 'space-between',
                            alignItems: 'center',
                            marginBottom: 4,
                          }}
                        >
                          <div
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              gap: 8,
                            }}
                          >
                            <div
                              style={{
                                width: 10,
                                height: 10,
                                borderRadius: 2,
                                background: col,
                                flexShrink: 0,
                              }}
                            />
                            <span style={{ fontSize: 13, fontWeight: 600 }}>
                              {type}
                            </span>
                            <span style={{ fontSize: 10, color: C.muted }}>
                              ({tp.length} partners)
                            </span>
                          </div>
                          <span
                            style={{
                              fontSize: 12,
                              fontFamily: 'monospace',
                              color: col,
                              fontWeight: 700,
                            }}
                          >
                            {pct}%
                          </span>
                        </div>
                        <div
                          style={{
                            background: C.faint,
                            borderRadius: 4,
                            height: 10,
                            overflow: 'hidden',
                          }}
                        >
                          <div
                            style={{
                              width: `${pct}%`,
                              height: '100%',
                              background: col,
                              borderRadius: 4,
                              opacity: 0.8,
                            }}
                          />
                        </div>
                        <div
                          style={{ fontSize: 10, color: C.muted, marginTop: 2 }}
                        >
                          {fmtA(spent)} spent of {fmtA(alloc)}
                        </div>
                      </div>
                    );
                  }).filter(Boolean);
                })()}
              </div>
            </div>
            {/* Request status + utilization */}
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: 16,
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 14,
                  padding: 22,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>
                  Requests by Status
                </div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 16 }}>
                  Current distribution
                </div>
                {(() => {
                  const dashReqs = requests.filter((r) =>
                    dashPartners.some((p) => p.name === r.partner)
                  );
                  const statuses = [
                    {
                      key: 'request_submitted',
                      label: 'Submitted',
                      color: C.warning,
                    },
                    { key: 'approved', label: 'Approved', color: '#f97316' },
                    {
                      key: 'sent_for_signature',
                      label: 'Sent for Signature',
                      color: C.purple,
                    },
                    { key: 'signed', label: 'Signed', color: C.success },
                    { key: 'po_raised', label: 'PO Raised', color: C.accent },
                    { key: 'rejected', label: 'Rejected', color: C.danger },
                    {
                      key: 'cancelled_by_partner',
                      label: 'Cancelled by Partner',
                      color: C.danger,
                    },
                    { key: 'postponed', label: 'Postponed', color: C.warning },
                  ];
                  const total = dashReqs.length || 1;
                  return statuses.map((s) => {
                    const count = dashReqs.filter(
                      (r) => r.status === s.key
                    ).length;
                    const amt = dashReqs
                      .filter((r) => r.status === s.key)
                      .reduce(
                        (sum, r) =>
                          (r.items || []).reduce((ss, i) => ss + i.amount, sum),
                        0
                      );
                    const pct = Math.round((count / total) * 100);
                    return (
                      <div key={s.key} style={{ marginBottom: 12 }}>
                        <div
                          style={{
                            display: 'flex',
                            justifyContent: 'space-between',
                            alignItems: 'center',
                            marginBottom: 4,
                          }}
                        >
                          <div
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              gap: 8,
                            }}
                          >
                            <div
                              style={{
                                width: 10,
                                height: 10,
                                borderRadius: '50%',
                                background: s.color,
                              }}
                            />
                            <span style={{ fontSize: 12, fontWeight: 600 }}>
                              {s.label}
                            </span>
                          </div>
                          <div
                            style={{
                              display: 'flex',
                              gap: 12,
                              alignItems: 'center',
                            }}
                          >
                            <span
                              style={{
                                fontSize: 11,
                                color: C.muted,
                                fontFamily: 'monospace',
                              }}
                            >
                              {fmtA(amt)}
                            </span>
                            <span
                              style={{
                                fontSize: 12,
                                fontWeight: 700,
                                color: s.color,
                                minWidth: 32,
                                textAlign: 'right',
                              }}
                            >
                              {count}
                            </span>
                          </div>
                        </div>
                        <div
                          style={{
                            background: C.faint,
                            borderRadius: 4,
                            height: 6,
                          }}
                        >
                          <div
                            style={{
                              width: `${pct}%`,
                              height: '100%',
                              background: s.color,
                              borderRadius: 4,
                              opacity: 0.8,
                            }}
                          />
                        </div>
                      </div>
                    );
                  });
                })()}
              </div>
              <div
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 14,
                  padding: 22,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>
                  Budget Utilization by Tier
                </div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 16 }}>
                  Allocated vs Spent vs Pending
                </div>
                {(() => {
                  const tierColors = {
                    Platinum: '#f59e0b',
                    Gold: C.accent,
                    Silver: C.muted,
                  };
                  return ['Platinum', 'Gold', 'Silver']
                    .map((tier) => {
                      const tp = dashPartners.filter((p) => p.tier === tier);
                      if (!tp.length) return null;
                      const alloc = tp.reduce((s, p) => s + p.allocated, 0);
                      const spent = tp.reduce((s, p) => s + p.spent, 0);
                      const pend = tp.reduce((s, p) => s + p.pending, 0);
                      const util = Math.round(
                        ((spent + pend) / (alloc || 1)) * 100
                      );
                      const col = tierColors[tier] || C.accent;
                      return (
                        <div key={tier} style={{ marginBottom: 18 }}>
                          <div
                            style={{
                              display: 'flex',
                              justifyContent: 'space-between',
                              marginBottom: 6,
                            }}
                          >
                            <div
                              style={{
                                display: 'flex',
                                alignItems: 'center',
                                gap: 8,
                              }}
                            >
                              <span
                                style={{
                                  fontWeight: 700,
                                  fontSize: 13,
                                  color: col,
                                }}
                              >
                                {tier}
                              </span>
                              <span style={{ fontSize: 10, color: C.muted }}>
                                ({tp.length} partners)
                              </span>
                            </div>
                            <span
                              style={{
                                fontSize: 12,
                                fontFamily: 'monospace',
                                fontWeight: 700,
                                color:
                                  util > 80
                                    ? C.danger
                                    : util > 60
                                    ? C.warning
                                    : C.success,
                              }}
                            >
                              {util}% used
                            </span>
                          </div>
                          <div
                            style={{
                              background: C.faint,
                              borderRadius: 6,
                              height: 12,
                              overflow: 'hidden',
                              position: 'relative',
                            }}
                          >
                            <div
                              style={{
                                width: `${Math.round((spent / alloc) * 100)}%`,
                                height: '100%',
                                background: col,
                                position: 'absolute',
                                borderRadius: 6,
                              }}
                            />
                            <div
                              style={{
                                left: `${Math.round((spent / alloc) * 100)}%`,
                                width: `${Math.round((pend / alloc) * 100)}%`,
                                height: '100%',
                                background: col + '60',
                                position: 'absolute',
                              }}
                            />
                          </div>
                          <div
                            style={{
                              display: 'flex',
                              gap: 12,
                              marginTop: 4,
                              fontSize: 10,
                              color: C.muted,
                            }}
                          >
                            <span>Spent: {fmtA(spent)}</span>
                            <span>Pending: {fmtA(pend)}</span>
                            <span>Available: {fmtA(alloc - spent - pend)}</span>
                          </div>
                        </div>
                      );
                    })
                    .filter(Boolean);
                })()}
              </div>
            </div>

            {/* ── Partner Performance: Top & At Risk ── */}
            {(() => {
              const ROI_TARGET = 30;
              const partnerPerf = dashPartners.map(p => {
                const partnerItems = requests
                  .filter(r => r.partner === p.name)
                  .flatMap(r => r.items || []);
                const totalPipeline = partnerItems.reduce((s, it) =>
                  s + (pipelineData[it.campaignId] || pipelineData[it.id] || 0), 0);
                const mdfAllocated = requests
                  .filter(r => r.partner === p.name && ['signed','po_raised'].includes(r.status))
                  .flatMap(r => r.items || [])
                  .reduce((s, it) => s + Number(it.mdfRequest || 0), 0);
                const mdfBase = mdfAllocated > 0 ? mdfAllocated : p.allocated;
                const ratio = mdfBase > 0 ? totalPipeline / mdfBase : 0;
                return { ...p, totalPipeline, mdfBase, ratio, hasPipeline: totalPipeline > 0, hasCommitted: mdfAllocated > 0 };
              });
              const topPerformers = [...partnerPerf].filter(p => p.hasPipeline).sort((a,b) => b.ratio - a.ratio).slice(0,5);
              const atRisk = [...partnerPerf].filter(p => p.hasCommitted && p.ratio < ROI_TARGET * 0.5).sort((a,b) => b.mdfBase - a.mdfBase).slice(0,5);
              if (topPerformers.length === 0 && atRisk.length === 0) return null;
              return (
                <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14, marginTop:14 }}>
                  {topPerformers.length > 0 && (
                    <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:14, padding:'16px 20px' }}>
                      <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:14 }}>
                        <span style={{ fontSize:18 }}>🏆</span>
                        <span style={{ fontFamily:"'Syne',sans-serif", fontWeight:800, fontSize:14 }}>Top Performers</span>
                        <span style={{ fontSize:10, color:C.muted, marginLeft:'auto' }}>by ROI ratio</span>
                      </div>
                      {topPerformers.map(p => <DashPartnerRow key={p.name} p={p} />)}
                    </div>
                  )}
                  {atRisk.length > 0 && (
                    <div style={{ background:C.surface, border:`1px solid #ef444430`, borderRadius:14, padding:'16px 20px' }}>
                      <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:14 }}>
                        <span style={{ fontSize:18 }}>⚠️</span>
                        <span style={{ fontFamily:"'Syne',sans-serif", fontWeight:800, fontSize:14, color:'#ef4444' }}>At Risk</span>
                        <span style={{ fontSize:10, color:C.muted, marginLeft:'auto' }}>MDF committed · low pipeline</span>
                      </div>
                      {atRisk.map(p => <DashPartnerRow key={p.name} p={p} />)}
                    </div>
                  )}
                </div>
              );
            })()}

          </div>
        )}

        {/* MDF Overview */}
        {tab === 'analytics' && (
          <div
            style={{
              flex: 1,
              display: 'flex',
              flexDirection: 'column',
              overflow: 'hidden',
              padding: 0,
            }}
          >
            <AnalyticsTab
              partners={partners}
              requests={requests}
              claims={claims}
              pipelineData={pipelineData}
              currency={currency}
              fmtA={fmtA}
              fmtB={fmtB}
              toUSD={toUSD}
              isDark={isDark}
              onImportPipeline={() => setShowPipelineImport(true)}
              onSavePipeline={(campaignId, val) => {
                setPipelineData(prev => {
                  const next = { ...prev, [campaignId]: val };
                  // Persist to localStorage
                  try {
                    const cached = JSON.parse(localStorage.getItem('mdf_manager_data') || '{}');
                    localStorage.setItem('mdf_manager_data', JSON.stringify({ ...cached, pipelineData: next }));
                  } catch {}
                  return next;
                });
                // Persist to Supabase so Partner Portal can read it
                dbSavePipeline(campaignId, val);
              }}
              onExport={(rows) =>
                doExport(
                  'analytics',
                  partners,
                  requests,
                  null,
                  null,
                  rows,
                  rate,
                  toUSD
                )
              }
              onSaveOverride={(rowKey, field, val) => {
                // rowKey is item.id - find the request containing this item and update it
                setRequests(prev => prev.map(r => ({
                  ...r,
                  items: (r.items || []).map(it =>
                    (it.id || `${r.id}-${r.items.indexOf(it)}`) === rowKey
                      ? { ...it, [field]: val }
                      : it
                  ),
                })));
                // BUG-005 FIX: when Campaign ID is entered, migrate any pipeline stored
                // under rowKey to the new campaignId key so they stay linked
                if (field === 'campaignId' && val) {
                  setPipelineData(prev => {
                    if (prev[rowKey] !== undefined && prev[val] === undefined) {
                      const next = { ...prev, [val]: prev[rowKey] };
                      delete next[rowKey];
                      try {
                        const cached = JSON.parse(localStorage.getItem('mdf_manager_data') || '{}');
                        localStorage.setItem('mdf_manager_data', JSON.stringify({ ...cached, pipelineData: next }));
                      } catch {}
                      return next;
                    }
                    return prev;
                  });
                }
              }}
            />
          </div>
        )}

        {/* Partners */}
        {tab === 'partners' && (
          <div
            style={{
              animation: 'slideIn 0.3s ease',
              padding: 28,
              overflowY: 'auto',
              height: '100%',
            }}
          >
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'flex-end',
                marginBottom: 14,
              }}
            >
              <div>
                <div
                  style={{
                    fontFamily: "'Syne',sans-serif",
                    fontSize: 26,
                    fontWeight: 800,
                    marginBottom: 4,
                  }}
                >
                  Partner MDF Overview
                </div>
                <div style={{ color: C.muted, fontSize: 13 }}>
                  {filteredP.length} of {partners.length} partners
                </div>
              </div>
              <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
                <span
                  style={{
                    fontFamily: "'Syne',sans-serif",
                    fontWeight: 800,
                    fontSize: 16,
                    color: C.accent,
                    letterSpacing: '0.05em',
                  }}
                >
                  OT
                </span>
                <div style={{ width: 1, height: 20, background: C.border }} />
                <button
                  onClick={() => setEditPartner({ _new: true })}
                  style={{
                    background: C.accent,
                    color: '#fff',
                    border: 'none',
                    borderRadius: 10,
                    padding: '7px 14px',
                    fontWeight: 700,
                    fontSize: 12,
                    cursor: 'pointer',
                  }}
                >
                  + Add Partner
                </button>
                <button
                  onClick={() => setShowImport(true)}
                  style={{
                    background: C.faint,
                    color: C.text,
                    border: `1px solid ${C.border}`,
                    borderRadius: 10,
                    padding: '7px 14px',
                    fontWeight: 600,
                    fontSize: 12,
                    cursor: 'pointer',
                  }}
                >
                  Import Partners
                </button>
                <button
                  onClick={() =>
                    doExport(
                      'partners',
                      partners,
                      requests,
                      filteredP,
                      null,
                      null,
                      rate,
                      toUSD
                    )
                  }
                  style={{
                    background: C.faint,
                    color: C.text,
                    border: `1px solid ${C.border}`,
                    borderRadius: 10,
                    padding: '7px 14px',
                    fontWeight: 600,
                    fontSize: 12,
                    cursor: 'pointer',
                  }}
                >
                  Excel
                </button>
                <button
                  onClick={() =>
                    doPPTExport(partners, requests, rate, currency, fmtA)
                  }
                  data-ppt-btn
                  style={{
                    background: C.accent,
                    color: '#fff',
                    border: 'none',
                    borderRadius: 10,
                    padding: '7px 14px',
                    fontWeight: 700,
                    fontSize: 12,
                    cursor: 'pointer',
                  }}
                >
                  PPT
                </button>
              </div>
            </div>
            <div
              style={{
                display: 'flex',
                gap: 10,
                marginBottom: 14,
                flexWrap: 'wrap',
                alignItems: 'center',
              }}
            >
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 10,
                  padding: '6px 12px',
                }}
              >
                <span style={{ fontSize: 11, fontWeight: 700, color: C.muted }}>
                  FY:
                </span>
                <MultiSelect
                  value={
                    partnerFY === 'All'
                      ? []
                      : Array.isArray(partnerFY)
                      ? partnerFY
                      : [partnerFY]
                  }
                  onChange={(v) => setPartnerFY(v.length === 0 ? 'All' : v)}
                  placeholder="All Years"
                  options={['FY25', 'FY26', 'FY27', 'FY28']}
                />
                {(filterMacro.length > 0 ||
                  filterRegion.length > 0 ||
                  filterTier.length > 0 ||
                  filterType.length > 0 ||
                  filterPAM.length > 0 ||
                  filterStatus.length > 0 ||
                  filterPartnerReqStatus.length > 0 ||
                  (Array.isArray(partnerFY)
                    ? partnerFY.length > 0
                    : partnerFY !== 'All') ||
                  search) && (
                  <button
                    onClick={() => {
                      setFilterMacro([]);
                      setFilterRegion([]);
                      setFilterTier([]);
                      setFilterType([]);
                      setFilterStatus([]);
                      setFilterPartnerReqStatus([]);
                      setPartnerFY('All');
                      setSearch('');
                    }}
                    style={{
                      background: 'transparent',
                      border: `1px solid ${C.border}`,
                      color: C.muted,
                      borderRadius: 8,
                      padding: '5px 12px',
                      fontSize: 12,
                      cursor: 'pointer',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    Clear
                  </button>
                )}
              </div>
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search partners..."
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  color: C.text,
                  borderRadius: 10,
                  padding: '7px 14px',
                  fontSize: 13,
                  width: 200,
                }}
              />
              <MultiSelect
                value={filterMacro}
                onChange={(v) => {
                  setFilterMacro(v);
                  setFilterRegion([]);
                }}
                placeholder="All Regions"
                options={ALL_MACROS}
              />
              <MultiSelect
                value={filterRegion}
                onChange={setFilterRegion}
                placeholder="All Sub-Regions"
                options={
                  filterMacro.length === 0
                    ? ALL_SUBREGIONS
                    : getSubregions(filterMacro[0])
                }
              />
              <MultiSelect
                value={filterTier}
                onChange={setFilterTier}
                placeholder="All Levels"
                options={['Platinum', 'Gold', 'Silver']}
              />
              <MultiSelect
                value={filterType}
                onChange={setFilterType}
                placeholder="All Types"
                options={PARTNER_TYPES}
              />
               <MultiSelect
                 value={filterPAM}
                 onChange={setFilterPAM}
                 placeholder="All PAMs"
                 options={[...new Set(partners.map(p => p.accountManager).filter(Boolean))].sort()}
               />
              <MultiSelect
                value={filterPartnerReqStatus}
                onChange={setFilterPartnerReqStatus}
                placeholder="All Request Statuses"
                options={[
                  { value: 'request_submitted', label: 'Submitted' },
                  { value: 'approved', label: 'Approved' },
                  { value: 'sent_for_signature', label: 'Sent for Signature' },
                  { value: 'signed', label: 'Signed' },
                  { value: 'po_raised', label: 'PO Raised' },
                  { value: 'rejected', label: 'Rejected' },
                  {
                    value: 'cancelled_by_partner',
                    label: 'Cancelled by Partner',
                  },
                  { value: 'postponed', label: 'Postponed' },
                ]}
              />
            </div>
            <div
              style={{
                background: C.surface,
                border: `1px solid ${C.border}`,
                borderRadius: 14,
                overflow: 'hidden',
              }}
            >
              <div style={{ overflowX: 'auto' }}>
                <table
                  style={{
                    width: '100%',
                    borderCollapse: 'collapse',
                    minWidth: 900,
                  }}
                >
                  <thead>
                    <tr
                      style={{
                        borderBottom: `2px solid ${C.border}`,
                        background: C.faint,
                      }}
                    >
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Region
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Sub-Region
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Country
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                          minWidth: 140,
                        }}
                      >
                        Partner Name
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Type
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Level
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'left',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        PAM
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'right',
                          fontSize: 10,
                          color: isDark ? C.muted : '#00008b',
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Allocated
                      </th>
                      {['Q1', 'Q2', 'Q3', 'Q4'].map((q) => (
                        <th
                          key={q}
                          style={{
                            padding: '10px 14px',
                            textAlign: 'right',
                            fontSize: 10,
                            color: C.accent,
                            fontWeight: 700,
                            letterSpacing: '0.08em',
                            textTransform: 'uppercase',
                            whiteSpace: 'nowrap',
                            background: C.accentGlow,
                          }}
                        >
                          {q} (USD)
                        </th>
                      ))}
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'right',
                          fontSize: 10,
                          color: C.success,
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Total (USD)
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'right',
                          fontSize: 10,
                          color: C.cyan || C.accent,
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Pipeline
                      </th>
                      <th
                        style={{
                          padding: '10px 14px',
                          textAlign: 'right',
                          fontSize: 10,
                          color: C.success,
                          fontWeight: 700,
                          letterSpacing: '0.08em',
                          textTransform: 'uppercase',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        ROI{' '}
                        <span style={{ fontWeight: 400, color: C.muted }}>
                          goal 1:30
                        </span>
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredP.map((p, i) => {
                      const partnerReqs = requests.filter(
                        (r) => r.partner === p.name
                      );
                      const qData = { Q1: 0, Q2: 0, Q3: 0, Q4: 0 };
                      partnerReqs.forEach((r) => {
                        (r.items || []).forEach((item) => {
                          const q = item.fyQuarter || '';
                          const fyMatch =
                            partnerFY === 'All' ||
                            (Array.isArray(partnerFY)
                              ? partnerFY.length === 0 ||
                                partnerFY.some((fy) =>
                                  (item.fyHalf || '').includes(fy)
                                )
                              : (item.fyHalf || '').includes(partnerFY));
                          if (fyMatch && qData[q] !== undefined)
                            qData[q] += toUSD(
                              item.mdfRequest ||
                                Math.round((item.amount || 0) * 0.5),
                              item.localCurrency || 'EUR'
                            );
                        });
                      });
                      const totalReq = Object.values(qData).reduce(
                        (s, v) => s + v,
                        0
                      );
                      const tierColors = {
                        Platinum: '#f59e0b',
                        Gold: C.accent,
                        Silver: C.muted,
                      };
                      return (
                        <tr
                          key={p.id}
                          style={{
                            borderBottom:
                              i < filteredP.length - 1
                                ? `1px solid ${C.border}20`
                                : 'none',
                            transition: 'background 0.1s',
                          }}
                        >
                          <td
                            style={{
                              padding: '10px 14px',
                              fontSize: 12,
                              color: C.muted,
                              whiteSpace: 'nowrap',
                            }}
                          >
                            {p.region}
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              fontSize: 12,
                              color: C.muted,
                              whiteSpace: 'nowrap',
                            }}
                          >
                            {p.subregion || '-'}
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              fontSize: 12,
                              color: C.muted,
                              whiteSpace: 'nowrap',
                            }}
                          >
                            {p.country}
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              fontWeight: 700,
                              fontSize: 13,
                              whiteSpace: 'nowrap',
                            }}
                          >
                            {p.name}
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              fontSize: 12,
                              whiteSpace: 'nowrap',
                            }}
                          >
                            <span
                              style={{
                                background: C.faint,
                                borderRadius: 6,
                                padding: '2px 8px',
                                fontSize: 11,
                                color: C.muted,
                              }}
                            >
                              {p.type || '-'}
                            </span>
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              whiteSpace: 'nowrap',
                            }}
                          >
                            <span
                              style={{
                                color: tierColors[p.tier] || C.muted,
                                fontWeight: 700,
                                fontSize: 12,
                              }}
                            >
                              {p.tier}
                            </span>
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              fontSize: 12,
                              color: C.muted,
                              whiteSpace: 'nowrap',
                            }}
                          >
                            {p.accountManager || '-'}
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              textAlign: 'right',
                              fontFamily: 'monospace',
                              fontSize: 12,
                              fontWeight: 700,
                              color: C.accent,
                            }}
                          >
                            {fmtA(p.allocated)}
                          </td>
                          {['Q1', 'Q2', 'Q3', 'Q4'].map((q) => (
                            <td
                              key={q}
                              style={{
                                padding: '10px 14px',
                                textAlign: 'right',
                                fontFamily: 'monospace',
                                fontSize: 12,
                                color: qData[q] > 0 ? C.text : C.muted,
                                background: C.accentGlow + '50',
                              }}
                            >
                              {qData[q] > 0
                                ? `USD ${qData[q].toLocaleString('en-US')}`
                                : '-'}
                            </td>
                          ))}
                          <td
                            style={{
                              padding: '10px 14px',
                              textAlign: 'right',
                              fontFamily: 'monospace',
                              fontSize: 12,
                              fontWeight: 700,
                              color: C.success,
                            }}
                          >
                            {totalReq > 0
                              ? `USD ${totalReq.toLocaleString('en-US')}`
                              : '-'}
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              textAlign: 'right',
                              fontFamily: 'monospace',
                              fontSize: 12,
                              color: p.pipeline > 0 ? (C.cyan || C.accent) : C.muted,
                            }}
                          >
                            {p.pipeline > 0
                              ? `USD ${Math.round(p.pipeline).toLocaleString('en-US')}`
                              : '-'}
                          </td>
                          <td
                            style={{
                              padding: '10px 14px',
                              textAlign: 'right',
                              fontSize: 12,
                            }}
                          >
                            {totalReq > 0 ? (
                              (() => {
                                const target = totalReq * 30;
                                const pipe = p.pipeline || 0;
                                const ratio = totalReq > 0 && pipe > 0 ? pipe / totalReq : null;
                                const onTarget = ratio !== null && ratio >= 30;
                                return (
                                  <span style={{ fontSize: 11, color: ratio !== null ? (onTarget ? C.success : C.danger) : C.muted }}>
                                    {ratio !== null ? `1:${Math.round(ratio)}` : '—'}{' '}
                                    <span style={{ color: C.muted, fontWeight: 400 }}>
                                      / {(target).toLocaleString('en-US')}
                                    </span>
                                  </span>
                                );
                              })()
                            ) : (
                              <span style={{ color: C.muted }}>-</span>
                            )}
                          </td>
                          <td style={{ padding: '8px 10px', whiteSpace: 'nowrap' }}>
                            <div style={{ display: 'flex', gap: 6 }}>
                              <button
                                onClick={() => setEditPartner(p)}
                                style={{ background: C.accent + '18', color: C.accent, border: `1px solid ${C.accent}30`, borderRadius: 7, padding: '4px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                              >
                                Edit
                              </button>
                              <button
                                onClick={() => {
                                  if (window.confirm(
                                      (() => {
                                        const aReqs = requests.filter(r=>r.partner===p.name&&!['rejected','cancelled_by_partner'].includes(r.status)).length;
                                        const aClaims = claims.filter(c=>c.partner===p.name&&!['rejected','cancelled'].includes(c.status)).length;
                                        return aReqs > 0 || aClaims > 0
                                          ? `WARNING: ${p.name} has ${aReqs} active request(s) and ${aClaims} active claim(s). Deleting will leave orphan records. Continue anyway?`
                                          : `Delete ${p.name}? This cannot be undone.`;
                                      })()
                                    )) {
                                    setPartners(prev => prev.filter(x => x.id !== p.id));
                                    addHistory(`Partner ${p.name} deleted`, p.id, 'edit');
                                    toast(`${p.name} removed`);
                                    supa.delete('partners', 'id', p.id).catch(() => {});
                                  }
                                }}
                                style={{ background: C.danger + '18', color: C.danger, border: `1px solid ${C.danger}30`, borderRadius: 7, padding: '4px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                              >
                                Delete
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {filteredP.length > 0 &&
                      (() => {
                        const qTotals = { Q1: 0, Q2: 0, Q3: 0, Q4: 0 };
                        filteredP.forEach((p) => {
                          requests
                            .filter((r) => r.partner === p.name)
                            .forEach((r) => {
                              (r.items || []).forEach((item) => {
                                const q = item.fyQuarter || '';
                                const fyMatch =
                                  partnerFY === 'All' ||
                                  (Array.isArray(partnerFY)
                                    ? partnerFY.length === 0 ||
                                      partnerFY.some((fy) =>
                                        (item.fyHalf || '').includes(fy)
                                      )
                                    : (item.fyHalf || '').includes(partnerFY));
                                if (fyMatch && qTotals[q] !== undefined)
                                  qTotals[q] += toUSD(
                                    item.mdfRequest ||
                                      Math.round((item.amount || 0) * 0.5),
                                    item.localCurrency || 'EUR'
                                  );
                              });
                            });
                        });
                        const grandTotal = Object.values(qTotals).reduce(
                          (s, v) => s + v,
                          0
                        );
                        const totalAlloc2 = filteredP.reduce(
                          (s, p) => s + p.allocated,
                          0
                        );
                        const totalPipeline2 = filteredP.reduce(
                          (s, p) => s + (p.pipeline || 0),
                          0
                        );
                        return (
                          <tr
                            style={{
                              borderTop: `2px solid ${C.border}`,
                              background: C.faint,
                              fontWeight: 800,
                            }}
                          >
                            <td
                              colSpan={7}
                              style={{
                                padding: '10px 14px',
                                fontSize: 12,
                                fontWeight: 800,
                              }}
                            >
                              TOTAL ({filteredP.length} partners)
                            </td>
                            <td
                              style={{
                                padding: '10px 14px',
                                textAlign: 'right',
                                fontFamily: 'monospace',
                                fontSize: 12,
                                fontWeight: 800,
                                color: C.accent,
                              }}
                            >
                              {fmtA(totalAlloc2)}
                            </td>
                            {['Q1', 'Q2', 'Q3', 'Q4'].map((q) => (
                              <td
                                key={q}
                                style={{
                                  padding: '10px 14px',
                                  textAlign: 'right',
                                  fontFamily: 'monospace',
                                  fontSize: 12,
                                  fontWeight: 800,
                                  color: C.accent,
                                  background: C.accentGlow + '50',
                                }}
                              >
                                {qTotals[q] > 0
                                  ? `USD ${qTotals[q].toLocaleString('en-US')}`
                                  : '-'}
                              </td>
                            ))}
                            <td
                              style={{
                                padding: '10px 14px',
                                textAlign: 'right',
                                fontFamily: 'monospace',
                                fontSize: 12,
                                fontWeight: 800,
                                color: C.success,
                              }}
                            >{`USD ${grandTotal.toLocaleString('en-US')}`}</td>
                            <td
                              style={{
                                padding: '10px 14px',
                                textAlign: 'right',
                                fontFamily: 'monospace',
                                fontSize: 12,
                                fontWeight: 800,
                                color: totalPipeline2 > 0 ? (C.cyan || C.accent) : C.muted,
                              }}
                            >
                              {totalPipeline2 > 0
                                ? `USD ${Math.round(totalPipeline2).toLocaleString('en-US')}`
                                : '-'}
                            </td>
                            <td
                              style={{
                                padding: '10px 14px',
                                textAlign: 'right',
                                fontFamily: 'monospace',
                                fontSize: 12,
                                color: C.muted,
                              }}
                            >
                              {totalPipeline2 > 0 ? (
                                <span style={{ fontSize: 11, color: totalPipeline2 >= grandTotal * 30 ? C.success : C.danger, fontWeight: 700 }}>
                                  1:{Math.round(totalPipeline2 / (grandTotal || 1))}{' '}
                                  <span style={{ color: C.muted, fontWeight: 400 }}>
                                    (target 1:30)
                                  </span>
                                </span>
                              ) : (
                                <span style={{ fontSize: 11 }}>
                                  Target: USD{' '}
                                  {(grandTotal * 30).toLocaleString('en-US')}
                                </span>
                              )}
                            </td>
                          </tr>
                        );
                      })()}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Requests */}
        {tab === 'requests' && (
          <div
            style={{
              animation: 'slideIn 0.3s ease',
              padding: 28,
              overflowY: 'auto',
              height: '100%',
            }}
          >

            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginBottom: 14,
              }}
            >
              <div>
                <div
                  style={{
                    fontFamily: "'Syne',sans-serif",
                    fontSize: 26,
                    fontWeight: 800,
                    marginBottom: 4,
                  }}
                >
                  MDF Requests
                </div>
                <div style={{ color: C.muted, fontSize: 13 }}>
                  {filteredR.length} requests .{' '}
                  {filteredR.reduce((s, r) => s + (r.items || []).length, 0)}{' '}
                  activities
                </div>
              </div>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <span
                  style={{
                    fontFamily: "'Syne',sans-serif",
                    fontWeight: 800,
                    fontSize: 16,
                    color: C.accent,
                    letterSpacing: '0.05em',
                  }}
                >
                  OT
                </span>
                <div
                  style={{
                    width: 1,
                    height: 20,
                    background: C.border,
                    flexShrink: 0,
                  }}
                />
                <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                  {Object.keys(combinedBPSel).length > 0 && (
                    <button
                      onClick={() => setShowCombinedBP(true)}
                      style={{
                        background: C.success,
                        color: '#000',
                        border: 'none',
                        borderRadius: 10,
                        padding: '7px 14px',
                        fontWeight: 800,
                        fontSize: 12,
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        gap: 6,
                      }}
                    >
                      Generate BP ({Object.keys(combinedBPSel).length}{' '}
                      {Object.keys(combinedBPSel).length === 1
                        ? 'activity'
                        : 'activities'}
                      )
                    </button>
                  )}
                  <button
                    onClick={() =>
                      doExport(
                        'requests',
                        partners,
                        requests,
                        null,
                        filteredR,
                        null,
                        rate,
                        toUSD
                      )
                    }
                    style={{
                      background: C.faint,
                      color: C.text,
                      border: `1px solid ${C.border}`,
                      borderRadius: 10,
                      padding: '7px 14px',
                      fontWeight: 600,
                      fontSize: 12,
                      cursor: 'pointer',
                    }}
                  >
                    Excel
                  </button>
                  <button
                    onClick={() =>
                      doPPTExport(partners, requests, rate, currency, fmtA)
                    }
                    data-ppt-btn
                    style={{
                      background: C.accent,
                      color: '#fff',
                      border: 'none',
                      borderRadius: 10,
                      padding: '7px 14px',
                      fontWeight: 700,
                      fontSize: 12,
                      cursor: 'pointer',
                    }}
                  >
                    PPT
                  </button>
                </div>
                <button
                  onClick={() => setShowNewReq(true)}
                  style={{
                    background: C.accent,
                    color: '#fff',
                    border: 'none',
                    borderRadius: 10,
                    padding: '10px 20px',
                    fontWeight: 700,
                    fontSize: 13,
                    cursor: 'pointer',
                  }}
                >
                  + New Request
                </button>
              </div>
            </div>
            <div
              style={{
                display: 'flex',
                gap: 8,
                marginBottom: 14,
                flexWrap: 'wrap',
                alignItems: 'center',
              }}
            >
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search..."
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  color: C.text,
                  borderRadius: 10,
                  padding: '7px 12px',
                  fontSize: 13,
                  width: 160,
                }}
              />
              <MultiSelect
                value={filterReqPartner}
                onChange={setFilterReqPartner}
                placeholder="All Partners"
                options={[...new Set(requests.map((r) => r.partner))].sort()}
              />
              <MultiSelect
                value={filterReqStatus}
                onChange={setFilterReqStatus}
                placeholder="All Statuses"
                options={[
                  { value: 'request_submitted', label: 'Submitted' },
                  { value: 'approved', label: 'Approved' },
                  { value: 'sent_for_signature', label: 'Sent for Signature' },
                  { value: 'signed', label: 'Signed' },
                  { value: 'po_raised', label: 'PO Raised' },
                  { value: 'rejected', label: 'Rejected' },
                  {
                    value: 'cancelled_by_partner',
                    label: 'Cancelled by Partner',
                  },
                  { value: 'postponed', label: 'Postponed' },
                  { value: 'bp_sent', label: 'BP Sent to Partner' },
                ]}
              />
              <MultiSelect
                value={filterReqPartnerType}
                onChange={setFilterReqPartnerType}
                placeholder="All Types"
                options={PARTNER_TYPES}
              />
              <MultiSelect
                value={filterReqFY}
                onChange={(v) => {
                  setFilterReqFY(v);
                  setFilterReqQ([]);
                  setFilterReqMonth([]);
                }}
                placeholder="All FY"
                options={[
                  ...new Set(
                    requests
                      .flatMap((r) => (r.items || []).map((it) => it.fyHalf))
                      .filter(Boolean)
                  ),
                ].sort()}
              />
              <MultiSelect
                value={filterReqQ}
                onChange={(v) => {
                  setFilterReqQ(v);
                  setFilterReqMonth([]);
                }}
                placeholder="All Quarters"
                options={['Q1', 'Q2', 'Q3', 'Q4']}
              />
              <MultiSelect
                value={filterReqMonth}
                onChange={setFilterReqMonth}
                placeholder="All Months"
                options={[
                  'July',
                  'August',
                  'September',
                  'October',
                  'November',
                  'December',
                  'January',
                  'February',
                  'March',
                  'April',
                  'May',
                  'June',
                ]}
              />
              <MultiSelect
                value={filterCMM}
                onChange={setFilterCMM}
                placeholder="All CMM"
                options={TEAM_MEMBERS}
              />
              {(filterReqStatus.length > 0 ||
                filterReqPartnerType.length > 0 ||
                filterReqPartner.length > 0 ||
                filterReqFY.length > 0 ||
                filterReqQ.length > 0 ||
                filterReqMonth.length > 0 ||
                filterCMM.length > 0 ||
                search) && (
                <button
                  onClick={() => {
                    setFilterReqStatus([]);
                    setFilterReqPartnerType([]);
                    setFilterReqPartner([]);
                    setFilterReqFY([]);
                    setFilterReqQ([]);
                    setFilterReqMonth([]);
                    setFilterCMM([]);
                    setSearch('');
                  }}
                  style={{
                    background: 'transparent',
                    border: `1px solid ${C.border}`,
                    color: C.muted,
                    borderRadius: 8,
                    padding: '6px 12px',
                    fontSize: 12,
                    cursor: 'pointer',
                  }}
                >
                  Clear
                </button>
              )}
              <button
                onClick={() =>
                  setSortReqDir((d) => (d === 'desc' ? 'asc' : 'desc'))
                }
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  color: C.muted,
                  borderRadius: 8,
                  padding: '6px 12px',
                  fontSize: 12,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: 5,
                  whiteSpace: 'nowrap',
                }}
              >
                Date {sortReqDir === 'desc' ? '- Newest' : '- Oldest'}
              </button>
              {Object.keys(combinedBPSel).length === 0 &&
                filteredR.some((r) =>
                  (r.items || []).some((it) => it.itemStatus === 'approved')
                ) && (
                  <span
                    style={{
                      fontSize: 11,
                      color: C.muted,
                      fontStyle: 'italic',
                    }}
                  >
                    Tick approved activities to generate a combined BP (same
                    partner only)
                  </span>
                )}
              {Object.keys(combinedBPSel).length > 0 && (
                <span
                  style={{ fontSize: 11, color: C.accent, fontWeight: 600 }}
                >
                  Partner locked:{' '}
                  {Object.values(combinedBPSel)[0].request.partner}
                </span>
              )}
              {Object.keys(combinedBPSel).length > 0 && (
                <button
                  onClick={() => setCombinedBPSel({})}
                  style={{
                    background: 'transparent',
                    border: `1px solid ${C.border}`,
                    color: C.muted,
                    borderRadius: 8,
                    padding: '5px 10px',
                    fontSize: 11,
                    cursor: 'pointer',
                  }}
                >
                  Clear selection
                </button>
              )}
            </div>
            <div
              style={{
                background: C.surface,
                border: `1px solid ${C.border}`,
                borderRadius: 14,
                overflow: 'hidden',
              }}
            >
              <div style={{ overflowX: 'auto' }}>
                <table
                  style={{
                    width: '100%',
                    borderCollapse: 'collapse',
                    minWidth: 900,
                  }}
                >
                  <thead>
                    <tr
                      style={{
                        borderBottom: `2px solid ${C.border}`,
                        background: C.faint,
                      }}
                    >
                      {[
                        'Activity ID',
                        'Partner',
                        'Region',
                        'CMM',
                        'Activity Description',
                        'Allocadia ID',
                        'Campaign ID',
                        'MDF Request (USD)',
                        'Status',
                        'PO Number',
                        'Submitted',
                        '',
                      ].map((h, i) => (
                        <th
                          key={i}
                          style={{
                            padding: '10px 14px',
                            textAlign: i >= 6 && i <= 8 ? 'right' : 'left',
                            fontSize: 10,
                            color: isDark ? C.muted : '#00008b',
                            fontWeight: 700,
                            letterSpacing: '0.08em',
                            textTransform: 'uppercase',
                            whiteSpace: 'nowrap',
                          }}
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredR.length === 0 && (
                      <tr>
                        <td colSpan={13} style={{ padding: 40, textAlign: 'center', color: C.muted }}>
                          No requests found
                        </td>
                      </tr>
                    )}
                    {filteredR.length > 0 && (() => {
                      const statusColors = {
                        request_submitted: C.warning,
                        approved: C.success,
                        sent_for_signature: C.purple,
                        signed: C.cyan || '#06b6d4',
                        po_raised: C.accent,
                        rejected: C.danger,
                      };
                      const statusLabels = {
                        request_submitted: 'Submitted',
                        approved: 'Approved',
                        sent_for_signature: 'Sent for Sign.',
                        signed: 'Signed',
                        po_raised: 'PO Raised',
                        rejected: 'Rejected',
                      };

                       // Split into needs-attention, in-progress, and completed
                       const COMPLETED_STATUSES = ['po_raised', 'rejected', 'cancelled_by_partner'];
                       const needsAttention = sortedR.filter(r =>
                         !COMPLETED_STATUSES.includes(r.status) && (
                           !r.assignedTo ||
                           (r.items || []).some(it => ['cancelled_by_partner','postponed'].includes(it.itemStatus))
                         )
                       );
                       const inProgress = sortedR.filter(r =>
                         !COMPLETED_STATUSES.includes(r.status) &&
                         r.assignedTo &&
                         !(r.items || []).some(it => ['cancelled_by_partner','postponed'].includes(it.itemStatus))
                       );
                       const completed = sortedR.filter(r =>
                         COMPLETED_STATUSES.includes(r.status)
                       );

                      const renderRows = (reqs, baseIdx) => reqs.flatMap((r, rIdx) => {
                        const partner = partners.find((p) => p.name === r.partner) || {};
                        return (r.items || []).map((item, itemIdx) => {
                          const i = baseIdx + rIdx * 10 + itemIdx;
                          // Show request-level status when it has advanced beyond item approval
                          const advancedStatuses = ['sent_for_signature','signed','po_raised','rejected','cancelled_by_partner'];
                          const itemStatus = advancedStatuses.includes(r.status)
                            ? r.status
                            : (item.itemStatus || r.status);
                          const sColor = statusColors[itemStatus] || C.muted;
                          const mdfUSD = toUSD(
                            item.mdfRequest || Math.round((item.amount || 0) * 0.5),
                            item.localCurrency || 'EUR'
                          );
                          const isAlertItem = ['cancelled_by_partner','postponed'].includes(itemStatus) && !item.acknowledged;
                          const isUnassigned = !r.assignedTo;
                          const rowBg = isAlertItem ? `${C.danger}08`
                            : isUnassigned ? `${C.warning}08`
                            : i % 2 === 0 ? C.faint : C.surface;
                          const rowBorderLeft = isAlertItem ? `3px solid ${C.danger}`
                            : isUnassigned ? `3px solid ${C.warning}`
                            : '3px solid transparent';
                          const rowHoverBg = isAlertItem ? `${C.danger}15`
                            : isUnassigned ? `${C.warning}15`
                            : C.accentGlow;
                          return (
                            <tr
                              key={item.id}
                              onClick={() => setSelReq(r)}
                              style={{ borderBottom: `1px solid ${C.border}20`, background: rowBg, borderLeft: rowBorderLeft, cursor: 'pointer', transition: 'background 0.1s' }}
                              onMouseEnter={e => e.currentTarget.style.background = rowHoverBg}
                              onMouseLeave={e => e.currentTarget.style.background = rowBg}
                            >
                              <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontSize: 11, color: C.accent, whiteSpace: 'nowrap' }}>{r.id}</td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap', fontWeight: 600 }}>{r.partner}</td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap', color: C.muted, fontSize: 11 }}>{partner.region || '-'}</td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                {/* CMM assignment - dropdown for unassigned, badge for assigned */}
                                {!r.assignedTo ? (
                                  <select
                                    onClick={e => e.stopPropagation()}
                                    value=""
                                    onChange={e => {
                                      if (!e.target.value) return;
                                      const cmm = e.target.value;
                                      setRequests(prev => {
                                        const next = prev.map(req =>
                                          req.id !== r.id ? req : {
                                            ...req,
                                            assignedTo: cmm,
                                            items: req.items.map(it => ({ ...it, assignedTo: cmm })),
                                          }
                                        );
                                        const updated = next.find(req => req.id === r.id);
                                        if (updated) dbSaveRequest(updated).catch(() => {});
                                        return next;
                                      });
                                      addHistory(`${r.id} assigned to ${cmm}`, r.id, 'edit');
                                      toast(`Assigned to ${cmm}`);
                                    }}
                                    style={{ background: C.warning + '20', border: `1px solid ${C.warning}`, color: C.warning, borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                                  >
                                    <option value="">Assign CMM...</option>
                                    {TEAM_MEMBERS.map(m => <option key={m} value={m}>{m}</option>)}
                                  </select>
                                ) : (
                                  (() => {
                                    const col = r.assignedTo === 'Decio A.' ? C.accent : r.assignedTo === 'Kaila' ? C.success : C.purple;
                                    return (
                                      <span style={{ background: col + '20', color: col, border: `1px solid ${col}40`, borderRadius: 20, padding: '2px 10px', fontSize: 11, fontWeight: 700 }}>
                                        {r.assignedTo}
                                      </span>
                                    );
                                  })()
                                )}
                              </td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis' }} title={item.title}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.title || '-'}</span>
                                </div>
                              </td>
                              <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontSize: 11, color: C.muted, whiteSpace: 'nowrap' }}>{item.allocadiaId || '-'}</td>
                              <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontSize: 11, color: C.muted, whiteSpace: 'nowrap' }}>{item.campaignId || '-'}</td>
                              <td style={{ padding: '10px 14px', textAlign: 'right', fontFamily: 'monospace', fontSize: 11, fontWeight: 700, color: '#f59e0b', whiteSpace: 'nowrap' }}>{`USD ${mdfUSD.toLocaleString('en-US')}`}</td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                                  <span style={{ background: sColor + '18', color: sColor, border: `1px solid ${sColor}30`, borderRadius: 6, padding: '3px 8px', fontSize: 10, fontWeight: 700 }}>
                                    {statusLabels[itemStatus] || itemStatus}
                                  </span>
                                  {r.partnerNotified && (r.status === 'signed' || r.status === 'po_raised') && (
                                    <span title={`Signed BP emailed to partner${r.notifiedAt ? ' on ' + r.notifiedAt : ''}`}
                                      style={{ background: C.success + '18', color: C.success, border: `1px solid ${C.success}30`, borderRadius: 6, padding: '3px 8px', fontSize: 9, fontWeight: 700 }}>
                                      BP Sent
                                    </span>
                                  )}
                                </div>
                              </td>
                              <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontSize: 10, color: C.muted, whiteSpace: 'nowrap' }}>{r.poNumber || '-'}</td>
                              <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontSize: 10, color: C.muted, whiteSpace: 'nowrap' }}>{r.submitted}</td>
                              <td style={{ padding: '10px 14px' }}>
                                <button
                                  onClick={e => { e.stopPropagation(); setSelReq(r); }}
                                  style={{ background: C.accent, color: '#fff', border: 'none', borderRadius: 8, padding: '5px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
                                >Review</button>
                              </td>
                            </tr>
                          );
                        });
                      });

                      return (
                        <>
                          {/* === NEEDS ATTENTION === */}
                          {needsAttention.length > 0 && (
                            <tr style={{ background: C.danger + '10' }}>
                              <td colSpan={13} style={{ padding: '8px 14px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                  <div style={{ width: 8, height: 8, borderRadius: '50%', background: C.danger, flexShrink: 0 }} />
                                  <span style={{ fontSize: 11, fontWeight: 800, color: C.danger, letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                                    Needs Attention — {needsAttention.length} request{needsAttention.length !== 1 ? 's' : ''}
                                  </span>
                                  <span style={{ fontSize: 10, color: C.muted }}>
                                    {needsAttention.filter(r => !r.assignedTo).length > 0 && `${needsAttention.filter(r => !r.assignedTo).length} unassigned`}
                                    {needsAttention.filter(r => !r.assignedTo).length > 0 && needsAttention.some(r => (r.items||[]).some(it => ['cancelled_by_partner','postponed'].includes(it.itemStatus) && !it.acknowledged)) && ' · '}
                                    {needsAttention.some(r => (r.items||[]).some(it => it.itemStatus === 'cancelled_by_partner' && !it.acknowledged)) && `${cancelledByPartnerCnt} cancelled by partner`}
                                    {cancelledByPartnerCnt > 0 && postponedByPartnerCnt > 0 && ' · '}
                                    {postponedByPartnerCnt > 0 && `${postponedByPartnerCnt} postponed`}
                                  </span>
                                </div>
                              </td>
                            </tr>
                          )}
                          {renderRows(needsAttention, 0)}

                          {/* === IN PROGRESS === */}
                          {inProgress.length > 0 && (
                            <tr style={{ background: C.success + '10' }}>
                              <td colSpan={13} style={{ padding: '8px 14px', borderTop: needsAttention.length > 0 ? `2px solid ${C.border}` : 'none' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                  <div style={{ width: 8, height: 8, borderRadius: '50%', background: C.success, flexShrink: 0 }} />
                                  <span style={{ fontSize: 11, fontWeight: 800, color: C.success, letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                                    In Progress — {inProgress.length} request{inProgress.length !== 1 ? 's' : ''}
                                  </span>
                                </div>
                              </td>
                            </tr>
                          )}
                          {renderRows(inProgress, 1000)}

                          {/* === COMPLETED === */}
                          {completed.length > 0 && (
                            <tr style={{ background: C.faint }}>
                              <td colSpan={13} style={{ padding: '8px 14px', borderTop: `2px solid ${C.border}` }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                  <div style={{ width: 8, height: 8, borderRadius: '50%', background: C.muted, flexShrink: 0 }} />
                                  <span style={{ fontSize: 11, fontWeight: 800, color: C.muted, letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                                    Completed — {completed.length} request{completed.length !== 1 ? 's' : ''}
                                  </span>
                                </div>
                              </td>
                            </tr>
                          )}
                          {renderRows(completed, 2000)}

                          {/* Totals footer */}
                          {filteredR.length > 0 && (() => {
                            const allItems = filteredR.flatMap(r => r.items || []);
                            const totUSD = allItems.reduce((s, it) => s + toUSD(it.mdfRequest || Math.round((it.amount || 0) * 0.5), it.localCurrency || 'EUR'), 0);
                            return (
                              <tr style={{ borderTop: `2px solid ${C.border}`, background: C.faint }}>
                                <td colSpan={10} style={{ padding: '10px 14px', fontWeight: 800, fontSize: 12 }}>
                                  TOTAL — {filteredR.length} requests · {allItems.length} activities
                                </td>
                                <td style={{ padding: '10px 14px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 800, color: '#f59e0b' }}>
                                  {`USD ${totUSD.toLocaleString('en-US')}`}
                                </td>
                                <td colSpan={4} />
                              </tr>
                            );
                          })()}
                        </>
                      );
                    })()}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Claims */}
        {tab === 'claims' && (
          <ClaimsTab
            claims={claims}
            setClaims={(updater) => {
              setClaims(prev => {
                const next = typeof updater === 'function' ? updater(prev) : updater;
                const changed = next.filter(c => {
                  const old = prev.find(p => p.id === c.id);
                  return !old || JSON.stringify(old) !== JSON.stringify(c);
                });
                changed.forEach(c => dbSaveClaim(c).catch(() => {}));
                return next;
              });
            }}
            partners={partners}
            requests={requests}
            addHistory={addHistory}
            toast={toast}
            toUSD={toUSD}
            fmtA={fmtA}
          />
        )}

        {/* History */}
        {tab === 'history' && (
          <div
            style={{
              animation: 'slideIn 0.3s ease',
              padding: 28,
              overflowY: 'auto',
              height: '100%',
            }}
          >
            <div
              style={{
                fontFamily: "'Syne',sans-serif",
                fontSize: 26,
                fontWeight: 800,
                marginBottom: 20,
              }}
            >
              Activity History
            </div>
            {history.length === 0 && (
              <div style={{ color: C.muted, fontSize: 13 }}>
                No activity yet. Changes will appear here.
              </div>
            )}
            {history.map((h) => (
              <div
                key={h.id}
                style={{
                  background: C.surface,
                  border: `1px solid ${C.border}`,
                  borderRadius: 10,
                  padding: '12px 16px',
                  marginBottom: 8,
                  display: 'flex',
                  gap: 12,
                  alignItems: 'center',
                }}
              >
                <div
                  style={{
                    width: 8,
                    height: 8,
                    borderRadius: '50%',
                    background:
                      h.type === 'approve'
                        ? C.success
                        : h.type === 'reject'
                        ? C.danger
                        : h.type === 'create'
                        ? C.accent
                        : C.muted,
                    flexShrink: 0,
                  }}
                />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600 }}>
                    {h.action}
                  </div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>
                    {h.user} . {h.ts}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Modals */}
      {selReq && (
        <RequestReviewModal
          request={selReq}
          partners={partners}
          requests={requests}
          setRequests={(updater) => {
              setRequests(prev => {
                const next = typeof updater === 'function' ? updater(prev) : updater;
                // Save changed requests to DB
                const changed = next.filter(r => {
                  const old = prev.find(p => p.id === r.id);
                  return !old || JSON.stringify(old) !== JSON.stringify(r);
                });
                changed.forEach(r => dbSaveRequest(r).catch(() => {}));
                return next;
              });
            }}
          selectedItems={selectedItems}
          setSelectedItems={setSelectedItems}
          addHistory={addHistory}
          toast={toast}
          openPlanModal={openPlanModal}
          setPlanModal={setPlanModal}
          onClose={() => {
            setSelReq(null);
            setSelectedItems({});
          }}
          fmtA={fmtA}
          currency={currency}
        />
      )}
      {showCombinedBP &&
        (() => {
          const selItems = Object.values(combinedBPSel);
          const partnerName = selItems[0]?.request?.partner || '';
          const partnerRecord = partners.find(
            (p) => p.name === partnerName
          ) || { name: partnerName };
          const firstReq = selItems[0]?.request || {};
          const partner = {
            ...partnerRecord,
            contactName:
              firstReq.partnerContact || partnerRecord.contactName || '',
            contactEmail:
              firstReq.partnerEmail || partnerRecord.contactEmail || '',
            accountManager:
              firstReq.partnerManager || partnerRecord.accountManager || '',
            type: firstReq.partnerType || partnerRecord.type || '',
            tier: firstReq.partnerTier || partnerRecord.tier || '',
          };
          const activities = selItems.map(({ item }) => ({
            ...item,
            selected: true,
          }));
          return (
            <PlanModal
              modal={{ partner, activities, reqId: null }}
              partners={partners}
              currentUser={currentUser}
              generateFn={generateMDFBusinessPlan}
              onUpdateStatus={(_reqId, _po, allocIds, campIds) => {
                // Save allocadia/campaign IDs back to each item's request
                // Single setRequests call processes ALL items at once (fixes race condition)
                setRequests((prev) => {
                  const updatedMap = {};
                  selItems.forEach(({ item, request }) => {
                    if (!updatedMap[request.id]) updatedMap[request.id] = {};
                    updatedMap[request.id][item.id] = {
                      allocadiaId: allocIds?.[item.id] || item.allocadiaId || '',
                      campaignId: campIds?.[item.id] || item.campaignId || '',
                    };
                  });
                  return prev.map((r) => {
                    if (!updatedMap[r.id]) return r;
                    return {
                      ...r,
                      items: r.items.map((it) =>
                        updatedMap[r.id][it.id]
                          ? { ...it, ...updatedMap[r.id][it.id] }
                          : it
                      ),
                    };
                  });
                });
                addHistory(
                  `Combined BP generated for ${partnerName} (${selItems.length} activities)`,
                  null,
                  'approve'
                );
                toast('Combined Business Plan generated!');
                setCombinedBPSel({});
              }}
              onClose={() => setShowCombinedBP(false)}
            />
          );
        })()}
      {planModal && (
        <PlanModal
          modal={planModal}
          partners={partners}
          currentUser={currentUser}
          generateFn={generateMDFBusinessPlan}
          onUpdateStatus={(reqId, po, allocIds, campIds) => {
            if (reqId) {
              const now = new Date().toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
              setRequests((prev) =>
                prev.map((r) => {
                  if (r.id !== reqId) return r;
                  const updatedItems = (r.items || []).map((it) => ({
                    ...it,
                    allocadiaId:
                      allocIds && allocIds[it.id]
                        ? allocIds[it.id]
                        : it.allocadiaId || '',
                    campaignId:
                      campIds && campIds[it.id]
                        ? campIds[it.id]
                        : it.campaignId || '',
                  }));
                  return { ...r, status: 'sent_for_signature', bpGeneratedAt: now, items: updatedItems };
                })
              );
              addHistory(
                `BP generated & sent for signature - PO: ${po || 'N/A'}`,
                reqId,
                'approve'
              );
            }
            toast(
              'Business Plan generated! Status updated to Sent for Signature.'
            );
          }}
          onClose={() => setPlanModal(null)}
        />
      )}
      {showNewReq && (
        <NewRequestModal
          partners={partners}
          currentUser={currentUser}
          portalPartner={portalPartner}
          isPartnerView={isPartnerView}
          onAdd={(req) => {
            setRequests((p) => [req, ...p]);
            addHistory(`New request ${req.id} from ${req.partner}`, req.id, 'create');
            toast('Request submitted!');
            dbSaveRequest(req).catch(() => {});
          }}
          onClose={() => setShowNewReq(false)}
        />
      )}
      {editPartner && (
          <EditPartnerModal
            key={editPartner._new ? 'new' : editPartner.id}
            partner={editPartner._new ? null : editPartner}
            onSave={(updated) => {
              if (editPartner._new) {
                setPartners(prev => [...prev, updated]);
                addHistory(`Partner ${updated.name} added`, updated.id, 'create');
                toast(`${updated.name} added`);
              } else {
                setPartners(prev => prev.map(p => p.id === updated.id ? updated : p));
                addHistory(`Partner ${updated.name} updated`, updated.id, 'edit');
                toast(`${updated.name} saved`);
              }
              // Save to Supabase so partner portal sees the change immediately
              dbSavePartner(updated).catch(() => {});
              setEditPartner(null);
            }}
            onClose={() => setEditPartner(null)}
          />
        )}
      {showPipelineImport && (
          <PipelineImportModal
            C={C}
            onImport={(data, count) => {
              setPipelineData(prev => ({ ...prev, ...data }));
              toast(`Pipeline imported: ${count} campaign IDs updated`);
              setTimeout(() => setShowPipelineImport(false), 1500);
            }}
            onClose={() => setShowPipelineImport(false)}
          />
        )}
      {showImport && (
        <ImportModal
          onClose={() => setShowImport(false)}
          onImport={(pts, reqs) => {
            const existingCount = partners.length;
            const importedNames = new Set(pts.map(p => p.name));
            const referencedNames = [...new Set([
              ...requests.map(r => r.partner),
              ...claims.map(c => c.partner),
            ])];
            const orphaned = referencedNames.filter(n => !importedNames.has(n));
            let msg = existingCount > 0
              ? `This will REPLACE all ${existingCount} existing partners with ${pts.length} imported partners.`
              : null;
            if (orphaned.length > 0) {
              msg = (msg || '') + `\n\nWARNING: ${orphaned.length} partner(s) have active requests/claims but are NOT in the import file:\n${orphaned.slice(0,5).join(', ')}${orphaned.length > 5 ? ` and ${orphaned.length - 5} more` : ''}.\nThose records will become orphaned.`;
            }
            if (msg && !window.confirm((msg || '') + '\n\nContinue?')) return;
            setPartners(pts);
            if (reqs) setRequests(reqs);
            toast(`Imported ${pts.length} partners`);
            // Save all imported partners to Supabase
            // BREAK-008 FIX: batch saves 10 at a time to avoid Supabase rate limits
            (async () => {
              for (let i = 0; i < pts.length; i += 10) {
                await Promise.all(pts.slice(i, i + 10).map(p => dbSavePartner(p).catch(() => {})));
              }
            })();
            addHistory(
              `Imported ${pts.length} partners via Excel (replaced ${existingCount})`,
              null,
              'create'
            );
          }}
        />
      )}
    </div>
  );
}
