// Campagnes: the monthly offer catalogue (about 50 articles for next month),
// the awareness campaign before it, and the agreements with the growers.
// Server: shadow-app/server/campaigns.js (/api/campaigns...). Articles come
// from our purchase invoices; the steady conditions live on reusable
// master article cards.
import React, { Fragment, useCallback, useEffect, useMemo, useState } from "react";
import { LibraryTab, MakeTab, MediaSettings, PostsTab, SocialSettings, ThemesEditor } from "./CampaignMedia.jsx";

const STAGE_LABELS = {
  selecting: "Choosing articles",
  growers: "Grower deals",
  content: "Making content",
  running: "Campaign running",
  offer: "Offer month",
  evaluating: "Evaluating",
  closed: "Closed",
  cancelled: "Cancelled",
};
const ITEM_LABELS = {
  shortlist: "Shortlist",
  selected: "Selected",
  deal_ok: "Grower deal OK",
  content_ok: "Content ready",
  live: "Live",
  evaluated: "Evaluated",
  dropped: "Dropped",
};
const DEAL_LABELS = {
  idea: "Idea",
  contacted: "Contacted",
  negotiating: "Negotiating",
  agreed: "Agreed",
  confirmed: "Confirmed",
  delivering: "Delivering",
  closed: "Closed",
  declined: "Declined",
};
const DEADLINE_LABELS = { articles: "Articles chosen", growers: "Grower deals done", content: "Content ready / start", evaluation: "Evaluation" };
const CONDITION_FIELDS = [
  ["length_cm", "Length (cm)", "number"],
  ["height_cm", "Plant height (cm)", "number"],
  ["pot_cm", "Pot size (cm)", "number"],
  ["stems_per_bunch", "Stems per bunch", "number"],
  ["packages", "Packages", "number"],
  ["per_package", "Per package", "number"],
  ["pieces", "Pieces", "number"],
  ["unit", "Unit (stems / pots / bunches)", "text"],
  ["fust", "Fust / packaging", "text"],
  ["quality", "Quality", "text"],
  ["colour", "Colour", "text"],
  ["origin", "Origin", "text"],
  ["ref_price", "Reference price (€)", "number"],
  ["min_order", "Minimum order", "number"],
  ["delivery_days", "Delivery days", "text"],
];

function monthLabel(month) {
  if (!/^\d{4}-\d{2}$/.test(month || "")) return month || "";
  const [year, mon] = month.split("-").map(Number);
  return new Date(Date.UTC(year, mon - 1, 1)).toLocaleDateString("en-GB", { month: "long", year: "numeric", timeZone: "UTC" });
}

function nextMonths(count) {
  const now = new Date();
  return Array.from({ length: count }, (_, index) => {
    const date = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + index, 1));
    return date.toISOString().slice(0, 7);
  });
}

function day(value) {
  if (!value) return "";
  const text = String(value).slice(0, 10);
  const [year, month, dayOfMonth] = text.split("-");
  return year && month && dayOfMonth ? `${dayOfMonth}-${month}-${year}` : text;
}

function money(value) {
  return value === null || value === undefined || value === "" ? "" : `€ ${Number(value).toFixed(2)}`;
}

function amount(value) {
  return value === null || value === undefined || value === "" ? "" : Number(value).toLocaleString("nl-NL");
}

// "55 cm · 10 × 50 = 500" -- the short line under an article's name.
function conditionsSummary(conditions = {}) {
  const parts = [];
  if (conditions.length_cm) parts.push(`${conditions.length_cm} cm`);
  if (conditions.pot_cm) parts.push(`pot ${conditions.pot_cm}`);
  if (conditions.height_cm) parts.push(`h ${conditions.height_cm}`);
  if (conditions.packages && conditions.per_package) parts.push(`${conditions.packages} × ${conditions.per_package} = ${amount(conditions.pieces ?? conditions.packages * conditions.per_package)}`);
  else if (conditions.per_package) parts.push(`${conditions.per_package} per package`);
  if (conditions.quality) parts.push(conditions.quality);
  return parts.join(" · ");
}

function Message({ error, notice }) {
  if (error) return <p className="cp-msg cp-error">{error}</p>;
  if (notice) return <p className="cp-msg cp-notice">{notice}</p>;
  return null;
}

// ---------------------------------------------------------------------------

export default function CampaignsPage({ apiJson }) {
  const [tab, setTab] = useState("planning");
  const [meta, setMeta] = useState(null);
  const [campaignId, setCampaignId] = useState(null);
  const [error, setError] = useState("");

  useEffect(() => {
    apiJson("/api/campaigns/settings").then(setMeta).catch((err) => setError(err.message));
  }, [apiJson]);

  const canManage = Boolean(meta?.can_manage);
  const openCampaign = useCallback((id) => {
    setCampaignId(id);
    setTab("catalogue");
  }, []);

  const tabs = [
    ["planning", "Planning"],
    ["catalogue", "Catalogue"],
    ["growers", "Growers"],
    ["articles", "Articles"],
    ["make", "Make"],
    ["library", "Library"],
    ["posts", "Posts"],
    ...(canManage ? [["settings", "Settings"]] : []),
  ];

  return (
    <div className="cp-page">
      <nav className="cp-tabs" aria-label="Campagnes">
        {tabs.map(([key, label]) => (
          <button key={key} type="button" className={tab === key ? "primary" : ""} onClick={() => setTab(key)}>{label}</button>
        ))}
      </nav>
      <Message error={error} />
      {meta && tab === "planning" && <PlanningTab apiJson={apiJson} meta={meta} canManage={canManage} onOpen={openCampaign} />}
      {meta && tab === "catalogue" && <CatalogueTab apiJson={apiJson} meta={meta} canManage={canManage} campaignId={campaignId} setCampaignId={setCampaignId} />}
      {meta && tab === "growers" && <GrowersTab apiJson={apiJson} meta={meta} canManage={canManage} campaignId={campaignId} />}
      {meta && tab === "articles" && <ArticlesTab apiJson={apiJson} meta={meta} canManage={canManage} />}
      {meta && tab === "make" && <MakeTab apiJson={apiJson} canManage={canManage} campaignId={campaignId} />}
      {meta && tab === "library" && <LibraryTab apiJson={apiJson} canManage={canManage} />}
      {meta && tab === "posts" && <PostsTab apiJson={apiJson} canManage={canManage} />}
      {meta && tab === "settings" && canManage && (
        <>
          <SettingsTab apiJson={apiJson} meta={meta} onSaved={(settings) => setMeta({ ...meta, settings })} />
          <MediaSettings apiJson={apiJson} />
          <SocialSettings apiJson={apiJson} />
          <ThemesEditor apiJson={apiJson} />
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Planning: the offer months on a timeline.

function PlanningTab({ apiJson, meta, canManage, onOpen }) {
  const [campaigns, setCampaigns] = useState(null);
  const [month, setMonth] = useState(nextMonths(3)[2]);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const target = meta.settings?.target_articles || 50;

  const load = useCallback(() => {
    apiJson("/api/campaigns").then((data) => setCampaigns(data.campaigns)).catch((err) => setError(err.message));
  }, [apiJson]);
  useEffect(load, [load]);

  async function create() {
    setBusy(true);
    setError("");
    try {
      const data = await apiJson("/api/campaigns", { method: "POST", body: JSON.stringify({ offer_month: month }) });
      load();
      onOpen(data.campaign.id);
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }

  const today = new Date().toISOString().slice(0, 10);
  return (
    <section className="cp-section">
      {canManage && (
        <div className="info-panel cp-new">
          <strong>New offer month</strong>
          <select value={month} onChange={(event) => setMonth(event.target.value)}>
            {nextMonths(8).map((option) => <option key={option} value={option}>{monthLabel(option)}</option>)}
          </select>
          <button type="button" className="primary" onClick={create} disabled={busy}>Plan this month</button>
          <span className="cp-hint">The campaign runs in the month{meta.settings?.lead_months > 1 ? "s" : ""} before; the deadlines can be changed after.</span>
        </div>
      )}
      <Message error={error} />
      {!campaigns ? <p>Loading…</p> : !campaigns.length ? <p className="cp-hint">No offer months planned yet.</p> : (
        <div className="cp-campaign-grid">
          {campaigns.map((campaign) => {
            const counts = campaign.counts || {};
            const active = Object.entries(counts).filter(([status]) => status !== "dropped").reduce((sum, [, n]) => sum + n, 0);
            const selected = active - (counts.shortlist || 0);
            const deals = campaign.deal_counts || {};
            const dealsOk = (deals.agreed || 0) + (deals.confirmed || 0) + (deals.delivering || 0) + (deals.closed || 0);
            const dealsTotal = Object.values(deals).reduce((sum, n) => sum + n, 0);
            return (
              <button key={campaign.id} type="button" className="fust-tile cp-campaign" onClick={() => onOpen(campaign.id)}>
                <span className="cp-campaign-head">
                  <strong>{monthLabel(campaign.offer_month)}</strong>
                  <span className={`cp-stage cp-stage-${campaign.stage}`}>{STAGE_LABELS[campaign.stage] || campaign.stage}</span>
                </span>
                <span className="cp-hint">Campaign {day(campaign.campaign_start)} – {day(campaign.campaign_end)}</span>
                <span className="cp-bar" title={`${selected} selected, ${counts.shortlist || 0} on the shortlist, target ${target}`}>
                  <span className="cp-bar-fill" style={{ width: `${Math.min(100, (selected / target) * 100)}%` }} />
                  <span className="cp-bar-short" style={{ width: `${Math.min(100, ((counts.shortlist || 0) / target) * 100)}%` }} />
                </span>
                <span>{selected} / {target} articles{counts.shortlist ? ` (+${counts.shortlist} shortlist)` : ""}</span>
                <span>{dealsOk} / {dealsTotal} grower deals agreed</span>
                <span className="cp-deadlines">
                  {Object.entries(campaign.deadlines || {}).map(([key, value]) => (
                    <span key={key} className={value < today ? "cp-past" : value <= addDaysIso(today, 7) ? "cp-soon" : ""}>
                      {DEADLINE_LABELS[key] || key}: {day(value)}
                    </span>
                  ))}
                </span>
              </button>
            );
          })}
        </div>
      )}
    </section>
  );
}

function addDaysIso(isoDay, days) {
  const date = new Date(`${isoDay}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

// ---------------------------------------------------------------------------
// Catalogue: the articles of one offer month.

function CatalogueTab({ apiJson, meta, canManage, campaignId, setCampaignId }) {
  const [campaigns, setCampaigns] = useState([]);
  const [data, setData] = useState(null);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [adding, setAdding] = useState(false);
  const [openItem, setOpenItem] = useState(null);
  const [showDropped, setShowDropped] = useState(false);

  useEffect(() => {
    apiJson("/api/campaigns").then((result) => {
      setCampaigns(result.campaigns);
      if (!campaignId && result.campaigns[0]) setCampaignId(result.campaigns[0].id);
    }).catch((err) => setError(err.message));
  }, [apiJson, campaignId, setCampaignId]);

  const load = useCallback(() => {
    if (!campaignId) return;
    apiJson(`/api/campaigns/${campaignId}`).then(setData).catch((err) => setError(err.message));
  }, [apiJson, campaignId]);
  useEffect(load, [load]);

  async function patchCampaign(fields) {
    try {
      await apiJson(`/api/campaigns/${campaignId}`, { method: "PATCH", body: JSON.stringify(fields) });
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  async function patchItem(item, fields) {
    try {
      await apiJson(`/api/campaigns/items/${item.id}`, { method: "PATCH", body: JSON.stringify(fields) });
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  async function removeItem(item) {
    if (!window.confirm(`Take "${item.article.name}" out of this offer month?`)) return;
    try {
      const result = await apiJson(`/api/campaigns/items/${item.id}`, { method: "DELETE" });
      setNotice(result.dropped ? "It has grower deals, so it is marked Dropped (the deals stay)." : "Taken out.");
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  async function copyFrom(otherId) {
    if (!otherId) return;
    try {
      const result = await apiJson(`/api/campaigns/${campaignId}/copy-from/${otherId}`, { method: "POST" });
      setNotice(`${result.added} article(s) added to the shortlist.`);
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  if (!campaigns.length) return <p className="cp-hint">Plan an offer month first (Planning).</p>;
  const campaign = data?.campaign;
  const items = (data?.items || []).filter((item) => showDropped || item.status !== "dropped");
  const target = meta.settings?.target_articles || 50;
  const activeCount = (data?.items || []).filter((item) => item.status !== "dropped").length;

  return (
    <section className="cp-section">
      <div className="cp-row">
        <select value={campaignId || ""} onChange={(event) => setCampaignId(Number(event.target.value))}>
          {campaigns.map((option) => <option key={option.id} value={option.id}>{monthLabel(option.offer_month)} — {STAGE_LABELS[option.stage]}</option>)}
        </select>
        {campaign && <span className="cp-hint">{activeCount} / {target} articles</span>}
        {canManage && campaign && <button type="button" className="primary" onClick={() => setAdding(true)}>+ Add articles</button>}
        {canManage && campaign && (
          <select value="" onChange={(event) => copyFrom(Number(event.target.value))}>
            <option value="">Copy articles from…</option>
            {campaigns.filter((option) => option.id !== campaignId).map((option) => <option key={option.id} value={option.id}>{monthLabel(option.offer_month)}</option>)}
          </select>
        )}
        <label className="cp-check"><input type="checkbox" checked={showDropped} onChange={(event) => setShowDropped(event.target.checked)} /> Show dropped</label>
      </div>
      <Message error={error} notice={notice} />
      {campaign && <CampaignHeader campaign={campaign} meta={meta} canManage={canManage} onPatch={patchCampaign} />}
      {campaign && (
        <div className="info-panel cp-table-wrap">
          <table className="ik-table cp-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Article</th>
                <th>Hoofdgroep</th>
                <th>VBN</th>
                <th>Conditions</th>
                <th>Offer price</th>
                <th>Planned qty</th>
                <th>Growers</th>
                <th>Status</th>
                {canManage && <th />}
              </tr>
            </thead>
            <tbody>
              {items.map((item, index) => (
                <tr key={item.id} className={item.status === "dropped" ? "cp-dropped" : ""}>
                  <td>{index + 1}</td>
                  <td>
                    <button type="button" className="link-button cp-name" onClick={() => setOpenItem(item)}>{item.article.name}</button>
                    {(item.article.name_fr || item.article.name_en) && <div className="cp-sub">FR {item.article.name_fr || "—"} · EN {item.article.name_en || "—"}</div>}
                  </td>
                  <td>{item.article.hoofdgroep}</td>
                  <td>{(item.article.vbn_codes || []).join(", ")}</td>
                  <td className="cp-wrap">{conditionsSummary(item.article.conditions)}</td>
                  <td>
                    {canManage ? <NumberCell value={item.offer_price} onSave={(value) => patchItem(item, { offer_price: value })} /> : money(item.offer_price)}
                  </td>
                  <td>
                    {canManage ? <NumberCell value={item.planned_qty} onSave={(value) => patchItem(item, { planned_qty: value })} /> : amount(item.planned_qty)}
                  </td>
                  <td className="cp-wrap">
                    {item.deals.length ? item.deals.map((deal) => (
                      <span key={deal.id} className={`cp-chip cp-deal-${deal.status}`} title={deal.next_action ? `Next: ${deal.next_action} (${day(deal.next_action_at)})` : ""}>
                        {deal.grower_name || deal.grower_code} · {DEAL_LABELS[deal.status]}{deal.agreed_price !== null && deal.agreed_price !== undefined ? ` · ${money(deal.agreed_price)}` : ""}
                      </span>
                    )) : <span className="cp-hint">no deal yet</span>}
                  </td>
                  <td>
                    {canManage ? (
                      <select value={item.status} onChange={(event) => patchItem(item, { status: event.target.value })}>
                        {meta.item_statuses.map((status) => <option key={status} value={status}>{ITEM_LABELS[status] || status}</option>)}
                      </select>
                    ) : ITEM_LABELS[item.status]}
                  </td>
                  {canManage && <td><button type="button" className="link-button" onClick={() => removeItem(item)} title="Take out of this month">✕</button></td>}
                </tr>
              ))}
              {!items.length && <tr><td colSpan={10} className="cp-hint">No articles yet — use “+ Add articles”.</td></tr>}
            </tbody>
          </table>
        </div>
      )}
      {adding && campaign && (
        <AddArticlesDialog apiJson={apiJson} campaign={campaign} existing={new Set((data?.items || []).map((item) => item.article_id))}
          onClose={() => { setAdding(false); load(); }} />
      )}
      {openItem && (
        <ItemDialog apiJson={apiJson} meta={meta} canManage={canManage} item={openItem} campaign={campaign}
          onClose={() => { setOpenItem(null); load(); }} />
      )}
    </section>
  );
}

function NumberCell({ value, onSave }) {
  const [text, setText] = useState(value ?? "");
  useEffect(() => setText(value ?? ""), [value]);
  return (
    <input className="cp-num" value={text} inputMode="decimal" onChange={(event) => setText(event.target.value)}
      onBlur={() => { if (String(text) !== String(value ?? "")) onSave(text === "" ? null : text); }} />
  );
}

function CampaignHeader({ campaign, meta, canManage, onPatch }) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState(campaign);
  useEffect(() => setDraft(campaign), [campaign]);
  if (!open) {
    return (
      <div className="info-panel cp-head">
        <strong>{monthLabel(campaign.offer_month)}</strong>
        <span className={`cp-stage cp-stage-${campaign.stage}`}>{STAGE_LABELS[campaign.stage]}</span>
        <span>Campaign {day(campaign.campaign_start)} – {day(campaign.campaign_end)}</span>
        {Object.entries(campaign.deadlines || {}).map(([key, value]) => <span key={key} className="cp-hint">{DEADLINE_LABELS[key] || key}: {day(value)}</span>)}
        {(campaign.title_fr || campaign.title_en) && <span className="cp-sub">FR “{campaign.title_fr}” · EN “{campaign.title_en}”</span>}
        {canManage && <button type="button" onClick={() => setOpen(true)}>Edit</button>}
      </div>
    );
  }
  const set = (key, value) => setDraft({ ...draft, [key]: value });
  return (
    <div className="info-panel cp-form">
      <label>Stage
        <select value={draft.stage} onChange={(event) => set("stage", event.target.value)}>
          {meta.stages.map((stage) => <option key={stage} value={stage}>{STAGE_LABELS[stage] || stage}</option>)}
        </select>
      </label>
      <label>Campaign from<input type="date" value={draft.campaign_start || ""} onChange={(event) => set("campaign_start", event.target.value)} /></label>
      <label>Campaign until<input type="date" value={draft.campaign_end || ""} onChange={(event) => set("campaign_end", event.target.value)} /></label>
      {Object.keys(DEADLINE_LABELS).map((key) => (
        <label key={key}>{DEADLINE_LABELS[key]}
          <input type="date" value={draft.deadlines?.[key] || ""} onChange={(event) => set("deadlines", { ...draft.deadlines, [key]: event.target.value })} />
        </label>
      ))}
      <label>Title (FR)<input value={draft.title_fr || ""} onChange={(event) => set("title_fr", event.target.value)} /></label>
      <label>Title (EN)<input value={draft.title_en || ""} onChange={(event) => set("title_en", event.target.value)} /></label>
      <label className="cp-wide">Notes<textarea rows={2} value={draft.notes || ""} onChange={(event) => set("notes", event.target.value)} /></label>
      <div className="cp-row cp-wide">
        <button type="button" className="primary" onClick={() => { onPatch({ stage: draft.stage, campaign_start: draft.campaign_start, campaign_end: draft.campaign_end,
          deadlines: draft.deadlines, title_fr: draft.title_fr, title_en: draft.title_en, notes: draft.notes }); setOpen(false); }}>Save</button>
        <button type="button" onClick={() => setOpen(false)}>Cancel</button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Adding articles: from our invoices, our own article cards, or the product
// export.

function invoiceRowToArticle(row) {
  const features = row.features || {};
  const length = Number(features.S20);
  return {
    name: row.description,
    vbn_codes: row.vbn_code ? [row.vbn_code] : [],
    hoofdgroep: row.hoofdgroep || "",
    conditions: {
      ...(Number.isFinite(length) && length > 0 ? { length_cm: length } : {}),
      ...(row.packages ? { packages: row.packages } : {}),
      ...(row.per_package ? { per_package: row.per_package } : {}),
      ...(row.last_price ? { ref_price: row.last_price } : {}),
      ...(Object.keys(features).length ? { features } : {}),
    },
    growers: [{
      name: row.supplier_name,
      code: row.grower_code || "",
      gln: row.supplier_gln || "",
      fh: row.supplier_fh_number || "",
      last_price: row.last_price,
      last_date: row.last_date,
      pieces: row.pieces,
    }],
  };
}

function AddArticlesDialog({ apiJson, campaign, existing, onClose }) {
  const [source, setSource] = useState("invoices");
  const [query, setQuery] = useState({ product: "", grower: "", vbn: "", months: 18 });
  const [rows, setRows] = useState([]);
  const [articles, setArticles] = useState([]);
  const [products, setProducts] = useState([]);
  const [error, setError] = useState("");
  const [added, setAdded] = useState([]);
  const [busy, setBusy] = useState(false);

  async function search() {
    setBusy(true);
    setError("");
    try {
      if (source === "invoices") {
        const params = new URLSearchParams({ product: query.product, grower: query.grower, vbn: query.vbn, months: String(query.months) });
        setRows((await apiJson(`/api/campaigns/source/invoices?${params}`)).rows);
      } else if (source === "articles") {
        setArticles((await apiJson(`/api/campaigns/articles?q=${encodeURIComponent(query.product)}`)).articles);
      } else {
        setProducts((await apiJson(`/api/campaigns/source/products?q=${encodeURIComponent(query.product)}`)).rows);
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }

  async function add(body, key) {
    setError("");
    try {
      await apiJson(`/api/campaigns/${campaign.id}/items`, { method: "POST", body: JSON.stringify(body) });
      setAdded((list) => [...list, key]);
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <div className="cp-overlay" role="dialog" aria-label="Add articles">
      <div className="cp-dialog">
        <div className="cp-row cp-dialog-head">
          <strong>Add articles to {monthLabel(campaign.offer_month)}</strong>
          <button type="button" onClick={onClose}>Done</button>
        </div>
        <div className="cp-row">
          {[["invoices", "From our invoices"], ["articles", "Our article cards"], ["products", "Product export"]].map(([key, label]) => (
            <button key={key} type="button" className={source === key ? "primary" : ""} onClick={() => setSource(key)}>{label}</button>
          ))}
        </div>
        <form className="cp-row" onSubmit={(event) => { event.preventDefault(); search(); }}>
          <input placeholder={source === "invoices" ? "Product name (e.g. Raphis)" : "Search"} value={query.product} onChange={(event) => setQuery({ ...query, product: event.target.value })} />
          {source === "invoices" && (
            <>
              <input placeholder="Grower" value={query.grower} onChange={(event) => setQuery({ ...query, grower: event.target.value })} />
              <input placeholder="VBN code" className="cp-num" value={query.vbn} onChange={(event) => setQuery({ ...query, vbn: event.target.value })} />
              <select value={query.months} onChange={(event) => setQuery({ ...query, months: Number(event.target.value) })}>
                {[3, 6, 12, 18, 24, 36].map((months) => <option key={months} value={months}>last {months} months</option>)}
              </select>
            </>
          )}
          <button type="submit" className="primary" disabled={busy}>Search</button>
        </form>
        <Message error={error} />
        <div className="cp-dialog-body">
          {source === "invoices" && (
            <table className="ik-table cp-table">
              <thead><tr><th>Invoice name</th><th>VBN</th><th>Grower</th><th>Hoofdgroep</th><th>Packing</th><th>Features</th><th>Pieces</th><th>Price (min–max · last)</th><th>Last bought</th><th /></tr></thead>
              <tbody>
                {rows.map((row) => {
                  const key = `${row.vbn_code}|${row.description}|${row.supplier_name}`;
                  return (
                    <tr key={key}>
                      <td>{row.description}</td>
                      <td>{row.vbn_code}</td>
                      <td>{row.supplier_name}{row.grower_code ? ` (${row.grower_code})` : ""}</td>
                      <td>{row.hoofdgroep}</td>
                      <td>{row.packages && row.per_package ? `${row.packages} × ${row.per_package}` : row.per_package || ""}</td>
                      <td>{Object.entries(row.features || {}).map(([code, value]) => `${code}=${value}`).join(" ")}</td>
                      <td>{amount(row.pieces)}</td>
                      <td>{money(row.min_price)}–{money(row.max_price)} · {money(row.last_price)}</td>
                      <td>{day(row.last_date)}</td>
                      <td>{added.includes(key) ? "✓ added" : <button type="button" onClick={() => add({ article: invoiceRowToArticle(row) }, key)}>Add</button>}</td>
                    </tr>
                  );
                })}
                {!rows.length && <tr><td colSpan={10} className="cp-hint">Search our purchases by name, grower or VBN code. Adding makes a new article card with the invoice's conditions.</td></tr>}
              </tbody>
            </table>
          )}
          {source === "articles" && (
            <table className="ik-table cp-table">
              <thead><tr><th>Article</th><th>Hoofdgroep</th><th>VBN</th><th>Conditions</th><th>Growers</th><th>Months offered</th><th /></tr></thead>
              <tbody>
                {articles.map((article) => (
                  <tr key={article.id}>
                    <td>{article.name}</td>
                    <td>{article.hoofdgroep}</td>
                    <td>{article.vbn_codes.join(", ")}</td>
                    <td>{conditionsSummary(article.conditions)}</td>
                    <td>{article.growers.map((grower) => grower.name || grower.code).join(", ")}</td>
                    <td>{article.campaigns}</td>
                    <td>{existing.has(article.id) || added.includes(`a${article.id}`) ? "✓ in this month" : <button type="button" onClick={() => add({ article_id: article.id }, `a${article.id}`)}>Add</button>}</td>
                  </tr>
                ))}
                {!articles.length && <tr><td colSpan={7} className="cp-hint">Search the article cards made before (they keep their conditions).</td></tr>}
              </tbody>
            </table>
          )}
          {source === "products" && (
            <table className="ik-table cp-table">
              <thead><tr><th>Product</th><th>Sales group</th><th>Height</th><th>Pot</th><th>Last sale</th><th>Last price</th><th /></tr></thead>
              <tbody>
                {products.map((product) => {
                  const key = `p${product.product_id}`;
                  return (
                    <tr key={product.product_id}>
                      <td>{product.description}</td>
                      <td>{product.sales_group}</td>
                      <td>{product.height || ""}</td>
                      <td>{product.pot_size || ""}</td>
                      <td>{day(product.last_sale_date)} {product.last_client || ""}</td>
                      <td>{money(product.last_price)}</td>
                      <td>{added.includes(key) ? "✓ added" : (
                        <button type="button" onClick={() => add({ article: { name: product.description, product_id: product.product_id, sales_group: product.sales_group,
                          conditions: { ...(product.height ? { height_cm: product.height } : {}), ...(product.pot_size ? { pot_cm: product.pot_size } : {}) } } }, key)}>Add</button>
                      )}</td>
                    </tr>
                  );
                })}
                {!products.length && <tr><td colSpan={7} className="cp-hint">Search our product export (height, pot size, sales group).</td></tr>}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// One article in a campaign: its card (steady conditions) + grower deals.

function ItemDialog({ apiJson, meta, canManage, item, campaign, onClose }) {
  const [deals, setDeals] = useState([]);
  const [openDeal, setOpenDeal] = useState(null);
  const [error, setError] = useState("");
  const [newDeal, setNewDeal] = useState(null);

  const loadDeals = useCallback(() => {
    apiJson(`/api/campaigns/deals?campaign_id=${campaign.id}`)
      .then((data) => setDeals(data.deals.filter((deal) => deal.item_id === item.id)))
      .catch((err) => setError(err.message));
  }, [apiJson, campaign.id, item.id]);
  useEffect(loadDeals, [loadDeals]);

  async function createDeal() {
    try {
      const data = await apiJson(`/api/campaigns/items/${item.id}/deals`, { method: "POST", body: JSON.stringify(newDeal) });
      setNewDeal(null);
      loadDeals();
      setOpenDeal(data.deal.id);
    } catch (err) {
      setError(err.message);
    }
  }

  const growers = item.article.growers || [];
  return (
    <div className="cp-overlay" role="dialog" aria-label={item.article.name}>
      <div className="cp-dialog">
        <div className="cp-row cp-dialog-head">
          <strong>{item.article.name} — {monthLabel(campaign.offer_month)}</strong>
          <button type="button" onClick={onClose}>Close</button>
        </div>
        <Message error={error} />
        <div className="cp-dialog-body">
          <ArticleEditor apiJson={apiJson} meta={meta} canManage={canManage} articleId={item.article_id} />
          <h3 className="cp-h3">Grower deals for this month</h3>
          <table className="ik-table cp-table">
            <thead><tr><th>Grower</th><th>Status</th><th>Agreed price</th><th>Valid</th><th>Next action</th><th>Contacts</th></tr></thead>
            <tbody>
              {deals.map((deal) => (
                <tr key={deal.id}>
                  <td><button type="button" className="link-button" onClick={() => setOpenDeal(deal.id)}>{deal.grower_name || deal.grower_code}</button>{deal.grower_code ? ` (${deal.grower_code})` : ""}</td>
                  <td><span className={`cp-chip cp-deal-${deal.status}`}>{DEAL_LABELS[deal.status]}</span></td>
                  <td>{money(deal.agreed_price)}</td>
                  <td>{day(deal.valid_from)}{deal.valid_to ? ` – ${day(deal.valid_to)}` : ""}</td>
                  <td>{deal.next_action ? `${deal.next_action} (${day(deal.next_action_at)})` : ""}</td>
                  <td>{deal.contacts || 0}</td>
                </tr>
              ))}
              {!deals.length && <tr><td colSpan={6} className="cp-hint">No grower deal yet.</td></tr>}
            </tbody>
          </table>
          {canManage && !newDeal && (
            <div className="cp-row">
              <span>New deal with:</span>
              {growers.map((grower) => (
                <button key={`${grower.name}|${grower.code}`} type="button" onClick={() => setNewDeal({ grower_name: grower.name, grower_code: grower.code, grower_gln: grower.gln, grower_fh: grower.fh, status: "idea" })}>
                  {grower.name || grower.code}
                </button>
              ))}
              <button type="button" onClick={() => setNewDeal({ grower_name: "", grower_code: "", status: "idea" })}>Other grower…</button>
            </div>
          )}
          {newDeal && (
            <div className="info-panel cp-form">
              <label>Grower name<input value={newDeal.grower_name} onChange={(event) => setNewDeal({ ...newDeal, grower_name: event.target.value })} /></label>
              <label>Connect code<input value={newDeal.grower_code} onChange={(event) => setNewDeal({ ...newDeal, grower_code: event.target.value })} /></label>
              <div className="cp-row cp-wide">
                <button type="button" className="primary" onClick={createDeal}>Make deal</button>
                <button type="button" onClick={() => setNewDeal(null)}>Cancel</button>
              </div>
            </div>
          )}
          {openDeal && <DealEditor apiJson={apiJson} meta={meta} canManage={canManage} dealId={openDeal} onChanged={loadDeals} onClose={() => setOpenDeal(null)} />}
        </div>
      </div>
    </div>
  );
}

// The master card: names (NL invoice / FR / EN), VBN codes, conditions,
// growers and the sales match rules.
function ArticleEditor({ apiJson, meta, canManage, articleId, onSaved }) {
  const [article, setArticle] = useState(null);
  const [draft, setDraft] = useState(null);
  const [error, setError] = useState("");
  const [saved, setSaved] = useState("");

  const load = useCallback(() => {
    apiJson(`/api/campaigns/articles/${articleId}`).then((data) => { setArticle(data.article); setDraft(data.article); }).catch((err) => setError(err.message));
  }, [apiJson, articleId]);
  useEffect(load, [load]);

  if (!draft) return <Message error={error} notice="Loading the article card…" />;
  const set = (key, value) => setDraft({ ...draft, [key]: value });
  const setCondition = (key, value) => setDraft({ ...draft, conditions: { ...draft.conditions, [key]: value } });
  const dirty = JSON.stringify(draft) !== JSON.stringify(article);

  async function save() {
    setError("");
    try {
      const conditions = { ...draft.conditions };
      // Pieces follow packages x per package unless typed separately.
      if (conditions.packages && conditions.per_package && String(article.conditions?.pieces ?? "") === String(conditions.pieces ?? "")) delete conditions.pieces;
      await apiJson(`/api/campaigns/articles/${articleId}`, { method: "PATCH", body: JSON.stringify({
        name: draft.name, name_fr: draft.name_fr, name_en: draft.name_en, vbn_codes: draft.vbn_codes, hoofdgroep: draft.hoofdgroep,
        conditions, notes: draft.notes, match_rules: draft.match_rules, growers: draft.growers,
      }) });
      setSaved("Saved.");
      load();
      if (onSaved) onSaved();
    } catch (err) {
      setError(err.message);
    }
  }

  const features = draft.conditions?.features || {};
  const disabled = !canManage;
  return (
    <div className="info-panel">
      <div className="cp-form">
        <label>Name (as on the invoice)<input disabled={disabled} value={draft.name} onChange={(event) => set("name", event.target.value)} /></label>
        <label>Name FR<input disabled={disabled} value={draft.name_fr} onChange={(event) => set("name_fr", event.target.value)} /></label>
        <label>Name EN<input disabled={disabled} value={draft.name_en} onChange={(event) => set("name_en", event.target.value)} /></label>
        <label>VBN code(s)<input disabled={disabled} value={(draft.vbn_codes || []).join(", ")} onChange={(event) => set("vbn_codes", event.target.value.split(/[\s,;]+/).filter(Boolean))} /></label>
        <label>Hoofdgroep<input disabled={disabled} value={draft.hoofdgroep} onChange={(event) => set("hoofdgroep", event.target.value)} /></label>
        {CONDITION_FIELDS.map(([key, label, type]) => (
          <label key={key}>{label}
            <input disabled={disabled} inputMode={type === "number" ? "decimal" : undefined} value={draft.conditions?.[key] ?? ""}
              placeholder={key === "pieces" && draft.conditions?.packages && draft.conditions?.per_package ? String(draft.conditions.packages * draft.conditions.per_package) : ""}
              onChange={(event) => setCondition(key, event.target.value)} />
          </label>
        ))}
        {Object.keys(features).length > 0 && (
          <div className="cp-wide cp-hint">
            VBN features from the invoice: {Object.entries(features).map(([code, value]) => `${meta.feature_labels?.[code] || code} = ${value}`).join(" · ")}
          </div>
        )}
        <label className="cp-wide">Notes<textarea disabled={disabled} rows={2} value={draft.notes} onChange={(event) => set("notes", event.target.value)} /></label>
        <div className="cp-wide">
          <strong>Growers that delivered it</strong>
          <div className="cp-row">
            {(draft.growers || []).map((grower, index) => (
              <span key={index} className="cp-chip">{grower.name}{grower.code ? ` (${grower.code})` : ""}{grower.last_price ? ` · last ${money(grower.last_price)}` : ""}</span>
            ))}
            {!draft.growers?.length && <span className="cp-hint">none yet</span>}
          </div>
        </div>
        <div className="cp-wide">
          <strong>Sales follow-up</strong> <span className="cp-hint">which lots count as this article (results tab, later): VBN codes, Connect grower codes, name terms</span>
          <div className="cp-form">
            <label>VBN codes<input disabled={disabled} value={(draft.match_rules?.vbn_codes || []).join(", ")}
              onChange={(event) => set("match_rules", { ...draft.match_rules, vbn_codes: event.target.value.split(/[\s,;]+/).filter(Boolean) })} /></label>
            <label>Grower codes<input disabled={disabled} value={(draft.match_rules?.grower_codes || []).join(", ")}
              onChange={(event) => set("match_rules", { ...draft.match_rules, grower_codes: event.target.value.split(/[\s,;]+/).filter(Boolean) })} /></label>
            <label>Name terms<input disabled={disabled} value={(draft.match_rules?.name_terms || []).join(", ")}
              onChange={(event) => set("match_rules", { ...draft.match_rules, name_terms: event.target.value.split(/\s*,\s*/).filter(Boolean) })} /></label>
          </div>
        </div>
      </div>
      {article.history?.length > 0 && (
        <p className="cp-hint">Offered in: {article.history.map((row) => `${monthLabel(row.offer_month)} (${ITEM_LABELS[row.status] || row.status}${row.offer_price ? `, ${money(row.offer_price)}` : ""})`).join(" · ")}</p>
      )}
      <Message error={error} notice={saved} />
      {canManage && <button type="button" className="primary" disabled={!dirty} onClick={save}>Save article card</button>}
    </div>
  );
}

// ---------------------------------------------------------------------------
// A grower deal: status, price, quantities per week, conditions, contact log.

function DealEditor({ apiJson, meta, canManage, dealId, onChanged, onClose }) {
  const [data, setData] = useState(null);
  const [draft, setDraft] = useState(null);
  const [contact, setContact] = useState({ kind: "call", note: "" });
  const [error, setError] = useState("");

  const load = useCallback(() => {
    apiJson(`/api/campaigns/deals/${dealId}`).then((result) => { setData(result); setDraft(result.deal); }).catch((err) => setError(err.message));
  }, [apiJson, dealId]);
  useEffect(load, [load]);

  if (!draft) return <Message error={error} notice="Loading the deal…" />;
  const set = (key, value) => setDraft({ ...draft, [key]: value });
  const disabled = !canManage;

  async function save() {
    setError("");
    try {
      await apiJson(`/api/campaigns/deals/${dealId}`, { method: "PATCH", body: JSON.stringify({
        grower_name: draft.grower_name, grower_code: draft.grower_code, status: draft.status, agreed_price: draft.agreed_price,
        valid_from: draft.valid_from, valid_to: draft.valid_to, next_action_at: draft.next_action_at, next_action: draft.next_action,
        qty_plan: draft.qty_plan, conditions: draft.conditions,
      }) });
      load();
      onChanged();
    } catch (err) {
      setError(err.message);
    }
  }

  async function addContact() {
    if (!contact.note.trim()) return;
    try {
      await apiJson(`/api/campaigns/deals/${dealId}/contacts`, { method: "POST", body: JSON.stringify(contact) });
      setContact({ kind: contact.kind, note: "" });
      load();
      onChanged();
    } catch (err) {
      setError(err.message);
    }
  }

  const qtyPlan = draft.qty_plan || [];
  return (
    <div className="info-panel cp-deal">
      <div className="cp-row cp-dialog-head">
        <strong>Deal: {draft.grower_name || draft.grower_code} — {draft.article_name} ({monthLabel(draft.offer_month)})</strong>
        <button type="button" onClick={onClose}>Close deal</button>
      </div>
      <div className="cp-form">
        <label>Status
          <select disabled={disabled} value={draft.status} onChange={(event) => set("status", event.target.value)}>
            {meta.deal_statuses.map((status) => <option key={status} value={status}>{DEAL_LABELS[status] || status}</option>)}
          </select>
        </label>
        <label>Grower<input disabled={disabled} value={draft.grower_name} onChange={(event) => set("grower_name", event.target.value)} /></label>
        <label>Connect code<input disabled={disabled} value={draft.grower_code} onChange={(event) => set("grower_code", event.target.value)} /></label>
        <label>Agreed price (€)<input disabled={disabled} inputMode="decimal" value={draft.agreed_price ?? ""} onChange={(event) => set("agreed_price", event.target.value)} /></label>
        <label>Valid from<input disabled={disabled} type="date" value={draft.valid_from || ""} onChange={(event) => set("valid_from", event.target.value)} /></label>
        <label>Valid until<input disabled={disabled} type="date" value={draft.valid_to || ""} onChange={(event) => set("valid_to", event.target.value)} /></label>
        <label>Next action<input disabled={disabled} value={draft.next_action || ""} placeholder="e.g. call back about price" onChange={(event) => set("next_action", event.target.value)} /></label>
        <label>Next action on<input disabled={disabled} type="date" value={draft.next_action_at || ""} onChange={(event) => set("next_action_at", event.target.value)} /></label>
        {CONDITION_FIELDS.filter(([key]) => ["length_cm", "packages", "per_package", "pieces", "quality", "fust", "min_order", "delivery_days"].includes(key)).map(([key, label]) => (
          <label key={key}>{label} (deal)
            <input disabled={disabled} value={draft.conditions?.[key] ?? ""} onChange={(event) => set("conditions", { ...draft.conditions, [key]: event.target.value })} />
          </label>
        ))}
        <div className="cp-wide">
          <strong>Quantities per week</strong>
          {qtyPlan.map((row, index) => (
            <div key={index} className="cp-row">
              <input disabled={disabled} placeholder="week (e.g. 2027-W09)" value={row.week} onChange={(event) => set("qty_plan", qtyPlan.map((other, at) => (at === index ? { ...other, week: event.target.value } : other)))} />
              <input disabled={disabled} className="cp-num" placeholder="qty" value={row.qty ?? ""} onChange={(event) => set("qty_plan", qtyPlan.map((other, at) => (at === index ? { ...other, qty: event.target.value } : other)))} />
              <input disabled={disabled} placeholder="note" value={row.note || ""} onChange={(event) => set("qty_plan", qtyPlan.map((other, at) => (at === index ? { ...other, note: event.target.value } : other)))} />
              {canManage && <button type="button" className="link-button" onClick={() => set("qty_plan", qtyPlan.filter((_, at) => at !== index))}>✕</button>}
            </div>
          ))}
          {canManage && <button type="button" onClick={() => set("qty_plan", [...qtyPlan, { week: "", qty: "", note: "" }])}>+ week</button>}
        </div>
      </div>
      <Message error={error} />
      {canManage && <button type="button" className="primary" onClick={save}>Save deal</button>}
      <h3 className="cp-h3">Contact log</h3>
      {canManage && (
        <div className="cp-row">
          <select value={contact.kind} onChange={(event) => setContact({ ...contact, kind: event.target.value })}>
            {meta.contact_kinds.map((kind) => <option key={kind} value={kind}>{kind}</option>)}
          </select>
          <input className="cp-grow" placeholder="What was said / agreed" value={contact.note} onChange={(event) => setContact({ ...contact, note: event.target.value })}
            onKeyDown={(event) => { if (event.key === "Enter") addContact(); }} />
          <button type="button" onClick={addContact}>Add</button>
        </div>
      )}
      <ul className="cp-log">
        {(data?.contacts || []).map((row) => (
          <li key={row.id}><span className="cp-hint">{new Date(row.at).toLocaleString("nl-NL")} · {row.kind} · {row.by_user}</span> {row.note}</li>
        ))}
        {!data?.contacts?.length && <li className="cp-hint">Nothing logged yet.</li>}
      </ul>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Growers: every deal, grouped per grower.

function GrowersTab({ apiJson, meta, canManage, campaignId }) {
  const [campaigns, setCampaigns] = useState([]);
  const [filter, setFilter] = useState({ campaign: campaignId || "", grower: "", open: true });
  const [deals, setDeals] = useState(null);
  const [openDeal, setOpenDeal] = useState(null);
  const [error, setError] = useState("");

  useEffect(() => {
    apiJson("/api/campaigns").then((data) => setCampaigns(data.campaigns)).catch((err) => setError(err.message));
  }, [apiJson]);

  const load = useCallback(() => {
    const params = new URLSearchParams();
    if (filter.campaign) params.set("campaign_id", filter.campaign);
    if (filter.grower) params.set("grower", filter.grower);
    if (filter.open) params.set("open", "1");
    apiJson(`/api/campaigns/deals?${params}`).then((data) => setDeals(data.deals)).catch((err) => setError(err.message));
  }, [apiJson, filter]);
  useEffect(load, [load]);

  const groups = useMemo(() => {
    const map = new Map();
    for (const deal of deals || []) {
      const key = deal.grower_code || deal.grower_name;
      if (!map.has(key)) map.set(key, { name: deal.grower_name, code: deal.grower_code, deals: [] });
      map.get(key).deals.push(deal);
    }
    return [...map.values()].sort((left, right) => left.name.localeCompare(right.name));
  }, [deals]);

  const today = new Date().toISOString().slice(0, 10);
  return (
    <section className="cp-section">
      <div className="cp-row">
        <select value={filter.campaign} onChange={(event) => setFilter({ ...filter, campaign: event.target.value })}>
          <option value="">All offer months</option>
          {campaigns.map((campaign) => <option key={campaign.id} value={campaign.id}>{monthLabel(campaign.offer_month)}</option>)}
        </select>
        <input placeholder="Grower" value={filter.grower} onChange={(event) => setFilter({ ...filter, grower: event.target.value })} />
        <label className="cp-check"><input type="checkbox" checked={filter.open} onChange={(event) => setFilter({ ...filter, open: event.target.checked })} /> Open deals only</label>
      </div>
      <Message error={error} />
      {!deals ? <p>Loading…</p> : !groups.length ? <p className="cp-hint">No grower deals here. Deals are made from an article in the Catalogue.</p> : groups.map((group) => (
        <div key={`${group.code}|${group.name}`} className="info-panel cp-grower">
          <div className="cp-row"><strong>{group.name}</strong>{group.code && <span className="cp-hint">{group.code}</span>}<span className="cp-hint">{group.deals.length} deal(s)</span></div>
          <table className="ik-table cp-table">
            <thead><tr><th>Offer month</th><th>Article</th><th>Status</th><th>Agreed price</th><th>Next action</th><th>Last contact</th></tr></thead>
            <tbody>
              {group.deals.map((deal) => (
                <Fragment key={deal.id}>
                  <tr>
                    <td>{monthLabel(deal.offer_month)}</td>
                    <td><button type="button" className="link-button" onClick={() => setOpenDeal(openDeal === deal.id ? null : deal.id)}>{deal.article_name}</button></td>
                    <td><span className={`cp-chip cp-deal-${deal.status}`}>{DEAL_LABELS[deal.status]}</span></td>
                    <td>{money(deal.agreed_price)}</td>
                    <td className={deal.next_action_at && deal.next_action_at <= today ? "cp-soon" : ""}>{deal.next_action ? `${deal.next_action} (${day(deal.next_action_at)})` : ""}</td>
                    <td>{deal.last_contact_at ? new Date(deal.last_contact_at).toLocaleDateString("nl-NL") : ""}</td>
                  </tr>
                  {openDeal === deal.id && (
                    <tr><td colSpan={6}><DealEditor apiJson={apiJson} meta={meta} canManage={canManage} dealId={deal.id} onChanged={load} onClose={() => setOpenDeal(null)} /></td></tr>
                  )}
                </Fragment>
              ))}
            </tbody>
          </table>
        </div>
      ))}
    </section>
  );
}

// ---------------------------------------------------------------------------
// Articles: all master cards.

function ArticlesTab({ apiJson, meta, canManage }) {
  const [query, setQuery] = useState("");
  const [archived, setArchived] = useState(false);
  const [articles, setArticles] = useState(null);
  const [openId, setOpenId] = useState(null);
  const [error, setError] = useState("");

  const load = useCallback(() => {
    apiJson(`/api/campaigns/articles?q=${encodeURIComponent(query)}${archived ? "&archived=1" : ""}`)
      .then((data) => setArticles(data.articles)).catch((err) => setError(err.message));
  }, [apiJson, query, archived]);
  useEffect(() => {
    const timer = setTimeout(load, 250);
    return () => clearTimeout(timer);
  }, [load]);

  async function setArchivedFlag(article, value) {
    try {
      await apiJson(`/api/campaigns/articles/${article.id}`, { method: "PATCH", body: JSON.stringify({ archived: value }) });
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <section className="cp-section">
      <div className="cp-row">
        <input placeholder="Search article cards" value={query} onChange={(event) => setQuery(event.target.value)} />
        <label className="cp-check"><input type="checkbox" checked={archived} onChange={(event) => setArchived(event.target.checked)} /> Include archived</label>
        <span className="cp-hint">New cards are made from the Catalogue (“+ Add articles”).</span>
      </div>
      <Message error={error} />
      {!articles ? <p>Loading…</p> : (
        <div className="info-panel cp-table-wrap">
          <table className="ik-table cp-table">
            <thead><tr><th>Article</th><th>FR / EN</th><th>Hoofdgroep</th><th>VBN</th><th>Conditions</th><th>Growers</th><th>Months</th>{canManage && <th />}</tr></thead>
            <tbody>
              {articles.map((article) => (
                <Fragment key={article.id}>
                  <tr className={article.archived ? "cp-dropped" : ""}>
                    <td><button type="button" className="link-button cp-name" onClick={() => setOpenId(openId === article.id ? null : article.id)}>{article.name}</button></td>
                    <td className="cp-wrap">{article.name_fr} / {article.name_en}</td>
                    <td>{article.hoofdgroep}</td>
                    <td>{article.vbn_codes.join(", ")}</td>
                    <td className="cp-wrap">{conditionsSummary(article.conditions)}</td>
                    <td className="cp-wrap">{article.growers.map((grower) => grower.name || grower.code).join(", ")}</td>
                    <td>{article.campaigns}</td>
                    {canManage && <td><button type="button" className="link-button" onClick={() => setArchivedFlag(article, !article.archived)}>{article.archived ? "Restore" : "Archive"}</button></td>}
                  </tr>
                  {openId === article.id && (
                    <tr><td colSpan={8}><ArticleEditor apiJson={apiJson} meta={meta} canManage={canManage} articleId={article.id} onSaved={load} /></td></tr>
                  )}
                </Fragment>
              ))}
              {!articles.length && <tr><td colSpan={8} className="cp-hint">No article cards found.</td></tr>}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

// ---------------------------------------------------------------------------
// Settings: who gets the tasks, lead time, target size.

function SettingsTab({ apiJson, meta, onSaved }) {
  const [draft, setDraft] = useState(meta.settings);
  const [directory, setDirectory] = useState({ users: [], teams: [] });
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    apiJson("/api/users/directory").then(setDirectory).catch(() => {});
  }, [apiJson]);

  async function save() {
    setError("");
    try {
      const data = await apiJson("/api/campaigns/settings", { method: "PUT", body: JSON.stringify(draft) });
      onSaved(data.settings);
      setMessage("Saved.");
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <section className="cp-section">
      <div className="info-panel cp-form">
        <label>Campaign tasks go to
          <select value={draft.task_target} onChange={(event) => setDraft({ ...draft, task_target: event.target.value })}>
            <option value="">Nobody (no tasks)</option>
            <optgroup label="Teams">
              {directory.teams.map((team) => <option key={team.key} value={`team:${team.key}`}>{team.name}</option>)}
            </optgroup>
            <optgroup label="Users">
              {directory.users.map((user) => <option key={user.username} value={`user:${user.username}`}>{user.username}</option>)}
            </optgroup>
          </select>
        </label>
        <label>Campaign runs … month(s) before the offer month
          <select value={draft.lead_months} onChange={(event) => setDraft({ ...draft, lead_months: Number(event.target.value) })}>
            {[1, 2, 3].map((months) => <option key={months} value={months}>{months}</option>)}
          </select>
        </label>
        <label>Articles per offer month (target)<input inputMode="numeric" value={draft.target_articles} onChange={(event) => setDraft({ ...draft, target_articles: event.target.value })} /></label>
        <label>Deadline task … days before<input inputMode="numeric" value={draft.deadline_task_days} onChange={(event) => setDraft({ ...draft, deadline_task_days: event.target.value })} /></label>
        <p className="cp-wide cp-hint">A team for campaign work (e.g. “Campagne”) can be added by an admin in the Tasks settings. Tasks are made for the deadlines of each offer month and for a grower deal's “next action” date.</p>
        <Message error={error} notice={message} />
        <div className="cp-wide"><button type="button" className="primary" onClick={save}>Save</button></div>
      </div>
    </section>
  );
}
