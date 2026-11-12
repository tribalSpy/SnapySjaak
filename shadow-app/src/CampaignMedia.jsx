// Campagnes media: Make (queue photos / videos on the BEAST PC), Library
// (the Social Drive photos), Posts (recorded by hand until the platform
// APIs are approved) and the media settings / themes.
// Server: shadow-app/server/media.js. Previews and downloads are Google
// Drive links (the BEAST shares its outputs by link), so no image or video
// passes through the app server.
import React, { Fragment, useCallback, useEffect, useMemo, useState } from "react";

const PLATFORM_LABELS = { instagram: "Instagram", facebook: "Facebook", tiktok: "TikTok", youtube: "YouTube", linkedin: "LinkedIn", pinterest: "Pinterest" };
const POST_LABELS = { planned: "Planned", approved: "Approved", published: "Published", failed: "Failed", cancelled: "Cancelled" };
const ITEM_STATUS = { draft: "Draft", queued: "On the BEAST", ready: "Ready", failed: "Failed", archived: "Archived" };
const VARIANT_LABELS = { "9x16": "9:16 Reel / Short / TikTok", "4x5": "4:5 feed", "1x1": "1:1 square" };

export const driveThumb = (id, size = 400) => (id ? `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w${size}` : "");
export const drivePreview = (id) => `https://drive.google.com/file/d/${encodeURIComponent(id)}/preview`;
export const driveView = (id) => `https://drive.google.com/file/d/${encodeURIComponent(id)}/view`;
export const driveDownload = (id) => `https://drive.google.com/uc?export=download&id=${encodeURIComponent(id)}`;

function Message({ error, notice }) {
  if (error) return <p className="cp-msg cp-error">{error}</p>;
  if (notice) return <p className="cp-msg cp-notice">{notice}</p>;
  return null;
}

function when(value) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString("nl-NL", { dateStyle: "short", timeStyle: "short" });
}

function localInput(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const pad = (number) => String(number).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function useMediaMeta(apiJson) {
  const [meta, setMeta] = useState(null);
  const [themes, setThemes] = useState([]);
  useEffect(() => {
    apiJson("/api/media/meta").then(setMeta).catch(() => {});
    apiJson("/api/media/themes").then((data) => setThemes(data.themes)).catch(() => {});
  }, [apiJson]);
  return { meta, themes };
}

// ---------------------------------------------------------------------------
// BEAST status line

function BeastStatus({ apiJson }) {
  const [status, setStatus] = useState(null);
  useEffect(() => {
    let stopped = false;
    const load = () => apiJson("/api/media/beast").then((data) => { if (!stopped) setStatus(data); }).catch(() => {});
    load();
    const timer = setInterval(load, 20000);
    return () => { stopped = true; clearInterval(timer); };
  }, [apiJson]);
  if (!status) return null;
  const online = status.agents.filter((agent) => agent.online);
  const waiting = status.queue.filter((row) => row.status === "pending").reduce((sum, row) => sum + row.n, 0);
  const working = status.queue.filter((row) => row.status === "claimed").reduce((sum, row) => sum + row.n, 0);
  return (
    <div className={`cp-beast ${online.length ? "cp-beast-on" : "cp-beast-off"}`}>
      <strong>BEAST PC</strong>
      {online.length
        ? <span>online ({online.map((agent) => agent.pc_name || agent.agent_name).join(", ")})</span>
        : <span>{status.agents.length ? `offline — last seen ${when(status.agents[0].last_seen_at)}` : "not connected yet (media-poller-app)"}</span>}
      <span className="cp-hint">{working} working · {waiting} waiting</span>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Make: queue a photo / video on the BEAST, see what was made.

export function MakeTab({ apiJson, canManage, campaignId }) {
  const { meta, themes } = useMediaMeta(apiJson);
  const [campaigns, setCampaigns] = useState([]);
  const [form, setForm] = useState({ campaign_id: campaignId || "", article_ids: [], theme_key: "", recipe: "", asset_ids: [], variants: ["9x16"],
    duration: 8, text_fr: "", text_en: "", show_conditions: true, show_price: false, fit: "auto", title: "", now: false, ai_prompt: "", art_styles: ["sketch", "watercolor", "oil"] });
  const [items, setItems] = useState([]);
  const [campaignItems, setCampaignItems] = useState([]);
  const [assets, setAssets] = useState([]);
  const [assetQuery, setAssetQuery] = useState("");
  const [onlyArticle, setOnlyArticle] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [busy, setBusy] = useState(false);
  const [openItem, setOpenItem] = useState(null);

  useEffect(() => {
    apiJson("/api/campaigns").then((data) => {
      setCampaigns(data.campaigns);
      setForm((current) => (current.campaign_id || !data.campaigns[0] ? current : { ...current, campaign_id: data.campaigns[0].id }));
    }).catch(() => {});
  }, [apiJson]);

  useEffect(() => {
    if (!form.campaign_id) { setCampaignItems([]); return; }
    apiJson(`/api/campaigns/${form.campaign_id}`).then((data) => setCampaignItems(data.items.filter((item) => item.status !== "dropped"))).catch(() => {});
  }, [apiJson, form.campaign_id]);

  const loadItems = useCallback(() => {
    apiJson(`/api/media/items${form.campaign_id ? `?campaign_id=${form.campaign_id}` : ""}`).then((data) => setItems(data.items)).catch((err) => setError(err.message));
  }, [apiJson, form.campaign_id]);
  useEffect(loadItems, [loadItems]);
  // While something is on the BEAST, look again every 10 seconds.
  useEffect(() => {
    if (!items.some((item) => item.status === "queued")) return undefined;
    const timer = setInterval(loadItems, 10000);
    return () => clearInterval(timer);
  }, [items, loadItems]);

  useEffect(() => {
    const timer = setTimeout(() => {
      const params = new URLSearchParams({ q: assetQuery });
      if (onlyArticle && form.article_ids[0]) params.set("article_id", form.article_ids[0]);
      apiJson(`/api/media/assets?${params}`).then((data) => setAssets(data.assets)).catch(() => {});
    }, 250);
    return () => clearTimeout(timer);
  }, [apiJson, assetQuery, onlyArticle, form.article_ids]);

  const theme = themes.find((entry) => entry.key === form.theme_key);
  const recipes = meta?.recipes || [];
  const recipe = recipes.find((entry) => entry.key === form.recipe);
  const orderedRecipes = useMemo(() => {
    const first = theme ? theme.recipes : [];
    return [...first.map((key) => recipes.find((entry) => entry.key === key)).filter(Boolean), ...recipes.filter((entry) => !first.includes(entry.key))];
  }, [theme, recipes]);

  const set = (key, value) => setForm((current) => ({ ...current, [key]: value }));
  const toggle = (key, value) => setForm((current) => ({ ...current, [key]: current[key].includes(value) ? current[key].filter((entry) => entry !== value) : [...current[key], value] }));

  function chooseRecipe(key) {
    const next = recipes.find((entry) => entry.key === key);
    setForm((current) => ({ ...current, recipe: key, variants: next?.kind === "photo" ? ["4x5"] : (meta?.settings?.default_variants?.length ? meta.settings.default_variants : ["9x16"]) }));
  }

  async function make() {
    setBusy(true);
    setError("");
    setNotice("");
    try {
      const articleNames = campaignItems.filter((item) => form.article_ids.includes(item.article_id)).map((item) => item.article.name);
      const result = await apiJson("/api/media/items", { method: "POST", body: JSON.stringify({
        title: form.title || [theme?.name, articleNames.join(", ")].filter(Boolean).join(" — "),
        theme_key: form.theme_key,
        recipe: form.recipe,
        campaign_id: form.campaign_id || null,
        article_ids: form.article_ids,
        asset_ids: form.asset_ids,
        options: { duration: form.duration, variants: form.variants, text_fr: form.text_fr, text_en: form.text_en, show_conditions: form.show_conditions, show_price: form.show_price, fit: form.fit,
          ...(recipe?.engine === "ai" ? { now: form.now, ai_prompt: form.ai_prompt, art_styles: form.art_styles } : {}) },
      }) });
      setNotice(`Sent to the BEAST (item ${result.item_id}). It appears below when it's ready.`);
      setForm((current) => ({ ...current, asset_ids: [], title: "", text_fr: "", text_en: "" }));
      loadItems();
    } catch (err) {
      setError(err.message);
      loadItems();
    } finally {
      setBusy(false);
    }
  }

  const multi = recipe?.photos === "many";
  const enough = form.recipe && form.asset_ids.length > 0 && (multi || form.asset_ids.length === 1);

  return (
    <section className="cp-section">
      <BeastStatus apiJson={apiJson} />
      {canManage && (
        <div className="info-panel cp-make">
          <div className="cp-form">
            <label>Offer month
              <select value={form.campaign_id} onChange={(event) => setForm({ ...form, campaign_id: Number(event.target.value) || "", article_ids: [] })}>
                <option value="">— none —</option>
                {campaigns.map((campaign) => <option key={campaign.id} value={campaign.id}>{campaign.offer_month}</option>)}
              </select>
            </label>
            <label>Theme
              <select value={form.theme_key} onChange={(event) => set("theme_key", event.target.value)}>
                <option value="">— choose —</option>
                {themes.filter((entry) => entry.active).map((entry) => <option key={entry.key} value={entry.key}>{entry.name}</option>)}
              </select>
            </label>
            <label>What to make
              <select value={form.recipe} onChange={(event) => chooseRecipe(event.target.value)}>
                <option value="">— choose —</option>
                {orderedRecipes.map((entry) => (
                  <option key={entry.key} value={entry.key} disabled={!entry.available}>
                    {entry.kind === "video" ? "🎬" : "🖼"} {entry.label}{entry.available ? "" : " (needs the BEAST's AI part)"}
                  </option>
                ))}
              </select>
            </label>
            {recipe?.kind === "video" && (
              <label>Length (seconds)<input inputMode="numeric" value={form.duration} onChange={(event) => set("duration", event.target.value)} /></label>
            )}
            <label>Photo fit
              <select value={form.fit} onChange={(event) => set("fit", event.target.value)}>
                <option value="auto">Auto (packshots whole)</option>
                <option value="contain">Whole photo</option>
                <option value="cover">Fill the frame</option>
              </select>
            </label>
            {recipe?.engine === "ai" && (
              <>
                <label className="cp-wide">Extra for the AI (optional, English)
                  <input value={form.ai_prompt} placeholder="e.g. pink petals, morning light" onChange={(event) => set("ai_prompt", event.target.value)} />
                </label>
                {form.recipe === "ai_art" && (
                  <div className="cp-wide cp-row">
                    <span>Styles:</span>
                    {[["sketch", "Pencil sketch"], ["watercolor", "Watercolour"], ["oil", "Oil painting"]].map(([key, label]) => (
                      <label key={key} className="cp-check"><input type="checkbox" checked={form.art_styles.includes(key)} onChange={() => toggle("art_styles", key)} /> {label}</label>
                    ))}
                  </div>
                )}
                <label className="cp-check cp-wide">
                  <input type="checkbox" checked={form.now} onChange={(event) => set("now", event.target.checked)} />
                  Make it now (else it waits for the BEAST's night window, when the GPU is free) — AI takes several minutes
                </label>
              </>
            )}
            <label>Text FR (optional)<input value={form.text_fr} placeholder="the article's FR name" onChange={(event) => set("text_fr", event.target.value)} /></label>
            <label>Text EN (optional)<input value={form.text_en} placeholder="the article's EN name" onChange={(event) => set("text_en", event.target.value)} /></label>
            <div className="cp-wide cp-row">
              <span>Formats:</span>
              {(meta?.variants || []).map((variant) => (
                <label key={variant} className="cp-check"><input type="checkbox" checked={form.variants.includes(variant)} onChange={() => toggle("variants", variant)} /> {VARIANT_LABELS[variant]}</label>
              ))}
              <label className="cp-check"><input type="checkbox" checked={form.show_conditions} onChange={(event) => set("show_conditions", event.target.checked)} /> Show conditions</label>
              <label className="cp-check"><input type="checkbox" checked={form.show_price} onChange={(event) => set("show_price", event.target.checked)} /> Show offer price</label>
            </div>
            {campaignItems.length > 0 && (
              <div className="cp-wide">
                <strong>Articles</strong> <span className="cp-hint">(their FR/EN names and conditions go on the picture{multi ? "; in order, one per photo" : ""})</span>
                <div className="cp-row cp-pick-row">
                  {campaignItems.map((item) => (
                    <button key={item.article_id} type="button" className={form.article_ids.includes(item.article_id) ? "primary" : ""} onClick={() => toggle("article_ids", item.article_id)}>
                      {form.article_ids.includes(item.article_id) ? `${form.article_ids.indexOf(item.article_id) + 1}. ` : ""}{item.article.name_fr || item.article.name}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>
          <div className="cp-row">
            <strong>Photos</strong>
            <span className="cp-hint">{multi ? "click in the order you want them" : "choose one"} · {form.asset_ids.length} chosen</span>
            <input placeholder="Search the library" value={assetQuery} onChange={(event) => setAssetQuery(event.target.value)} />
            {form.article_ids[0] && <label className="cp-check"><input type="checkbox" checked={onlyArticle} onChange={(event) => setOnlyArticle(event.target.checked)} /> only photos linked to the first article</label>}
            {form.asset_ids.length > 0 && <button type="button" className="link-button" onClick={() => set("asset_ids", [])}>clear</button>}
          </div>
          <div className="cp-thumbs">
            {assets.map((asset) => {
              const index = form.asset_ids.indexOf(asset.id);
              return (
                <button key={asset.id} type="button" className={`cp-thumb ${index >= 0 ? "cp-thumb-on" : ""}`} title={`${asset.folder_path ? `${asset.folder_path}/` : ""}${asset.name}`}
                  onClick={() => (multi ? toggle("asset_ids", asset.id) : set("asset_ids", index >= 0 ? [] : [asset.id]))}>
                  {asset.thumb_id ? <img src={driveThumb(asset.thumb_id, 240)} alt={asset.name} loading="lazy" /> : <span className="cp-hint">{asset.name}</span>}
                  {index >= 0 && <span className="cp-thumb-n">{index + 1}</span>}
                </button>
              );
            })}
            {!assets.length && <p className="cp-hint">No photos — drop HD photos in the Social Drive folder and scan it (Library tab).</p>}
          </div>
          <Message error={error} notice={notice} />
          <div className="cp-row">
            <button type="button" className="primary" disabled={!enough || busy || !form.variants.length} onClick={make}>Make it on the BEAST</button>
            {!enough && <span className="cp-hint">Choose what to make and {multi ? "the photos" : "a photo"}.</span>}
          </div>
        </div>
      )}
      <h3 className="cp-h3">Made {form.campaign_id ? "for this offer month" : ""}</h3>
      <div className="cp-items">
        {items.map((item) => <ItemCard key={item.id} item={item} themes={themes} recipes={recipes} onOpen={() => setOpenItem(item.id)} />)}
        {!items.length && <p className="cp-hint">Nothing made yet.</p>}
      </div>
      {openItem && (
        <ItemDetail apiJson={apiJson} canManage={canManage} itemId={openItem} meta={meta} themes={themes}
          onClose={() => { setOpenItem(null); loadItems(); }} />
      )}
    </section>
  );
}

function ItemCard({ item, themes, recipes, onOpen }) {
  const cover = item.outputs.find((output) => output.kind === "cover") || item.outputs.find((output) => output.kind === "photo");
  const videos = item.outputs.filter((output) => output.kind === "video").length;
  const photos = item.outputs.filter((output) => output.kind === "photo").length;
  return (
    <button type="button" className="cp-item" onClick={onOpen}>
      <span className="cp-item-img">
        {cover ? <img src={driveThumb(cover.drive_id, 300)} alt="" loading="lazy" /> : <span className="cp-hint">{ITEM_STATUS[item.status] || item.status}</span>}
      </span>
      <span className="cp-item-body">
        <strong>{item.title || recipes.find((entry) => entry.key === item.recipe)?.label}</strong>
        <span className="cp-hint">{themes.find((theme) => theme.key === item.theme_key)?.name || item.theme_key} · {recipes.find((entry) => entry.key === item.recipe)?.label || item.recipe}</span>
        <span className={`cp-chip cp-item-${item.status}`}>{ITEM_STATUS[item.status] || item.status}{item.status === "ready" ? ` · ${videos ? `${videos} video` : ""}${videos && photos ? " · " : ""}${photos ? `${photos} photo` : ""}` : ""}</span>
        {item.error && <span className="cp-error-text">{item.error}</span>}
        {item.posts?.length > 0 && <span className="cp-hint">{item.posts.map((post) => `${PLATFORM_LABELS[post.platform]} ${POST_LABELS[post.status]}`).join(" · ")}</span>}
      </span>
    </button>
  );
}

function ItemDetail({ apiJson, canManage, itemId, meta, onClose }) {
  const [item, setItem] = useState(null);
  const [draft, setDraft] = useState({ caption_fr: "", caption_en: "", hashtags: "" });
  const [post, setPost] = useState({ platforms: [], status: "planned", scheduled_at: "", post_url: "", lang: "fr" });
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  const load = useCallback(() => {
    apiJson("/api/media/items").then((data) => {
      const found = data.items.find((entry) => entry.id === itemId);
      setItem(found || null);
      if (found) setDraft({ caption_fr: found.caption_fr, caption_en: found.caption_en, hashtags: found.hashtags });
    }).catch((err) => setError(err.message));
  }, [apiJson, itemId]);
  useEffect(load, [load]);

  if (!item) return null;
  const videos = item.outputs.filter((output) => output.kind === "video");
  const photos = item.outputs.filter((output) => output.kind === "photo");
  const covers = item.outputs.filter((output) => output.kind === "cover");
  const captionFor = (lang) => [lang === "fr" ? draft.caption_fr : lang === "en" ? draft.caption_en : `${draft.caption_fr}\n\n${draft.caption_en}`, draft.hashtags].filter(Boolean).join("\n\n");

  async function saveCaptions() {
    try {
      await apiJson(`/api/media/items/${item.id}`, { method: "PATCH", body: JSON.stringify(draft) });
      setNotice("Captions saved.");
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  async function rerender() {
    try {
      await apiJson(`/api/media/items/${item.id}/render`, { method: "POST" });
      setNotice("Sent to the BEAST again.");
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  async function addPost() {
    setError("");
    try {
      await apiJson("/api/media/posts", { method: "POST", body: JSON.stringify({
        media_item_id: item.id,
        platforms: post.platforms,
        status: post.status,
        scheduled_at: post.scheduled_at ? new Date(post.scheduled_at).toISOString() : null,
        posted_at: post.status === "published" && post.scheduled_at ? new Date(post.scheduled_at).toISOString() : undefined,
        post_url: post.post_url,
        caption: captionFor(post.lang),
      }) });
      setNotice(post.status === "published" ? "Recorded as posted." : "Planned — see the Posts tab.");
      setPost({ ...post, platforms: [], post_url: "" });
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  function copy(textValue) {
    navigator.clipboard?.writeText(textValue).then(() => setNotice("Copied."), () => setError("Could not copy"));
  }

  return (
    <div className="cp-overlay" role="dialog" aria-label={item.title}>
      <div className="cp-dialog">
        <div className="cp-row cp-dialog-head">
          <strong>{item.title || item.recipe}</strong>
          <span className={`cp-chip cp-item-${item.status}`}>{ITEM_STATUS[item.status] || item.status}</span>
          {item.render_seconds ? <span className="cp-hint">made in {Math.round(item.render_seconds)} s</span> : null}
          <button type="button" onClick={onClose}>Close</button>
        </div>
        <Message error={error || item.error} notice={notice} />
        <div className="cp-dialog-body">
          {item.status === "queued" && <p className="cp-hint">Waiting for / being made on the BEAST PC…</p>}
          <div className="cp-previews">
            {videos.map((output) => (
              <div key={output.drive_id} className="cp-preview">
                <iframe title={output.name} src={drivePreview(output.drive_id)} allow="autoplay" className={`cp-frame cp-frame-${output.variant}`} />
                <div className="cp-row">
                  <span>{VARIANT_LABELS[output.variant] || output.variant}{output.duration ? ` · ${output.duration}s` : ""}</span>
                  <a className="button" href={driveDownload(output.drive_id)} target="_blank" rel="noreferrer">Download</a>
                  <a className="link-button" href={driveView(output.drive_id)} target="_blank" rel="noreferrer">Open in Drive</a>
                  {covers.filter((cover) => cover.variant === output.variant).map((cover) => (
                    <a key={cover.drive_id} className="link-button" href={driveDownload(cover.drive_id)} target="_blank" rel="noreferrer">Cover picture</a>
                  ))}
                </div>
              </div>
            ))}
            {photos.map((output) => (
              <div key={output.drive_id} className="cp-preview">
                <img src={driveThumb(output.drive_id, 800)} alt={output.name} className="cp-photo-out" />
                <div className="cp-row">
                  <span>{VARIANT_LABELS[output.variant] || output.variant}{output.index ? ` · ${output.index}` : ""}</span>
                  <a className="button" href={driveDownload(output.drive_id)} target="_blank" rel="noreferrer">Download</a>
                </div>
              </div>
            ))}
          </div>
          <div className="info-panel cp-form">
            <label className="cp-wide">Caption FR
              <textarea rows={3} disabled={!canManage} value={draft.caption_fr} onChange={(event) => setDraft({ ...draft, caption_fr: event.target.value })} />
            </label>
            <label className="cp-wide">Caption EN
              <textarea rows={3} disabled={!canManage} value={draft.caption_en} onChange={(event) => setDraft({ ...draft, caption_en: event.target.value })} />
            </label>
            <label className="cp-wide">Hashtags<input disabled={!canManage} value={draft.hashtags} onChange={(event) => setDraft({ ...draft, hashtags: event.target.value })} /></label>
            <div className="cp-wide cp-row">
              {canManage && <button type="button" className="primary" onClick={saveCaptions}>Save captions</button>}
              <button type="button" onClick={() => copy(captionFor("fr"))}>Copy FR + tags</button>
              <button type="button" onClick={() => copy(captionFor("en"))}>Copy EN + tags</button>
              <button type="button" onClick={() => copy(captionFor("both"))}>Copy FR + EN + tags</button>
              {canManage && <button type="button" onClick={rerender}>Make again</button>}
            </div>
          </div>
          {canManage && item.status === "ready" && (
            <div className="info-panel cp-form">
              <strong className="cp-wide">Post it (by hand for now: download, post from the phone / browser, paste the link)</strong>
              <div className="cp-wide cp-row">
                {(meta?.platforms || []).map((platform) => (
                  <label key={platform} className="cp-check">
                    <input type="checkbox" checked={post.platforms.includes(platform)}
                      onChange={() => setPost({ ...post, platforms: post.platforms.includes(platform) ? post.platforms.filter((entry) => entry !== platform) : [...post.platforms, platform] })} />
                    {PLATFORM_LABELS[platform]}
                  </label>
                ))}
              </div>
              <label>Status
                <select value={post.status} onChange={(event) => setPost({ ...post, status: event.target.value })}>
                  <option value="planned">Plan it</option>
                  <option value="published">Already posted</option>
                </select>
              </label>
              <label>{post.status === "published" ? "Posted on" : "Planned for"}
                <input type="datetime-local" value={post.scheduled_at} onChange={(event) => setPost({ ...post, scheduled_at: event.target.value })} />
              </label>
              <label>Caption language
                <select value={post.lang} onChange={(event) => setPost({ ...post, lang: event.target.value })}>
                  <option value="fr">French</option>
                  <option value="en">English</option>
                  <option value="both">French + English</option>
                </select>
              </label>
              {post.status === "published" && <label>Link to the post<input value={post.post_url} placeholder="https://www.instagram.com/p/…" onChange={(event) => setPost({ ...post, post_url: event.target.value })} /></label>}
              <div className="cp-wide"><button type="button" className="primary" disabled={!post.platforms.length} onClick={addPost}>{post.status === "published" ? "Record as posted" : "Plan"}</button></div>
            </div>
          )}
          {item.posts?.length > 0 && (
            <p className="cp-hint">Posts: {item.posts.map((entry) => `${PLATFORM_LABELS[entry.platform]} — ${POST_LABELS[entry.status]} ${when(entry.posted_at || entry.scheduled_at)}`).join(" · ")}</p>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Library: the Social Drive photos.

export function LibraryTab({ apiJson, canManage }) {
  const [assets, setAssets] = useState(null);
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(null);
  const [articles, setArticles] = useState([]);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  const load = useCallback(() => {
    apiJson(`/api/media/assets?q=${encodeURIComponent(query)}`).then((data) => setAssets(data.assets)).catch((err) => setError(err.message));
  }, [apiJson, query]);
  useEffect(() => {
    const timer = setTimeout(load, 250);
    return () => clearTimeout(timer);
  }, [load]);
  useEffect(() => {
    apiJson("/api/campaigns/articles").then((data) => setArticles(data.articles)).catch(() => {});
  }, [apiJson]);

  async function scan() {
    setError("");
    try {
      const result = await apiJson("/api/media/assets/scan", { method: "POST" });
      setNotice(result.already ? "A scan is already waiting for the BEAST." : "Scan sent to the BEAST — new photos appear here when it's done.");
    } catch (err) {
      setError(err.message);
    }
  }

  async function patch(asset, fields) {
    try {
      const data = await apiJson(`/api/media/assets/${asset.id}`, { method: "PATCH", body: JSON.stringify(fields) });
      setOpen(data.asset);
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <section className="cp-section">
      <BeastStatus apiJson={apiJson} />
      <div className="cp-row">
        <input placeholder="Search photos (name, folder, tag)" value={query} onChange={(event) => setQuery(event.target.value)} />
        {canManage && <button type="button" className="primary" onClick={scan}>Scan the Social folder</button>}
        <span className="cp-hint">{assets ? `${assets.length} photo(s)` : ""} · subfolders become tags</span>
      </div>
      <Message error={error} notice={notice} />
      <div className="cp-thumbs cp-thumbs-big">
        {(assets || []).map((asset) => (
          <button key={asset.id} type="button" className={`cp-thumb ${open?.id === asset.id ? "cp-thumb-on" : ""}`} onClick={() => setOpen(asset)} title={asset.name}>
            {asset.thumb_id ? <img src={driveThumb(asset.thumb_id, 300)} alt={asset.name} loading="lazy" /> : <span className="cp-hint">{asset.name}</span>}
            {asset.article_ids.length > 0 && <span className="cp-thumb-n">✓</span>}
          </button>
        ))}
      </div>
      {open && (
        <div className="info-panel cp-form">
          <div className="cp-wide cp-row">
            <strong>{open.name}</strong>
            <span className="cp-hint">{open.folder_path} · {open.width}×{open.height}</span>
            <a className="link-button" href={driveView(open.drive_id)} target="_blank" rel="noreferrer">Open in Drive</a>
            <button type="button" onClick={() => setOpen(null)}>Close</button>
          </div>
          <label className="cp-wide">Tags (comma separated)
            <input disabled={!canManage} defaultValue={open.tags.join(", ")} key={`t${open.id}`} onBlur={(event) => patch(open, { tags: event.target.value.split(",") })} />
          </label>
          <div className="cp-wide">
            <strong>Linked articles</strong> <span className="cp-hint">(Make can show only an article's photos)</span>
            <div className="cp-row cp-pick-row">
              {articles.map((article) => (
                <button key={article.id} type="button" disabled={!canManage} className={open.article_ids.includes(article.id) ? "primary" : ""}
                  onClick={() => patch(open, { article_ids: open.article_ids.includes(article.id) ? open.article_ids.filter((id) => id !== article.id) : [...open.article_ids, article.id] })}>
                  {article.name}
                </button>
              ))}
              {!articles.length && <span className="cp-hint">No article cards yet.</span>}
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

// ---------------------------------------------------------------------------
// Posts: planned and posted (by hand), with their stats.

export function PostsTab({ apiJson, canManage }) {
  const [posts, setPosts] = useState(null);
  const [range, setRange] = useState(() => {
    const today = new Date();
    const from = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const to = new Date(today.getFullYear(), today.getMonth() + 2, 0);
    const iso = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
    return { from: iso(from), to: iso(to) };
  });
  const [statsFor, setStatsFor] = useState(null);
  const [error, setError] = useState("");

  const load = useCallback(() => {
    apiJson(`/api/media/posts?from=${range.from}&to=${range.to}`).then((data) => setPosts(data.posts)).catch((err) => setError(err.message));
  }, [apiJson, range]);
  useEffect(load, [load]);

  async function patch(post, fields) {
    setError("");
    try {
      await apiJson(`/api/media/posts/${post.id}`, { method: "PATCH", body: JSON.stringify(fields) });
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  async function remove(post) {
    if (!window.confirm("Remove this planned post?")) return;
    try {
      await apiJson(`/api/media/posts/${post.id}`, { method: "DELETE" });
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  function markPosted(post) {
    const link = window.prompt("Link to the post (optional):", post.post_url || "");
    if (link === null) return;
    patch(post, { status: "published", post_url: link });
  }

  const totals = (post) => {
    const last = post.stats[post.stats.length - 1];
    return last ? ["views", "reach", "likes", "comments", "shares", "saves"].filter((key) => last[key] !== null && last[key] !== undefined).map((key) => `${key} ${last[key]}`).join(" · ") : "";
  };

  return (
    <section className="cp-section">
      <div className="cp-row">
        <label>From <input type="date" value={range.from} onChange={(event) => setRange({ ...range, from: event.target.value })} /></label>
        <label>Until <input type="date" value={range.to} onChange={(event) => setRange({ ...range, to: event.target.value })} /></label>
        <span className="cp-hint">Posts are planned from a made item (Make tab). Until the platforms' APIs are approved, post by hand and record the link here; type the stats from the platform's insights.</span>
      </div>
      <Message error={error} />
      <div className="info-panel cp-table-wrap">
        <table className="ik-table cp-table">
          <thead><tr><th>When</th><th>Platform</th><th>What</th><th>Status</th><th>Link</th><th>Stats (latest)</th>{canManage && <th />}</tr></thead>
          <tbody>
            {(posts || []).map((post) => (
              <Fragment key={post.id}>
                <tr>
                  <td>{canManage && post.status !== "published"
                    ? <input type="datetime-local" value={localInput(post.scheduled_at)} onChange={(event) => patch(post, { scheduled_at: event.target.value ? new Date(event.target.value).toISOString() : null })} />
                    : when(post.posted_at || post.scheduled_at)}</td>
                  <td>{PLATFORM_LABELS[post.platform] || post.platform}</td>
                  <td className="cp-wrap">{post.item_title || post.recipe || "—"}</td>
                  <td><span className={`cp-chip cp-post-${post.status}`}>{POST_LABELS[post.status]}</span></td>
                  <td>{post.post_url ? <a href={post.post_url} target="_blank" rel="noreferrer">open</a> : ""}</td>
                  <td className="cp-wrap">{totals(post)}</td>
                  {canManage && (
                    <td className="cp-row">
                      {post.status === "planned" && <button type="button" onClick={() => patch(post, { status: "approved" })}>Approve</button>}
                      {["planned", "approved"].includes(post.status) && <button type="button" onClick={() => markPosted(post)}>Posted ✓</button>}
                      {post.status === "published" && <button type="button" onClick={() => setStatsFor(statsFor === post.id ? null : post.id)}>Stats</button>}
                      {["planned", "approved"].includes(post.status) && <button type="button" className="link-button" onClick={() => remove(post)}>✕</button>}
                    </td>
                  )}
                </tr>
                {statsFor === post.id && <tr><td colSpan={7}><StatsForm apiJson={apiJson} post={post} onSaved={() => { setStatsFor(null); load(); }} /></td></tr>}
              </Fragment>
            ))}
            {posts && !posts.length && <tr><td colSpan={7} className="cp-hint">No posts in this period.</td></tr>}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function StatsForm({ apiJson, post, onSaved }) {
  const today = new Date();
  const iso = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
  const [form, setForm] = useState({ day: iso, views: "", reach: "", likes: "", comments: "", shares: "", saves: "", follows: "", clicks: "" });
  const [error, setError] = useState("");
  async function save() {
    try {
      await apiJson(`/api/media/posts/${post.id}/stats`, { method: "POST", body: JSON.stringify(form) });
      onSaved();
    } catch (err) {
      setError(err.message);
    }
  }
  return (
    <div className="cp-form">
      <label>Day<input type="date" value={form.day} onChange={(event) => setForm({ ...form, day: event.target.value })} /></label>
      {["views", "reach", "likes", "comments", "shares", "saves", "follows", "clicks"].map((key) => (
        <label key={key}>{key}<input inputMode="numeric" value={form[key]} onChange={(event) => setForm({ ...form, [key]: event.target.value })} /></label>
      ))}
      <Message error={error} />
      <div className="cp-wide"><button type="button" className="primary" onClick={save}>Save stats</button></div>
      {post.stats.length > 0 && (
        <p className="cp-wide cp-hint">Saved: {post.stats.map((row) => `${String(row.day).slice(0, 10)}: ${row.views ?? "–"} views, ${row.likes ?? "–"} likes`).join(" · ")}</p>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Settings: Drive folders, brand; themes.

export function MediaSettings({ apiJson }) {
  const [settings, setSettings] = useState(null);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  useEffect(() => {
    apiJson("/api/media/meta").then((data) => setSettings(data.settings)).catch((err) => setError(err.message));
  }, [apiJson]);
  if (!settings) return <Message error={error} />;
  const set = (key, value) => setSettings({ ...settings, [key]: value });
  async function save() {
    setError("");
    try {
      const data = await apiJson("/api/media/settings", { method: "PUT", body: JSON.stringify(settings) });
      setSettings(data.settings);
      setMessage("Saved.");
    } catch (err) {
      setError(err.message);
    }
  }
  return (
    <section className="cp-section">
      <h3 className="cp-h3">Photos &amp; videos (BEAST PC)</h3>
      <div className="info-panel cp-form">
        <label className="cp-wide">Social library Drive folder (link or id) — where people drop HD photos
          <input value={settings.library_folder_id} onChange={(event) => set("library_folder_id", event.target.value)} placeholder="https://drive.google.com/drive/folders/…" />
        </label>
        <label className="cp-wide">Output Drive folder (link or id) — where the BEAST puts what it makes
          <input value={settings.output_folder_id} onChange={(event) => set("output_folder_id", event.target.value)} placeholder="https://drive.google.com/drive/folders/…" />
        </label>
        <label>Brand name on pictures<input value={settings.brand_name} onChange={(event) => set("brand_name", event.target.value)} /></label>
        <label>Accent colour<input type="color" value={settings.accent} onChange={(event) => set("accent", event.target.value)} /></label>
        <label className="cp-check"><input type="checkbox" checked={settings.share_outputs} onChange={(event) => set("share_outputs", event.target.checked)} /> Share outputs by link (needed for previews here)</label>
        <div className="cp-wide cp-row">
          <span>Default video formats:</span>
          {Object.entries(VARIANT_LABELS).map(([variant, label]) => (
            <label key={variant} className="cp-check">
              <input type="checkbox" checked={settings.default_variants.includes(variant)}
                onChange={() => set("default_variants", settings.default_variants.includes(variant) ? settings.default_variants.filter((entry) => entry !== variant) : [...settings.default_variants, variant])} />
              {label}
            </label>
          ))}
        </div>
        <p className="cp-wide cp-hint">The BEAST's Drive login must be able to write in the output folder (a service account: put the folders in a Shared Drive). See media-poller-app/README.md.</p>
        <Message error={error} notice={message} />
        <div className="cp-wide"><button type="button" className="primary" onClick={save}>Save</button></div>
      </div>
    </section>
  );
}

export function ThemesEditor({ apiJson }) {
  const { meta } = useMediaMeta(apiJson);
  const [themes, setThemes] = useState([]);
  const [openId, setOpenId] = useState(null);
  const [newName, setNewName] = useState("");
  const [error, setError] = useState("");
  const load = useCallback(() => {
    apiJson("/api/media/themes").then((data) => setThemes(data.themes)).catch((err) => setError(err.message));
  }, [apiJson]);
  useEffect(load, [load]);

  async function save(theme) {
    setError("");
    try {
      await apiJson(`/api/media/themes/${theme.id}`, { method: "PATCH", body: JSON.stringify({
        name: theme.name, recipes: theme.recipes, caption_prompt: theme.caption_prompt, hashtags: theme.hashtags, music: theme.music, cadence: theme.cadence, active: theme.active,
      }) });
      setOpenId(null);
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  async function add() {
    if (!newName.trim()) return;
    try {
      await apiJson("/api/media/themes", { method: "POST", body: JSON.stringify({ name: newName, recipes: ["zoom_in", "photo_post"] }) });
      setNewName("");
      load();
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <section className="cp-section">
      <h3 className="cp-h3">Themes</h3>
      <Message error={error} />
      <div className="info-panel cp-table-wrap">
        <table className="ik-table cp-table">
          <thead><tr><th>Theme</th><th>Recipes</th><th>Caption brief</th><th>Active</th><th /></tr></thead>
          <tbody>
            {themes.map((theme) => (
              <Fragment key={theme.id}>
                <tr>
                  <td><strong>{theme.name}</strong></td>
                  <td className="cp-wrap">{theme.recipes.map((key) => meta?.recipes.find((recipe) => recipe.key === key)?.label || key).join(", ")}</td>
                  <td className="cp-wrap">{theme.caption_prompt}</td>
                  <td>{theme.active ? "yes" : "no"}</td>
                  <td><button type="button" className="link-button" onClick={() => setOpenId(openId === theme.id ? null : theme.id)}>Edit</button></td>
                </tr>
                {openId === theme.id && (
                  <tr><td colSpan={5}><ThemeForm theme={theme} recipes={meta?.recipes || []} onSave={save} onCancel={() => setOpenId(null)} /></td></tr>
                )}
              </Fragment>
            ))}
          </tbody>
        </table>
      </div>
      <div className="cp-row">
        <input placeholder="New theme name" value={newName} onChange={(event) => setNewName(event.target.value)} />
        <button type="button" onClick={add}>Add theme</button>
      </div>
    </section>
  );
}

function ThemeForm({ theme, recipes, onSave, onCancel }) {
  const [draft, setDraft] = useState(theme);
  const set = (key, value) => setDraft({ ...draft, [key]: value });
  return (
    <div className="cp-form">
      <label>Name<input value={draft.name} onChange={(event) => set("name", event.target.value)} /></label>
      <label>Music file (in the BEAST's music folder)<input value={draft.music} onChange={(event) => set("music", event.target.value)} /></label>
      <label>Cadence (e.g. 1st of the month)<input value={draft.cadence} onChange={(event) => set("cadence", event.target.value)} /></label>
      <label className="cp-check"><input type="checkbox" checked={draft.active} onChange={(event) => set("active", event.target.checked)} /> Active</label>
      <label className="cp-check" title="Autopilot comes in a later phase; every post is approved by hand until then."><input type="checkbox" checked={false} disabled readOnly /> Autopilot (later — off)</label>
      <div className="cp-wide cp-row">
        <span>Recipes:</span>
        {recipes.map((recipe) => (
          <label key={recipe.key} className="cp-check">
            <input type="checkbox" checked={draft.recipes.includes(recipe.key)} onChange={() => set("recipes", draft.recipes.includes(recipe.key) ? draft.recipes.filter((key) => key !== recipe.key) : [...draft.recipes, recipe.key])} />
            {recipe.label}
          </label>
        ))}
      </div>
      <label className="cp-wide">Caption brief (for the FR/EN captions)<textarea rows={2} value={draft.caption_prompt} onChange={(event) => set("caption_prompt", event.target.value)} /></label>
      <label className="cp-wide">Hashtags<input value={draft.hashtags} onChange={(event) => set("hashtags", event.target.value)} /></label>
      <div className="cp-wide cp-row">
        <button type="button" className="primary" onClick={() => onSave(draft)}>Save theme</button>
        <button type="button" onClick={onCancel}>Cancel</button>
      </div>
    </div>
  );
}
