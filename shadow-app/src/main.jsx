import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const REFRESH_INTERVAL_MS = 15000;

function formatTimestamp(value) {
  if (!value) {
    return "unknown time";
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return value;
  }

  return parsed.toLocaleString();
}

function imageUrl(image, run, retryKey = "") {
  const params = new URLSearchParams();
  params.set("id", image.id);
  params.set("account", String(run?.metadata?.drive_account || "default"));
  if (image.mime_type) {
    params.set("mime", image.mime_type);
  }
  if (retryKey) {
    params.set("retry", String(retryKey));
  }
  return `/api/image?${params.toString()}`;
}

function PhotoImage({ image, run, alt, loading = "eager" }) {
  const [attempt, setAttempt] = useState(0);
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    setAttempt(0);
    setFailed(false);
  }, [image.id, run?.folder_id]);

  useEffect(() => {
    if (!failed || attempt >= 3) {
      return undefined;
    }

    const timer = window.setTimeout(() => {
      setFailed(false);
      setAttempt((value) => value + 1);
    }, 1200);

    return () => window.clearTimeout(timer);
  }, [attempt, failed]);

  return (
    <img
      src={imageUrl(image, run, attempt)}
      alt={alt}
      loading={loading}
      onError={() => setFailed(true)}
    />
  );
}

async function apiJson(path, options = {}) {
  const response = await fetch(path, {
    ...options,
    headers: {
      ...(options.body ? { "content-type": "application/json" } : {}),
      ...(options.headers || {}),
    },
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || `Request failed with ${response.status}`);
  }
  return payload;
}

function useDashboardData(selectedDate, searchTerm, syncVersion, enabled) {
  const [state, setState] = useState({
    data: null,
    loading: true,
    error: "",
  });

  useEffect(() => {
    if (!enabled) {
      setState({ data: null, loading: false, error: "" });
      return undefined;
    }

    const controller = new AbortController();
    const params = new URLSearchParams();
    if (selectedDate) {
      params.set("date", selectedDate);
    }
    if (searchTerm.trim()) {
      params.set("search", searchTerm.trim());
    }

    setState((current) => ({ ...current, loading: true, error: "" }));
    fetch(`/api/data?${params.toString()}`, { signal: controller.signal })
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Data request failed with ${response.status}`);
        }
        return response.json();
      })
      .then((payload) => {
        setState({ data: payload, loading: false, error: "" });
      })
      .catch((error) => {
        if (error.name !== "AbortError") {
          setState({ data: null, loading: false, error: error.message });
        }
      });

    return () => controller.abort();
  }, [selectedDate, searchTerm, syncVersion, enabled]);

  return state;
}

function useSyncStatus(enabled) {
  const [status, setStatus] = useState({});

  useEffect(() => {
    if (!enabled) {
      setStatus({});
      return undefined;
    }

    let stopped = false;

    async function loadStatus() {
      try {
        const response = await fetch("/api/status");
        if (!response.ok) {
          return;
        }
        const payload = await response.json();
        if (!stopped) {
          setStatus(payload);
        }
      } catch {
        if (!stopped) {
          setStatus({});
        }
      }
    }

    loadStatus();
    const interval = window.setInterval(loadStatus, REFRESH_INTERVAL_MS);
    return () => {
      stopped = true;
      window.clearInterval(interval);
    };
  }, [enabled]);

  return status;
}

function App() {
  const [auth, setAuth] = useState({ loading: true, user: null, setupRequired: false });
  const [page, setPage] = useState("dashboard");
  const loggedIn = Boolean(auth.user);
  const syncStatus = useSyncStatus(loggedIn);
  const [selectedDate, setSelectedDate] = useState("");
  const [dateWasManuallySelected, setDateWasManuallySelected] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [expandedCustomers, setExpandedCustomers] = useState(() => new Set());
  const [lightbox, setLightbox] = useState(null);
  const [syncVersion, setSyncVersion] = useState(0);
  const requestedDate = dateWasManuallySelected ? selectedDate : "";
  const { data, loading, error } = useDashboardData(requestedDate, searchTerm, syncVersion, loggedIn && page === "dashboard");

  const dates = data?.dates || [];
  const activeDate = selectedDate || data?.selected_date || "";
  const firstDate = dates[0] || "";
  const lastDate = dates.at(-1) || "";
  const syncRunning = syncStatus?.state === "running";

  useEffect(() => {
    apiJson("/api/auth/me")
      .then((payload) => {
        setAuth({
          loading: false,
          user: payload.user,
          setupRequired: payload.setup_required,
        });
      })
      .catch(() => setAuth({ loading: false, user: null, setupRequired: false }));
  }, []);

  useEffect(() => {
    if (!dateWasManuallySelected && data?.selected_date) {
      setSelectedDate(data.selected_date);
    }
  }, [data?.selected_date, dateWasManuallySelected]);

  useEffect(() => {
    if (syncStatus?.state === "completed" || syncStatus?.state === "failed") {
      setSyncVersion((value) => value + 1);
    }
  }, [syncStatus?.state, syncStatus?.updated_at]);

  const flatPhotoGroups = useMemo(() => {
    const groups = [];
    for (const group of data?.groups || []) {
      for (const run of group.runs || []) {
        if (Array.isArray(run.images) && run.images.length) {
          groups.push(run.images.map((image) => ({ image, run })));
        }
      }
    }
    return groups;
  }, [data]);

  async function startSync(pathname) {
    const params = new URLSearchParams();
    if (pathname === "/api/refresh-date") {
      params.set("date", activeDate);
    }

    const response = await fetch(`${pathname}?${params.toString()}`, { method: "POST" });
    if (response.ok) {
      setSyncVersion((value) => value + 1);
    }
  }

  async function handleLogout() {
    await apiJson("/api/auth/logout", { method: "POST" });
    setAuth({ loading: false, user: null, setupRequired: false });
    setPage("dashboard");
  }

  function toggleCustomer(customerCode) {
    setExpandedCustomers((current) => {
      const next = new Set(current);
      if (next.has(customerCode)) {
        next.delete(customerCode);
      } else {
        next.add(customerCode);
      }
      return next;
    });
  }

  if (auth.loading) {
    return <AuthShell title="Loading..." />;
  }

  if (!auth.user) {
    return (
      <AuthShell title="Sjaak vd Vijver Expedition Photo Dashboard">
        {auth.setupRequired ? (
          <SetupForm onSetup={(user) => setAuth({ loading: false, user, setupRequired: false })} />
        ) : (
          <LoginForm onLogin={(user) => setAuth({ loading: false, user, setupRequired: false })} />
        )}
      </AuthShell>
    );
  }

  return (
    <>
      <aside className="sidebar">
        <div>
          <p className="eyebrow">SnappySjaak</p>
          <h1>Sjaak vd Vijver Expedition Photo Dashboard</h1>
        </div>

        <nav className="side-nav" aria-label="Shadow app pages">
          <button className={page === "dashboard" ? "active" : ""} onClick={() => setPage("dashboard")}>
            Photos
          </button>
          {auth.user.role === "admin" && (
            <button className={page === "users" ? "active" : ""} onClick={() => setPage("users")}>
              Users
            </button>
          )}
        </nav>

        <div className="control-stack">
          <button className="primary" disabled={syncRunning} onClick={() => startSync("/api/rebuild")}>
            Rebuild run index
          </button>
          <button disabled={syncRunning || !activeDate} onClick={() => startSync("/api/refresh-date")}>
            Refresh this date
          </button>
          <p className="sidebar-note">
            {data?.generated_at
              ? `Using saved run index from ${formatTimestamp(data.generated_at)}.`
              : "No saved run index loaded yet."}
          </p>
        </div>

        <SyncPanel status={syncStatus} />
        <ParseErrors errors={data?.parse_errors || []} />
        <div className="signed-in">
          <p>Signed in as <strong>{auth.user.username}</strong></p>
          <button onClick={handleLogout}>Log out</button>
        </div>
      </aside>

      <main className="workspace">
        <header className="page-header">
          <h1>{page === "users" ? "Users" : "Sjaak vd Vijver Expedition Photo Dashboard"}</h1>
        </header>

        {page === "users" ? (
          <UsersPage currentUser={auth.user} />
        ) : (
          <>
            <section className="toolbar" aria-label="Filters">
              <label>
                <span>Filter by date</span>
                <input
                  type="date"
                  value={activeDate}
                  min={firstDate}
                  max={lastDate}
                  onChange={(event) => {
                    setDateWasManuallySelected(true);
                    setSelectedDate(event.target.value);
                  }}
                />
              </label>

              <label>
                <span>Search customer code</span>
                <input
                  value={searchTerm}
                  onChange={(event) => setSearchTerm(event.target.value)}
                  placeholder="cust123"
                />
              </label>

              <Metric label="Customers" value={data?.metrics?.customers || 0} />
              <Metric label="Runs" value={data?.metrics?.runs || 0} />
              <Metric label="Images" value={data?.metrics?.images || 0} />
            </section>

            {loading && <div className="notice">Loading run data...</div>}
            {error && <div className="notice danger">Unable to load data: {error}</div>}
            {data?.cache_missing && (
              <div className="notice">
                No saved run index found. Start a rebuild, or open the Streamlit app once to create the shared cache.
              </div>
            )}
            {data?.auto_sync && (
              <div className="notice">
                Automatic background sync started
                {data.auto_sync.date ? ` for ${data.auto_sync.date}` : ""}.
              </div>
            )}

            {!loading && !error && data && data.groups.length === 0 && (
              <div className="notice">No runs match the selected filters.</div>
            )}

            <section className="customer-list">
              {(data?.groups || []).map((group) => {
                const isExpanded = expandedCustomers.has(group.customer_code);
                return (
                  <CustomerGroup
                    key={group.customer_code}
                    group={group}
                    expanded={isExpanded}
                    onToggle={() => toggleCustomer(group.customer_code)}
                    onOpenPhoto={(run, imageIndex) => {
                      const photos = run.images.map((image) => ({ image, run }));
                      setLightbox({ photos, index: imageIndex });
                    }}
                  />
                );
              })}
            </section>
          </>
        )}
      </main>

      {lightbox && (
        <Lightbox
          photos={lightbox.photos}
          index={lightbox.index}
          onChange={(index) => setLightbox((current) => ({ ...current, index }))}
          onClose={() => setLightbox(null)}
        />
      )}

      <div className="sr-only" aria-live="polite">
        {flatPhotoGroups.length} photo groups loaded
      </div>
    </>
  );
}

function AuthShell({ title, children }) {
  return (
    <main className="auth-page">
      <section className="auth-panel">
        <p className="eyebrow">SnappySjaak</p>
        <h1>{title}</h1>
        {children}
      </section>
    </main>
  );
}

function LoginForm({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  async function submit(event) {
    event.preventDefault();
    setError("");
    try {
      const payload = await apiJson("/api/auth/login", {
        method: "POST",
        body: JSON.stringify({ username, password }),
      });
      onLogin(payload.user);
    } catch (loginError) {
      setError(loginError.message);
    }
  }

  return (
    <form className="auth-form" onSubmit={submit}>
      <label>
        <span>Username</span>
        <input value={username} onChange={(event) => setUsername(event.target.value)} autoComplete="username" />
      </label>
      <label>
        <span>Password</span>
        <input
          type="password"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          autoComplete="current-password"
        />
      </label>
      {error && <div className="notice danger">{error}</div>}
      <button className="primary" type="submit">Log in</button>
    </form>
  );
}

function SetupForm({ onSetup }) {
  const [username, setUsername] = useState("admin");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  async function submit(event) {
    event.preventDefault();
    setError("");
    try {
      const payload = await apiJson("/api/auth/setup", {
        method: "POST",
        body: JSON.stringify({ username, password }),
      });
      onSetup(payload.user);
    } catch (setupError) {
      setError(setupError.message);
    }
  }

  return (
    <form className="auth-form" onSubmit={submit}>
      <p className="sidebar-note">Create the first admin account for this shadow app.</p>
      <label>
        <span>Admin username</span>
        <input value={username} onChange={(event) => setUsername(event.target.value)} autoComplete="username" />
      </label>
      <label>
        <span>Admin password</span>
        <input
          type="password"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          autoComplete="new-password"
        />
      </label>
      {error && <div className="notice danger">{error}</div>}
      <button className="primary" type="submit">Create admin</button>
    </form>
  );
}

function UsersPage({ currentUser }) {
  const [users, setUsers] = useState([]);
  const [form, setForm] = useState({ username: "", password: "", role: "viewer" });
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");

  async function loadUsers() {
    const payload = await apiJson("/api/users");
    setUsers(payload.users || []);
  }

  useEffect(() => {
    loadUsers().catch((loadError) => setError(loadError.message));
  }, []);

  async function addUser(event) {
    event.preventDefault();
    setError("");
    setMessage("");
    try {
      await apiJson("/api/users", {
        method: "POST",
        body: JSON.stringify(form),
      });
      setForm({ username: "", password: "", role: "viewer" });
      setMessage("User added.");
      await loadUsers();
    } catch (addError) {
      setError(addError.message);
    }
  }

  async function updateUser(username, changes) {
    setError("");
    setMessage("");
    try {
      await apiJson(`/api/users/${encodeURIComponent(username)}`, {
        method: "PATCH",
        body: JSON.stringify(changes),
      });
      setMessage("User updated.");
      await loadUsers();
    } catch (updateError) {
      setError(updateError.message);
    }
  }

  async function deleteUser(username) {
    setError("");
    setMessage("");
    try {
      await apiJson(`/api/users/${encodeURIComponent(username)}`, { method: "DELETE" });
      setMessage("User deleted.");
      await loadUsers();
    } catch (deleteError) {
      setError(deleteError.message);
    }
  }

  return (
    <section className="users-page">
      <form className="user-form" onSubmit={addUser}>
        <label>
          <span>Username</span>
          <input value={form.username} onChange={(event) => setForm({ ...form, username: event.target.value })} />
        </label>
        <label>
          <span>Password</span>
          <input
            type="password"
            value={form.password}
            onChange={(event) => setForm({ ...form, password: event.target.value })}
            autoComplete="new-password"
          />
        </label>
        <label>
          <span>Role</span>
          <select value={form.role} onChange={(event) => setForm({ ...form, role: event.target.value })}>
            <option value="viewer">Viewer</option>
            <option value="admin">Admin</option>
          </select>
        </label>
        <button className="primary" type="submit">Add user</button>
      </form>

      {error && <div className="notice danger">{error}</div>}
      {message && <div className="notice">{message}</div>}

      <div className="users-list">
        {users.map((user) => (
          <UserRow
            key={user.username}
            user={user}
            currentUser={currentUser}
            onRoleChange={(role) => updateUser(user.username, { role })}
            onPasswordChange={(password) => updateUser(user.username, { password })}
            onDelete={() => deleteUser(user.username)}
          />
        ))}
      </div>
    </section>
  );
}

function UserRow({ user, currentUser, onRoleChange, onPasswordChange, onDelete }) {
  const [password, setPassword] = useState("");

  return (
    <article className="user-row">
      <div>
        <strong>{user.username}</strong>
        <span>{user.role} | created {formatTimestamp(user.created_at)}</span>
      </div>
      <select value={user.role} onChange={(event) => onRoleChange(event.target.value)}>
        <option value="viewer">Viewer</option>
        <option value="admin">Admin</option>
      </select>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          if (password) {
            onPasswordChange(password);
            setPassword("");
          }
        }}
      >
        <input
          type="password"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          placeholder="New password"
          autoComplete="new-password"
        />
        <button type="submit">Set password</button>
      </form>
      <button disabled={user.username === currentUser.username} onClick={onDelete}>Delete</button>
    </article>
  );
}

function SyncPanel({ status }) {
  if (!status?.state) {
    return <p className="sync idle">Background sync idle</p>;
  }

  if (status.state === "running") {
    return <p className="sync running">Running {String(status.mode || "").replaceAll("_", " ")} since {formatTimestamp(status.started_at)}</p>;
  }

  if (status.state === "completed") {
    return <p className="sync done">Background sync completed at {formatTimestamp(status.updated_at)}</p>;
  }

  if (status.state === "failed") {
    return <p className="sync failed">Background sync failed: {status.error || "unknown error"}</p>;
  }

  return <p className="sync idle">Background sync {status.state}</p>;
}

function ParseErrors({ errors }) {
  if (!errors.length) {
    return <p className="sidebar-note">No skipped folders.</p>;
  }

  return (
    <details className="parse-errors">
      <summary>{errors.length} skipped folders</summary>
      <div>
        {errors.slice(0, 80).map((error) => (
          <p key={`${error.folder_id}-${error.reason}`}>
            <strong>{error.folder_name}</strong>
            <span>{error.carrier}</span>
          </p>
        ))}
      </div>
    </details>
  );
}

function Metric({ label, value }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function CustomerGroup({ group, expanded, onToggle, onOpenPhoto }) {
  const totalImages = group.runs.reduce((total, run) => total + (run.images?.length || 0), 0);
  const carriers = [...new Set(group.runs.map((run) => run.carrier).filter(Boolean))].sort();

  return (
    <article className={`customer ${expanded ? "open" : ""}`}>
      <button className="customer-header" onClick={onToggle} aria-expanded={expanded}>
        <span>
          <strong>{group.customer_code}</strong>
          <small>
            {group.runs.length} runs | {totalImages} images | Carriers: {carriers.join(", ")}
          </small>
        </span>
        <b>{expanded ? "Collapse" : "Expand"}</b>
      </button>

      {expanded && (
        <div className="run-list">
          {group.runs.map((run) => (
            <RunCard key={run.folder_id} run={run} onOpenPhoto={onOpenPhoto} />
          ))}
        </div>
      )}
    </article>
  );
}

function RunCard({ run, onOpenPhoto }) {
  const hasQrInfo = run.qr_info && run.qr_info !== "No QR info found";

  return (
    <section className="run-card">
      <header>
        <h2>{run.carrier} | {run.run_id || "No run ID"}</h2>
        <p><strong>Carrier:</strong> {run.carrier}</p>
        <p><strong>Run ID:</strong> {run.run_id || "N/A"}</p>
        <p><strong>Folder:</strong> <code>{run.folder_name}</code></p>
      </header>

      {hasQrInfo && (
        <div className="qr">
          <strong>QR Info</strong>
          {run.qr_source === "filename" ? <code>{run.qr_info}</code> : <pre>{run.qr_info}</pre>}
        </div>
      )}

      {run.images?.length ? (
        <div className="gallery">
          {run.images.map((image, index) => (
            <button
              className="photo-card"
              key={image.id}
              onClick={() => onOpenPhoto(run, index)}
              aria-label={`Open ${image.name}`}
            >
              <PhotoImage image={image} run={run} alt={image.name} loading="lazy" />
              <span>{image.name}</span>
            </button>
          ))}
        </div>
      ) : (
        <p className="empty">No image files found in this run folder.</p>
      )}
    </section>
  );
}

function Lightbox({ photos, index, onChange, onClose }) {
  const active = photos[index];

  useEffect(() => {
    function handleKeydown(event) {
      if (event.key === "ArrowLeft") {
        event.preventDefault();
        onChange((index - 1 + photos.length) % photos.length);
      }
      if (event.key === "ArrowRight") {
        event.preventDefault();
        onChange((index + 1) % photos.length);
      }
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
      }
    }

    window.addEventListener("keydown", handleKeydown);
    return () => window.removeEventListener("keydown", handleKeydown);
  }, [index, onChange, onClose, photos.length]);

  function previous() {
    onChange((index - 1 + photos.length) % photos.length);
  }

  function next() {
    onChange((index + 1) % photos.length);
  }

  return (
    <div className="lightbox" role="dialog" aria-modal="true" onClick={onClose}>
      <button className="lightbox-close" onClick={onClose} aria-label="Close photo viewer">x</button>
      <button className="lightbox-nav previous" onClick={(event) => { event.stopPropagation(); previous(); }} aria-label="Previous photo">
        ‹
      </button>
      <figure onClick={(event) => event.stopPropagation()}>
        <PhotoImage image={active.image} run={active.run} alt={active.image.name} />
        <figcaption>
          {index + 1} / {photos.length} - {active.image.name}
        </figcaption>
      </figure>
      <button className="lightbox-nav next" onClick={(event) => { event.stopPropagation(); next(); }} aria-label="Next photo">
        ›
      </button>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
