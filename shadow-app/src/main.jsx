import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const REFRESH_INTERVAL_MS = 15000;
const PERMISSIONS = {
  PHOTOS_VIEW: "photos:view",
  FUST_VIEW: "fust:view",
  FUST_IN: "fust:in",
  FUST_OUT: "fust:out",
  FUST_OVERVIEW: "fust:overview",
  CLOCK_VIEW: "clock:view",
  CLOCK_MANAGE: "clock:manage",
  USERS_MANAGE: "users:manage",
  SETTINGS_MANAGE: "settings:manage",
};
const ALL_PERMISSIONS = Object.values(PERMISSIONS);
const DEFAULT_PERMISSIONS_BY_ROLE = {
  admin: ALL_PERMISSIONS,
  viewer: [PERMISSIONS.PHOTOS_VIEW],
};
const PAGE_DEFINITIONS = [
  { key: "dashboard", label: "Photos", permission: PERMISSIONS.PHOTOS_VIEW },
  { key: "fust", label: "Fust", permission: PERMISSIONS.FUST_VIEW },
  { key: "clock", label: "Inklokken", permission: PERMISSIONS.CLOCK_VIEW },
  { key: "users", label: "Users", permission: PERMISSIONS.USERS_MANAGE },
  { key: "settings", label: "Settings", permission: PERMISSIONS.SETTINGS_MANAGE },
];

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

function normalizePermissions(role, permissions) {
  if (role === "admin") {
    return [...DEFAULT_PERMISSIONS_BY_ROLE.admin];
  }
  const defaults = DEFAULT_PERMISSIONS_BY_ROLE[role] || DEFAULT_PERMISSIONS_BY_ROLE.viewer;
  if (!Array.isArray(permissions)) {
    return [...defaults];
  }

  const normalized = [...new Set(
    permissions
      .map((value) => String(value || "").trim())
      .filter((value) => ALL_PERMISSIONS.includes(value)),
  )];

  return normalized.length ? normalized : [...defaults];
}

function hasPermission(user, permission) {
  if (!user) {
    return false;
  }

  const permissions = normalizePermissions(user.role, user.permissions);
  return permissions.includes(permission);
}

function availablePagesForUser(user) {
  return PAGE_DEFINITIONS.filter((page) => hasPermission(user, page.permission));
}

function defaultPageForUser(user) {
  return availablePagesForUser(user)[0]?.key || "dashboard";
}

function pageHeading(page) {
  switch (page) {
    case "fust":
      return {
        title: "Fust Management",
        caption: "Capture IN and OUT movements, then review balances and recent actions.",
      };
    case "clock":
      return {
        title: "Inklokken",
        caption: "Scan badge codes, add manual corrections, and export clocked times online.",
      };
    case "users":
      return {
        title: "Users",
        caption: "Control which menus and actions each account can use.",
      };
    case "settings":
      return {
        title: "Settings",
        caption: "Prepare email recipients, spreadsheet mapping, and business master data.",
      };
    default:
      return {
        title: "Sjaak vd Vijver Expedition Photo Dashboard",
        caption: "Choose departure date and optionally filter by customer code.",
      };
  }
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

function emptyFustMetrics() {
  return { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 };
}

function DocumentStatus({ action }) {
  const documentKind = action?.type === "IN" ? "fustbon" : "cmr";
  const document = action?.type === "IN" ? action?.fustbon || {} : action?.cmr || {};
  const label = action?.type === "IN" ? "Fustbon" : "CMR";
  if (!action || !["IN", "OUT"].includes(action.type)) {
    return "-";
  }
  if (document.status === "uploaded" && document.file_id) {
    const href = `/api/fust/actions/${encodeURIComponent(action.id)}/document/${documentKind}`;
    return <a href={href} target="_blank" rel="noreferrer">Open {label}</a>;
  }
  if (document.status === "uploaded" && document.web_link) {
    return <a href={document.web_link} target="_blank" rel="noreferrer">Open {label}</a>;
  }
  if (document.status === "skipped") {
    return "Skipped";
  }
  if (document.status === "failed") {
    return document.error || "Upload failed";
  }
  return "Missing";
}


function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || "").split(",").pop() || "");
    reader.onerror = () => reject(reader.error || new Error("Could not read file"));
    reader.readAsDataURL(file);
  });
}

function useFustMeta(enabled) {
  const [state, setState] = useState({ loading: true, data: null, error: "" });

  useEffect(() => {
    if (!enabled) {
      setState({ loading: false, data: null, error: "" });
      return undefined;
    }

    let cancelled = false;
    setState((current) => ({ ...current, loading: true, error: "" }));
    apiJson("/api/fust/meta")
      .then((payload) => {
        if (!cancelled) {
          setState({ loading: false, data: payload, error: "" });
        }
      })
      .catch((error) => {
        if (!cancelled) {
          setState({ loading: false, data: null, error: error.message });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [enabled]);

  return state;
}

function useFustActions(enabled) {
  const [refreshKey, setRefreshKey] = useState(0);
  const [state, setState] = useState({ loading: true, data: null, error: "" });

  useEffect(() => {
    if (!enabled) {
      setState({ loading: false, data: null, error: "" });
      return undefined;
    }

    let cancelled = false;
    setState((current) => ({ ...current, loading: true, error: "" }));
    apiJson("/api/fust/actions")
      .then((payload) => {
        if (!cancelled) {
          setState({ loading: false, data: payload, error: "" });
        }
      })
      .catch((error) => {
        if (!cancelled) {
          setState({ loading: false, data: null, error: error.message });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [enabled, refreshKey]);

  return {
    ...state,
    refresh: () => setRefreshKey((value) => value + 1),
  };
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
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [fustMenuVersion, setFustMenuVersion] = useState(0);
  const loggedIn = Boolean(auth.user);
  const syncStatus = useSyncStatus(loggedIn);
  const [selectedDate, setSelectedDate] = useState("");
  const [dateWasManuallySelected, setDateWasManuallySelected] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [expandedCustomers, setExpandedCustomers] = useState(() => new Set());
  const [lightbox, setLightbox] = useState(null);
  const [syncVersion, setSyncVersion] = useState(0);
  const requestedDate = dateWasManuallySelected ? selectedDate : "";
  const canViewPhotos = hasPermission(auth.user, PERMISSIONS.PHOTOS_VIEW);
  const visiblePages = availablePagesForUser(auth.user);
  const onPhotosPage = page === "dashboard" && canViewPhotos;
  const { data, loading, error } = useDashboardData(requestedDate, searchTerm, syncVersion, loggedIn && onPhotosPage);

  const dates = data?.dates || [];
  const activeDate = selectedDate || data?.selected_date || "";
  const firstDate = dates[0] || "";
  const lastDate = dates.at(-1) || "";
  const syncRunning = syncStatus?.state === "running";

  useEffect(() => {
    apiJson("/api/auth/me")
      .then((payload) => {
        const nextPage = defaultPageForUser(payload.user);
        setAuth({
          loading: false,
          user: payload.user,
          setupRequired: payload.setup_required,
        });
        if (payload.user) {
          setPage(nextPage);
          setSidebarOpen(true);
        }
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

  useEffect(() => {
    if (!auth.user) {
      return;
    }
    const fallbackPage = defaultPageForUser(auth.user);
    if (!visiblePages.some((item) => item.key === page)) {
      setPage(fallbackPage);
    }
  }, [auth.user, page, visiblePages]);

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
    setSidebarOpen(false);
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
          <SetupForm onSetup={(user) => {
            setAuth({ loading: false, user, setupRequired: false });
            setPage(defaultPageForUser(user));
            setSidebarOpen(true);
          }}
          />
        ) : (
          <LoginForm onLogin={(user) => {
            setAuth({ loading: false, user, setupRequired: false });
            setPage(defaultPageForUser(user));
            setSidebarOpen(true);
          }}
          />
        )}
      </AuthShell>
    );
  }

  const heading = pageHeading(page);

  return (
    <>
      <button
        type="button"
        className="sidebar-toggle"
        aria-label={sidebarOpen ? "Close menu" : "Open menu"}
        aria-expanded={sidebarOpen}
        onClick={() => setSidebarOpen((open) => !open)}
      >
        <span />
        <span />
        <span />
      </button>
      {sidebarOpen && (
        <button
          type="button"
          className="sidebar-backdrop"
          aria-label="Close menu"
          onClick={() => setSidebarOpen(false)}
        />
      )}
      <aside className={`sidebar ${sidebarOpen ? "open" : ""}`}>
        <div>
          <p className="eyebrow">SnappySjaak</p>
          <h1>Sjaak vd Vijver Expedition Shadow App</h1>
          <p className="sidebar-note">Photos stays separate while Fust gets its own protected workspace.</p>
        </div>

        <nav className="side-nav" aria-label="Shadow app pages">
          {visiblePages.map((item) => (
            <button
              key={item.key}
              className={page === item.key ? "active" : ""}
              onClick={(event) => {
                event.currentTarget.blur();
                setPage(item.key);
                if (item.key === "fust") {
                  setFustMenuVersion((value) => value + 1);
                }
                setSidebarOpen(false);
              }}
            >
              {item.label}
            </button>
          ))}
        </nav>

        {canViewPhotos && (
          <>
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
          </>
        )}

        <div className="signed-in">
          <p>Signed in as <strong>{auth.user.username}</strong></p>
          <button onClick={handleLogout}>Log out</button>
        </div>
      </aside>

      <main className="workspace">
        <header className="page-header">
          <h1>{heading.title}</h1>
          <p>{heading.caption}</p>
        </header>

        {page === "users" && <UsersPage currentUser={auth.user} />}
        {page === "fust" && <FustPage currentUser={auth.user} menuVersion={fustMenuVersion} />}
        {page === "clock" && <ClockPage currentUser={auth.user} />}
        {page === "settings" && <SettingsPage currentUser={auth.user} />}
        {page === "dashboard" && canViewPhotos && (
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
  const [form, setForm] = useState({
    username: "",
    password: "",
    role: "viewer",
    permissions: [...DEFAULT_PERMISSIONS_BY_ROLE.viewer],
  });
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
      setForm({
        username: "",
        password: "",
        role: "viewer",
        permissions: [...DEFAULT_PERMISSIONS_BY_ROLE.viewer],
      });
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
          <select
            value={form.role}
            onChange={(event) => {
              const role = event.target.value;
              setForm({ ...form, role, permissions: [...DEFAULT_PERMISSIONS_BY_ROLE[role]] });
            }}
          >
            <option value="viewer">Viewer</option>
            <option value="admin">Admin</option>
          </select>
        </label>
        <button className="primary" type="submit">Add user</button>
        <PermissionChecklist
          title="Menu access"
          permissions={form.permissions}
          onChange={(permissions) => setForm({ ...form, permissions })}
        />
      </form>

      {error && <div className="notice danger">{error}</div>}
      {message && <div className="notice">{message}</div>}

      <div className="users-list">
        {users.map((user) => (
          <UserRow
            key={user.username}
            user={user}
            currentUser={currentUser}
            onSave={(changes) => updateUser(user.username, changes)}
            onDelete={() => deleteUser(user.username)}
          />
        ))}
      </div>
    </section>
  );
}

function UserRow({ user, currentUser, onSave, onDelete }) {
  const [password, setPassword] = useState("");
  const [role, setRole] = useState(user.role);
  const [permissions, setPermissions] = useState(normalizePermissions(user.role, user.permissions));

  useEffect(() => {
    setRole(user.role);
    setPermissions(normalizePermissions(user.role, user.permissions));
  }, [user.permissions, user.role]);

  return (
    <article className="user-row">
      <div>
        <strong>{user.username}</strong>
        <span>{user.role} | created {formatTimestamp(user.created_at)}</span>
      </div>

      <div className="user-row-actions">
        <label>
          <span>Role</span>
          <select
            value={role}
            onChange={(event) => {
              const nextRole = event.target.value;
              setRole(nextRole);
              setPermissions([...DEFAULT_PERMISSIONS_BY_ROLE[nextRole]]);
            }}
          >
            <option value="viewer">Viewer</option>
            <option value="admin">Admin</option>
          </select>
        </label>
        <button onClick={() => onSave({ role, permissions })}>Save access</button>
      </div>

      <PermissionChecklist title="Allowed menus" permissions={permissions} onChange={setPermissions} />

      <form
        className="password-form"
        onSubmit={(event) => {
          event.preventDefault();
          if (password) {
            onSave({ password });
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

function PermissionChecklist({ title, permissions, onChange }) {
  return (
    <fieldset className="permission-grid">
      <legend>{title}</legend>
      {ALL_PERMISSIONS.map((permission) => (
        <label key={permission} className="permission-option">
          <input
            type="checkbox"
            checked={permissions.includes(permission)}
            onChange={(event) => {
              if (event.target.checked) {
                onChange([...permissions, permission]);
                return;
              }
              onChange(permissions.filter((item) => item !== permission));
            }}
          />
          <span>{permission}</span>
        </label>
      ))}
    </fieldset>
  );
}

function fustTileLabel(tab) {
  if (tab === "last-actions") {
    return "Last actions";
  }
  return tab.toUpperCase();
}

function FustPage({ currentUser, menuVersion }) {
  const visibleTabs = [
    hasPermission(currentUser, PERMISSIONS.FUST_IN) ? "in" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OUT) ? "out" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OVERVIEW) ? "overview" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OVERVIEW) ? "last-actions" : null,
  ].filter(Boolean);
  const [activeTab, setActiveTab] = useState("");
  const { loading: metaLoading, data: metaData, error: metaError } = useFustMeta(Boolean(currentUser));
  const { loading: actionsLoading, data: actionsData, error: actionsError, refresh } = useFustActions(Boolean(currentUser));

  useEffect(() => {
    setActiveTab("");
  }, [menuVersion]);

  useEffect(() => {
    if (activeTab && !visibleTabs.includes(activeTab)) {
      setActiveTab("");
    }
  }, [activeTab, visibleTabs]);

  if (!visibleTabs.length) {
    return <div className="notice">This account can open Fust but does not yet have an IN, OUT, or Overview action assigned.</div>;
  }

  return (
    <section className="fust-page">
      {metaError && <div className="notice danger">Unable to load Fust master data: {metaError}</div>}
      {actionsError && <div className="notice danger">Unable to load Fust actions: {actionsError}</div>}

      {!activeTab && (
        <div className="fust-tile-grid" aria-label="Fust menu">
          {visibleTabs.map((tab) => (
            <button
              key={tab}
              type="button"
              className="fust-tile"
              onClick={() => setActiveTab(tab)}
            >
              <strong>{fustTileLabel(tab)}</strong>
            </button>
          ))}
        </div>
      )}

      {activeTab && (
        <div className="section-header fust-section-nav">
          <h2>{fustTileLabel(activeTab)}</h2>
          <button type="button" onClick={() => setActiveTab("")}>Fust menu</button>
        </div>
      )}

      {activeTab === "in" && (
        <FustActionForm
          type="IN"
          metaData={metaData}
          loading={metaLoading}
          onSaved={refresh}
        />
      )}

      {activeTab === "out" && (
        <FustActionForm
          type="OUT"
          metaData={metaData}
          loading={metaLoading}
          onSaved={refresh}
        />
      )}

      {activeTab === "overview" && (
        <FustOverview
          loading={actionsLoading}
          actions={actionsData?.actions || []}
          overview={actionsData?.overview || []}
          sourceDebug={actionsData?.source_debug || null}
          onRefresh={refresh}
        />
      )}

      {activeTab === "last-actions" && (
        <FustLastActions
          loading={actionsLoading}
          actions={actionsData?.actions || []}
          onRefresh={refresh}
        />
      )}
    </section>
  );
}

function FustActionForm({ type, metaData, loading, onSaved }) {
  const [form, setForm] = useState({
    action_date: new Date().toISOString().slice(0, 10),
    country: "",
    customer_name: "",
    customer_code: "",
    connect_name: "",
    remark: "",
    metrics: emptyFustMetrics(),
  });
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const [pendingCmrAction, setPendingCmrAction] = useState(null);
  const [cmrFile, setCmrFile] = useState(null);
  const [cmrBusy, setCmrBusy] = useState(false);
  const records = metaData?.records || [];
  const countries = metaData?.countries || [];
  const customerOptions = records.filter((record) => record.country === form.country);
  const customerNames = [...new Set(customerOptions.map((record) => record.customer_name))].sort((left, right) => left.localeCompare(right));
  const connectOptions = customerOptions.filter((record) => record.customer_name === form.customer_name);
  const documentLabel = type === "IN" ? "Fustbon" : "CMR";
  const documentEndpoint = type === "IN" ? "fustbon" : "cmr";

  function resetActionEntry() {
    setForm((current) => ({
      ...current,
      remark: "",
      metrics: emptyFustMetrics(),
    }));
  }

  useEffect(() => {
    if (countries.length && !form.country) {
      setForm((current) => ({ ...current, country: countries[0] }));
    }
  }, [countries, form.country]);

  useEffect(() => {
    if (customerNames.length && !customerNames.includes(form.customer_name)) {
      setForm((current) => ({
        ...current,
        customer_name: customerNames[0] || "",
        connect_name: "",
        customer_code: "",
      }));
    }
  }, [customerNames, form.customer_name]);

  useEffect(() => {
    const activeConnect = connectOptions.find((record) => record.connect_name === form.connect_name);
    if (!activeConnect && connectOptions.length) {
      const first = connectOptions[0];
      setForm((current) => ({
        ...current,
        connect_name: first.connect_name,
        customer_code: first.customer_code,
      }));
    } else if (activeConnect) {
      setForm((current) => ({
        ...current,
        customer_code: activeConnect.customer_code,
      }));
    }
  }, [connectOptions, form.connect_name]);

  async function submit(event) {
    event.preventDefault();
    if (pendingCmrAction) {
      setError(`Finish the ${documentLabel} step first, or skip ${documentLabel}.`);
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson("/api/fust/submit", {
        method: "POST",
        body: JSON.stringify({ ...form, type }),
      });
      setMessage(
        `${type} saved. Sheet sync: ${payload.action.sheet_sync.ok ? "ok" : payload.action.sheet_sync.error}. Email: ${payload.action.email_sync.ok ? "ok" : payload.action.email_sync.error}`,
      );
      setPendingCmrAction(payload.action);
      onSaved();
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      setSaving(false);
    }
  }


  async function uploadCmr() {
    if (!pendingCmrAction || !cmrFile) {
      setError(`Choose a ${documentLabel} file first, or skip ${documentLabel}.`);
      return;
    }
    setCmrBusy(true);
    setError("");
    try {
      const contentBase64 = await fileToBase64(cmrFile);
      const payload = await apiJson(`/api/fust/actions/${encodeURIComponent(pendingCmrAction.id)}/${documentEndpoint}-upload`, {
        method: "PATCH",
        body: JSON.stringify({
          file: {
            name: cmrFile.name,
            type: cmrFile.type || "application/octet-stream",
            content_base64: contentBase64,
          },
        }),
      });
      setPendingCmrAction(null);
      setCmrFile(null);
      resetActionEntry();
      const uploadedDocument = type === "IN" ? payload.action.fustbon : payload.action.cmr;
      setMessage(`${documentLabel} uploaded: ${uploadedDocument?.file_name || cmrFile.name}`);
      onSaved();
    } catch (uploadError) {
      setError(uploadError.message);
      onSaved();
    } finally {
      setCmrBusy(false);
    }
  }

  async function skipCmr() {
    if (!pendingCmrAction) {
      return;
    }
    setCmrBusy(true);
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(pendingCmrAction.id)}/${documentEndpoint}-skip`, {
        method: "PATCH",
      });
      setPendingCmrAction(null);
      setCmrFile(null);
      resetActionEntry();
      setMessage(`${documentLabel} skipped for this ${type} action.`);
      onSaved();
    } catch (skipError) {
      setError(skipError.message);
    } finally {
      setCmrBusy(false);
    }
  }

  return (
    <div className="fust-form-layout">
      <form className="fust-form" onSubmit={submit}>
        <div className="form-grid">
          <label>
            <span>Date</span>
            <input
              type="date"
              value={form.action_date}
              onChange={(event) => setForm({ ...form, action_date: event.target.value })}
            />
          </label>
          <label>
            <span>Country</span>
            <select
              value={form.country}
              onChange={(event) => setForm({
                ...form,
                country: event.target.value,
                customer_name: "",
                connect_name: "",
                customer_code: "",
              })}
            >
              <option value="">Choose country</option>
              {countries.map((country) => <option key={country} value={country}>{country}</option>)}
            </select>
          </label>
          <label>
            <span>Klantnaam / carrier</span>
            <select
              value={form.customer_name}
              onChange={(event) => setForm({
                ...form,
                customer_name: event.target.value,
                connect_name: "",
                customer_code: "",
              })}
              disabled={!form.country}
            >
              <option value="">Choose klantnaam</option>
              {customerNames.map((name) => <option key={name} value={name}>{name}</option>)}
            </select>
          </label>
          <label>
            <span>Connect</span>
            <select
              value={form.connect_name}
              onChange={(event) => {
                const selected = connectOptions.find((record) => record.connect_name === event.target.value);
                setForm({
                  ...form,
                  connect_name: event.target.value,
                  customer_code: selected?.customer_code || "",
                });
              }}
              disabled={!form.customer_name}
            >
              <option value="">Choose connect</option>
              {connectOptions.map((option) => (
                <option key={`${option.customer_name}-${option.connect_name}`} value={option.connect_name}>
                  {option.connect_name}
                </option>
              ))}
            </select>
          </label>
          <label className="wide">
            <span>Remark</span>
            <input
              value={form.remark}
              onChange={(event) => setForm({ ...form, remark: event.target.value })}
              placeholder="alleen fusten"
            />
          </label>
        </div>

        <div className="metrics-grid">
          {Object.entries(form.metrics).map(([key, value]) => (
            <label key={key}>
              <span>{key.toUpperCase()}</span>
              <input
                type="number"
                min="0"
                value={value || ""}
                onChange={(event) => setForm({
                  ...form,
                  metrics: {
                    ...form.metrics,
                    [key]: Number(event.target.value || 0),
                  },
                })}
              />
            </label>
          ))}
        </div>

        {loading && <div className="notice">Loading Data sheet options...</div>}
        {message && <div className="notice">{message}</div>}
        {error && <div className="notice danger">{error}</div>}

        {pendingCmrAction && (
          <div className="cmr-panel">
            <h3>{documentLabel} for saved {type} action</h3>
            <input
              type="file"
              accept="image/*,.pdf"
              capture="environment"
              onChange={(event) => setCmrFile(event.target.files?.[0] || null)}
            />
            <div className="cmr-actions">
              <button type="button" className="primary" disabled={cmrBusy || !cmrFile} onClick={uploadCmr}>
                {cmrBusy ? `Saving ${documentLabel}...` : `Upload ${documentLabel}`}
              </button>
              <button type="button" disabled={cmrBusy} onClick={skipCmr}>Skip {documentLabel}</button>
            </div>
          </div>
        )}

        {!pendingCmrAction && (
          <button className="primary" type="submit" disabled={saving || loading}>
            {saving ? `Saving ${type}...` : `Save ${type}`}
          </button>
        )}
      </form>
    </div>
  );
}

function FustOverview({ loading, actions, overview, sourceDebug, onRefresh }) {
  const [selectedWeek, setSelectedWeek] = useState("");
  const [selectedFromDate, setSelectedFromDate] = useState("");
  const [selectedToDate, setSelectedToDate] = useState("");
  const [selectedCountry, setSelectedCountry] = useState("");
  const [selectedCustomer, setSelectedCustomer] = useState("");
  const [expandedCountryWeek, setExpandedCountryWeek] = useState(false);
  const [showTransactionRecords, setShowTransactionRecords] = useState(false);

  function emptyTotals() {
    return {
      out: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
      in: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
      balance: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
    };
  }

  function sumOverviewEntries(entries) {
    return entries.reduce((sum, entry) => ({
      out: {
        dc: sum.out.dc + entry.out.dc,
        cctag: sum.out.cctag + entry.out.cctag,
        dcs: sum.out.dcs + entry.out.dcs,
        dco: sum.out.dco + entry.out.dco,
        pal: sum.out.pal + entry.out.pal,
        vk: sum.out.vk + entry.out.vk,
      },
      in: {
        dc: sum.in.dc + entry.in.dc,
        cctag: sum.in.cctag + entry.in.cctag,
        dcs: sum.in.dcs + entry.in.dcs,
        dco: sum.in.dco + entry.in.dco,
        pal: sum.in.pal + entry.in.pal,
        vk: sum.in.vk + entry.in.vk,
      },
      balance: {
        dc: sum.balance.dc + entry.balance.dc,
        cctag: sum.balance.cctag + entry.balance.cctag,
        dcs: sum.balance.dcs + entry.balance.dcs,
        dco: sum.balance.dco + entry.balance.dco,
        pal: sum.balance.pal + entry.balance.pal,
        vk: sum.balance.vk + entry.balance.vk,
      },
    }), emptyTotals());
  }

  function buildOverviewFromActions(items) {
    const grouped = new Map();
    for (const action of items) {
      const key = `${action.week || ""}__${action.country}__${action.customer_name}`;
      if (!grouped.has(key)) {
        grouped.set(key, {
          week: action.week || "",
          country: action.country,
          customer_name: action.customer_name,
          out: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
          in: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
          balance: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
        });
      }
      const entry = grouped.get(key);
      const target = action.type === "OUT" ? entry.out : entry.in;
      target.dc += Number(action.metrics?.dc || 0);
      target.cctag += Number(action.metrics?.cctag || 0);
      target.dcs += Number(action.metrics?.dcs || 0);
      target.dco += Number(action.metrics?.dco || 0);
      target.pal += Number(action.metrics?.pal || 0);
      target.vk += Number(action.metrics?.vk || 0);
      entry.balance = {
        dc: entry.in.dc - entry.out.dc,
        cctag: entry.in.cctag - entry.out.cctag,
        dcs: entry.in.dcs - entry.out.dcs,
        dco: entry.in.dco - entry.out.dco,
        pal: entry.in.pal - entry.out.pal,
        vk: entry.in.vk - entry.out.vk,
      };
    }

    return [...grouped.values()].sort((left, right) => (
      String(left.customer_name).localeCompare(String(right.customer_name))
    ));
  }

  function downloadExcelFriendlyTable(filename, headers, rows) {
    const tsv = [
      headers.join("\t"),
      ...rows.map((row) => row.map((value) => String(value ?? "").replaceAll("\t", " ")).join("\t")),
    ].join("\n");
    const blob = new Blob([tsv], { type: "text/tab-separated-values;charset=utf-8;" });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);
  }

  const weekOptions = [...new Set(actions.map((action) => String(action.week || "")).filter(Boolean))]
    .sort((left, right) => Number(right) - Number(left));
  const datedActions = actions.filter((action) => {
    const actionDate = String(action.action_date || "");
    if (selectedFromDate && actionDate && actionDate < selectedFromDate) {
      return false;
    }
    if (selectedToDate && actionDate && actionDate > selectedToDate) {
      return false;
    }
    return true;
  });
  const scopedActions = datedActions.filter((action) => !selectedWeek || String(action.week || "") === selectedWeek);
  const weekFilteredOverview = buildOverviewFromActions(scopedActions);
  const countryOptions = [...new Set(weekFilteredOverview.map((entry) => entry.country).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right));
  const countryFilteredOverview = weekFilteredOverview.filter((entry) => !selectedCountry || entry.country === selectedCountry);
  const customerOptions = [...new Set(countryFilteredOverview.map((entry) => entry.customer_name).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right));

  useEffect(() => {
    if (selectedCountry && !countryOptions.includes(selectedCountry)) {
      setSelectedCountry("");
    }
  }, [countryOptions, selectedCountry]);

  useEffect(() => {
    if (selectedCustomer && !customerOptions.includes(selectedCustomer)) {
      setSelectedCustomer("");
      setShowTransactionRecords(false);
    }
  }, [customerOptions, selectedCustomer]);

  const countryRows = useMemo(() => {
    const grouped = new Map();
    for (const entry of weekFilteredOverview) {
      if (!grouped.has(entry.country)) {
        grouped.set(entry.country, []);
      }
      grouped.get(entry.country).push(entry);
    }

    return [...grouped.entries()]
      .map(([country, entries]) => ({
        week: selectedWeek || "All weeks",
        country,
        customer_name: `${entries.length} cust/transports`,
        ...sumOverviewEntries(entries),
        row_count: entries.length,
      }))
      .sort((left, right) => left.country.localeCompare(right.country));
  }, [selectedWeek, weekFilteredOverview]);

  const weekRows = useMemo(() => {
    const grouped = new Map();
    for (const entry of countryFilteredOverview) {
      const week = String(entry.week || "");
      if (!grouped.has(week)) {
        grouped.set(week, []);
      }
      grouped.get(week).push(entry);
    }

    return [...grouped.entries()]
      .map(([week, entries]) => ({
        week,
        country: selectedCountry,
        customer_name: `${entries.length} cust/transports`,
        ...sumOverviewEntries(entries),
        row_count: entries.length,
      }))
      .sort((left, right) => Number(right.week || 0) - Number(left.week || 0));
  }, [countryFilteredOverview, selectedCountry]);

  const customerRows = countryFilteredOverview
    .sort((left, right) => left.customer_name.localeCompare(right.customer_name));
  const collapsedCountryWeekRows = selectedCountry && selectedWeek ? weekRows : [];
  const scopedOverview = selectedCountry
    ? (selectedWeek && !expandedCountryWeek ? collapsedCountryWeekRows : (selectedWeek ? customerRows : weekRows))
    : countryRows;
  const visibleOverview = selectedCustomer
    ? customerRows.filter((entry) => entry.customer_name === selectedCustomer)
    : scopedOverview;
  const totals = sumOverviewEntries(visibleOverview);
  const transactionRecords = datedActions
    .filter((action) => showTransactionRecords)
    .filter((action) => selectedWeek && String(action.week || "") === selectedWeek)
    .filter((action) => selectedCountry && action.country === selectedCountry)
    .filter((action) => selectedCustomer && action.customer_name === selectedCustomer)
    .sort((left, right) => {
      const rightDate = String(right.action_date || right.created_at || "");
      const leftDate = String(left.action_date || left.created_at || "");
      return rightDate.localeCompare(leftDate);
    });
  const exportRows = visibleOverview.map((entry, index) => [
    selectedCountry && selectedWeek && !selectedCustomer && index > 0 ? "" : (entry.week || ""),
    selectedCountry && selectedWeek && !selectedCustomer && index > 0 ? "" : entry.country,
    entry.customer_name,
    entry.out.dc,
    entry.in.dc,
    entry.balance.dc,
    entry.out.cctag,
    entry.in.cctag,
    entry.balance.cctag,
    entry.out.dcs,
    entry.in.dcs,
    entry.balance.dcs,
    entry.out.dco,
    entry.in.dco,
    entry.balance.dco,
    entry.out.vk,
    entry.in.vk,
    entry.balance.vk,
    entry.out.pal,
  ]);

  if (loading) {
    return <div className="notice">Loading Fust overview...</div>;
  }

  return (
    <div className="overview-stack">
      <div className="data-table-card">
        <div className="overview-filters">
          <label>
            <span>Week</span>
            <select
              value={selectedWeek}
              onChange={(event) => {
                setSelectedWeek(event.target.value);
                setSelectedCustomer("");
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            >
              <option value="">All weeks</option>
              {weekOptions.map((week) => <option key={week} value={week}>{week}</option>)}
            </select>
          </label>
          <label>
            <span>From date</span>
            <input
              type="date"
              value={selectedFromDate}
              onChange={(event) => {
                setSelectedFromDate(event.target.value);
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            />
          </label>
          <label>
            <span>To date</span>
            <input
              type="date"
              value={selectedToDate}
              onChange={(event) => {
                setSelectedToDate(event.target.value);
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            />
          </label>
          <label>
            <span>Country</span>
            <select
              value={selectedCountry}
              onChange={(event) => {
                setSelectedCountry(event.target.value);
                setSelectedCustomer("");
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            >
              <option value="">All countries</option>
              {countryOptions.map((country) => <option key={country} value={country}>{country}</option>)}
            </select>
          </label>
          <label>
            <span>Cust/transport</span>
            <select
              value={selectedCustomer}
              onChange={(event) => {
                setSelectedCustomer(event.target.value);
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
              disabled={!selectedCountry}
            >
              <option value="">All cust/transports</option>
              {customerOptions.map((customer) => <option key={customer} value={customer}>{customer}</option>)}
            </select>
          </label>
        </div>
        <div className="section-header">
          <h2>
            {!selectedCountry && "Country totals"}
            {selectedCountry && !selectedWeek && "Week totals"}
            {selectedCountry && selectedWeek && !expandedCountryWeek && !selectedCustomer && "Country total"}
            {selectedCountry && selectedWeek && (expandedCountryWeek || selectedCustomer) && "Customer totals"}
          </h2>
          <button
            type="button"
            onClick={() => downloadExcelFriendlyTable(
              `fust-overview-${selectedWeek || "all-weeks"}-${selectedCountry || "all-countries"}-${selectedCustomer || "active-table"}.tsv`,
              ["Week", "Country", "Cust/transport", "DC out", "DC in", "DC balance", "CCTag out", "CCTag in", "CCTag balance", "DCS out", "DCS in", "DCS balance", "DCO out", "DCO in", "DCO balance", "VK out", "VK in", "VK balance", "pal out"],
              exportRows,
            )}
            disabled={!visibleOverview.length}
          >
            Export active table
          </button>
        </div>
        <div className="table-wrap">
          <table className="data-table balance-table">
            <thead>
              <tr>
                <th>Week</th>
                <th>Country</th>
                <th>Cust/transport</th>
                <th>DC out</th>
                <th>DC in</th>
                <th>DC balance</th>
                <th>CCTag out</th>
                <th>CCTag in</th>
                <th>CCTag balance</th>
                <th>DCS out</th>
                <th>DCS in</th>
                <th>DCS balance</th>
                <th>DCO out</th>
                <th>DCO in</th>
                <th>DCO balance</th>
                <th>VK out</th>
                <th>VK in</th>
                <th>VK balance</th>
                <th>pal out</th>
              </tr>
            </thead>
            <tbody>
              {visibleOverview.map((entry) => (
                <tr
                  key={`${entry.week}-${entry.country}-${entry.customer_name}`}
                  className={selectedCustomer === entry.customer_name ? "clickable-row active-row" : "clickable-row"}
                  onClick={() => {
                    if (!selectedCountry) {
                      setSelectedCountry(entry.country);
                      setSelectedCustomer("");
                      setShowTransactionRecords(false);
                      return;
                    }
                    if (selectedCountry && !selectedWeek) {
                      setSelectedWeek(String(entry.week || ""));
                      setSelectedCustomer("");
                      setExpandedCountryWeek(true);
                      setShowTransactionRecords(false);
                      return;
                    }
                    if (selectedCountry && selectedWeek && !expandedCountryWeek && !selectedCustomer) {
                      setExpandedCountryWeek(true);
                      setShowTransactionRecords(false);
                      return;
                    }
                    if (selectedCountry && selectedWeek && selectedCustomer) {
                      setShowTransactionRecords((current) => !current);
                      return;
                    }
                    setSelectedCustomer(entry.customer_name);
                    setShowTransactionRecords(false);
                  }}
                >
                  <td>{entry.week || ""}</td>
                  <td>{entry.country}</td>
                  <td>{entry.customer_name}</td>
                  <td>{entry.out.dc}</td>
                  <td>{entry.in.dc}</td>
                  <td>{entry.balance.dc}</td>
                  <td>{entry.out.cctag}</td>
                  <td>{entry.in.cctag}</td>
                  <td>{entry.balance.cctag}</td>
                  <td>{entry.out.dcs}</td>
                  <td>{entry.in.dcs}</td>
                  <td>{entry.balance.dcs}</td>
                  <td>{entry.out.dco}</td>
                  <td>{entry.in.dco}</td>
                  <td>{entry.balance.dco}</td>
                  <td>{entry.out.vk}</td>
                  <td>{entry.in.vk}</td>
                  <td>{entry.balance.vk}</td>
                  <td>{entry.out.pal}</td>
                </tr>
              ))}
              {!!visibleOverview.length && (
                <tr className="summary-row">
                  <td colSpan="3"><strong>Total</strong></td>
                  <td><strong>{totals.out.dc}</strong></td>
                  <td><strong>{totals.in.dc}</strong></td>
                  <td><strong>{totals.balance.dc}</strong></td>
                  <td><strong>{totals.out.cctag}</strong></td>
                  <td><strong>{totals.in.cctag}</strong></td>
                  <td><strong>{totals.balance.cctag}</strong></td>
                  <td><strong>{totals.out.dcs}</strong></td>
                  <td><strong>{totals.in.dcs}</strong></td>
                  <td><strong>{totals.balance.dcs}</strong></td>
                  <td><strong>{totals.out.dco}</strong></td>
                  <td><strong>{totals.in.dco}</strong></td>
                  <td><strong>{totals.balance.dco}</strong></td>
                  <td><strong>{totals.out.vk}</strong></td>
                  <td><strong>{totals.in.vk}</strong></td>
                  <td><strong>{totals.balance.vk}</strong></td>
                  <td><strong>{totals.out.pal}</strong></td>
                </tr>
              )}
              {!visibleOverview.length && (
                <tr>
                  <td colSpan="19">No overview rows found for the selected filters.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        {showTransactionRecords && selectedWeek && selectedCountry && selectedCustomer && (
          <div className="transaction-detail">
            <div className="section-header">
              <h3>Transactions</h3>
              <button type="button" onClick={() => setShowTransactionRecords(false)}>Close</button>
            </div>
            <div className="table-wrap">
              <table className="data-table action-table">
                <thead>
                  <tr>
                    <th>Type</th>
                    <th>Date</th>
                    <th>Country</th>
                    <th>Cust/transport</th>
                    <th>Connect</th>
                    <th>DC</th>
                    <th>CCTag</th>
                    <th>DCS</th>
                    <th>DCO</th>
                    <th>PAL</th>
                    <th>VK</th>
                    <th>Document</th>
                    <th>Remark</th>
                  </tr>
                </thead>
                <tbody>
                  {transactionRecords.map((action) => (
                    <tr key={action.id}>
                      <td>{action.type}</td>
                      <td>{action.action_date}</td>
                      <td>{action.country}</td>
                      <td>{action.customer_name}</td>
                      <td>{action.connect_name}</td>
                      <td>{action.metrics?.dc || 0}</td>
                      <td>{action.metrics?.cctag || 0}</td>
                      <td>{action.metrics?.dcs || 0}</td>
                      <td>{action.metrics?.dco || 0}</td>
                      <td>{action.metrics?.pal || 0}</td>
                      <td>{action.metrics?.vk || 0}</td>
                      <td><DocumentStatus action={action} /></td>
                      <td>{action.remark || "-"}</td>
                    </tr>
                  ))}
                  {!transactionRecords.length && (
                    <tr>
                      <td colSpan="13">No transactions were found for this filter.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function FustLastActions({ loading, actions, onRefresh }) {
  const [busyActionId, setBusyActionId] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [editingActionId, setEditingActionId] = useState("");
  const [editForm, setEditForm] = useState(null);

  function startEdit(action) {
    setEditingActionId(action.id);
    setEditForm({
      type: action.type,
      action_date: action.action_date || "",
      country: action.country || "",
      customer_name: action.customer_name || "",
      connect_name: action.connect_name || "",
      customer_code: action.customer_code || "",
      remark: action.remark || "",
      metrics: {
        dc: Number(action.metrics?.dc || 0),
        cctag: Number(action.metrics?.cctag || 0),
        dcs: Number(action.metrics?.dcs || 0),
        dco: Number(action.metrics?.dco || 0),
        pal: Number(action.metrics?.pal || 0),
        vk: Number(action.metrics?.vk || 0),
      },
    });
  }

  function cancelEdit() {
    setEditingActionId("");
    setEditForm(null);
  }

  async function saveEdit(actionId) {
    setBusyActionId(`${actionId}:save`);
    setMessage("");
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(actionId)}`, {
        method: "PUT",
        body: JSON.stringify(editForm),
      });
      setMessage("Action updated.");
      cancelEdit();
      onRefresh();
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setBusyActionId("");
    }
  }

  async function retryAction(actionId, kind) {
    setBusyActionId(`${actionId}:${kind}`);
    setMessage("");
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(actionId)}/${kind}`, {
        method: "POST",
      });
      setMessage(kind === "retry-sheet" ? "Sheet sync retried." : "Email resend retried.");
      onRefresh();
    } catch (retryError) {
      setError(retryError.message);
    } finally {
      setBusyActionId("");
    }
  }

  async function deleteLocalAction(actionId) {
    setBusyActionId(`${actionId}:delete`);
    setMessage("");
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(actionId)}`, {
        method: "DELETE",
      });
      setMessage("Local action deleted.");
      onRefresh();
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setBusyActionId("");
    }
  }

  if (loading) {
    return <div className="notice">Loading last actions...</div>;
  }

  return (
    <div className="overview-stack">
      {message && <div className="notice">{message}</div>}
      {error && <div className="notice danger">{error}</div>}

      <div className="data-table-card">
        <div className="section-header">
          <h2>Last actions</h2>
        </div>
        <div className="table-wrap">
          <table className="data-table action-table">
            <thead>
              <tr>
                <th>Type</th>
                <th>Date</th>
                <th>Country</th>
                <th>Klantnaam</th>
                <th>Connect</th>
                <th>DC</th>
                <th>DCS</th>
                <th>DCO</th>
                <th>CCTag</th>
                <th>PAL</th>
                <th>VK</th>
                <th>Document</th>
                <th>Remark</th>
                <th>Sheet</th>
                <th>Email</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {actions.map((action) => {
                const isEditing = editingActionId === action.id && editForm;
                const canModify = action.created_by !== "spreadsheet";
                return (
                  <tr key={action.id}>
                    <td>
                      {isEditing ? (
                        <select value={editForm.type} onChange={(event) => setEditForm({ ...editForm, type: event.target.value })}>
                          <option value="IN">IN</option>
                          <option value="OUT">OUT</option>
                        </select>
                      ) : action.type}
                    </td>
                    <td>
                      {isEditing ? (
                        <input type="date" value={editForm.action_date} onChange={(event) => setEditForm({ ...editForm, action_date: event.target.value })} />
                      ) : action.action_date}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.country} onChange={(event) => setEditForm({ ...editForm, country: event.target.value })} />
                      ) : action.country}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.customer_name} onChange={(event) => setEditForm({ ...editForm, customer_name: event.target.value })} />
                      ) : action.customer_name}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.connect_name} onChange={(event) => setEditForm({ ...editForm, connect_name: event.target.value })} />
                      ) : action.connect_name}
                    </td>
                    {["dc", "dcs", "dco", "cctag", "pal", "vk"].map((metric) => (
                      <td key={metric}>
                        {isEditing ? (
                          <input
                            className="metric-edit"
                            type="number"
                            min="0"
                            value={editForm.metrics[metric]}
                            onChange={(event) => setEditForm({
                              ...editForm,
                              metrics: {
                                ...editForm.metrics,
                                [metric]: Number(event.target.value || 0),
                              },
                            })}
                          />
                        ) : (action.metrics?.[metric] || 0)}
                      </td>
                    ))}
                    <td><DocumentStatus action={action} /></td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.remark} onChange={(event) => setEditForm({ ...editForm, remark: event.target.value })} />
                      ) : (action.remark || "-")}
                    </td>
                    <td>{action.sheet_sync?.ok ? "ok" : action.sheet_sync?.error || "-"}</td>
                    <td>{action.email_sync?.ok ? "ok" : action.email_sync?.error || "-"}</td>
                    <td>
                      <div className="retry-actions">
                        {isEditing ? (
                          <>
                            <button
                              type="button"
                              disabled={busyActionId === `${action.id}:save`}
                              onClick={() => saveEdit(action.id)}
                            >
                              Save
                            </button>
                            <button type="button" onClick={cancelEdit}>Cancel</button>
                          </>
                        ) : canModify && (
                          <button type="button" onClick={() => startEdit(action)}>Edit</button>
                        )}
                        {!action.sheet_sync?.ok && (
                          <button
                            type="button"
                            disabled={busyActionId === `${action.id}:retry-sheet`}
                            onClick={() => retryAction(action.id, "retry-sheet")}
                          >
                            Retry sheet
                          </button>
                        )}
                        {!action.email_sync?.ok && (
                          <button
                            type="button"
                            disabled={busyActionId === `${action.id}:retry-email`}
                            onClick={() => retryAction(action.id, "retry-email")}
                          >
                            Retry email
                          </button>
                        )}
                        {canModify && !isEditing && (
                          <button
                            type="button"
                            disabled={busyActionId === `${action.id}:delete`}
                            onClick={() => deleteLocalAction(action.id)}
                          >
                            Delete local
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {!actions.length && (
                <tr>
                  <td colSpan="16">No actions were found.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function cmrFolderLines(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return "";
  }
  return Object.entries(value)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([country, folderId]) => `${country}=${folderId}`)
    .join("\n");
}

function parseCmrFolderLines(value) {
  return Object.fromEntries(
    String(value || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        const [country, ...folderParts] = line.split("=");
        return [String(country || "").trim().toUpperCase(), folderParts.join("=").trim()];
      })
      .filter(([country, folderId]) => country && folderId),
  );
}


function todayInputValue() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function timeInputValue() {
  return new Date().toLocaleTimeString("nl-NL", { hour12: false }).slice(0, 5);
}

function employeeOptionLabel(employee) {
  return `${employee.name} (${employee.tbnr})`;
}

function workedTimeToMinutes(value) {
  const match = String(value || "").match(/^(\d+):(\d{2})$/);
  if (!match) {
    return 0;
  }
  return (Number(match[1]) * 60) + Number(match[2]);
}

function minutesToWorkedTime(minutes) {
  const rounded = Math.max(0, Math.round(minutes));
  const hours = Math.floor(rounded / 60);
  const remainder = rounded % 60;
  return `${hours}:${String(remainder).padStart(2, "0")}`;
}

function ClockPage({ currentUser }) {
  const canManage = hasPermission(currentUser, PERMISSIONS.CLOCK_MANAGE);
  const [activeTab, setActiveTab] = useState("clock");
  const [employees, setEmployees] = useState([]);
  const [records, setRecords] = useState([]);
  const [sessions, setSessions] = useState([]);
  const [selectedDate, setSelectedDate] = useState(todayInputValue());
  const [exportFrom, setExportFrom] = useState(todayInputValue());
  const [exportTo, setExportTo] = useState(todayInputValue());
  const [scanCode, setScanCode] = useState("");
  const [manual, setManual] = useState({ employeeKey: "", action_date: todayInputValue(), action_time: timeInputValue(), direction: "IN" });
  const [editingId, setEditingId] = useState("");
  const [editForm, setEditForm] = useState({});
  const [loading, setLoading] = useState(false);
  const [employeeLoading, setEmployeeLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [nowLabel, setNowLabel] = useState(timeInputValue());
  const scanInputRef = useRef(null);

  async function loadEmployees() {
    setEmployeeLoading(true);
    try {
      const payload = await apiJson("/api/clock/employees");
      setEmployees(payload.employees || []);
    } catch (loadError) {
      setError(loadError.message);
    } finally {
      setEmployeeLoading(false);
    }
  }

  async function loadRecords(date = selectedDate) {
    setLoading(true);
    try {
      const payload = await apiJson(`/api/clock/records?date=${encodeURIComponent(date)}`);
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
    } catch (loadError) {
      setError(loadError.message);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadEmployees();
  }, []);

  useEffect(() => {
    loadRecords(selectedDate);
  }, [selectedDate]);

  useEffect(() => {
    const interval = window.setInterval(() => setNowLabel(timeInputValue()), 1000);
    return () => window.clearInterval(interval);
  }, []);

  const employeeByKey = useMemo(() => {
    const map = new Map();
    for (const employee of employees) {
      map.set(employee.tbnr, employee);
      map.set(employeeOptionLabel(employee), employee);
    }
    return map;
  }, [employees]);

  const selectedDateTotalWorked = useMemo(() => {
    const totalMinutes = sessions.reduce((total, session) => total + workedTimeToMinutes(session.row?.[6]), 0);
    return minutesToWorkedTime(totalMinutes);
  }, [sessions]);

  async function submitScan(event) {
    event.preventDefault();
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const scanDate = todayInputValue();
      const payload = await apiJson("/api/clock/scan", {
        method: "POST",
        body: JSON.stringify({ code: scanCode, action_date: scanDate, action_time: `${timeInputValue()}:00` }),
      });
      setSelectedDate(scanDate);
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
      setScanCode("");
      setMessage(`${payload.record.name}: ${payload.record.action_time} ${payload.record.direction}`);
      window.setTimeout(() => scanInputRef.current?.focus(), 0);
    } catch (scanError) {
      setError(scanError.message);
    } finally {
      setBusy(false);
    }
  }

  async function submitManual(event) {
    event.preventDefault();
    if (!canManage) {
      return;
    }
    const employee = employeeByKey.get(manual.employeeKey);
    if (!employee) {
      setError("Choose a valid employee");
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/clock/records", {
        method: "POST",
        body: JSON.stringify({
          employee,
          action_date: manual.action_date,
          action_time: manual.action_time.length === 5 ? `${manual.action_time}:00` : manual.action_time,
          direction: manual.direction,
        }),
      });
      setSelectedDate(manual.action_date);
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
      setMessage(`Manual ${manual.direction} saved for ${employee.name}`);
    } catch (manualError) {
      setError(manualError.message);
    } finally {
      setBusy(false);
    }
  }

  function startEdit(record) {
    setEditingId(record.id);
    setEditForm({
      action_date: record.action_date,
      action_time: String(record.action_time || "").slice(0, 5),
      employeeKey: record.tbnr,
      direction: record.direction,
    });
  }

  async function saveEdit(record) {
    const employee = employeeByKey.get(editForm.employeeKey) || {
      tbnr: record.tbnr,
      name: record.name,
      type: record.employee_type,
    };
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson(`/api/clock/records/${encodeURIComponent(record.id)}`, {
        method: "PATCH",
        body: JSON.stringify({
          action_date: editForm.action_date,
          action_time: editForm.action_time.length === 5 ? `${editForm.action_time}:00` : editForm.action_time,
          tbnr: employee.tbnr,
          name: employee.name,
          employee_type: employee.type || employee.employee_type || "",
          direction: editForm.direction,
        }),
      });
      setEditingId("");
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
      setSelectedDate(editForm.action_date);
      setMessage("Clock record updated");
    } catch (editError) {
      setError(editError.message);
    } finally {
      setBusy(false);
    }
  }

  async function deleteRecord(record) {
    if (!window.confirm(`Delete ${record.name} ${record.action_time} ${record.direction}?`)) {
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    try {
      await apiJson(`/api/clock/records/${encodeURIComponent(record.id)}`, { method: "DELETE" });
      await loadRecords(selectedDate);
      setMessage("Clock record deleted");
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setBusy(false);
    }
  }

  async function refreshDateFromBackup() {
    const date = manual.action_date || selectedDate;
    if (!window.confirm(`Replace app records for ${date} with the spreadsheet backup rows for that date?`)) {
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/clock/records/import-backup", {
        method: "POST",
        body: JSON.stringify({ date }),
      });
      setSelectedDate(date);
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
      setMessage(`Backup refreshed for ${date}: ${payload.imported_count || 0} records loaded`);
    } catch (backupError) {
      setError(backupError.message);
    } finally {
      setBusy(false);
    }
  }

  const exportUrl = `/api/clock/records/export?from=${encodeURIComponent(exportFrom)}&to=${encodeURIComponent(exportTo)}`;

  return (
    <section className="overview-stack clock-page">
      {canManage && (
        <div className="tab-strip clock-tabs">
          <button type="button" className={activeTab === "clock" ? "active" : ""} onClick={() => setActiveTab("clock")}>Clock</button>
          <button type="button" className={activeTab === "manual" ? "active" : ""} onClick={() => setActiveTab("manual")}>Manual</button>
          <button type="button" className={activeTab === "export" ? "active" : ""} onClick={() => setActiveTab("export")}>Export</button>
        </div>
      )}

      {error && <div className="notice danger">{error}</div>}
      {message && <div className="notice clock-result">{message}</div>}

      {activeTab === "clock" && (
        <div className="clock-hero">
          <div className="clock-face">
            <strong>{nowLabel}</strong>
          </div>
          <form className="clock-scan-surface" onSubmit={submitScan}>
            <label>
              <span>Badge code</span>
              <input
                ref={scanInputRef}
                value={scanCode}
                onChange={(event) => setScanCode(event.target.value.toUpperCase())}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    submitScan(event);
                  }
                }}
                placeholder="Scan badge"
                autoFocus
              />
            </label>
          </form>
        </div>
      )}

      {activeTab === "manual" && canManage && (
        <div className="data-table-card">
          <h2>Manual correction</h2>
          <form className="clock-manual-form" onSubmit={submitManual}>
            <label className="wide">
              <span>Employee</span>
              <input
                list="clock-employees"
                value={manual.employeeKey}
                onChange={(event) => setManual({ ...manual, employeeKey: event.target.value })}
                placeholder="Name or badge"
              />
              <datalist id="clock-employees">
                {employees.map((employee) => (
                  <option key={employee.tbnr} value={employeeOptionLabel(employee)} />
                ))}
              </datalist>
            </label>
            <label>
              <span>Date</span>
              <input
                type="date"
                value={manual.action_date}
                onChange={(event) => {
                  setManual({ ...manual, action_date: event.target.value });
                  setSelectedDate(event.target.value);
                }}
              />
            </label>
            <label>
              <span>Time</span>
              <input type="time" value={manual.action_time} onChange={(event) => setManual({ ...manual, action_time: event.target.value })} />
            </label>
            <label>
              <span>Direction</span>
              <select value={manual.direction} onChange={(event) => setManual({ ...manual, direction: event.target.value })}>
                <option value="IN">IN</option>
                <option value="OUT">OUT</option>
              </select>
            </label>
            <button className="primary" type="submit" disabled={busy}>Add manual</button>
          </form>

          <div className="clock-manual-overview">
            <div className="section-header">
              <h2>Selected date overview</h2>
              <div className="clock-overview-actions">
                <strong>{selectedDateTotalWorked}</strong>
                <button type="button" onClick={refreshDateFromBackup} disabled={busy}>Refresh from backup</button>
              </div>
            </div>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Time</th>
                    <th>TBNR</th>
                    <th>Name</th>
                    <th>Type</th>
                    <th>Direction</th>
                    <th>Source</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {records.map((record) => {
                    const isEditing = editingId === record.id;
                    return (
                      <tr key={record.id}>
                        <td>
                          {isEditing ? (
                            <input type="time" value={editForm.action_time || ""} onChange={(event) => setEditForm({ ...editForm, action_time: event.target.value })} />
                          ) : record.action_time}
                        </td>
                        <td>{isEditing ? <input value={editForm.employeeKey || ""} onChange={(event) => setEditForm({ ...editForm, employeeKey: event.target.value.toUpperCase() })} /> : record.tbnr}</td>
                        <td>{record.name}</td>
                        <td>{record.employee_type}</td>
                        <td>
                          {isEditing ? (
                            <select value={editForm.direction || "IN"} onChange={(event) => setEditForm({ ...editForm, direction: event.target.value })}>
                              <option value="IN">IN</option>
                              <option value="OUT">OUT</option>
                            </select>
                          ) : record.direction}
                        </td>
                        <td>{record.source}</td>
                        <td className="row-actions">
                          {isEditing ? (
                            <>
                              <button type="button" onClick={() => saveEdit(record)} disabled={busy}>Save</button>
                              <button type="button" onClick={() => setEditingId("")} disabled={busy}>Cancel</button>
                            </>
                          ) : (
                            <>
                              <button type="button" onClick={() => startEdit(record)} disabled={busy}>Edit</button>
                              <button type="button" onClick={() => deleteRecord(record)} disabled={busy}>Delete</button>
                            </>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                  {!records.length && !loading && (
                    <tr>
                      <td colSpan="7">No clock records for this date.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {activeTab === "export" && (
        <div className="data-table-card">
          <div className="section-header">
            <h2>Export and records</h2>
            <button type="button" onClick={() => loadRecords(selectedDate)} disabled={loading}>{loading ? "Loading..." : "Refresh"}</button>
          </div>
          <div className="clock-export-controls">
            <label>
              <span>From date</span>
              <input type="date" value={exportFrom} onChange={(event) => setExportFrom(event.target.value)} />
            </label>
            <label>
              <span>To date</span>
              <input type="date" value={exportTo} onChange={(event) => setExportTo(event.target.value)} />
            </label>
            <a className="button-link" href={exportUrl}>Export range</a>
            <label>
              <span>Show day</span>
              <input type="date" value={selectedDate} onChange={(event) => setSelectedDate(event.target.value)} />
            </label>
          </div>
          {loading && <div className="notice">Loading clock records...</div>}
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>TBNR</th>
                  <th>Name</th>
                  <th>Type</th>
                  <th>IN</th>
                  <th>OUT</th>
                  <th>Worked</th>
                  <th>Source</th>
                </tr>
              </thead>
              <tbody>
                {sessions.map((session, index) => {
                const row = session.row || [];
                return (
                  <tr key={`${row[0]}-${row[1]}-${row[4]}-${row[5]}-${index}`}>
                    <td>{row[0] || "-"}</td>
                    <td>{row[1] || "-"}</td>
                    <td>{row[2] || "-"}</td>
                    <td>{row[3] || "-"}</td>
                    <td>{row[4] || "-"}</td>
                    <td>{row[5] || "-"}</td>
                    <td>{row[6] || "-"}</td>
                    <td>{row[7] || "-"}</td>
                  </tr>
                );
              })}
              {!sessions.length && !loading && (
                  <tr>
                    <td colSpan="8">No clock records for this date.</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </section>
  );
}

function SettingsPage({ currentUser }) {
  const [form, setForm] = useState(null);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const [saving, setSaving] = useState(false);
  const { loading: metaLoading, data: metaData, error: metaError } = useFustMeta(Boolean(currentUser));
  const [backups, setBackups] = useState([]);
  const [backupBusy, setBackupBusy] = useState(false);
  const [connectionTest, setConnectionTest] = useState(null);
  const [connectionBusy, setConnectionBusy] = useState(false);
  const [googleBusy, setGoogleBusy] = useState(false);

  async function loadBackups() {
    const payload = await apiJson("/api/fust/backups");
    setBackups(payload.backups || []);
  }

  async function loadConnectionTest() {
    setConnectionBusy(true);
    try {
      const payload = await apiJson("/api/fust/connection-test");
      setConnectionTest(payload);
    } catch (connectionError) {
      setConnectionTest({
        account: { client_email: "", project_id: "" },
        spreadsheet_id: form?.spreadsheet_id || "",
        sheet_name: form?.data_sheet_name || "",
        read_ok: false,
        row_count: 0,
        headers: [],
        error: connectionError.message,
      });
    } finally {
      setConnectionBusy(false);
    }
  }

  useEffect(() => {
    apiJson("/api/fust/settings")
      .then((payload) => setForm({
        ...payload.settings,
        cmr_country_folders_text: cmrFolderLines(payload.settings.cmr_country_folders),
      }))
      .catch((settingsError) => setError(settingsError.message));

    loadBackups().catch((backupError) => setError(backupError.message));
    loadConnectionTest().catch(() => {});
  }, []);

  async function submit(event) {
    event.preventDefault();
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/fust/settings", {
        method: "PATCH",
        body: JSON.stringify({
          ...form,
          email_recipients: String(form.email_recipients || "")
            .split(/[\n,;]/)
            .map((value) => value.trim())
            .filter(Boolean),
          cmr_country_folders: parseCmrFolderLines(form.cmr_country_folders_text),
        }),
      });
      setForm({
        ...payload.settings,
        email_recipients: payload.settings.email_recipients.join("\n"),
        cmr_country_folders_text: cmrFolderLines(payload.settings.cmr_country_folders),
      });
      setMessage("Fust settings saved.");
      await loadConnectionTest();
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  async function createBackup() {
    setBackupBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/fust/backups", {
        method: "POST",
      });
      setBackups(payload.backups || []);
      setMessage(`Backup created: ${payload.backup?.filename || "snapshot saved"}`);
    } catch (backupError) {
      setError(backupError.message);
    } finally {
      setBackupBusy(false);
    }
  }


  async function connectGoogleDrive() {
    setGoogleBusy(true);
    setError("");
    setMessage("");
    try {
      const saved = await apiJson("/api/fust/settings", {
        method: "PATCH",
        body: JSON.stringify({
          ...form,
          email_recipients: String(form.email_recipients || "")
            .split(/[\n,;]/)
            .map((value) => value.trim())
            .filter(Boolean),
          cmr_country_folders: parseCmrFolderLines(form.cmr_country_folders_text),
        }),
      });
      setForm({
        ...saved.settings,
        email_recipients: saved.settings.email_recipients.join("\n"),
        cmr_country_folders_text: cmrFolderLines(saved.settings.cmr_country_folders),
      });
      const payload = await apiJson("/api/fust/google/auth-url");
      window.location.href = payload.auth_url;
    } catch (googleError) {
      setError(googleError.message);
      setGoogleBusy(false);
    }
  }

  if (!form) {
    return <div className="notice">Loading settings...</div>;
  }

  return (
    <section className="settings-page">
      <form className="settings-form" onSubmit={submit}>
        <div className="form-grid">
          <label>
            <span>Spreadsheet ID</span>
            <input
              value={form.spreadsheet_id}
              onChange={(event) => setForm({ ...form, spreadsheet_id: event.target.value })}
            />
          </label>
          <label>
            <span>Data tab</span>
            <input
              value={form.data_sheet_name}
              onChange={(event) => setForm({ ...form, data_sheet_name: event.target.value })}
            />
          </label>
          <label>
            <span>IN tab</span>
            <input
              value={form.in_sheet_name}
              onChange={(event) => setForm({ ...form, in_sheet_name: event.target.value })}
            />
          </label>
          <label>
            <span>OUT tab</span>
            <input
              value={form.out_sheet_name}
              onChange={(event) => setForm({ ...form, out_sheet_name: event.target.value })}
            />
          </label>
          <label>
            <span>Dashboard tab</span>
            <input
              value={form.dashboard_sheet_name}
              onChange={(event) => setForm({ ...form, dashboard_sheet_name: event.target.value })}
            />
          </label>
          <label className="wide">
            <span>Clock spreadsheet ID</span>
            <input
              value={form.clock_spreadsheet_id || ""}
              onChange={(event) => setForm({ ...form, clock_spreadsheet_id: event.target.value })}
              placeholder="Spreadsheet ID for badges and backup"
            />
          </label>
          <label>
            <span>Clock employee tab</span>
            <input
              value={form.clock_employee_sheet_name || ""}
              onChange={(event) => setForm({ ...form, clock_employee_sheet_name: event.target.value })}
              placeholder="badges"
            />
          </label>
          <label>
            <span>Clock records tab</span>
            <input
              value={form.clock_records_sheet_name || ""}
              onChange={(event) => setForm({ ...form, clock_records_sheet_name: event.target.value })}
              placeholder="backup"
            />
          </label>
          <label className="wide">
            <span>Document country folder IDs</span>
            <textarea
              value={form.cmr_country_folders_text || ""}
              onChange={(event) => setForm({ ...form, cmr_country_folders_text: event.target.value })}
              rows={8}
              placeholder={"FR=google-drive-folder-id\nPT=google-drive-folder-id\nGB=google-drive-folder-id"}
            />
          </label>
          <label className="wide">
            <span>Document fallback folder ID</span>
            <input
              value={form.cmr_fallback_folder_id || ""}
              onChange={(event) => setForm({ ...form, cmr_fallback_folder_id: event.target.value })}
              placeholder="optional fallback folder id"
            />
          </label>
          <label>
            <span>Google OAuth client ID</span>
            <input
              value={form.cmr_google_client_id || ""}
              onChange={(event) => setForm({ ...form, cmr_google_client_id: event.target.value })}
              placeholder="client id"
            />
          </label>
          <label>
            <span>Google OAuth client secret</span>
            <input
              type="password"
              value={form.cmr_google_client_secret || ""}
              onChange={(event) => setForm({ ...form, cmr_google_client_secret: event.target.value })}
              placeholder="client secret"
            />
          </label>
          <div className="wide oauth-connect-row">
            <span>Google Drive upload account: {form.cmr_google_connected_email || "not connected"}</span>
            <button type="button" onClick={connectGoogleDrive} disabled={googleBusy}>
              {googleBusy ? "Connecting..." : "Connect Google Drive"}
            </button>
          </div>
          <label className="wide">
            <span>Target email recipients</span>
            <textarea
              value={Array.isArray(form.email_recipients) ? form.email_recipients.join("\n") : form.email_recipients}
              onChange={(event) => setForm({ ...form, email_recipients: event.target.value })}
              rows={6}
              placeholder={"name@example.com\nother@example.com"}
            />
          </label>
          <label>
            <span>SMTP host</span>
            <input
              value={form.smtp_host || ""}
              onChange={(event) => setForm({ ...form, smtp_host: event.target.value })}
              placeholder="smtp.gmail.com"
            />
          </label>
          <label>
            <span>SMTP port</span>
            <input
              value={form.smtp_port ?? 587}
              onChange={(event) => setForm({ ...form, smtp_port: event.target.value })}
              placeholder="587"
            />
          </label>
          <label>
            <span>SMTP username</span>
            <input
              value={form.smtp_username || ""}
              onChange={(event) => setForm({ ...form, smtp_username: event.target.value })}
              placeholder="ftereso@gmail.com"
            />
          </label>
          <label>
            <span>SMTP sender email</span>
            <input
              value={form.smtp_from || ""}
              onChange={(event) => setForm({ ...form, smtp_from: event.target.value })}
              placeholder="ftereso@gmail.com"
            />
          </label>
          <label className="wide">
            <span>SMTP app password</span>
            <input
              type="password"
              value={form.smtp_password || ""}
              onChange={(event) => setForm({ ...form, smtp_password: event.target.value })}
              placeholder="16-digit app password"
            />
          </label>
          <label className="checkbox-row">
            <input
              type="checkbox"
              checked={form.smtp_starttls !== false}
              onChange={(event) => setForm({ ...form, smtp_starttls: event.target.checked })}
            />
            <span>Use STARTTLS</span>
          </label>
        </div>
        {message && <div className="notice">{message}</div>}
        {error && <div className="notice danger">{error}</div>}
        <button className="primary" type="submit" disabled={saving}>
          {saving ? "Saving..." : "Save settings"}
        </button>
      </form>

      <InfoPanel
        title="Current live behavior"
        lines={[
          `User storage stays on the Render disk once created by ${currentUser.username}.`,
          "Secret Files are best kept as a seed or restore snapshot.",
          "Fust actions save locally first, then try Sheets and email.",
          "If one sync step fails, the saved action still stays in Render storage.",
          "Target email recipients are where notifications go.",
          "SMTP sender settings are how the app sends the email.",
        ]}
      />

      <div className="data-table-card">
        <div className="section-header">
          <h2>Spreadsheet connection test</h2>
          <button type="button" onClick={loadConnectionTest} disabled={connectionBusy}>
            {connectionBusy ? "Testing..." : "Test again"}
          </button>
        </div>
        {connectionTest && (
          <>
            <p className="sidebar-note">
              Runtime client email: {connectionTest.account?.client_email || "unknown"}
            </p>
            <p className="sidebar-note">
              Project: {connectionTest.account?.project_id || "unknown"}
            </p>
            <p className="sidebar-note">
              Spreadsheet ID: {connectionTest.spreadsheet_id || "(empty)"} | Tab: {connectionTest.sheet_name || "(empty)"}
            </p>
            <p className="sidebar-note">
              Result: {connectionTest.read_ok ? `read ok (${connectionTest.row_count} rows)` : "read failed"}
            </p>
            <p className="sidebar-note">
              Headers: {(connectionTest.headers || []).join(", ") || "none"}
            </p>
            {connectionTest.error && <div className="notice danger">{connectionTest.error}</div>}
          </>
        )}
      </div>

      <div className="data-table-card">
        <h2>Data tab debug</h2>
        {metaLoading && <div className="notice">Loading Data tab preview...</div>}
        {metaError && <div className="notice danger">Unable to read Data tab: {metaError}</div>}
        {!metaLoading && !metaError && (
          <>
            <p className="sidebar-note">
              Source: {metaData?.source || "unknown"} | Rows read: {metaData?.raw_row_count || 0} | Usable records: {metaData?.records?.length || 0}
            </p>
            <p className="sidebar-note">
              Headers: {(metaData?.headers || []).join(", ") || "none detected"}
            </p>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>klantnaam</th>
                    <th>Country</th>
                    <th>klantcode connect</th>
                  </tr>
                </thead>
                <tbody>
                  {(metaData?.sample_records || []).map((record, index) => (
                    <tr key={`${record.customer_name}-${record.country}-${record.customer_code}-${index}`}>
                      <td>{record.customer_name}</td>
                      <td>{record.country}</td>
                      <td>{record.customer_code || record.connect_name}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>

      <div className="data-table-card">
        <div className="section-header">
          <h2>Fust backups</h2>
          <button type="button" className="primary" onClick={createBackup} disabled={backupBusy}>
            {backupBusy ? "Creating..." : "Create backup"}
          </button>
        </div>
        <p className="sidebar-note">
          Snapshots are stored on the Render persistent disk and include current Fust settings plus all saved actions.
        </p>
        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>
                <th>Filename</th>
                <th>Created</th>
                <th>Size</th>
                <th>Download</th>
              </tr>
            </thead>
            <tbody>
              {backups.map((backup) => (
                <tr key={backup.filename}>
                  <td>{backup.filename}</td>
                  <td>{formatTimestamp(backup.created_at)}</td>
                  <td>{Math.round((backup.size_bytes || 0) / 1024)} KB</td>
                  <td>
                    <a href={backup.download_path}>Download</a>
                  </td>
                </tr>
              ))}
              {!backups.length && (
                <tr>
                  <td colSpan="4">No Fust backups created yet.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
}

function InfoPanel({ title, lines }) {
  return (
    <article className="info-panel">
      <h2>{title}</h2>
      <div>
        {lines.map((line) => (
          <p key={line}>{line}</p>
        ))}
      </div>
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
