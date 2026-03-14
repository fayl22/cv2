const STORAGE_KEY = "spIssueTrackerConfig";

const el = {
  siteUrl: document.getElementById("siteUrl"),
  listTitle: document.getElementById("listTitle"),
  authToken: document.getElementById("authToken"),
  saveConfigBtn: document.getElementById("saveConfigBtn"),
  loadIssuesBtn: document.getElementById("loadIssuesBtn"),
  configStatus: document.getElementById("configStatus"),
  issueForm: document.getElementById("issueForm"),
  title: document.getElementById("title"),
  description: document.getElementById("description"),
  priority: document.getElementById("priority"),
  status: document.getElementById("status"),
  createStatus: document.getElementById("createStatus"),
  issuesBody: document.getElementById("issuesBody"),
  issuesStatus: document.getElementById("issuesStatus"),
};

let requestDigestCache = { value: null, expiresAt: 0 };

function setMessage(node, message, type = "") {
  node.textContent = message;
  node.className = `status-message ${type}`.trim();
}

function loadConfig() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) {
      return;
    }

    const config = JSON.parse(saved);
    el.siteUrl.value = config.siteUrl || "";
    el.listTitle.value = config.listTitle || "Issues";
    el.authToken.value = config.authToken || "";
  } catch {
    setMessage(el.configStatus, "Could not parse saved configuration.", "error");
  }
}

function saveConfig() {
  const config = getConfig();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
  setMessage(el.configStatus, "Configuration saved.", "success");
}

function getConfig() {
  return {
    siteUrl: el.siteUrl.value.trim().replace(/\/$/, ""),
    listTitle: el.listTitle.value.trim(),
    authToken: el.authToken.value.trim(),
  };
}

function ensureConfig() {
  const config = getConfig();
  if (!config.siteUrl || !config.listTitle) {
    throw new Error("Site URL and list title are required.");
  }
  return config;
}

function buildHeaders(config, includeJson = true) {
  const headers = {};
  if (includeJson) {
    headers["Accept"] = "application/json;odata=verbose";
    headers["Content-Type"] = "application/json;odata=verbose";
  }

  if (config.authToken) {
    headers.Authorization = `Bearer ${config.authToken}`;
  }

  return headers;
}

async function getListItemEntityType(config) {
  const headers = buildHeaders(config, false);
  headers.Accept = "application/json;odata=verbose";

  const res = await fetch(
    `${config.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.listTitle)}')?$select=ListItemEntityTypeFullName`,
    { headers }
  );

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Could not get list metadata (${res.status}). ${txt}`);
  }

  const data = await res.json();
  return data.d.ListItemEntityTypeFullName;
}

async function getRequestDigest(config) {
  const now = Date.now();
  if (requestDigestCache.value && requestDigestCache.expiresAt > now) {
    return requestDigestCache.value;
  }

  const headers = buildHeaders(config, false);
  headers.Accept = "application/json;odata=verbose";

  const response = await fetch(`${config.siteUrl}/_api/contextinfo`, {
    method: "POST",
    headers,
  });

  if (!response.ok) {
    const txt = await response.text();
    throw new Error(`Unable to obtain request digest (${response.status}). ${txt}`);
  }

  const payload = await response.json();
  const timeoutSeconds = Number(payload.d.GetContextWebInformation.FormDigestTimeoutSeconds || 1200);
  requestDigestCache = {
    value: payload.d.GetContextWebInformation.FormDigestValue,
    expiresAt: now + Math.max(30, timeoutSeconds - 30) * 1000,
  };

  return requestDigestCache.value;
}

async function loadIssues() {
  const config = ensureConfig();
  setMessage(el.issuesStatus, "Loading issues...");

  const query = "$select=Id,Title,Description,Priority,IssueStatus,Created&$orderby=Id desc&$top=200";
  const url = `${config.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.listTitle)}')/items?${query}`;

  const res = await fetch(url, { headers: buildHeaders(config, false) });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Failed to load issues (${res.status}). ${txt}`);
  }

  const data = await res.json();
  renderIssues(data.d.results || []);
  setMessage(el.issuesStatus, `Loaded ${data.d.results.length} issue(s).`, "success");
}

function renderIssues(items) {
  if (!items.length) {
    el.issuesBody.innerHTML = `<tr><td colspan="7" class="small">No issues found.</td></tr>`;
    return;
  }

  el.issuesBody.innerHTML = items
    .map(
      (item) => `
      <tr>
        <td>${item.Id}</td>
        <td>${escapeHtml(item.Title || "")}</td>
        <td>${escapeHtml(item.Description || "")}</td>
        <td><span class="pill">${escapeHtml(item.Priority || "-")}</span></td>
        <td>${escapeHtml(item.IssueStatus || "-")}</td>
        <td>${new Date(item.Created).toLocaleString()}</td>
        <td>
          <button type="button" class="ghost" data-action="toggle-status" data-id="${item.Id}" data-status="${escapeHtml(
            item.IssueStatus || "Open"
          )}">Toggle Open/Closed</button>
          <button type="button" class="ghost" data-action="delete" data-id="${item.Id}">Delete</button>
        </td>
      </tr>`
    )
    .join("");
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function createIssue(event) {
  event.preventDefault();
  const config = ensureConfig();

  if (!el.title.value.trim()) {
    setMessage(el.createStatus, "Title is required.", "error");
    return;
  }

  setMessage(el.createStatus, "Creating issue...");

  const [entityType, digest] = await Promise.all([
    getListItemEntityType(config),
    getRequestDigest(config),
  ]);

  const payload = {
    __metadata: { type: entityType },
    Title: el.title.value.trim(),
    Description: el.description.value.trim(),
    Priority: el.priority.value,
    IssueStatus: el.status.value,
  };

  const headers = buildHeaders(config);
  headers["X-RequestDigest"] = digest;

  const res = await fetch(
    `${config.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.listTitle)}')/items`,
    {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    }
  );

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Failed to create issue (${res.status}). ${txt}`);
  }

  el.issueForm.reset();
  el.priority.value = "Medium";
  el.status.value = "Open";
  setMessage(el.createStatus, "Issue created successfully.", "success");
  await loadIssues();
}

async function toggleStatus(id, currentStatus) {
  const config = ensureConfig();
  const [entityType, digest] = await Promise.all([
    getListItemEntityType(config),
    getRequestDigest(config),
  ]);

  const nextStatus = currentStatus === "Closed" ? "Open" : "Closed";
  const payload = {
    __metadata: { type: entityType },
    IssueStatus: nextStatus,
  };

  const headers = buildHeaders(config);
  headers["X-RequestDigest"] = digest;
  headers["X-HTTP-Method"] = "MERGE";
  headers["IF-MATCH"] = "*";

  const res = await fetch(
    `${config.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.listTitle)}')/items(${id})`,
    {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    }
  );

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Failed to update issue ${id} (${res.status}). ${txt}`);
  }

  setMessage(el.issuesStatus, `Issue ${id} marked as ${nextStatus}.`, "success");
  await loadIssues();
}

async function deleteIssue(id) {
  const config = ensureConfig();
  const digest = await getRequestDigest(config);

  const headers = buildHeaders(config, false);
  headers["X-RequestDigest"] = digest;
  headers["X-HTTP-Method"] = "DELETE";
  headers["IF-MATCH"] = "*";

  const res = await fetch(
    `${config.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.listTitle)}')/items(${id})`,
    { method: "POST", headers }
  );

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Failed to delete issue ${id} (${res.status}). ${txt}`);
  }

  setMessage(el.issuesStatus, `Issue ${id} deleted.`, "success");
  await loadIssues();
}

el.saveConfigBtn.addEventListener("click", saveConfig);
el.loadIssuesBtn.addEventListener("click", async () => {
  try {
    await loadIssues();
  } catch (error) {
    setMessage(el.issuesStatus, error.message, "error");
  }
});

el.issueForm.addEventListener("submit", async (event) => {
  try {
    await createIssue(event);
  } catch (error) {
    setMessage(el.createStatus, error.message, "error");
  }
});

el.issuesBody.addEventListener("click", async (event) => {
  const button = event.target.closest("button[data-action]");
  if (!button) {
    return;
  }

  try {
    const id = Number(button.dataset.id);
    if (button.dataset.action === "toggle-status") {
      await toggleStatus(id, button.dataset.status);
    }

    if (button.dataset.action === "delete") {
      const shouldDelete = window.confirm(`Delete issue #${id}?`);
      if (shouldDelete) {
        await deleteIssue(id);
      }
    }
  } catch (error) {
    setMessage(el.issuesStatus, error.message, "error");
  }
});

loadConfig();
