/* global Office, MTO */
(function () {
  let config = null;
  let meeting = null; // {title, attendees[], organizer, organizerEmail, start, location, body}

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;
    document.addEventListener("DOMContentLoaded", init);
    if (document.readyState !== "loading") init();
  });

  let inited = false;
  function init() {
    if (inited) return;
    inited = true;

    config = MTO.settings.load();
    seedSelf();

    bindUi();
    extractMeeting()
      .then(() => {
        renderPreview();
        renderProfiles();
      })
      .catch((err) => showStatus("Could not read the meeting: " + err, true));
  }

  // --- seed excludeSelf with the signed-in user so "don't list me" just works ---
  function seedSelf() {
    try {
      const me = Office.context.mailbox.userProfile.emailAddress;
      const list = config.attendee.excludeSelf || [];
      if (me && !list.map((s) => s.toLowerCase()).includes(me.toLowerCase())) {
        config.attendee.excludeSelf = list.concat([me]);
      }
    } catch (e) {
      /* userProfile not available in some hosts; ignore */
    }
  }

  // --- read the open appointment (read mode = sync arrays; compose = async) ---
  function extractMeeting() {
    const item = Office.context.mailbox.item;
    meeting = {
      title: "",
      attendees: [],
      organizer: "",
      organizerEmail: "",
      start: null,
      location: "",
      body: "",
    };

    const tasks = [];

    // Subject
    tasks.push(readField(item.subject).then((v) => (meeting.title = v || "")));
    // Location
    tasks.push(readField(item.location).then((v) => (meeting.location = v || "")));
    // Start
    tasks.push(readField(item.start).then((v) => (meeting.start = v ? new Date(v) : null)));

    // Organizer (read mode only; compose has no organizer getter)
    if (item.organizer) {
      tasks.push(
        readField(item.organizer).then((org) => {
          if (org) {
            meeting.organizer = org.displayName || "";
            meeting.organizerEmail = org.emailAddress || "";
          }
        })
      );
    }

    // Attendees: read mode exposes arrays; compose exposes async getters.
    tasks.push(readAttendees(item.requiredAttendees));
    tasks.push(readAttendees(item.optionalAttendees));

    // Body (text), capped to keep URIs sane.
    tasks.push(
      new Promise((resolve) => {
        if (item.body && item.body.getAsync) {
          item.body.getAsync(Office.CoercionType.Text, (res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              meeting.body = (res.value || "").trim();
            }
            resolve();
          });
        } else {
          resolve();
        }
      })
    );

    return Promise.all(tasks).then(() => {
      // Include the organizer in the attendee list (dedupe handles overlap).
      if (meeting.organizer || meeting.organizerEmail) {
        meeting.attendees.unshift({
          name: meeting.organizer,
          email: meeting.organizerEmail,
        });
      }
    });
  }

  // A field may be a plain value (read) or an object with getAsync (compose).
  function readField(field) {
    return new Promise((resolve) => {
      if (field && typeof field.getAsync === "function") {
        field.getAsync((res) =>
          resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : null)
        );
      } else {
        resolve(field);
      }
    });
  }

  function readAttendees(field) {
    return readField(field).then((val) => {
      const arr = Array.isArray(val) ? val : [];
      for (const a of arr) {
        meeting.attendees.push({
          name: a.displayName || "",
          email: a.emailAddress || "",
        });
      }
    });
  }

  // --- placeholder substitution ------------------------------------------------
  function formatDate(d, fmt) {
    if (!d) return "";
    const p = (n) => String(n).padStart(2, "0");
    return fmt
      .replace(/YYYY/g, d.getFullYear())
      .replace(/MM/g, p(d.getMonth() + 1))
      .replace(/DD/g, p(d.getDate()))
      .replace(/HH/g, p(d.getHours()))
      .replace(/mm/g, p(d.getMinutes()));
  }

  function rawValues() {
    return {
      vault: config.vault || "",
      title: meeting.title,
      date: formatDate(meeting.start, config.dateFormat),
      location: meeting.location,
      organizer: meeting.organizer,
      attendees: MTO.format.formatAttendees(meeting.attendees, config.attendee),
      body: meeting.body,
    };
  }

  function buildUri(profile) {
    const values = rawValues();
    // substitute {{key}} -> encodeURIComponent(value); unknown keys -> "".
    return profile.uriTemplate.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) =>
      encodeURIComponent(values[key] != null ? String(values[key]) : "")
    );
  }

  // --- launching obsidian:// (layered fallbacks) -------------------------------
  function launch(uri) {
    // A task-pane webview usually won't pass obsidian:// to the OS, so we bounce
    // through an HTTPS launcher page opened in the real browser, which then
    // redirects to the obsidian:// URI (the system browser knows the protocol).
    const launcher =
      new URL("launch.html", window.location.href).href +
      "#" +
      encodeURIComponent(uri);

    // Primary: Office API opens the default browser (http/https only — hence the
    // launcher bounce). Supported on Outlook Windows/Mac/web.
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(launcher);
        showStatus("Opening in your browser → Obsidian…");
        return;
      }
    } catch (e) {
      /* fall through */
    }

    // Fallback 1: a normal popup to the launcher.
    const w = window.open(launcher, "_blank");
    if (w) {
      showStatus("Opening in your browser → Obsidian…");
      return;
    }

    // Fallback 2: direct anchor click (works in Outlook on the web).
    const a = document.createElement("a");
    a.href = uri;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showStatus("Opening Obsidian… If nothing happens, use “Copy URI”.");
  }

  function copy(text, label) {
    navigator.clipboard.writeText(text).then(
      () => showStatus(label + " copied."),
      () => showStatus("Clipboard blocked - select and copy manually.", true)
    );
  }

  // --- UI ----------------------------------------------------------------------
  function el(id) {
    return document.getElementById(id);
  }

  function renderPreview() {
    const v = rawValues();
    el("m-title").textContent = v.title || "(no subject)";
    el("m-when").textContent = v.date || "";
    el("m-attendees").textContent = v.attendees || "(none / hidden)";
  }

  function launcherUrl(obsidianUri) {
    return (
      new URL("launch.html", window.location.href).href +
      "#" +
      encodeURIComponent(obsidianUri)
    );
  }

  function renderProfiles() {
    const wrap = el("profiles");
    wrap.innerHTML = "";
    if (!config.vault) {
      showStatus("Set your vault name in Settings first.", true);
    }
    (config.profiles || []).forEach((p) => {
      // A real anchor + genuine user click is what reliably reaches the system
      // browser (which knows obsidian://). target=_blank opens the launcher,
      // which then redirects to the obsidian:// URI.
      const a = document.createElement("a");
      a.className = "primary";
      a.textContent = p.name;
      a.href = launcherUrl(buildUri(p));
      a.target = "_blank";
      a.rel = "noopener";
      a.onclick = () => showStatus("Opening in your browser → Obsidian…");
      wrap.appendChild(a);
    });
  }

  function bindUi() {
    el("btn-copy-attendees").onclick = () =>
      copy(MTO.format.formatAttendees(meeting.attendees, config.attendee), "Attendees");
    el("btn-settings").onclick = () => toggleSettings(true);
    el("btn-settings-cancel").onclick = () => toggleSettings(false);
    el("btn-settings-save").onclick = saveSettings;
  }

  function toggleSettings(show) {
    el("settings").classList.toggle("hidden", !show);
    el("main").classList.toggle("hidden", show);
    if (show) {
      el("s-vault").value = config.vault || "";
      el("s-nameformat").value = config.attendee.nameFormat || "firstLast";
      el("s-separator").value = config.attendee.separator || "; ";
      el("s-includeemail").checked = !!config.attendee.includeEmail;
      el("s-dateformat").value = config.dateFormat || "";
    }
  }

  function saveSettings() {
    config.vault = el("s-vault").value.trim();
    config.attendee.nameFormat = el("s-nameformat").value;
    config.attendee.separator = el("s-separator").value;
    config.attendee.includeEmail = el("s-includeemail").checked;
    config.dateFormat = el("s-dateformat").value.trim();

    MTO.settings.save(config).then(
      () => {
        seedSelf();
        toggleSettings(false);
        renderPreview();
        renderProfiles();
        showStatus("Settings saved.");
      },
      (err) => showStatus("Save failed: " + (err && err.message ? err.message : err), true)
    );
  }

  let statusTimer = null;
  function showStatus(msg, isError) {
    const s = el("status");
    s.textContent = msg;
    s.className = isError ? "status error" : "status";
    if (statusTimer) clearTimeout(statusTimer);
    if (!isError) statusTimer = setTimeout(() => (s.textContent = ""), 4000);
  }
})();
