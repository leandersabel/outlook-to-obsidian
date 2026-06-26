/* global Office, MTO */
(function () {
  let config = null;
  let meeting = null; // {title, attendees[], organizer, organizerEmail, start, end, location, body}

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
      end: null,
      location: "",
      body: "",
    };

    const tasks = [];

    // Subject
    tasks.push(readField(item.subject).then((v) => (meeting.title = v || "")));
    // Location
    tasks.push(readField(item.location).then((v) => (meeting.location = v || "")));
    // Start / End
    tasks.push(readField(item.start).then((v) => (meeting.start = v ? new Date(v) : null)));
    tasks.push(readField(item.end).then((v) => (meeting.end = v ? new Date(v) : null)));

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

    return Promise.all(tasks);
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
    const attendeesStr = MTO.format.formatAttendees(meeting.attendees, config.attendee);
    return {
      vault: config.vault || "",
      title: meeting.title,
      date: formatDate(meeting.start, config.dateFormat),
      start: formatDate(meeting.start, config.dateFormat),
      end: formatDate(meeting.end, config.dateFormat),
      location: meeting.location,
      organizer: meeting.organizer,
      organizerEmail: meeting.organizerEmail,
      attendees: attendeesStr,
      body: meeting.body,
    };
  }

  // substitute {{key}} -> encodeURIComponent(value). Unknown keys -> "".
  function fillTemplate(template, values, encode) {
    return template.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) => {
      const v = values[key] != null ? String(values[key]) : "";
      return encode ? encodeURIComponent(v) : v;
    });
  }

  function buildUri(profile) {
    const values = rawValues();
    // {{content}} is itself a template (markdown) -> expand first (unencoded),
    // then it gets encoded when substituted into the URI.
    values.content = fillTemplate(config.contentTemplate || "", values, false);
    return fillTemplate(profile.uriTemplate, values, true);
  }

  // --- launching obsidian:// (layered fallbacks) -------------------------------
  function launch(uri) {
    // Primary: a real user-gesture anchor click (best honored by OS handlers).
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
    el("m-when").textContent = v.start || "";
    el("m-attendees").textContent = v.attendees || "(none / hidden)";
  }

  function renderProfiles() {
    const wrap = el("profiles");
    wrap.innerHTML = "";
    if (!config.vault) {
      showStatus("Set your vault name in Settings first.", true);
    }
    (config.profiles || []).forEach((p) => {
      const btn = document.createElement("button");
      btn.className = "primary";
      btn.textContent = p.name;
      btn.onclick = () => launch(buildUri(p));
      wrap.appendChild(btn);
    });
  }

  function bindUi() {
    el("btn-copy-attendees").onclick = () =>
      copy(MTO.format.formatAttendees(meeting.attendees, config.attendee), "Attendees");
    el("btn-copy-uri").onclick = () => {
      const p = (config.profiles || [])[0];
      if (p) copy(buildUri(p), "URI");
    };
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
      el("s-content").value = config.contentTemplate || "";
      el("s-profiles").value = JSON.stringify(config.profiles || [], null, 2);
    }
  }

  function saveSettings() {
    let profiles;
    try {
      profiles = JSON.parse(el("s-profiles").value);
      if (!Array.isArray(profiles)) throw new Error("profiles must be an array");
    } catch (e) {
      showStatus("Profiles JSON invalid: " + e.message, true);
      return;
    }
    config.vault = el("s-vault").value.trim();
    config.attendee.nameFormat = el("s-nameformat").value;
    config.attendee.separator = el("s-separator").value;
    config.attendee.includeEmail = el("s-includeemail").checked;
    config.dateFormat = el("s-dateformat").value.trim();
    config.contentTemplate = el("s-content").value;
    config.profiles = profiles;

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
