/**
 * format.js - attendee name normalization for the Outlook-to-Obsidian add-in.
 *
 * Division of labour (by design):
 *   - Outlook owns NAMES. This file turns Office.js attendee objects into a
 *     clean, de-duplicated, delimited string of display names.
 *   - It deliberately does NOT wrap names in [[ ]] - wikilinks are an Obsidian
 *     concept and are added on the Obsidian side (prepare_meeting.js).
 *
 * Exposed globally (no bundler): window.MTO.format
 */
(function () {
  // Outlook display names are usually clean; we only reshape, not aggressively
  // re-case. nameFormat decides the shape.
  function normalizeName(displayName, nameFormat) {
    let name = (displayName || "").trim();
    if (!name) return "";

    // Handle "Last, First [Middle]" form from some directories.
    let first, last, middleAware;
    if (name.includes(",")) {
      const [l, rest] = name.split(",").map((s) => s.trim());
      const restWords = rest ? rest.split(/\s+/).filter(Boolean) : [];
      first = restWords[0] || "";
      last = l;
      middleAware = [first, ...restWords.slice(1), last];
    } else {
      const words = name.split(/\s+/).filter(Boolean);
      first = words[0] || "";
      last = words.length > 1 ? words[words.length - 1] : "";
      middleAware = words;
    }

    switch (nameFormat) {
      case "display":
        return middleAware.join(" ");
      case "lastFirst":
        return last ? `${last}, ${first}` : first;
      case "firstLast":
      default:
        return last ? `${first} ${last}` : first;
    }
  }

  /**
   * @param {Array<{name:string,email:string}>} attendees
   * @param {Object} opts {nameFormat, includeEmail, separator, dedupe, excludeSelf[]}
   * @returns {string}
   */
  function formatAttendees(attendees, opts) {
    const o = opts || {};
    const exclude = new Set((o.excludeSelf || []).map((e) => (e || "").toLowerCase()));
    const seen = new Set();
    const out = [];

    for (const a of attendees || []) {
      const email = (a.email || "").toLowerCase();
      if (email && exclude.has(email)) continue;

      let name = normalizeName(a.name || email.split("@")[0], o.nameFormat);
      if (!name) continue;

      const key = (email || name.toLowerCase());
      if (o.dedupe !== false) {
        if (seen.has(key)) continue;
        seen.add(key);
      }

      out.push(o.includeEmail && a.email ? `${name} <${a.email}>` : name);
    }

    return out.join(o.separator != null ? o.separator : "; ");
  }

  window.MTO = window.MTO || {};
  window.MTO.format = { formatAttendees };
})();
