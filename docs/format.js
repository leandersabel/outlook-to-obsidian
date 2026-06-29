/**
 * format.js - attendee name normalization for the Outlook-to-Obsidian add-in.
 *
 * Division of labour (by design):
 *   - Outlook owns NAMES. This file turns Office.js attendee objects into a
 *     clean, de-duplicated, delimited string of names.
 *   - It does NOT wrap names in [[ ]] - wikilinks are an Obsidian concept,
 *     added on the Obsidian side (prepare_meeting.js).
 *
 * Outlook's names are unreliable: the display name can be just an email address,
 * or have first/last flipped. The email local part (first.last) is the steadier
 * signal, so we take ordering from it and use the display name only to recover
 * casing and expand initials. Mirrors the vault's parse_attendees.js.
 *
 * Exposed globally (no bundler): window.MTO.format
 */
(function () {
  const PARTICLES = new Set(["von", "van", "de", "der", "den", "di", "du", "le", "la", "el"]);

  // Capitalize a word, but leave intentional internal caps (McDonald, O'Brien,
  // DeShawn) alone - only re-case all-lower or all-upper input.
  function capWord(w) {
    if (!w) return w;
    if (/[a-z]/.test(w) && /[A-Z]/.test(w.slice(1))) return w;
    if (w.includes("-")) return w.split("-").map(capWord).join("-");
    if (w.includes("'")) return w.split("'").map(capWord).join("'");
    return w[0].toUpperCase() + w.slice(1).toLowerCase();
  }

  // Capitalize parts in order; a particle (von, van, de...) stays lowercase
  // unless it leads the name.
  const capParts = (parts) =>
    parts.map((w, i) => (i > 0 && PARTICLES.has(w.toLowerCase()) ? w.toLowerCase() : capWord(w)));

  // Fold a display word to the ASCII form an email would use, so "Müller"
  // matches the local part "mueller" (German ä/ö/ü/ß transliteration + accents).
  const fold = (s) =>
    s
      .toLowerCase()
      .replace(/ä/g, "ae").replace(/ö/g, "oe").replace(/ü/g, "ue").replace(/ß/g, "ss")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  // first.last in the email is authoritative for ordering. Use the display
  // words only to recover real casing/spelling and expand single-letter initials.
  function enrichFromDisplay(emailParts, display) {
    const dw = display ? display.replace(/,/g, " ").split(/\s+/).filter(Boolean) : [];
    const resolved = emailParts.map((ep) => {
      const m = dw.find((w) => fold(w) === ep || fold(w).startsWith(ep));
      return m && (ep.length === 1 || fold(m) === ep) ? m : ep;
    });
    // Drop middle parts (e.g. anna.marie.muller) down to first + last.
    return resolved.length > 2 ? [resolved[0], resolved[resolved.length - 1]] : resolved;
  }

  // "Last, First Middle" -> [First, Middle, Last]; "First Last" -> [First, Last].
  function displayToParts(display) {
    if (display.includes(",")) {
      const [last, rest] = display.split(",").map((s) => s.trim());
      const restWords = rest ? rest.split(/\s+/).filter(Boolean) : [];
      return [...restWords, last].filter(Boolean);
    }
    return display.split(/\s+/).filter(Boolean);
  }

  // Resolve an attendee to an ordered, capitalized [first, ...middle, last].
  function nameParts(display, email) {
    const local = email ? email.split("@")[0].replace(/\d+$/g, "") : "";
    const emailParts = local.split(/[._-]+/).filter(Boolean);

    let parts;
    if (emailParts.length >= 2) parts = enrichFromDisplay(emailParts, display);
    else if (display) parts = displayToParts(display);
    else if (local) parts = [local]; // unsplittable local part (e.g. jsmith)
    else return [];

    return capParts(parts);
  }

  function normalizeName(displayName, email, nameFormat) {
    const display = (displayName || "").trim();
    const parts = nameParts(display, (email || "").toLowerCase());
    if (!parts.length) return display;

    const first = parts[0];
    const last = parts.length > 1 ? parts[parts.length - 1] : "";
    switch (nameFormat) {
      case "display":
        return display || parts.join(" "); // Outlook's label verbatim; derive only if absent
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

      const name = normalizeName(a.name, a.email, o.nameFormat);
      if (!name) continue;

      const key = email || name.toLowerCase();
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
