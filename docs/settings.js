/**
 * settings.js - per-mailbox configuration for Outlook-to-Obsidian.
 *
 * Storage: Office.context.roamingSettings -> persisted in the user's mailbox,
 * roams across Mac / Windows / Outlook-on-the-web. NOT stored on GitHub Pages
 * (that only serves the static code) and NOT in localStorage (per-device).
 *
 * Exposed globally as window.MTO.settings
 */
(function () {
  const KEY = "mto.config.v1";

  // No personal data baked in: this ships blank so anyone can adopt it.
  const DEFAULTS = {
    vault: "",
    dateFormat: "YYYY-MM-DD HH:mm",
    attendee: {
      nameFormat: "firstLast", // firstLast | display | lastFirst
      includeEmail: false,
      separator: "; ",
      dedupe: true,
      excludeSelf: [], // seeded at runtime with the signed-in user's address
    },
    // Each profile = one button. Triggers a QuickAdd macro, passing NAMED
    // variables via the `value-<name>=` syntax (bare {{VALUE}} can't be filled
    // from a URI and would prompt). The macro picks the project, wraps
    // attendees, then creates the note from Templates/Meeting (Outlook).md.
    profiles: [
      {
        name: "Create meeting note",
        uriTemplate:
          "obsidian://quickadd?vault={{vault}}&choice=Meeting%20(Outlook)&value-title={{title}}&value-attendees={{attendees}}&value-date={{date}}&value-organizer={{organizer}}&value-location={{location}}&value-agenda={{body}}",
      },
    ],
  };

  function deepMerge(base, override) {
    if (Array.isArray(base)) return override != null ? override : base;
    if (typeof base === "object" && base) {
      const out = {};
      for (const k of Object.keys(base)) {
        out[k] = k in (override || {}) ? deepMerge(base[k], override[k]) : base[k];
      }
      // keep any extra keys the user added
      for (const k of Object.keys(override || {})) if (!(k in out)) out[k] = override[k];
      return out;
    }
    return override != null ? override : base;
  }

  function load() {
    let raw = null;
    try {
      raw = Office.context.roamingSettings.get(KEY);
    } catch (e) {
      raw = null;
    }
    let parsed = null;
    if (typeof raw === "string") {
      try {
        parsed = JSON.parse(raw);
      } catch (e) {
        parsed = null;
      }
    } else if (raw && typeof raw === "object") {
      parsed = raw;
    }
    const merged = deepMerge(DEFAULTS, parsed || {});
    // Profiles are code-defined, never stored, so they always reflect the
    // shipped defaults (no manual JSON editing in settings).
    merged.profiles = DEFAULTS.profiles;
    return merged;
  }

  function save(config) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.roamingSettings.set(KEY, JSON.stringify(config));
        Office.context.roamingSettings.saveAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(res.error);
        });
      } catch (e) {
        reject(e);
      }
    });
  }

  window.MTO = window.MTO || {};
  window.MTO.settings = { load, save, DEFAULTS };
})();
