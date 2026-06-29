/**
 * prepare_meeting.js - QuickAdd user script for the "Outlook to Obsidian" add-in.
 *
 * Before the Meeting (Outlook) template runs, this prepares three QuickAdd
 * variables from the add-in's data:
 *   - {{VALUE:title}}: note-name-safe (illegal filename characters stripped).
 *   - {{VALUE:attendeesYaml}}: names as a YAML wikilink list for frontmatter.
 *   - {{VALUE:agenda}}: invite body, minus the Teams join block, with stray
 *     {{ }} / <% %> defused so they can't crash the template engine.
 */

// Illegal note-name characters (e.g. ":") crash QuickAdd's file creation.
const noteNameSafe = (s) => (s || "").replace(/[\\/:*?"<>|#^[\]]/g, " ").replace(/\s+/g, " ").trim();

// Keep everything before the Teams invite block - it opens with a long divider
// line or the "Microsoft Teams meeting" header - then defuse template tokens.
const cleanAgenda = (s) =>
  (s || "")
    .split(/\r?\n[ \t]*[-_—]{10,}|\r?\nMicrosoft Teams meeting/i)[0]
    .replace(/<%|%>|\{\{|\}\}/g, (m) => m[0] + " " + m[1])
    .trim();

module.exports = ({ variables: v }) => {
  if (!v) return;
  v.title = noteNameSafe(v.title);
  v.agenda = cleanAgenda(v.agenda);
  // Leading "\n  " per name continues the template's "attendees:" line as a YAML list.
  v.attendeesYaml = (v.attendees || "").split(/\s*[;\n\r]\s*/).filter(Boolean).map((n) => `\n  - "[[${n}]]"`).join("");
};
