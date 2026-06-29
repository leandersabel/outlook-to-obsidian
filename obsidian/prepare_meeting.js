/**
 * prepare_meeting.js - QuickAdd user script for the "Outlook to Obsidian" add-in.
 *
 * Before the Meeting (Outlook) template runs, this prepares three QuickAdd
 * variables from the add-in's data:
 *   - {{VALUE:title}}: note-name-safe (illegal filename characters stripped).
 *   - {{VALUE:attendeesYaml}}: names as a YAML wikilink list for frontmatter.
 *   - {{VALUE:agenda}}: invite body, minus the Teams join block, with stray
 *     {{ }} / <% %> defused so they can't crash the template engine.
 *
 * Every field here is attacker-controlled (a meeting invite's title, body,
 * location, organizer and attendee names all come from whoever sent it), so the
 * sanitizers below are the trust boundary into the vault.
 */

// Strip template tokens so untrusted text can't reach the QuickAdd/Templater engine.
const defuseTokens = (s) => (s || "").replace(/<%|%>|\{\{|\}\}/g, (m) => m[0] + " " + m[1]);

// Make a value safe to sit inside a double-quoted YAML scalar: a stray " or
// newline would break out of the string and inject arbitrary frontmatter keys.
const yamlSafe = (s) => defuseTokens(s).replace(/["\r\n]/g, " ").replace(/\s+/g, " ").trim();

// Illegal note-name characters (e.g. ":") crash QuickAdd's file creation.
const noteNameSafe = (s) => (s || "").replace(/[\\/:*?"<>|#^[\]]/g, " ").replace(/\s+/g, " ").trim();

// Keep everything before the Teams invite block - it opens with a long divider
// line or the "Microsoft Teams meeting" header - then defuse template tokens.
const cleanAgenda = (s) =>
  defuseTokens((s || "").split(/\r?\n[ \t]*[-_—]{10,}|\r?\nMicrosoft Teams meeting/i)[0]).trim();

module.exports = ({ variables: v }) => {
  if (!v) return;
  v.title = noteNameSafe(v.title);
  v.agenda = cleanAgenda(v.agenda);
  v.location = yamlSafe(v.location);
  v.organizer = yamlSafe(v.organizer);
  // Leading "\n  " per name continues the template's "attendees:" line as a YAML
  // list. Strip "/[/] from each name so it can't break out of the "[[ ]]" scalar.
  v.attendeesYaml = (v.attendees || "")
    .split(/\s*[;\n\r]\s*/)
    .filter(Boolean)
    .map((n) => `\n  - "[[${yamlSafe(n).replace(/[[\]]/g, " ").replace(/\s+/g, " ").trim()}]]"`)
    .join("");
};
