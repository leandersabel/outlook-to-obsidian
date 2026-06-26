# Outlook to Obsidian

An Outlook add-in that sends the **open meeting's title, attendees, time and agenda**
into Obsidian via an `obsidian://` URI. Works in **new Outlook for Mac**, **Outlook on
Windows**, and **Outlook on the web** (so also Linux, in a browser).

No Azure app registration. No calendar sync. The add-in reads the appointment you're
looking at via Office.js and hands the data to Obsidian — you decide how, with editable
URI templates.

## How it splits the work

| Concern | Owner |
| --- | --- |
| Correct attendee **names** (clean, de-duplicated, "First Last" etc.) | The Outlook add-in — Outlook knows names best |
| Wrapping names in `[[ ]]` wikilinks | Obsidian (`wrap_attendees.js`) — wikilinks are an Obsidian concept |

The add-in emits **plain names** (`Anna Müller; Bob Smith`). Whether/how they become
wikilinks is up to the Obsidian side.

## Architecture

- **Code** is static HTML/JS hosted on **GitHub Pages** (public, read-only).
- **Configuration** (vault name, profiles, formatting) lives in
  `Office.context.roamingSettings` — stored in **your mailbox**, private, and roaming
  across all your devices. Nothing personal is in this repo.

## Install (sideload)

1. **Host the code.** Two options:
   - *Fork & host:* fork this repo, enable **GitHub Pages** (Settings → Pages → Source:
     `main` / `/docs`). Your files land at
     `https://<you>.github.io/outlook-to-obsidian/`. Update the URLs in `manifest.xml`
     to your username.
   - *Shared instance:* point `manifest.xml` at someone else's Pages URL. Config is
     per-mailbox, so there are no collisions.
2. **Sideload the manifest** — easiest via **Outlook on the web**:
   Settings → *Get Add-ins* → **My add-ins** → **Add a custom add-in** →
   **Add from file…** → choose `manifest.xml`.
   Once registered to your mailbox it also appears in **new Outlook for Mac/Windows**
   (they share the account's add-in registry).
3. Open a calendar event → ribbon → **Send to Obsidian**.

> If your tenant blocks custom sideloading, ask IT to *centrally deploy* this manifest
> (it's read-only and needs only `ReadItem`), or use the **Copy attendees / Copy URI**
> fallback buttons with your existing clipboard flow.

## First run

The add-in ships blank. Open **⚙ Settings** and set your **vault name** (exactly as
Obsidian shows it). Your own address is auto-excluded from attendee lists.

## Placeholders

Use these in any profile's `uriTemplate` and in the note content template:

`{{vault}}` `{{title}}` `{{date}}` `{{start}}` `{{end}}` `{{location}}`
`{{organizer}}` `{{organizerEmail}}` `{{attendees}}` `{{body}}` `{{content}}`

Each value is URL-encoded on substitution. `{{content}}` is the expanded "note content
template" (Markdown) from Settings.

## Obsidian-side recipes

### 1. New note — no plugins (works out of the box)
Default profile:
```
obsidian://new?vault={{vault}}&name={{title}}&content={{content}}
```
Creates a note titled with the subject; body comes from the content template. Attendees
appear as plain names (no wrapping — nothing runs to wrap them).

### 2. QuickAdd + `wrap_attendees.js` (wikilinked attendees)
Profile:
```
obsidian://quickadd?choice=Meeting&value-title={{title}}&value-attendees={{attendees}}&value-date={{date}}&value-organizer={{organizer}}&value-agenda={{body}}
```
QuickAdd passes URI data as `value-<name>=` parameters into **named** variables
(`{{VALUE:title}}` …). Bare `{{VALUE}}` placeholders can't be filled from a URI and
will prompt instead — so every variable the choice uses must be one we pass.

In Obsidian:
1. Copy `wrap_attendees.js` into your QuickAdd user-scripts folder.
2. Create a QuickAdd **Macro** choice named **`Meeting`** with two steps:
   - Step 1: **User Script** → `wrap_attendees.js`
     (reads the `attendees` variable, sets `{{VALUE:attendeesWikilinks}}`, and — if
     enabled — writes the `attendees` frontmatter of the active note).
   - Step 2: a **Template** choice that creates the note and inserts
     `{{VALUE:title}}`, `{{VALUE:attendeesWikilinks}}`, `{{VALUE:date}}`,
     `{{VALUE:organizer}}`, `{{VALUE:agenda}}`.

### 3. Advanced URI (frontmatter / run a command)
The [Advanced URI](https://github.com/Vinzent03/obsidian-advanced-uri) plugin exposes
params for `filepath`, content, `mode`, frontmatter writes and `commandid=` to trigger
QuickAdd/Templater. Point a profile at `obsidian://adv-uri?...` mapping placeholders to
its params. Wrapping still happens in `wrap_attendees.js` when you route through QuickAdd.

## Files

```
manifest.xml            sideloaded; points at the Pages URLs
docs/                   GitHub Pages root
  taskpane.html/.css    UI
  taskpane.js           reads the meeting, builds + launches the URI
  format.js             attendee name normalization (NO wikilink wrapping)
  settings.js           RoamingSettings config (defaults, load/save)
  commands.html         placeholder FunctionFile
  assets/icon-*.png     ribbon icons
```
`wrap_attendees.js` lives in your Obsidian vault's QuickAdd scripts folder, not here.

## Limitations

- Very large meetings may hide the attendee list; the add-in shows what Outlook exposes.
- Long agendas can exceed practical URI lengths — trim the body or use the clipboard
  fallback.
- Custom-protocol launch from a desktop task pane can be blocked by some hosts; the
  anchor-click primary plus **Copy URI** fallback cover that. Outlook-on-the-web is the
  most reliable launcher.
