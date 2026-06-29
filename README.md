# Outlook to Obsidian

An Outlook add-in that turns the open meeting into an Obsidian note — title,
attendees, time and agenda — via an `obsidian://` URI. Works in **new Outlook for
Mac**, **Outlook on Windows**, and **Outlook on the web** (so also Linux, in a browser).

No Azure app registration. No calendar sync. The add-in reads the appointment via
Office.js and hands the data to a QuickAdd macro in Obsidian.

## How it splits the work

| Concern | Owner |
| --- | --- |
| Correct attendee **names** (clean, de-duplicated, "First Last") | The Outlook add-in — Outlook knows names best |
| Wrapping names in `[[ ]]`, building the note | Obsidian (QuickAdd + `prepare_meeting.js`) |

The add-in emits **plain names** (`Anna Müller; Bob Smith`); Obsidian wraps them.

## Architecture

- **Code** is static HTML/JS on **GitHub Pages** (public, read-only).
- **Config** (vault name, attendee formatting) lives in
  `Office.context.roamingSettings` — stored in **your mailbox**, private, roaming across
  devices. The profile/button is code-defined, not stored. Nothing personal is in this repo.

## Install (sideload)

1. **Host the code:** fork this repo, enable **GitHub Pages** (Settings → Pages → Source
   `main` / `/docs`); files land at `https://<you>.github.io/outlook-to-obsidian/`. Update
   the URLs in `manifest.xml` to your username. (Or point the manifest at an existing
   instance — config is per-mailbox, no collisions.)
2. **Sideload `manifest.xml`:** in Outlook → open an email → **Apps / Get Apps** →
   **Add a custom add-in → From File…**. Registered to your mailbox, it appears across
   Mac/Windows/web.
3. Open a calendar event → ribbon → **Send to Obsidian** → **⚙ Settings**, set your
   **vault name** → **Save**.

> If your tenant blocks custom sideloading, ask IT to centrally deploy the manifest
> (read-only, `ReadItem` only).

## Obsidian setup

The button fires an `obsidian://quickadd?...` URI into a QuickAdd macro. Ready-to-use
vault files and step-by-step setup are in **[`obsidian/`](obsidian/)** — only the
QuickAdd plugin is required.

## Files

```
manifest.xml          sideloaded; points at the Pages URLs
docs/                 GitHub Pages root — the add-in itself
  taskpane.html/.css  UI
  taskpane.js         reads the meeting, builds + launches the URI
  format.js           attendee name normalization (no wikilink wrapping)
  settings.js         RoamingSettings config + code-defined profile
  launch.html         HTTPS bounce page → obsidian://
  commands.html       placeholder FunctionFile
  assets/icon-*.png   ribbon icons
obsidian/             vault-side companion (NOT the add-in)
  prepare_meeting.js  QuickAdd user script
  Meeting (Outlook).md  note template
  README.md           QuickAdd setup
```

## Limitations

- Very large meetings may hide the attendee list; the add-in shows what Outlook exposes.
- Very long agendas can exceed practical URI lengths — trim the invite body.
- Desktop task panes can't pass `obsidian://` to the OS, so launches bounce through
  `launch.html` in your default browser. Outlook-on-the-web launches most reliably.
