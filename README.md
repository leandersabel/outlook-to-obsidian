# Outlook to Obsidian

An Outlook add-in that turns the open meeting into an Obsidian note — title,
attendees, time, and agenda — via an `obsidian://` URI.

The add-in reads the appointment with Office.js and hands the data to a QuickAdd macro
in Obsidian.

## How it splits the work

| Concern | Owner |
| --- | --- |
| Correct attendee **names** (clean, de-duplicated, "First Last") | The Outlook add-in |
| Wrapping names in `[[ ]]`, building the note | Obsidian (QuickAdd + `prepare_meeting.js`) |

The add-in emits **plain names** (`Anna Müller; Bob Smith`); Obsidian wraps them.

## Install

1. Download **`manifest.xml`** (it points at the hosted add-in).
2. In Outlook, open an email → **Apps / Get Apps** → **Add a custom add-in → From File…**
   → pick `manifest.xml`.
3. Open a calendar event → ribbon → **Send to Obsidian** → **⚙ Settings**, set your
   **vault name** → **Save**.

## Obsidian setup

The button fires an `obsidian://quickadd?...` URI into a QuickAdd macro. The vault files
and setup are in **[`obsidian/`](obsidian/)**.

## Self-host (optional)

Steps 1–2 use the hosted add-in, which runs whatever this repo currently deploys. To run
code you control instead: fork the repo, enable **GitHub Pages** (Settings → Pages →
Source `main` / `/docs`), update the URLs in `manifest.xml` to your username, and
sideload that manifest.

## Files

```
manifest.xml          sideloaded; points at the Pages URLs
docs/                 GitHub Pages root — the add-in itself
  taskpane.html/.css  UI
  taskpane.js         reads the meeting, builds + launches the URI
  format.js           attendee name normalization
  settings.js         config (vault name, attendee formatting)
  launch.html         bounce page → obsidian://
  commands.html       placeholder FunctionFile
  assets/icon-*.png   ribbon icons
obsidian/             vault-side companion (NOT the add-in)
  prepare_meeting.js  QuickAdd user script
  Meeting (Outlook).md  note template
  README.md           QuickAdd setup
```

## License

[MIT](LICENSE)

## Limitations

- Large meetings may hide the attendee list; the add-in shows what Outlook exposes.
- Long agendas can exceed practical URI lengths — trim the invite body.
- Desktop task panes can't open `obsidian://` directly, so launches bounce through
  `launch.html` in your browser.
