# Obsidian side

These files are **not part of the add-in** — they live in your Obsidian vault and
receive what the add-in sends. The add-in fires an `obsidian://quickadd?...` URI; a
QuickAdd macro turns it into a note.

| File | Goes in | Does |
| --- | --- | --- |
| `prepare_meeting.js` | your QuickAdd user-scripts folder | sanitizes the title, builds the attendee wikilink list, strips the Teams join block from the agenda |
| `Meeting (Outlook).md` | your Templater/QuickAdd templates folder | the note layout the macro fills in |

This is a minimal, generic version. It needs **only the [QuickAdd] plugin** — no
Templater, no project picker, no Dataview.

[QuickAdd]: https://github.com/chhoumann/quickadd

## Setup

1. **Copy the two files** into your vault (anywhere; you'll point QuickAdd at them).

2. **Create a QuickAdd Macro** named exactly **`Meeting (Outlook)`**
   (QuickAdd → Manage Macros → add), with two commands in order:
   - **User Script** → `prepare_meeting.js`
   - **Nested Choice → Template**, configured as:
     - Template path: `Meeting (Outlook).md`
     - File Name Format: `{{VALUE:title}}`
     - Target folder: wherever your meeting notes live (e.g. `Meetings`)
     - If a note can already exist with that name, set "file already exists" to
       *Increment the file name*.

3. **Add the QuickAdd choice as a command** (toggle it on) so the URI can trigger it.

4. **Point the add-in at it.** In the Outlook add-in's ⚙ Settings, set your **vault
   name**. The add-in's button already fires:
   ```
   obsidian://quickadd?vault=<vault>&choice=Meeting%20(Outlook)&value-title=…&value-attendees=…&value-date=…&value-organizer=…&value-location=…&value-agenda=…
   ```

That's it. Open a meeting in Outlook → **Create meeting note**.

## Variables the add-in passes

`prepare_meeting.js` and the template consume these (QuickAdd fills `{{VALUE:name}}`
from each `value-name=` URI parameter):

| Variable | From the meeting |
| --- | --- |
| `{{VALUE:title}}` | subject (also the file name) |
| `{{VALUE:attendees}}` | `Name; Name` — the script turns it into `{{VALUE:attendeesYaml}}` |
| `{{VALUE:date}}` | start time |
| `{{VALUE:organizer}}` | organizer |
| `{{VALUE:location}}` | location |
| `{{VALUE:agenda}}` | invite body (Teams block stripped) |

## Customizing

Edit `Meeting (Outlook).md` to taste — add frontmatter, reorder sections. Only one
rule: keep `attendees: {{VALUE:attendeesYaml}}` on a single line, because the script
emits an indented YAML list that continues that line.

For date-nested folders (`Meetings/2026/06-June/…`), a heading, or extra pickers, add a
**Templater** step or use Templater in the template — that's how the author's own vault
does it, left out here to keep the dependency list to QuickAdd alone.
