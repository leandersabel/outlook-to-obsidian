# Obsidian side

These files live in your vault and receive what the add-in sends. The add-in fires an
`obsidian://quickadd?...` URI; a QuickAdd macro turns it into a note. Requires the
[QuickAdd] plugin.

[QuickAdd]: https://github.com/chhoumann/quickadd

| File | Goes in | Does |
| --- | --- | --- |
| `prepare_meeting.js` | your QuickAdd user-scripts folder | sanitizes the title, builds the attendee wikilink list, strips the Teams join block from the agenda |
| `Meeting (Outlook).md` | your templates folder | the note layout the macro fills in |

## Setup

1. Copy the two files into your vault.
2. Create a QuickAdd **Macro** named exactly **`Meeting (Outlook)`**, with two commands
   in order:
   - **User Script** → `prepare_meeting.js`
   - **Nested Choice → Template**:
     - Template path: `Meeting (Outlook).md`
     - File Name Format: `{{VALUE:title}}`
     - Target folder: where your meeting notes live
     - "File already exists" → *Increment the file name*
3. Toggle the choice on as a command so the URI can trigger it.
4. In the add-in's ⚙ Settings, set your **vault name**.

Open a meeting in Outlook → **Create meeting note**.

## Variables the add-in passes

QuickAdd fills `{{VALUE:name}}` from each `value-name=` URI parameter.

| Variable | From the meeting |
| --- | --- |
| `{{VALUE:title}}` | subject (also the file name) |
| `{{VALUE:attendees}}` | `Name; Name` — the script turns it into `{{VALUE:attendeesYaml}}` |
| `{{VALUE:date}}` | start time |
| `{{VALUE:organizer}}` | organizer |
| `{{VALUE:location}}` | location |
| `{{VALUE:agenda}}` | invite body |

## The one rule

Keep `attendees: {{VALUE:attendeesYaml}}` on a single line — the script emits an
indented YAML list that continues it.
