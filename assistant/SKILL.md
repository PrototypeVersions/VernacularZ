---
name: vernacular-operation
description: Apply a specific tested date-field change to an editable PowerPoint deck using the operation exported from Vernacular.
---

Use this capability when the user asks to change the date field named in operation.json in an editable PowerPoint file.

1. Read operation.json to check that the requested field matches the exported operation.
2. Ask for the input deck or the intended date if it is missing. Express the date as YYYY-MM-DD.
3. Run the included script with Python 3. It writes a new copy and refuses to overwrite an existing file.

    python apply.py input.pptx YYYY-MM-DD updated.pptx

4. Deliver the editable output and report the before and after values printed by the script. If the script reports a missing or ambiguous field, describe that result and ask for a deck with a unique matching label.
5. Recommend opening the resulting copy in PowerPoint for visual inspection when layout matters. The script checks the file structure and target value.

The downloaded operation is scoped to a single date field with a stable label. It does not change other dates, insert slides, or send invitations. The exported package can be run manually or by an assistant with file and Python execution access.
