# Vernacular

Teach your AI how to work with your tools.

Vernacular is a browser-based prototype for teaching a narrow PowerPoint operation from examples, testing it on a fresh deck, and exporting it for reuse with a file-capable assistant.

## Try the workflow

1. Open **Teach**. The bundled examples illustrate changing the launch date while preserving the rehearsal date.
2. Select **Test on another file**, choose a new date, and run the test. The sample test deck starts from a different date.
3. Download the editable PowerPoint copy. The site checks the resulting XML and verifies that the other package parts are unchanged.
4. Save the operation to the browser library. **Export for assistant** downloads the operation, a skill description, and a Python runner. The runner can also be used directly.

Use your own examples by saving an original .pptx and a copy with one labeled date edited. The label and date must be in the same text paragraph, for example "Launch date: Nov 12". Try the inferred operation on a third deck with that label once. The app accepts English month names and ISO dates. It stops when a field is missing or ambiguous. It previews user files as slide text. Open the editable download in PowerPoint for visual verification.

This prototype uses a local inference rule with a bounded PowerPoint editor. It does not call a hosted model or upload files. The exported runner requires Python 3 and uses the standard library. Live assistant integration and support for other file formats are future work.

## Run locally

    python -m http.server 8000

Then open http://localhost:8000. The project is static and supports GitHub Pages from the root of main. JSZip is bundled with its license so the site can process files without fetching third-party JavaScript. Sample project files and artwork are illustrative.
