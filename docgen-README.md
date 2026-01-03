# docgen.js

This file provides browser DOCX generation using the docx UMD bundle. It creates an email-styled .docx containing credentials (blue-styled table), credit-card test details, and clickable Dashboard/API Docs links. It falls back to an HTML-as-.doc when docx isn't available.

How to use:
1. Add `docgen.js` to the repository root or `public/`.
2. Include it in `index.html` after the docx UMD script tag:

```html
<script src="https://unpkg.com/docx@9.0.2/build/index.umd.js"></script>
<script src="/docgen.js"></script>
```

3. Use the existing button with id `previewGenerate` to trigger generation, or call `window.MadfuDocGen.generate()`.

Notes:
- Colors can be adjusted at the top of `docgen.js` (CREDENTIAL_HEADER_FILL, CREDENTIAL_VALUE_FILL).
- The script uses ExternalHyperlink so Dashboard/API Docs entries are clickable in Word.
- If docx isn't available or generation fails, an HTML-as-.doc fallback is downloaded.
