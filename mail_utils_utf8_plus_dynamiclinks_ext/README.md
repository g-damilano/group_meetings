# mail_utils_utf8_plus_dynamiclinks_ext

This package extends the Outlook mailer to **dynamically inject current hyperlinks** into an **external HTML template**.

## How It Works
1. Place your HTML template as `template.html` (UTF‑8 encoded).
2. Use placeholders like `{{LINK_1}}`, `{{LINK_2}}` in the HTML.
3. Maintain `links_map.txt` with the current URL mapping.
4. Run `test_dynamic_links.vbs` — it will load the HTML, replace placeholders, and display the composed mail in Outlook.

## Files
- `mail_utils_mailer_dynamiclinks.vbs` — Outlook mailer with dynamic link mapping.
- `test_dynamic_links.vbs` — test script loading external HTML.
- `template.html` — example HTML template.
- `links_map.txt` — URL mapping file.
- `README.md` — this file.

## Usage
1. Extract all files to a folder.
2. Edit `template.html` and `links_map.txt` as needed.
3. Double‑click `test_dynamic_links.vbs` to preview the composed message.
4. Update only the `links_map.txt` to keep URLs current — no need to edit the HTML again.
