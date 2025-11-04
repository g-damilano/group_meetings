# InductionMailer Developer Notes

This document explains how the agent should use the templates and mail utility.

Requirements:
- The implementation must call `mail_utils_utf8_plus_dynamiclinks_ext.vbs` for all sends.
- The mail utility is expected to accept parameters: to, cc, subject, body_html_path (or body text), attachments[], isHTML, update_links=True.
- The script must resolve recipient display names via Outlook Address Book for greeting injection.
- For group sends (EMAIL_6), the 'To' should list all inductee addresses; for individual sends, each inductee receives a separate message.

Post-send behavior:
- After EMAIL_6, automatically generate and send `Induction/Reception Notification - auto.htm` content to chemistryreception@imperial.ac.uk listing dates and names.
- Log send status to `Induction/sent_log.csv` with columns: timestamp, option, template, to, cc, status, message_id (if available).

Notes on templates:
- The htm files inside Induction/ are the canonical templates. The mailer should open them, refresh any dynamic links via the mail_utils helper, inject the greeting or group header, and then pass to the mail utility.

Security & privacy:
- Do not persist any photos or sensitive personal data outside the secure SharePoint or local secure folder.
- Access forms saved by inductees should be moved only to the Brandt-Talbot Safety Folder (SharePoint).
