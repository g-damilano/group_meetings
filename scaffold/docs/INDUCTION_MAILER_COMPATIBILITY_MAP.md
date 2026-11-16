# Induction Mailer Compatibility Map

The legacy induction mailer remains a frozen behavioural reference. Stage 3
scaffold work expresses the reusable behaviours through VBS-first contracts and
anonymised fixtures.

## Preserved Behavioural Concepts

- Template-driven communication: templates provide subject metadata and message
  body content.
- Delivery modes:
  - `broadcast`: one message to many recipients;
  - `personal`: one message per recipient with recipient-specific placeholders.
- Recipient-derived placeholders are available to templates, including first
  name, surname, full name, email, alias, start date, and cohort.
- Recipient aliases in template metadata can be resolved for `cc` and `bcc`
  without the workflow opening Outlook or the address book.
- Attachments and inline images remain core composition concerns, not workflow
  concerns. Template manifests name attachment base folders, attachment lists,
  and inline-image folders; the workflow delegates resolution to the core
  facade.
- Template delivery policy metadata maps to `DeliveryPolicy` fields, including
  display-first dry-run actions, defer minutes, and urgent post-submit
  verification flags.
- Dry-run/display-first operation is the scaffold default.
- Enabled templates with missing files are reported as `TEMPLATE_MISSING` and
  produce `ACTION_REQUIRED` when no messages can be planned.

## Scaffold Mapping

| Legacy concern | Scaffold location |
| --- | --- |
| Template selection | template manifest fixture and input adapter |
| Recipient capture/resolution package | recipient CSV fixture and input adapter |
| Broadcast vs personal branching | induction message-plan adapter |
| Placeholder context construction | induction message-plan adapter plus core template engine |
| CC/BCC alias resolution | recipient directory plus core recipient resolver |
| Attachment and inline-image resolution | core workflow facade helpers |
| Deferred delivery and urgent verification metadata | `DeliveryPolicy` contracts |
| Missing template diagnostics | induction message-plan adapter skip records |
| Message composition | core workflow facade and mail composer |
| Managed identity and dispatch states | core contracts and managed dispatch |
| User intent gathering | dry-run VBS entrypoint |

## Deliberate Stage 3 Limits

- No Outlook COM is created by induction workflow or orchestration modules.
- No real email is sent.
- No live address-book or GAL lookup is performed.
- No private templates, addresses, links, logs, or institutional identifiers are
  copied into scaffold files.
- Live display/draft/send modes remain blocked until later guarded migration
  work explicitly enables them.
