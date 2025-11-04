Brandt-Talbot Induction Agent Package

Structure:
- AGENT.md: Operator guide
- blueprint.json: Machine-readable blueprint
- Induction/: Email templates and form placeholders
- InductionMailer_developer_notes.md: Integration notes and expectations
- package_manifest.json: list of packaged files

Next steps for integrator:
- Ensure mail_utils_utf8_plus_dynamiclinks_ext.vbs exists in repository and exposes the required interface.
- Implement InductionMailer.vbs to read blueprint.json, resolve names, load templates, inject greetings, and call mail_utils.
