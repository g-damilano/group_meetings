# mail_name_resolution Modules Usage Guide

The mail name resolution helpers are now split across four focused modules (`mail_name_resolution_prompt.vbs`, `mail_name_resolution_text.vbs`, `mail_name_resolution_directory.vbs`, and `mail_name_resolution_core.vbs`).  Together they centralize the directory lookups used by both **InductionMailer** and the stand-alone GAL utilities.  The most common entry point is `ResolveNameLookup`, which accepts the raw user input together with the requested search field.  Supported field labels include:

| Lookup prefix/value | Targeted field |
| ------------------- | -------------- |
| `email`, `mail`, `smtp`, `address` | SMTP/email address |
| `first`, `given`, `forename` | Given name |
| `surname`, `last`, `lastname`, `family` | Surname |
| `display`, `name` | Display name |
| `full`, `fullname` | Full name (first + surname) |
| `auto` *(default)* | Automatic detection based on the input pattern |

Every resolver returns a dictionary containing the following keys:

- `Found` — Boolean flag indicating whether a directory entry produced authoritative data.
- `First`, `Last`, `Full` — Resolved given name, surname, and display/full name.
- `Email` — The best SMTP-style address surfaced from Exchange, contacts, or proxy attributes.
- `DirectorySource`, `DirectoryAttempted` — Descriptive strings that capture where the winning result came from and which lists were queried.

When the `auto` mode is used, `ResolveNameLookup` delegates to `ResolveNameUsingHeuristics`, which now performs the following steps:

1. **Email recognition** – the helper `IsValidEmail` (now provided by the shared `mail_input_validation.vbs` module) applies a practical regular-expression check so that true email addresses are always resolved via `ResolveNameFromEmail`.
2. **Angle-bracket contact preference** – Outlook-style entries such as `Display Name <alias@example.com>` are detected within semicolon-delimited lists and the parsed SMTP/display data is returned before attempting any other heuristic.  When a matching directory entry cannot be located, the parsed contact details are still surfaced with `Found = True`, `DirectorySource = "Parsed recipient input"`, and the parsed SMTP address so that scripts like **InductionMailer** do not prompt for manual overrides.
3. **Name token analysis** – free-form values (with spaces, commas, dots, underscores, or hyphens) are tokenized and reshuffled into multiple candidate combinations for full-name and display-name searches.
4. **Outlook recipient resolution** – when the direct address list queries cannot identify a candidate, the resolver now falls back to Outlook's built-in recipient resolution using the original name text so that reversed display names (for example, `Surname, Given`) are still matched.
5. **Component fallbacks** – individual given-name and surname lookups are attempted using both the raw input and the inferred first/last name tokens.
6. **Last resort** – if no candidate produced a resolved entry, the final attempt falls back to the best-effort given-name lookup so that the caller still receives a consistent dictionary structure.

For manual calls, prefer routing through `ResolveNameLookup(<input>, "auto")` unless you are explicitly constraining the directory property to inspect.

Additional helper routines worth noting:

- `PromptRecipientLookupValues()` / `PromptRecipientLookupValuesWithOptions(prompt, allowEmpty, warning)` – prompts the console user for one or more semicolon-delimited lookup strings, normalizes commas/newlines, deduplicates repeated entries, and returns a VBScript array ready to be iterated by caller scripts.  `HasRecipientLookupValues(values)` provides a quick way to check whether the returned array contains at least one entry without worrying about `UBound` edge cases.
- `PromptNameResolutionLine(prompt)` – central console reader that trims the standard input result before returning it to the caller; useful when a script just needs a single raw lookup expression (for example, `name:Jane Doe`).
- `ParseRecipientLookupList(text)` – tokenizes raw recipient strings (including those sourced from files) into a deduplicated array using the same normalisation as the interactive prompt helper.

- `IsValidEmail(value)` – returns a Boolean indicating whether the supplied text resembles a SMTP address.
- `ApplyEmailCandidate(info, value)` – updates the resolution dictionary with the strongest SMTP-style email surfaced by a lookup while preserving previously collected values when appropriate.
- `CollectFlexibleNameTokens(text)` – extracts individual name tokens after normalizing common separators (spaces, commas, dots, underscores, hyphens, slashes, and tabs).
- `AttemptRecipientResolveForValue(ns, rawValue, info, descriptor)` – reuses Outlook's native recipient resolution for arbitrary search strings and applies any additional directory metadata returned by the address entry.
- `MaybeRefineNameUsingGivenName(emailAddress, info)` – supplements an email-derived lookup using the directory when the alias matches the resolved first name.

These helpers are safe to reuse in future scripts that need to interpret loosely formatted people identifiers before delivering mail merges or personal induction packs.
