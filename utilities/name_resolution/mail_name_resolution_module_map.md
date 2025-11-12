# Mail Name Resolution Module Map

To keep the name-resolution helpers maintainable, every VBScript module under `utilities/name_resolution/` focuses on a single responsibility.  Include `utilities/MNS_loader.vbs` from console scripts so that this layout is honoured automatically.  When introducing new helpers, match their intent to the module described below so that we continue to avoid monolithic catch-all files.

| Module | Primary responsibility | Representative routines |
| ------ | ---------------------- | ----------------------- |
| `MNS_prompt.vbs` | Console prompts and typed-input collection. | `PromptRecipientResolutionPackage`, `PromptNameResolutionLine`. |
| `MNS_text.vbs` | Tokenisation and text normalisation shared across resolvers. | `CollectFlexibleNameTokens`, `JoinTokens`, `BuildReversedTokenDisplay`. |
| `MNS_directory_recording.vbs` | Capturing Outlook/GAL responses for diagnostics and replay. | `SetLookupResponseRecorder`, `RecordLookupResponse`, `ReportAddressListLookup`. |
| `MNS_directory_lists.vbs` | Discovering address list names and ensuring they are valid lookup targets. | `BuildDirectorySearchList`, `CollectLocalContactsListNames`, `ShouldTreatAsLocalContactsList`. |
| `MNS_directory_population.vbs` | Hydrating the resolution dictionary from contacts, Exchange users, and address entries. | `PopulateNameFromAddressLists`, `PopulateNameFromPropertyAccessor`, `AttemptAliasRecipientResolve`. |
| `MNS_directory_matching.vbs` | Aggregating potential directory candidates, scoring them, and recording match reasons. | `CollectAddressEntriesByFields`, `ChooseBestAddressEntry`, `EnumeratePartialMatches`, `OutputMatchReasons`. |
| `MNS_contacts.vbs` | Rendering contact diagnostics and checking whether contact items satisfy lookup tokens. | `OutputContactSummary`, `ContactItemMatchesTokens`, `BuildMatchTokenList`. |
| `MNS_collections.vbs` | Shared collection/dictionary helpers used by multiple modules. | `CreateArrayList`, `CreateTextDictionary`, `GetCollectionCount`. |
| `MNS_directory.vbs` | Legacy shim that documents the module split (no logic). | N/A |
| `MNS_core.vbs` | Public resolver entry points and heuristics. | `ResolveNameLookup`, `ResolveNameUsingHeuristics`, `ResolveNameFromEmail`. |
| `MNS_console.vbs` | Console orchestration that ties the modules together. | `DumpResolutionBlock`, `DumpRecordedOutlookResponses`. |

## Contributing guidelines

1. **Keep responsibilities narrow.** If a helper touches the Outlook object model to gather data, it belongs in `MNS_directory_population.vbs`.  If it merely formats diagnostic output, it should live with the other output helpers.
2. **Favour reuse.** When adding a helper that more than one script might call, prefer extending the relevant module above instead of duplicating code within individual diagnostics.
3. **Add documentation.** Update this map and `mail_name_resolution_usage.md` whenever a new module is introduced or a helper moves categories so future contributors can navigate the utilities quickly.  If a new `.vbs` file is added under `utilities/name_resolution/`, the loader will pick it up automatically, so documenting the intent is the main guard against sprawling modules.
