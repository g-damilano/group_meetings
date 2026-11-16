# VBS trace + Python port map

This file is a *trace* of the VBS code in the provided zip, listing exported symbols and where they live in the Python port.

> Note: `InductionMailer.vbs` references a `mail_utils_mailer_dynamiclinks.vbs` library that is **not present** in the zip. The Python port implements the equivalent delivery/composition features using the included `mail_composer.vbs` / `mail_delivery_core.vbs`.

## `InductionMailer.vbs`

**Ported to:** induction_mailer/cli.py (InductionMailerApp) + induction_mailer/template/*

**Symbols:**

- `BootstrapIncludeEntry`
- `LoadUtilityModules`
- `Main`
- `PromptLine`
- `PromptOperationMode`
- `InitializeEnvironment`
- `LoadMailerLibrary`
- `EnsureLinkMapAvailable`
- `LoadTemplates`
- `CollectTemplates`
- `ParseTemplate`
- `ExtractHeaderAndBody`
- `ParseHeaderBlock`
- `BuildTemplateLabel`
- `BuildStageIndex`
- `SelectTemplatesByStage`
- `SelectTemplatesIndividually`
- `ResolveDeliveryMode`
- `ProcessTemplate`
- `HandleBroadcast`
- `HandlePersonal`
- `GetRecipientResolutionArrayUpperBound`
- `NormalizeNameParts`
- `NormalizeNameKey`
- `GetNormalizedNameValue`
- `EnsureCapitalizedNameText`
- `FindDictionaryKey`
- `CapitalizeNamePart`
- `ReplaceWholeWordCaseInsensitive`
- `IsWordBoundary`
- `IsWordCharacter`
- `ReviewResolvedName`
- `LoadLinkMapContext`
- `DetermineAction`
- `LogEvent`

## `group_meetings_manager.vbs`

**Ported to:**

- `induction_mailer/group_meetings/*` (workflow, workbook I/O, content builders)
- `induction_mailer/scripts/group_meetings_manager.py` (CLI entrypoint)

**Key mappings:**

- `Main` / `Coordinator` → `GroupMeetingsManager.run()` / `_coordinator()`
- `SyncMeetingWorkbook` → `group_meetings.sync.rclone_copy()`
- `LoadMeetingWorkbook`/`CloseMeetingWorkbook` → `group_meetings.workbook.MeetingWorkbook` (openpyxl)
- `EmailExists*` / `HasAlertBeenSentToday` → `OutlookClient.mail_exists*` / `OutlookClient.has_sent_today()`
- `DistributeEmail` / `AlertEmailWithAttachment` → `GroupMeetingsManager._send_plain_email()`

## `utilities/MNS_loader.vbs`

**Ported to:** Package structure under induction_mailer/name_resolution/*

**Symbols:**

- `LoadNameResolutionModules`
- `ResolveNameResolutionRoot`
- `ResolveFolderPath`
- `CollectNameResolutionModules`
- `SortNameResolutionModules`
- `ShouldSwapModules`
- `ModulePriority`

## `utilities/bootstrap_entry.vbs`

**Ported to:** (VBS-only bootstrap; Python uses normal imports)

**Symbols:**

- `BootstrapEnsureLoaded`
- `BootstrapReadAllText`

## `utilities/bootstrap_loader.vbs`

**Ported to:** (VBS-only bootstrap; Python uses normal imports)

**Symbols:**

- `BootstrapLocateNearest`
- `BootstrapReadAllText`
- `EnsureCommonBootstrapLoaded`

## `utilities/common_bootstrap.vbs`

**Ported to:** (VBS-only bootstrap; Python uses normal imports)

**Symbols:**

- `GetScriptFolder`
- `LoadBootstrapModules`
- `LoadModule`
- `NormalizeModulePath`
- `ResolveModulePath`
- `RemoveByteOrderMark`

## `utilities/common_console.vbs`

**Ported to:** Console prompting inside induction_mailer/cli.py

**Symbols:**

- `EnsureConsoleHost`
- `DescribeBoolean`
- `DescribeField`

## `utilities/common_outlook.vbs`

**Ported to:** induction_mailer/outlook/client.py

**Symbols:**

- `GetOutlookApp`
- `EmbedImagesFromFolder`
- `DefaultLogPath`
- `LogAction`

## `utilities/gal_address_lookup.vbs`

**Ported to:** induction_mailer/scripts/gal_lookup.py

**Symbols:**

- `BootstrapIncludeEntry`
- `Main`
- `RequireBootstrapModules`

## `utilities/haoran_qin_directory_probe.vbs`

**Ported to:** induction_mailer/scripts/directory_probe.py

**Symbols:**

- `BootstrapIncludeEntry`
- `Main`
- `RequireBootstrapModules`
- `RunEmailBaselineCheck`
- `RunRecipientResolveProbes`
- `RunGalDisplayEnumerationProbes`
- `ScanGalAddressListsForMatches`
- `ScanContactFoldersForTokens`
- `OutputAddressEntrySummary`
- `DumpResolutionBlock`
- `DumpResolutionDetails`
- `DumpRecordedOutlookResponses`
- `BuildCandidateProbeList`
- `AddProbe`
- `SummarizeBaselineComparison`
- `baselineLookup`
- `RecordBaselineEntry`
- `AppendAddressEntries`
- `CollectAddressEntriesByProperty`
- `GetGalListNames`
- `OpenAddressList`
- `DumpVariantStructure`
- `DumpDictionaryVariant`
- `DumpArrayListVariant`
- `DumpArrayVariant`
- `ShouldExpandVariant`
- `SafeTypeName`
- `DescribeVariant`

## `utilities/mail_collections.vbs`

**Ported to:** Python lists / small helpers in induction_mailer/template/repository.py

**Symbols:**

- `CreateCaseInsensitiveDictionary`
- `GetDictionaryValue`
- `CloneDictionary`
- `ArrayFromArrayList`
- `ContainsInArrayList`
- `Capitalize`
- `TemplateSortValue`
- `SortTemplateList`
- `BuildOrderedDictionary`
- `ParseNumberList`

## `utilities/mail_composer.vbs`

**Ported to:** induction_mailer/outlook/composer.py + induction_mailer/template/link_map.py + induction_mailer/template/assets.py

**Symbols:**

- `ComposeAndSendMail`
- `ComposerApplyBodyTransforms`
- `ComposerResolveSignatureHtml`
- `ComposerBuildHtmlDocument`
- `ComposerAddAttachments`
- `ComposerAddAttachmentPath`
- `ComposerAddInlineImages`
- `ApplyLinkMap`
- `LoadHtmlFromSource`
- `ReadUTF8File`
- `HttpGetUtf8`
- `GetOutlookApp`
- `ComposerLog`
- `ComposerGetOption`

## `utilities/mail_delivery_core.vbs`

**Ported to:** induction_mailer/outlook/composer.py + induction_mailer/template/assets.py

**Symbols:**

- `MailDeliveryExecute`
- `MailDeliveryResolveBodyContext`
- `MailDeliveryLoadHtml`
- `MailDeliveryApplyLinkCallback`
- `MailDeliveryApplyLinkReplacements`
- `MailDeliveryPrepareInlineSupport`
- `MailDeliveryRewriteInlineReferences`
- `MailDeliverySanitizeContentId`
- `MailDeliveryResolveAttachmentsOption`
- `MailDeliveryResolveAttachmentList`
- `MailDeliveryResolveAttachmentPath`
- `MailDeliveryAppendPath`
- `MailDeliveryIsAbsolutePath`
- `MailDeliveryResolveLogPath`
- `MailDeliveryDefaultLogPath`
- `MailDeliveryScriptFolder`
- `MailDeliveryResolveLogHandler`
- `MailDeliveryLogViaHandler`
- `MailDeliveryLog`
- `MailDeliveryEnsureComposer`
- `MailDeliveryLoadSignatureAutoCharset`
- `MailDeliveryReadHtmlAuto`
- `MailDeliveryDetectHtmlCharset`
- `MailDeliveryMarkerCount`
- `MailDeliveryCountOf`
- `MailDeliverySanitizeInlineCid`
- `MailDeliveryGetOption`
- `MailDeliveryCopyIfExists`

## `utilities/mail_delivery_utils.vbs`

**Ported to:** induction_mailer/template/assets.py

**Symbols:**

- `ResolveAttachments`
- `IsAbsolutePath`
- `PrepareInlineImageSupport`
- `RewriteInlineReferences`
- `SanitizeContentId`

## `utilities/mail_input_validation.vbs`

**Ported to:** induction_mailer/validation.py

**Symbols:**

- `IsValidEmail`

## `utilities/mail_template_utils.vbs`

**Ported to:** induction_mailer/template/repository.py + induction_mailer/template/placeholders.py

**Symbols:**

- `ReadFileSmart`
- `NormalizeLineEndings`
- `TrimLeadingBlankLines`
- `BuildContextFromMeta`
- `PromptForEmptyTags`
- `EnsurePlaceholderValues`
- `ExtractPlaceholders`
- `ApplyContextToText`
- `MergeRecipients`
- `AppendAddressesFromString`
- `AppendAddress`
- `SafeText`

## `utilities/name_resolution/MNS_collections.vbs`

**Ported to:** induction_mailer/name_resolution/collections.py + induction_mailer/utils.py (CaseInsensitiveDict)

**Symbols:**

- `CreateArrayList`
- `EnsureArrayList`
- `CreateTextDictionary`
- `GetCollectionCount`
- `GetDictionaryCount`

## `utilities/name_resolution/MNS_console.vbs`

**Ported to:** (folded into cli logging / future expansion)

**Symbols:**

- `ParseLookupExpression`
- `DescribeLookupField`
- `DisplayResolution`
- `DisplayResolutionSummary`
- `DisplayResolutionFields`

## `utilities/name_resolution/MNS_contacts.vbs`

**Ported to:** induction_mailer/name_resolution/contacts.py

**Symbols:**

- `BuildMatchTokenList`
- `OutputComponentSummary`
- `OutputContactSummary`
- `ContactItemMatchesTokens`
- `ContactFieldValue`
- `IsContactLikeItem`
- `DescribeItemClass`
- `ContainsAllTokens`

## `utilities/name_resolution/MNS_core.vbs`

**Ported to:** induction_mailer/name_resolution/resolver.py

**Symbols:**

- `CreateEmptyResolutionInfo`
- `AcquireResolutionNamespace`
- `InitializeResolutionInfo`
- `EnsureFullNameField`
- `ResolveNameLookup`
- `ResolveNameFromEmail`
- `ResolveNameUsingHeuristics`
- `ResolveNameUsingHeuristicsWithOptions`
- `MergeParsedContactIntoInfo`
- `ResolveNameFromGivenName`
- `ResolveNameFromSurname`
- `ResolveNameFromDisplayName`
- `ResolveNameFromFullName`
- `ResolutionWasSuccessful`

## `utilities/name_resolution/MNS_directory.vbs`

**Ported to:** (not explicitly mapped; see package)

_No Function/Sub definitions detected._

## `utilities/name_resolution/MNS_directory_lists.vbs`

**Ported to:** induction_mailer/name_resolution/directory_lists.py

**Symbols:**

- `BuildDirectorySearchList`
- `AddListNameIfMissing`
- `CollectLocalContactsListNames`
- `AppendAvailableContactList`
- `ShouldTreatAsLocalContactsList`
- `GetDefaultContactsAddressListName`

## `utilities/name_resolution/MNS_directory_matching.vbs`

**Ported to:** induction_mailer/name_resolution/directory_matching.py

**Symbols:**

- `CollectAddressEntriesByFilters`
- `CollectAddressEntriesByFields`
- `AddAddressEntryToCollection`
- `MergeAddressEntryMatches`
- `DescribeAddressEntry`
- `BuildAddressEntryTrackingKey`
- `AppendAddressEntryResults`
- `FindAddressEntryByFullName`
- `CollectAddressEntriesByFullName`
- `EscapeRestrictionValue`
- `FindAddressEntryByFields`
- `BuildAddressEntryCandidate`
- `CallerProfileHasOrgDetails`
- `ScoreEntryCandidate`
- `ChooseBestAddressEntry`
- `BuildSelectionSummary`
- `GetCurrentUserProfileInfo`
- `AppendReasonIfMissing`
- `EnsureAddressEntryKey`
- `EnsureReasonList`
- `AddEntryWithReasonCollection`
- `BuildAddressEntrySearchBlob`
- `EnumeratePartialMatches`
- `OutputMatchReasons`
- `OutputSelectionSummary`
- `BuildCandidateProbeList`
- `AddProbe`

## `utilities/name_resolution/MNS_directory_population.vbs`

**Ported to:** induction_mailer/name_resolution/directory_population.py

**Symbols:**

- `FindContactInAddressList`
- `ContactMatchesEmail`
- `PopulateNameFromContactItem`
- `PopulateNameFromPropertyAccessor`
- `PopulateNameFromAddressLists`
- `PopulateNameFromAddressListsByGivenName`
- `PopulateNameFromAddressListsBySurname`
- `PopulateNameFromAddressListsByDisplayName`
- `EnumerateDirectoryMatchesByDisplayName`
- `PopulateNameFromAddressListsByFullName`
- `AttemptAliasRecipientResolve`
- `AttemptRecipientResolveForValue`
- `MaybeRefineNameUsingGivenName`

## `utilities/name_resolution/MNS_directory_recording.vbs`

**Ported to:** induction_mailer/name_resolution/recording.py

**Symbols:**

- `SetLookupResponseRecorder`
- `ClearLookupResponseRecorder`
- `RecordLookupResponse`
- `RecordRecipientResolutionResponse`
- `BuildLookupResponseSummary`
- `RecordDirectoryAttempt`
- `ReportAddressListLookup`

## `utilities/name_resolution/MNS_prompt.vbs`

**Ported to:** induction_mailer/name_resolution/prompt.py

**Symbols:**

- `PromptNameResolutionLine`
- `ParseRecipientLookupList`
- `PromptRecipientLookupValues`
- `PromptRecipientLookupValuesWithOptions`
- `PromptRecipientResolutionPackage`
- `PromptRecipientResolutionPackageWithOptions`
- `BuildRecipientResolutionPackage`
- `HasRecipientResolutionEntries`
- `GetRecipientResolutionAddresses`
- `GetRecipientResolutionEntries`
- `GetRecipientResolutionCount`
- `CreateRecipientResolutionEntry`
- `DeterminePreferredRecipientAddress`
- `NormalizeLookupField`
- `HasRecipientLookupValues`
- `ArrayFromList`

## `utilities/name_resolution/MNS_text.vbs`

**Ported to:** induction_mailer/name_resolution/text.py

**Symbols:**

- `CollectNameTokens`
- `CollectFlexibleNameTokens`
- `ExtractDisplayNameAndEmail`
- `JoinTokens`
- `BuildReversedTokenDisplay`
- `AddCandidateValue`
- `ParseFullNameComponents`
- `NormalizeEmailText`
- `ApplyEmailCandidate`
- `ApplyFullNameCandidate`
- `ApplyFirstNameCandidate`
- `ShouldReplaceFirstName`
- `ExtractEmailLocalPart`
- `SafeLower`
- `GetAccessorString`
