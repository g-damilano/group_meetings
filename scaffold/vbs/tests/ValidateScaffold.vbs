Option Explicit
SetLocale 1033

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim scaffoldRoot
scaffoldRoot = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder))

Dim bootstrapPath
bootstrapPath = fso.BuildPath(scaffoldRoot, "vbs\core\bootstrap.vbs")
If Not fso.FileExists(bootstrapPath) Then
  WScript.Echo "FAILED: bootstrap not found: " & bootstrapPath
  WScript.Quit 1
End If

Dim ts
Set ts = fso.OpenTextFile(bootstrapPath, 1)
ExecuteGlobal ts.ReadAll
ts.Close

SC_LoadCoreModules

Dim failures
failures = 0

Sub RequireFile(relativePath)
  Dim path
  path = fso.BuildPath(scaffoldRoot, relativePath)
  If fso.FileExists(path) Then
    WScript.Echo "OK: " & relativePath
  Else
    WScript.Echo "MISSING: " & relativePath
    failures = failures + 1
  End If
End Sub

Sub RequireState(stateText)
  If SC_IsKnownDispatchState(stateText) Then
    WScript.Echo "OK state: " & stateText
  Else
    WScript.Echo "FAILED state: " & stateText
    failures = failures + 1
  End If
End Sub

RequireFile "README.md"
RequireFile "LICENSE"
RequireFile "Scaffold.vbs"
RequireFile "RunGroupMeetings.vbs"
RequireFile "RunInductionMailer.vbs"
RequireFile "OutlookWatchdog.vbs"
RequireFile "ValidateWorkflow.vbs"
RequireFile "docs\ARCHITECTURE.md"
RequireFile "docs\CONTRACTS.md"
RequireFile "docs\DISPATCH_STATE_MACHINE.md"
RequireFile "docs\WATCHDOG_SAFETY_DOCTRINE.md"
RequireFile "docs\WATCHDOG_SIDECAR.md"
RequireFile "docs\WATCHDOG_MANAGED_REPAIR.md"
RequireFile "docs\MIGRATION_SEQUENCE.md"
RequireFile "docs\GROUP_MEETINGS_COMPATIBILITY_MAP.md"
RequireFile "docs\GROUP_MEETINGS_WORKBOOK_CONVENTIONS.md"
RequireFile "docs\GROUP_MEETINGS_WORKFLOW_RUNNER.md"
RequireFile "docs\GROUP_MEETINGS_DRY_RUN_ENTRYPOINT.md"
RequireFile "docs\INDUCTION_MAILER_COMPATIBILITY_MAP.md"
RequireFile "docs\INDUCTION_MAILER_CONVENTIONS.md"
RequireFile "docs\INDUCTION_MAILER_DRY_RUN.md"
RequireFile "docs\LEGACY_ALTERNATE_MAP.md"
RequireFile "docs\LEGACY_MIRROR.md"
RequireFile "docs\REPO_BOUNDARY_AUDIT.md"
RequireFile "docs\SHARED_EXAMPLES.md"
RequireFile "docs\PYTHON_VALIDATION.md"
RequireFile "docs\PUBLIC_RELEASE_CHECKLIST.md"
RequireFile "docs\PUBLICATION_BOUNDARY.md"
RequireFile "docs\RELEASE_AUDIT.md"
RequireFile "docs\RELEASE_CANDIDATE_PLAN.md"
RequireFile "docs\RELEASE_EXPORT_PLAN.md"
RequireFile "docs\RELEASE_GATE.md"
RequireFile "docs\RELEASE_MANIFEST.md"
RequireFile "docs\RELEASE_NOTES_TEMPLATE.md"
RequireFile "docs\SCAFFOLD_ORCHESTRATION.md"
RequireFile "docs\VALIDATION_RUNBOOK.md"
RequireFile ".gitignore"
RequireFile "contracts\workflow_schema.json"
RequireFile "fixtures\contracts\stage4_shared_contract_fixture.json"
RequireFile "fixtures\dispatch_states\states.json"
RequireFile "fixtures\quickstart\dry_run_routes.json"
RequireFile "fixtures\sanitisation\blocked_patterns.json"
RequireFile "fixtures\watchdog\report_only_health_fixture.json"
RequireFile "fixtures\watchdog\managed_repair_guardrail_fixture.json"
RequireFile "migration\legacy_alternate_map.json"
RequireFile "examples\workbooks\README.md"
RequireFile "examples\workbooks\group_meetings_rota_example.csv"
RequireFile "examples\workbooks\induction_recipients_example.csv"
RequireFile "examples\workbooks\induction_template_manifest_example.csv"
RequireFile "examples\templates\README.md"
RequireFile "examples\templates\induction_assets\overview.txt"
RequireFile "examples\templates\induction_assets\setup checklist.txt"
RequireFile "examples\templates\induction_broadcast_example.html"
RequireFile "examples\templates\induction_personal_example.html"
RequireFile "examples\templates\induction_personal_example.files\logo one.png"
RequireFile "examples\task_scheduler\watchdog_report_only_task.xml"
RequireFile "python\validate_scaffold.py"
RequireFile "python\contract_fixture_report.py"
RequireFile "python\workflow_input_report.py"
RequireFile "python\command_route_report.py"
RequireFile "python\legacy_intake_audit.py"
RequireFile "python\legacy_mirror_audit.py"
RequireFile "python\legacy_sanitisation_plan.py"
RequireFile "python\repo_boundary_audit.py"
RequireFile "python\release_audit.py"
RequireFile "python\release_candidate_plan.py"
RequireFile "python\release_export_plan.py"
RequireFile "python\release_gate.py"
RequireFile "python\release_manifest.py"
RequireFile "python\tests\test_validate_scaffold.py"
RequireFile "python\tests\test_contract_fixture_report.py"
RequireFile "python\tests\test_workflow_input_report.py"
RequireFile "python\tests\test_command_route_report.py"
RequireFile "python\tests\test_legacy_intake_audit.py"
RequireFile "python\tests\test_legacy_mirror_audit.py"
RequireFile "python\tests\test_legacy_sanitisation_plan.py"
RequireFile "python\tests\test_repo_boundary_audit.py"
RequireFile "python\tests\test_release_audit.py"
RequireFile "python\tests\test_release_candidate_plan.py"
RequireFile "python\tests\test_release_export_plan.py"
RequireFile "python\tests\test_release_gate.py"
RequireFile "python\tests\test_release_manifest.py"
RequireFile "python\tests\test_validation_runbook.py"
RequireFile "vbs\core\contracts.vbs"
RequireFile "vbs\core\logger.vbs"
RequireFile "vbs\core\session.vbs"
RequireFile "vbs\core\folder_access.vbs"
RequireFile "vbs\core\recipient_resolver.vbs"
RequireFile "vbs\core\template_engine.vbs"
RequireFile "vbs\core\mail_assets.vbs"
RequireFile "vbs\core\mail_composer.vbs"
RequireFile "vbs\core\managed_dispatch.vbs"
RequireFile "vbs\core\outbox_inspector.vbs"
RequireFile "vbs\core\health_diagnostics.vbs"
RequireFile "vbs\core\workflow_facade.vbs"
RequireFile "vbs\core\bootstrap.vbs"
RequireFile "vbs\entrypoints\scaffold_launcher.vbs"
RequireFile "vbs\entrypoints\run_group_meetings_dry_run.vbs"
RequireFile "vbs\entrypoints\run_induction_mailer_dry_run.vbs"
RequireFile "vbs\entrypoints\run_watchdog_report_only.vbs"
RequireFile "vbs\workflows\workflow_helpers.vbs"
RequireFile "vbs\workflows\group_meetings_workflow.vbs"
RequireFile "vbs\workflows\group_meetings_message_plan.vbs"
RequireFile "vbs\workflows\group_meetings_rota_adapter.vbs"
RequireFile "vbs\workflows\group_meetings_runner.vbs"
RequireFile "vbs\workflows\induction_mailer_workflow.vbs"
RequireFile "vbs\workflows\induction_mailer_input_adapter.vbs"
RequireFile "vbs\workflows\induction_mailer_message_plan.vbs"
RequireFile "vbs\workflows\induction_mailer_runner.vbs"
RequireFile "vbs\watchdog\outlook_watchdog_sidecar.vbs"
RequireFile "vbs\watchdog\managed_outbox_repair.vbs"
RequireFile "vbs\tests\TemplateAndRecipientSmokeTest.vbs"
RequireFile "vbs\tests\ContractFixtureSmokeTest.vbs"
RequireFile "vbs\tests\MailAssetsSmokeTest.vbs"
RequireFile "vbs\tests\MailComposerSmokeTest.vbs"
RequireFile "vbs\tests\ManagedOutboxDecisionSmokeTest.vbs"
RequireFile "vbs\tests\OutboxInspectorSmokeTest.vbs"
RequireFile "vbs\tests\FolderAccessSmokeTest.vbs"
RequireFile "vbs\tests\WorkflowFacadeSmokeTest.vbs"
RequireFile "vbs\tests\WorkflowSkeletonSmokeTest.vbs"
RequireFile "vbs\tests\GroupMeetingsMessagePlanSmokeTest.vbs"
RequireFile "vbs\tests\GroupMeetingsRotaAdapterSmokeTest.vbs"
RequireFile "vbs\tests\GroupMeetingsRunnerSmokeTest.vbs"
RequireFile "vbs\tests\GroupMeetingsEntrypointSmokeTest.vbs"
RequireFile "vbs\tests\InductionMailerStage3SmokeTest.vbs"
RequireFile "vbs\tests\InductionMailerEntrypointSmokeTest.vbs"
RequireFile "vbs\tests\WatchdogSidecarSmokeTest.vbs"
RequireFile "vbs\tests\ManagedOutboxRepairSmokeTest.vbs"
RequireFile "vbs\tests\PublicReleaseGateSmokeTest.vbs"
RequireFile "vbs\tests\ReleaseManifestSmokeTest.vbs"
RequireFile "vbs\tests\RunAllSmokeTests.vbs"
RequireFile "vbs\tests\ScaffoldAlternateEntrypointsSmokeTest.vbs"
RequireFile "vbs\tests\ScaffoldRouterSmokeTest.vbs"
RequireFile "vbs\tests\fixtures\attachment.txt"
RequireFile "vbs\tests\fixtures\sample_template.html"
RequireFile "vbs\tests\fixtures\sample_template.files\logo one.png"
RequireFile "vbs\tests\fixtures\group_meetings_rota_fixture.csv"
RequireFile "vbs\tests\fixtures\group_meetings_recipients_fixture.csv"
RequireFile "vbs\tests\fixtures\induction_recipients_fixture.csv"
RequireFile "vbs\tests\fixtures\induction_recipients_invalid_fixture.csv"
RequireFile "vbs\tests\fixtures\induction_template_manifest_fixture.csv"
RequireFile "vbs\tests\fixtures\induction_template_manifest_invalid_fixture.csv"
RequireFile "vbs\tests\fixtures\induction_template_manifest_missing_template_fixture.csv"
RequireFile "vbs\tests\fixtures\induction_assets\overview.txt"
RequireFile "vbs\tests\fixtures\induction_assets\setup checklist.txt"
RequireFile "vbs\tests\fixtures\induction_templates\broadcast_overview.html"
RequireFile "vbs\tests\fixtures\induction_templates\personal_welcome.html"
RequireFile "vbs\tests\fixtures\induction_templates\personal_welcome.files\logo one.png"
RequireFile "vbs\tests\fixtures\induction_templates\disabled_template.html"

RequireState SC_STATE_NOT_REQUESTED
RequireState SC_STATE_SEND_REQUESTED
RequireState SC_STATE_QUEUED_OUTBOX
RequireState SC_STATE_DUPLICATE_SUPPRESSED
RequireState SC_STATE_ACTION_REQUIRED
RequireState SC_STATE_FAILED_SAFE

Dim expectedKey
expectedKey = "ORG|GROUP-MEETINGS|2026-05-29|REGULAR|TEAM"
If SC_BuildManagedEmailKey("org", "group meetings", #5/29/2026#, "regular", "team") <> expectedKey Then
  WScript.Echo "FAILED: deterministic managed key generation."
  failures = failures + 1
Else
  WScript.Echo "OK: deterministic managed key generation."
End If

If failures > 0 Then
  WScript.Echo "Validation failed: " & failures & " issue(s)."
  WScript.Quit 1
End If

WScript.Echo "Validation passed."
