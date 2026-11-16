# Group Meetings Workflow Runner

This Stage 2 runner wires the scaffold group-meetings workflow into a dry-run
orchestration path. It combines:

- rota/workbook-shaped CSV loading;
- meeting-record creation;
- message-plan generation;
- core workflow facade composition and managed-key dry-run dispatch;
- a single workflow result for entry scripts or tests.

The runner deliberately does not open Excel, open Outlook, send mail, delete
mail, or repair Outbox items. It is intended as the operational shape that a
future VBS entrypoint can call after gathering user intent.

## Contract Shape

`SC_GM_RunGroupMeetingsCsvDryRun` returns a dictionary with:

- `state`: dispatch state for the overall workflow run;
- `reason`: short status explanation;
- `action_required`: true when an operator must review the rota or workflow;
- `dry_run`: copied from `WorkflowContext`;
- `target_date` and `rota_source`;
- `row_count`;
- `plan_state`, `plan_reason`, `planned_count`, `skipped_count`;
- `composed_count`, `failed_count`;
- `load_result`, `plan`, `skips`, and `dispatch_results`.

The runner uses the same shared dispatch states as the rest of the scaffold:

- `COMPOSED` when rota loading, planning, and dry-run composition succeed;
- `ACTION_REQUIRED` when the rota is missing, incomplete, or intentionally
  produces a manager alert;
- `FAILED_SAFE` when dry-run is disabled or a composition step fails safe.

## Safety Boundary

The runner is dry-run only. If `WorkflowContext("dry_run")` is false, the
runner returns `FAILED_SAFE` before loading or composing messages. This keeps
Stage 2 focused on behavioural migration and prevents live mailbox effects.

Workflow code remains layered:

- rota adapter owns workbook-shaped data conversion;
- message-plan adapter owns group-meetings workflow decisions;
- core modules own recipient resolution, message composition contracts, managed
  keys, dispatch states, and logging;
- the runner only coordinates those pieces.

## Future Entry Point Use

A VBS entrypoint can call this runner after selecting a target date and rota
source. Live display/draft/send modes should be enabled only after the Stage 2
behavioural checks are complete and the dispatch path has explicit operator
controls.
