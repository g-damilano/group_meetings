# Python Validation

Python support begins in Stage 5 as a development-side validator and test
harness. It is not required for operational VBS deployment.

The validator uses only the Python standard library:

```text
python -B .\python\validate_scaffold.py --root .
python -B .\python\contract_fixture_report.py --root . --format json --fail-if-invalid
python -B .\python\workflow_input_report.py --root . --format json --fail-if-invalid
python -B .\python\command_route_report.py --root . --format json --fail-if-invalid
python -B -m unittest discover -s .\python\tests
```

It checks:

- JSON contract files parse and contain expected structures;
- shared contract fixtures represent `MessageSpec`, `RecipientSpec`,
  `DeliveryPolicy`, `ManagedEmailKey`, `DispatchResult`, `WorkflowContext`, and
  `HealthReport` with dry-run-safe example data;
- dispatch state fixtures match the documented state machine;
- VBS dispatch state constants, JSON schema enums, and dispatch-state fixtures
  stay in parity;
- the watchdog report-only health fixture keeps report-only, no-live-Outlook,
  no-mutation semantics while modelling managed, unmarked, stale, deferred, and
  uncertain Outbox items;
- the watchdog managed-repair guardrail fixture keeps repair validation
  fake-object-only, blocks live Outlook access and sending, and proves unmarked
  or uncertain mail cannot become mutable;
- workbook-shaped CSV examples contain required headers;
- group-meetings and induction example inputs are route-ready for dry-run use
  through `Scaffold.vbs`;
- the shared quickstart command fixture contains only scaffold-local dry-run,
  report-only, or deliberately blocked fail-safe routes;
- the root `Scaffold.vbs` group-meetings, induction, and watchdog routes return
  dry-run/report-only output matching the shared examples;
- induction template manifest paths, attachment paths, inline-image folders,
  delivery actions, deferred-delivery values, and urgent verification flags
  validate;
- public scaffold text, including XML examples, does not contain non-example
  email addresses, private path markers, or live SharePoint-style URLs.

The validator complements VBS smoke tests. It must not become the sole source of
architectural truth.

`contract_fixture_report.py` is an optional development-side report over the
shared contract fixture. It is report-only, performs no Outlook work, sends no
mail, and writes no files.

`workflow_input_report.py` is an optional development-side report over the
public workflow input examples. It estimates dry-run message counts and reports
input issues without calling VBS, Outlook, or mailbox APIs.

`command_route_report.py` is an optional cross-runtime report over the root
router. It runs only public help, validation, dry-run/report-only
`Scaffold.vbs` commands and blocked fail-safe guardrail commands declared in
`fixtures/quickstart/dry_run_routes.json`, then compares their key-value output
with `workflow_input_report.py`.
