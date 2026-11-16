"""Cross-runtime dry-run route report for the root scaffold command router."""

from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path
from typing import Any

import workflow_input_report


SCHEMA = "office-communication-scaffold.command-route-report.v1"
QUICKSTART_ROUTE_SCHEMA = "office-communication-scaffold.quickstart-routes.v1"
QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH = "fixtures/quickstart/dry_run_routes.json"


def parse_key_value_output(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for line in text.splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if key:
            result[key] = value.strip()
    return result


def run_scaffold(root: Path, args: list[str]) -> dict[str, Any]:
    command = ["cscript.exe", "//NoLogo", ".\\Scaffold.vbs", *args]
    completed = subprocess.run(command, check=False, capture_output=True, text=True, cwd=root)
    output = completed.stdout + completed.stderr
    return {
        "command": " ".join(command),
        "exit_code": completed.returncode,
        "output": output,
        "fields": parse_key_value_output(output),
    }


def windows_relative_arg(relative_path: str) -> str:
    return ".\\" + relative_path.replace("/", "\\")


def scaffold_cli_arg(value: object) -> str:
    text = str(value)
    if text.startswith(("examples/", "fixtures/", "contracts/", "docs/")):
        return windows_relative_arg(text)
    return text


def int_field(fields: dict[str, str], name: str, default: int = 0) -> int:
    try:
        return int(fields.get(name, str(default)))
    except ValueError:
        return default


def bool_field(fields: dict[str, str], name: str) -> bool:
    return fields.get(name, "").strip().lower() == "true"


def add_error(errors: list[str], route: str, message: str) -> None:
    errors.append(f"{route}: {message}")


def load_quickstart_fixture(root: Path) -> dict[str, Any]:
    path = root / QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    return data if isinstance(data, dict) else {}


def command_path_tokens(route: dict[str, Any]) -> list[str]:
    command = route.get("command", [])
    if not isinstance(command, list):
        return []
    return [str(token) for token in command]


def expected_count(input_report: dict[str, Any], source: str) -> int | None:
    current: Any = input_report.get("workflow_inputs", {})
    for part in source.split("."):
        if not isinstance(current, dict) or part not in current:
            return None
        current = current[part]
    try:
        return int(current)
    except (TypeError, ValueError):
        return None


def validate_quickstart_route_definitions(root: Path, data: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if data.get("schema") != QUICKSTART_ROUTE_SCHEMA:
        errors.append(f"{QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH} schema mismatch")
    for flag in ["report_only", "mutations_performed", "real_email_sending_performed"]:
        value = data.get(flag)
        if flag == "report_only" and value is not True:
            errors.append(f"{QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH} must declare report_only=true")
        if flag != "report_only" and value is not False:
            errors.append(f"{QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH} must declare {flag}=false")

    seen_names: set[str] = set()
    for section_name in ["routes", "blocked_routes"]:
        routes = data.get(section_name)
        if not isinstance(routes, list) or not routes:
            errors.append(f"{QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH} missing {section_name}")
            continue
        for index, route in enumerate(routes, start=1):
            prefix = f"{QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH}:{section_name}[{index}]"
            if not isinstance(route, dict):
                errors.append(f"{prefix} is not an object")
                continue
            name = str(route.get("name", "")).strip()
            if not name:
                errors.append(f"{prefix} missing name")
            elif name in seen_names:
                errors.append(f"{prefix} duplicates route name: {name}")
            else:
                seen_names.add(name)

            command = route.get("command")
            if not isinstance(command, list) or not command or not all(isinstance(part, str) and part for part in command):
                errors.append(f"{prefix} command must be a non-empty string list")
            else:
                first = str(command[0]).lower()
                if first in {"cscript", "cscript.exe", "wscript", "wscript.exe", "scaffold.vbs", ".\\scaffold.vbs"}:
                    errors.append(f"{prefix} command must contain Scaffold.vbs arguments only")
                for token in command:
                    if token.startswith(("examples/", "fixtures/")) and not (root / token).exists():
                        errors.append(f"{prefix} path argument does not exist: {token}")

            if not isinstance(route.get("expected_exit_code"), int):
                errors.append(f"{prefix} expected_exit_code must be an integer")
            expected_fields = route.get("expected_fields", {})
            expected_output = route.get("expected_output_contains", [])
            if not isinstance(expected_fields, dict):
                errors.append(f"{prefix} expected_fields must be an object when present")
            elif expected_fields and not all(
                isinstance(key, str) and isinstance(value, str) for key, value in expected_fields.items()
            ):
                errors.append(f"{prefix} expected_fields values must be strings")
            if not isinstance(expected_output, list):
                errors.append(f"{prefix} expected_output_contains must be a list when present")
            elif expected_output and not all(isinstance(value, str) and value for value in expected_output):
                errors.append(f"{prefix} expected_output_contains values must be non-empty strings")
            if not expected_fields and not expected_output:
                errors.append(f"{prefix} must define expected_fields or expected_output_contains")

            count_field = route.get("expected_count_field")
            count_source = route.get("expected_count_source")
            if count_field is not None and not isinstance(count_field, str):
                errors.append(f"{prefix} expected_count_field must be a string when present")
            if count_source is not None and not isinstance(count_source, str):
                errors.append(f"{prefix} expected_count_source must be a string when present")
            if (count_field is None) != (count_source is None):
                errors.append(f"{prefix} count field and count source must be provided together")

    return errors


def validate_quickstart_fixture(root: Path) -> list[str]:
    path = root / QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH
    if not path.exists():
        return [f"missing quickstart route fixture: {QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH}"]
    try:
        data = load_quickstart_fixture(root)
    except json.JSONDecodeError as exc:
        return [f"{QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH} is not valid JSON: {exc}"]
    return validate_quickstart_route_definitions(root, data)


def route_definitions(fixture: dict[str, Any], section_name: str) -> list[dict[str, Any]]:
    routes = fixture.get(section_name, [])
    if not isinstance(routes, list):
        return []
    return [route for route in routes if isinstance(route, dict)]


def run_route_definitions(root: Path, routes: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    results: dict[str, dict[str, Any]] = {}
    for route in routes:
        name = str(route.get("name", "")).strip()
        if not name:
            continue
        results[name] = run_scaffold(root, [scaffold_cli_arg(token) for token in command_path_tokens(route)])
    return results


def validate_route_result(
    errors: list[str],
    route: dict[str, Any],
    result: dict[str, Any],
    input_report: dict[str, Any],
) -> None:
    name = str(route.get("name", "")).strip() or "unnamed_route"
    expected_exit_code = route.get("expected_exit_code")
    if result["exit_code"] != expected_exit_code:
        add_error(errors, name, f"route exited {result['exit_code']} instead of {expected_exit_code}")

    fields = result["fields"]
    expected_fields = route.get("expected_fields", {})
    if isinstance(expected_fields, dict):
        for key, expected in expected_fields.items():
            actual = fields.get(str(key), "")
            if actual != str(expected):
                add_error(errors, name, f"field {key} was {actual!r} instead of {expected!r}")

    expected_output = route.get("expected_output_contains", [])
    if isinstance(expected_output, list):
        output = result.get("output", "")
        for expected_text in expected_output:
            if str(expected_text) not in output:
                add_error(errors, name, f"output did not contain {expected_text!r}")

    count_field = route.get("expected_count_field")
    count_source = route.get("expected_count_source")
    if isinstance(count_field, str) and isinstance(count_source, str):
        expected = expected_count(input_report, count_source)
        actual = int_field(fields, count_field, default=-1)
        if expected is None:
            add_error(errors, name, f"unsupported expected_count_source: {count_source}")
        elif actual != expected:
            add_error(errors, name, f"{count_field} {actual} did not match expected {expected}")


def build_report(root: Path) -> dict[str, Any]:
    root = root.resolve()
    input_report = workflow_input_report.build_report(root)
    errors: list[str] = []
    fixture_errors = validate_quickstart_fixture(root)
    errors.extend(fixture_errors)
    errors.extend(str(error) for error in input_report["validation_errors"])

    fixture = load_quickstart_fixture(root) if not fixture_errors else {}
    public_route_defs = route_definitions(fixture, "routes")
    blocked_route_defs = route_definitions(fixture, "blocked_routes")
    routes = run_route_definitions(root, public_route_defs)
    blocked_routes = run_route_definitions(root, blocked_route_defs)

    for route in public_route_defs:
        name = str(route.get("name", "")).strip()
        if name in routes:
            validate_route_result(errors, route, routes[name], input_report)
    for route in blocked_route_defs:
        name = str(route.get("name", "")).strip()
        if name in blocked_routes:
            validate_route_result(errors, route, blocked_routes[name], input_report)

    return {
        "schema": SCHEMA,
        "report_only": True,
        "mutations_performed": False,
        "real_email_sending_performed": False,
        "quickstart_fixture": {
            "path": QUICKSTART_ROUTE_FIXTURE_RELATIVE_PATH,
            "schema": fixture.get("schema", "") if fixture else "",
        },
        "routes": routes,
        "blocked_routes": blocked_routes,
        "expected_counts": {
            "group_meetings_planned_count": expected_count(
                input_report, "group_meetings.first_route_planned_message_estimate"
            ),
            "induction_planned_count": expected_count(input_report, "induction.planned_message_estimate"),
            "watchdog_repair_actions_taken": 0,
        },
        "validation_errors": errors,
    }


def validate_default_routes(root: Path) -> list[str]:
    return list(build_report(root)["validation_errors"])


def print_text(report: dict[str, Any]) -> None:
    print("Command route report")
    print(f"schema: {report['schema']}")
    print(f"report_only: {str(report['report_only']).lower()}")
    print(f"quickstart_fixture: {report['quickstart_fixture']['path']}")
    for name in report["routes"]:
        route = report["routes"][name]
        fields = route["fields"]
        print(f"{name}_exit_code: {route['exit_code']}")
        print(f"{name}_state: {fields.get('state', '')}")
    for name in report["blocked_routes"]:
        route = report["blocked_routes"][name]
        fields = route["fields"]
        print(f"{name}_exit_code: {route['exit_code']}")
        print(f"{name}_state: {fields.get('state', '')}")
    if report["validation_errors"]:
        print("validation_errors:")
        for error in report["validation_errors"]:
            print(f"- {error}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate Scaffold.vbs dry-run/report-only command routes.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--format", choices=["text", "json"], default="text")
    parser.add_argument("--fail-if-invalid", action="store_true")
    args = parser.parse_args()

    report = build_report(args.root)
    if args.format == "json":
        print(json.dumps(report, indent=2, sort_keys=True))
    else:
        print_text(report)

    if args.fail_if_invalid and report["validation_errors"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
