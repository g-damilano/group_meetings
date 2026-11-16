from __future__ import annotations

import argparse
from typing import List, Optional

from ..outlook.client import OutlookClient
from ..name_resolution.directory_matching import collect_address_entries_by_fields, describe_address_entry


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Lookup matching entries in the Global Address List")
    p.add_argument("query", help="Search value (exact match on Name/DisplayName)")
    p.add_argument("--max", type=int, default=20, help="Max matches to print")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    client = OutlookClient.connect()
    ns = client.namespace

    try:
        gal = ns.AddressLists("Global Address List")
    except Exception:
        print("Global Address List not available.")
        return 2

    matches = []
    for fields in [["Name"], ["DisplayName"], ["FullName"]]:
        matches.extend(collect_address_entries_by_fields(gal, fields, args.query))

    if not matches:
        print("No matches.")
        return 0

    for e in matches[: args.max]:
        print(describe_address_entry(e))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
