from __future__ import annotations

import argparse
from typing import List, Optional

from ..outlook.client import OutlookClient
from ..name_resolution.directory_lists import build_directory_search_list


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Probe available address lists and entry counts")
    p.add_argument("--max", type=int, default=5, help="Max entries to sample per list")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    client = OutlookClient.connect()
    ns = client.namespace

    lists = build_directory_search_list(ns)
    print("Address lists:")
    for name in lists:
        try:
            al = ns.AddressLists(name)
        except Exception:
            print(f"- {name}: (unavailable)")
            continue
        try:
            entries = al.AddressEntries
            count = int(entries.Count)
        except Exception:
            count = -1
        print(f"- {name}: {count}")
        if count > 0:
            sample = 0
            for e in al.AddressEntries:
                print(f"    * {getattr(e, 'Name', '')} <{getattr(e, 'Address', '')}>")
                sample += 1
                if sample >= args.max:
                    break

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
