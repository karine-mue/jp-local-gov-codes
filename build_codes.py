import argparse
import csv
import datetime
import json
import os
import re
import tempfile
import unicodedata
import urllib.request
from pathlib import Path
from typing import Any

import openpyxl

RECORD_KEYS = [
    "code",
    "code5",
    "prefecture_code",
    "prefecture_name",
    "prefecture_kana",
    "prefecture_kana_halfwidth",
    "municipality_code",
    "municipality_name",
    "municipality_kana",
    "municipality_kana_halfwidth",
    "type",
]

EVENT_KEYS = [
    "group_id",
    "date",
    "prefecture_name",
    "change_type",
    "code_before",
    "name_before",
    "code_after",
    "name_after",
    "reason",
]

LIFESPAN_KEYS = [
    "code",
    "name",
    "prefecture_name",
    "valid_from",
    "valid_to",
]

_ERA_BASES = {
    "R": 2018,
    "H": 1988,
    "S": 1925,
}
_ERA_RE = re.compile(r"^([RHS])(\d+)[.．](\d+)[.．](\d+)$")

def check_digit(first5: str) -> int:
    if not isinstance(first5, str) or len(first5) != 5 or not first5.isdigit():
        raise ValueError(f"Invalid first5: {first5!r}")
    weights = [6, 5, 4, 3, 2]
    total = 0
    for i, ch in enumerate(first5):
        total += int(ch) * weights[i]
    if total < 11:
        return 11 - total
    return (11 - (total % 11)) % 10

def parse_date_cell(value: Any) -> datetime.date | None:
    if value is None or value == "〃":
        return None
    if isinstance(value, bool):
        raise ValueError(f"Invalid date value: {value!r}")
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)):
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(value))
    if isinstance(value, str):
        stripped = value.strip()
        if stripped == "〃":
            return None
        m = _ERA_RE.match(stripped)
        if m:
            era = m.group(1)
            year = _ERA_BASES[era] + int(m.group(2))
            month = int(m.group(3))
            day = int(m.group(4))
            return datetime.date(year, month, day)
    raise ValueError(f"Invalid date value: {value!r}")

def to_fullwidth_kana(s: str | None) -> str | None:
    if s is None:
        return None
    return unicodedata.normalize("NFKC", s)

def classify_type(prefecture_name: str | None, municipality_name: str | None, in_designated_sheet: bool) -> str:
    if municipality_name is None:
        return "prefecture"
    if in_designated_sheet:
        return "ward" if municipality_name.endswith("区") else "designated_city"
    if prefecture_name == "東京都" and municipality_name.endswith("区"):
        return "special_ward"
    if municipality_name.endswith("市"):
        return "city"
    if municipality_name.endswith("町"):
        return "town"
    if municipality_name.endswith("村"):
        return "village"
    raise ValueError(f"Cannot classify municipality name: {municipality_name!r}")

def _clean_cell_str(value: Any) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    return s if s else None

def _extract_sheet_rows(ws) -> list[tuple[str, str | None, str | None, str | None, str | None]]:
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or len(row) < 5:
            continue
        code_raw = row[0]
        if code_raw is None:
            continue
        code = str(code_raw).strip()
        if len(code) != 6 or not code.isdigit():
            continue
        pref_name = _clean_cell_str(row[1])
        muni_name = _clean_cell_str(row[2])
        pref_kana_hw = _clean_cell_str(row[3])
        muni_kana_hw = _clean_cell_str(row[4])
        rows.append((code, pref_name, muni_name, pref_kana_hw, muni_kana_hw))
    return rows

def build_records(path: str | os.PathLike) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        designated_sheet = None
        main_sheet = None
        for ws in wb.worksheets:
            if "政令指定都市" in ws.title:
                designated_sheet = ws
            else:
                main_sheet = ws
        if designated_sheet is None or main_sheet is None:
            raise ValueError("Expected exactly two sheets: a main sheet and a designated-city sheet")
        merged: dict[str, dict[str, Any]] = {}
        for row in _extract_sheet_rows(main_sheet):
            code, pref_name, muni_name, pref_kana_hw, muni_kana_hw = row
            merged[code] = {
                "prefecture_name": pref_name,
                "municipality_name": muni_name,
                "prefecture_kana_halfwidth": pref_kana_hw,
                "municipality_kana_halfwidth": muni_kana_hw,
                "in_designated": False,
            }
        for row in _extract_sheet_rows(designated_sheet):
            code, pref_name, muni_name, pref_kana_hw, muni_kana_hw = row
            if code in merged:
                merged[code]["in_designated"] = True
            else:
                merged[code] = {
                    "prefecture_name": pref_name,
                    "municipality_name": muni_name,
                    "prefecture_kana_halfwidth": pref_kana_hw,
                    "municipality_kana_halfwidth": muni_kana_hw,
                    "in_designated": True,
                }
        records: list[dict] = []
        for code in sorted(merged.keys()):
            item = merged[code]
            pref_name = item["prefecture_name"]
            muni_name = item["municipality_name"]
            pref_kana_hw = item["prefecture_kana_halfwidth"]
            muni_kana_hw = item["municipality_kana_halfwidth"] if muni_name is not None else None
            in_designated = item["in_designated"]
            is_pref = muni_name is None
            rec = {
                "code": code,
                "code5": code[:5],
                "prefecture_code": code[:2],
                "prefecture_name": pref_name,
                "prefecture_kana": to_fullwidth_kana(pref_kana_hw),
                "prefecture_kana_halfwidth": pref_kana_hw,
                "municipality_code": None if is_pref else code[2:5],
                "municipality_name": None if is_pref else muni_name,
                "municipality_kana": None if is_pref else to_fullwidth_kana(muni_kana_hw),
                "municipality_kana_halfwidth": None if is_pref else muni_kana_hw,
                "type": classify_type(pref_name, muni_name, in_designated),
            }
            records.append(rec)
        return records
    finally:
        wb.close()

def validate_records(records: list[dict]) -> None:
    failing = []
    for rec in records:
        code = rec["code"]
        if int(code[5]) != check_digit(code[:5]):
            failing.append(code)
    if failing:
        raise ValueError(f"Invalid check digits for codes: {', '.join(failing)}")

_UNSET = object()

def _resolve_before_field(raw: str | None, prev_val, field_name: str) -> str | None:
    if raw is None:
        return None
    if raw == "〃":
        if prev_val is _UNSET:
            raise ValueError(f"First row of block uses '〃' for {field_name}")
        return prev_val
    return raw

def _resolve_after_field(raw: str | None, own_before: str | None, prev_after, field_name: str) -> str | None:
    if raw is None:
        return None
    if raw == "同左":
        return own_before
    if raw == "削除":
        return None
    if raw == "〃":
        if prev_after is _UNSET:
            raise ValueError(f"First row of block uses '〃' for {field_name}")
        return prev_after
    return raw

def _is_data_code_cell(value: str | None) -> bool:
    if value is None:
        return False
    if value in ("〃", "同左", "削除"):
        return True
    return len(value) == 6 and value.isdigit()

def build_history_events(path: str | os.PathLike) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb["改正一覧表"]
        events: list[dict] = []
        blocks: list[dict[str, Any]] = []
        current_block: dict[str, Any] | None = None

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None or len(row) < 14:
                continue
            pref_cell = row[4]
            code_before_cell = row[5]
            code_after_cell = row[9]
            if not _is_data_code_cell(_clean_cell_str(code_before_cell)) and not _is_data_code_cell(_clean_cell_str(code_after_cell)):
                continue
            pref_str = _clean_cell_str(pref_cell)
            if pref_str is not None:
                current_block = {
                    "group_id": len(blocks) + 1,
                    "prefecture_name": pref_str,
                    "rows": [],
                    "dates": set(),
                }
                blocks.append(current_block)
            if current_block is None:
                raise ValueError("First data row has no prefecture name")
            r = {
                "prefecture_name": pref_str,
                "code_before": _clean_cell_str(code_before_cell),
                "name_before": _clean_cell_str(row[6]),
                "kind": _clean_cell_str(row[8]),
                "code_after": _clean_cell_str(code_after_cell),
                "name_after": _clean_cell_str(row[10]),
                "reason": row[12] if row[12] is None or isinstance(row[12], str) else str(row[12]),
                "date_cell": row[13],
            }
            current_block["rows"].append(r)
            parsed = parse_date_cell(r["date_cell"])
            if parsed is not None:
                current_block["dates"].add(parsed)

        for block in blocks:
            dates = block["dates"]
            if len(dates) != 1:
                raise ValueError(
                    f"Block {block['group_id']} must have exactly one distinct date, got {len(dates)}: {sorted(dates)}"
                )
            block_date = next(iter(dates))
            block_date_iso = block_date.isoformat()
            prev_kind: str | None = None
            prev_cb = _UNSET
            prev_nb = _UNSET
            prev_ca = _UNSET
            prev_na = _UNSET
            for r in block["rows"]:
                raw_kind = r.get("kind")
                if raw_kind == "新設":
                    change_type = "created"
                elif raw_kind == "欠番":
                    change_type = "retired"
                elif raw_kind == "〃":
                    if prev_kind is None:
                        raise ValueError(
                            f"Block {block['group_id']}: '〃' kind on first row of block"
                        )
                    change_type = prev_kind
                else:
                    change_type = "continued"
                code_before = _resolve_before_field(r["code_before"], prev_cb, "code_before")
                name_before = _resolve_before_field(r["name_before"], prev_nb, "name_before")
                code_after = _resolve_after_field(r["code_after"], code_before, prev_ca, "code_after")
                if r["name_after"] == "〃" and r["code_after"] == "削除":
                    name_after = None
                elif r["name_after"] == "〃" and r["code_after"] == "同左":
                    name_after = name_before
                else:
                    name_after = _resolve_after_field(r["name_after"], name_before, prev_na, "name_after")
                prev_kind = change_type
                prev_cb = code_before
                prev_nb = name_before
                prev_ca = code_after
                prev_na = name_after
                events.append(
                    {
                        "group_id": block["group_id"],
                        "date": block_date_iso,
                        "prefecture_name": block["prefecture_name"],
                        "change_type": change_type,
                        "code_before": code_before,
                        "name_before": name_before,
                        "code_after": code_after,
                        "name_after": name_after,
                        "reason": r["reason"],
                    }
                )
        return events
    finally:
        wb.close()

def derive_lifespans(records: list[dict], events: list[dict]) -> list[dict]:
    lifespans: dict[str, dict] = {}
    for rec in records:
        code = rec["code"]
        name = rec["municipality_name"] if rec["municipality_name"] is not None else rec["prefecture_name"]
        lifespans[code] = {
            "code": code,
            "name": name,
            "prefecture_name": rec["prefecture_name"],
            "valid_from": None,
            "valid_to": None,
        }
    for ev in events:
        ctype = ev["change_type"]
        ev_date = ev["date"]
        pref_name = ev["prefecture_name"]
        if ctype == "created" and ev.get("code_after") is not None:
            code = ev["code_after"]
            if code not in lifespans:
                lifespans[code] = {
                    "code": code,
                    "name": ev.get("name_after"),
                    "prefecture_name": pref_name,
                    "valid_from": None,
                    "valid_to": None,
                }
            if lifespans[code]["valid_from"] is None:
                lifespans[code]["valid_from"] = ev_date
        elif ctype == "retired" and ev.get("code_before") is not None:
            code = ev["code_before"]
            if code not in lifespans:
                lifespans[code] = {
                    "code": code,
                    "name": ev.get("name_before"),
                    "prefecture_name": pref_name,
                    "valid_from": None,
                    "valid_to": None,
                }
            lifespans[code]["valid_to"] = ev_date
    return [lifespans[code] for code in sorted(lifespans.keys())]

def _write_csv(path: Path, keys: list[str], rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=keys, lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: ("" if row.get(k) is None else row.get(k)) for k in keys})

def _write_json(path: Path, data: Any) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _build_schema() -> dict:
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "required": RECORD_KEYS,
        "properties": {
            "code": {"type": "string", "pattern": r"^\d{6}$"},
            "code5": {"type": "string", "pattern": r"^\d{5}$"},
            "prefecture_code": {"type": "string", "pattern": r"^\d{2}$"},
            "prefecture_name": {"type": "string"},
            "prefecture_kana": {"type": "string"},
            "prefecture_kana_halfwidth": {"type": "string"},
            "municipality_code": {"type": ["string", "null"], "pattern": r"^\d{3}$"},
            "municipality_name": {"type": ["string", "null"]},
            "municipality_kana": {"type": ["string", "null"]},
            "municipality_kana_halfwidth": {"type": ["string", "null"]},
            "type": {"type": "string", "enum": ["prefecture", "designated_city", "ward", "special_ward", "city", "town", "village"]},
        },
    }

def _resolve_input_path(source: str, temp_files: list[str]) -> str:
    if source.startswith("http://") or source.startswith("https://"):
        suffix = Path(source).suffix or ".xlsx"
        fd, tmp_path = tempfile.mkstemp(suffix=suffix)
        os.close(fd)
        temp_files.append(tmp_path)
        urllib.request.urlretrieve(source, tmp_path)
        return tmp_path
    return source

def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Build local government code datasets.")
    parser.add_argument(
        "--current",
        default="https://www.soumu.go.jp/main_content/000925835.xlsx",
        help="Path or URL to current local government codes workbook",
    )
    parser.add_argument(
        "--history",
        default="https://www.soumu.go.jp/main_content/000875488.xlsx",
        help="Path or URL to revision history workbook",
    )
    parser.add_argument(
        "--out-dir",
        default=".",
        help="Output directory",
    )
    args = parser.parse_args(argv)
    temp_files: list[str] = []
    try:
        current_path = _resolve_input_path(args.current, temp_files)
        history_path = _resolve_input_path(args.history, temp_files)
        records = build_records(current_path)
        validate_records(records)
        events = build_history_events(history_path)
        lifespans = derive_lifespans(records, events)
        out_dir = Path(args.out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        _write_csv(out_dir / "local_gov_codes.csv", RECORD_KEYS, records)
        _write_json(out_dir / "local_gov_codes.json", records)
        _write_csv(out_dir / "history_events.csv", EVENT_KEYS, events)
        _write_json(out_dir / "history_events.json", events)
        _write_csv(out_dir / "code_lifespans.csv", LIFESPAN_KEYS, lifespans)
        _write_json(out_dir / "code_lifespans.json", lifespans)
        _write_json(out_dir / "schema.json", _build_schema())
        print(f"records={len(records)} events={len(events)} lifespans={len(lifespans)}")
    finally:
        for temp_file in temp_files:
            try:
                os.remove(temp_file)
            except OSError:
                pass

if __name__ == "__main__":
    main()
