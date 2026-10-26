# SPDX-License-Identifier: MIT
"""2989：生成自造酒店月表 JSON 模型和独立逐记录期望，不创建 XLSX。"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import date, timedelta
from decimal import Decimal
import json
from pathlib import Path

FIELDS = ["booking_id", "first_name", "last_name", "check_in", "check_out", "price", "notes"]
ALIASES = {
    "booking_id": ["Rezervasyon No", "Rez.No", "Booking No"],
    "first_name": ["Ad", "Adı", "İsim", "First Name"],
    "last_name": ["Soyad", "Soyadı", "Last Name"],
    "check_in": ["Giriş", "Giriş Tarihi", "Check-in"],
    "check_out": ["Çıkış", "Çıkış Tarihi", "Check-out"],
    "price": ["Fiyat", "Ücret", "Tutar", "Price"],
    "notes": ["Not", "Notlar", "Açıklama", "Notes"],
}
NAMES = [
    ("İpek", "Yılmaz", ["İPEK", "ipek", "  İpek  "], ["YILMAZ", "yılmaz", "  Yılmaz  "]),
    ("Işık", "Şahin", ["IŞIK", "ışık", " Işık "], ["ŞAHİN", "şahin", " Şahin "]),
    ("Özge", "Çelik", ["ÖZGE", "özge", " Özge "], ["ÇELİK", "çelik", " Çelik "]),
    ("Çağrı", "Öztürk", ["ÇAĞRI", "çağrı", " Çağrı "], ["ÖZTÜRK", "öztürk", " Öztürk "]),
    ("Şule", "Güneş", ["ŞULE", "şule", " Şule "], ["GÜNEŞ", "güneş", " Güneş "]),
    ("Derya", "Koç", ["DERYA", "derya", " Derya "], ["KOÇ", "koç", " Koç "]),
    ("Mehmet", "Akın", ["MEHMET", "mehmet", " Mehmet "], ["AKIN", "akın", " Akın "]),
    ("Ayşe", "İnce", ["AYŞE", "ayşe", " Ayşe "], ["İNCE", "ince", " İnce "]),
]
EXPECTED = dict(input_order_rows=1540, accepted_rows=1430, review_rows=80, duplicate_rows=30)


def date_cell(value, style):
    if style % 3 == 0:
        return {"$date": value.isoformat()}
    if style % 3 == 1 and value.day > 12:
        return value.strftime("%d.%m.%Y")
    return value.isoformat()


def price_cell(value, style):
    options = [value, f"USD {value}", f"{value} USD", f"${value}.00", f"USD {value},00"]
    return options[style % len(options)]


def build_model():
    """逐项写死生成规则；oracle 不调用任何生产引擎函数。"""
    sheets, layouts, order_rows, accepted, base = [], [], [], [], []
    for month_index in range(48):
        year, month = 2023 + month_index // 12, month_index % 12 + 1
        name = f"2989_{year:04d}-{month:02d}"
        if month_index == 47:
            rows = [["NOTLAR"], [], ["TOPLAM"]]
            sheets.append(dict(name=name + "_NOTLAR", rows=rows))
            layouts.extend(dict(source_sheet=name + "_NOTLAR", source_row=i + 1, status="metadata" if row else "blank") for i, row in enumerate(rows))
            continue
        fields = FIELDS[month_index % 7:] + FIELDS[:month_index % 7]
        if month_index % 2:
            fields = list(reversed(fields))
        header = [ALIASES[field][month_index % len(ALIASES[field])] for field in fields]
        rows = [[f"2989 自造酒店月表 {year:04d}-{month:02d}"] for _ in range(month_index % 3)]
        layouts.extend(dict(source_sheet=name, source_row=i + 1, status="metadata") for i in range(len(rows)))

        def nonorder(cells, status):
            rows.append(cells)
            layouts.append(dict(source_sheet=name, source_row=len(rows), status=status))

        def add_order(raw, identity, status, category, normalized=None):
            rows.append([raw[field] for field in fields])
            position = dict(source_sheet=name, source_row=len(rows))
            order_rows.append(dict(**position, identity=identity, booking_id=raw["booking_id"], status=status, category=category))
            layouts.append(dict(**position, status=status))
            if normalized is not None and status == "accepted":
                accepted.append(dict(normalized, **position))

        nonorder(header, "header")
        total_in_month = 32 if month_index < 43 else 31
        monthly_base = []
        for within in range(total_in_month):
            index = len(base)
            identity = f"2989_B{index + 1:04d}"
            first, last, raw_first, raw_last = NAMES[index % len(NAMES)]
            check_in = date(year, month, 15) + timedelta(days=index % 11)
            nights = 1 + index % (4 if 40 <= index < 50 else 5)
            check_out = check_in + timedelta(days=nights)
            price = nights * 40
            raw = dict(booking_id=identity if index != 1499 else "", first_name=raw_first[index % 3],
                       last_name=raw_last[(index // 3) % 3], check_in=date_cell(check_in, index),
                       check_out=date_cell(check_out, index + 1), price=price_cell(price, index), notes=f"2989 DEMO {index + 1:04d}")
            normalized = dict(booking_id=raw["booking_id"], first_name=first, last_name=last,
                              check_in=check_in.isoformat(), check_out=check_out.isoformat(), nights=nights,
                              price_usd=price, notes=raw["notes"])
            base.append(dict(raw=raw.copy(), normalized=normalized.copy(), identity=identity, month_index=month_index))
            status, category = "accepted", "valid"
            if 40 <= index < 50:
                status, category = "review", "same_id_conflict_original"
            elif 100 <= index < 160:
                status = "review"
                category = ["ambiguous_slash_date", "TRY_currency", "price_nights_mismatch", "missing_check_out", "missing_last_name", "negative_price"][(index - 100) // 10]
                if category == "ambiguous_slash_date":
                    raw["check_in"] = "01/02/2023"
                elif category == "TRY_currency":
                    raw["price"] = f"TRY {price}"
                elif category == "price_nights_mismatch":
                    raw["price"] = f"USD {price + 40}.00"
                elif category == "missing_check_out":
                    raw["check_out"] = ""
                elif category == "missing_last_name":
                    raw["last_name"] = ""
                else:
                    raw["price"] = f"USD -{price}.00"
            add_order(raw, identity, status, category, normalized)
            monthly_base.append(base[-1])
            if within == total_in_month // 2:
                nonorder([], "blank")
                nonorder(header.copy(), "repeated_header")
        if month_index == 0:
            for index, item in enumerate(base[:30]):
                raw = item["raw"].copy()
                norm = item["normalized"]
                first, last, raw_first, raw_last = NAMES[index % len(NAMES)]
                raw.update(first_name=raw_first[(index + 1) % 3], last_name=raw_last[(index + 1) % 3],
                           check_in={"$date": norm["check_in"]}, check_out=norm["check_out"],
                           price=price_cell(norm["price_usd"], index + 2))
                add_order(raw, item["identity"], "duplicate", "same_id_equivalent_copy")
        if month_index == 1:
            for item in base[40:50]:
                raw = item["raw"].copy()
                norm = item["normalized"]
                raw["check_out"] = {"$date": (date.fromisoformat(norm["check_out"]) + timedelta(days=1)).isoformat()}
                raw["price"] = f"USD {norm['price_usd'] + 40},00"
                add_order(raw, item["identity"], "review", "same_id_conflict_added")
        nonorder(["TOPLAM"] + [""] * (len(fields) - 1), "total")
        sheets.append(dict(name=name, rows=rows))
    assert len(base) == 1500 and len(order_rows) == 1540 and len(accepted) == 1430 and len(sheets) == 48
    counts = Counter(row["status"] for row in order_rows)
    assert counts == Counter(accepted=1430, review=80, duplicate=30)
    totals = {}
    for row in accepted:
        for period in (row["check_in"][:7], row["check_in"][:4]):
            summary = totals.setdefault(period, dict(period=period, bookings=0, revenue_usd=0))
            summary["bookings"] += 1
            summary["revenue_usd"] += row["price_usd"]
    oracle = dict(schema_version=1, synthetic=True, expected_counts=EXPECTED.copy(),
                  identity_policy="Stable synthetic B-number; blank booking_id sample retains its oracle identity in notes only",
                  base_bookings=1500, source_sheets=48, monthly_business_sheets=47,
                  physical_order_rows=order_rows, expected_records=accepted, expected_summary=[totals[k] for k in sorted(totals)],
                  all_source_rows=layouts, total_revenue_usd=sum(row["price_usd"] for row in accepted),
                  generation_rules=dict(months="2023-01 through 2026-12; final NOTLAR has no bookings", first_43_months=32,
                                        next_4_months=31, conflict_base_zero_based_indices=[40, 49], invalid_base_zero_based_indices=[100, 159],
                                        no_id_unique_identity="2989_B1500", nightly_usd=40,
                                        summary_policy="Entire confirmed price assigned to check-in month/year; no split by nights"))
    return dict(sheets=sheets, oracle=oracle)


def generate(output):
    output = Path(output)
    if output.exists() or not output.parent.is_dir():
        raise ValueError("模型目标须为已有父目录下的新文件，不覆盖")
    model = build_model()
    with output.open("x", encoding="utf-8") as stream:
        json.dump(model, stream, ensure_ascii=False, indent=2)
    return dict(source_sheets=len(model["sheets"]), **model["oracle"]["expected_counts"],
                total_revenue_usd=model["oracle"]["total_revenue_usd"])


def verify_model(actual, fixture=None):
    """按源坐标逐项核验，不依赖生产解析、规范化或归集函数。"""
    fixture = fixture or build_model()
    oracle = fixture["oracle"]
    counts = actual["report"]["counts"]
    for key, expected in [("accepted", 1430), ("review", 80), ("duplicate", 30), ("booking_rows", 1540), ("sheets", 48)]:
        if counts[key] != expected:
            raise AssertionError(f"独立计数不符：{key}: {counts[key]} != {expected}")
    ledger = {(row["source_sheet"], row["source_row"]): row for row in actual["ledger"]}
    positions = {(sheet["name"], index + 1) for sheet in fixture["sheets"] for index in range(len(sheet["rows"]))}
    if len(actual["ledger"]) != len(ledger) or not positions <= set(ledger):
        raise AssertionError("来源记录出现重复坐标或缺少模型原始行")
    extras = set(ledger) - positions
    # 独立SDK渲染为已知NOTLAR页补了固定5个样式空行，不能容许任意额外来源。
    note_sheet = fixture["sheets"][-1]["name"]
    allowed_style_rows = {(note_sheet, row) for row in range(4, 9)}
    if extras and extras != allowed_style_rows:
        raise AssertionError("出现已知NOTLAR第4至8行以外的额外来源")
    for position in extras:
        row = ledger[position]
        raw = json.loads(row["raw_json"])
        if row["status"] != "blank" or not isinstance(raw, list) or any(value not in (None, "") for value in raw):
            raise AssertionError("额外样式行含值或不是blank状态")
    for row in oracle["physical_order_rows"]:
        if ledger[row["source_sheet"], row["source_row"]]["status"] != row["status"]:
            raise AssertionError(f"订单来源状态不符：{row['identity']}")
    accepted = {(row["source_sheet"], row["source_row"]): row for row in actual["accepted"]}
    expected = {(row["source_sheet"], row["source_row"]): row for row in oracle["expected_records"]}
    if len(actual["accepted"]) != 1430 or set(accepted) != set(expected):
        raise AssertionError("有效预订来源集合不符")
    for position, row in expected.items():
        for field, value in row.items():
            answer = accepted[position].get(field)
            mismatch = Decimal(str(answer)) != Decimal(str(value)) if field == "price_usd" else answer != value
            if mismatch:
                raise AssertionError(f"有效字段不符：{position} {field}: {answer!r} != {value!r}")
    summaries = {row["period"]: row for row in actual["summary"]}
    if len(summaries) != len(actual["summary"]):
        raise AssertionError("期间统计出现重复period")
    expected_summaries = {row["period"]: row for row in oracle["expected_summary"]}
    for period, expected_row in expected_summaries.items():
        row = summaries.get(period)
        if row is None or row["bookings"] != expected_row["bookings"] or Decimal(str(row["revenue_usd"])) != expected_row["revenue_usd"]:
            raise AssertionError(f"月/年统计不符：{period}")
    for period in set(summaries) - set(expected_summaries):
        if summaries[period]["bookings"] != 0 or Decimal(str(summaries[period]["revenue_usd"])) != 0:
            raise AssertionError("不存在有效订单的补全期间不是零")
    if sum(row["status"] == "review" for row in actual["ledger"]) != 80 or actual["review"] != [row for row in actual["ledger"] if row["status"] == "review"]:
        raise AssertionError("复核集合与来源状态不一致")
    return dict(accepted_rows=1430, review_rows=80, duplicate_rows=30, input_order_rows=1540,
                verified_accepted_records=1430, verified_physical_rows=len(ledger), verified_summary_periods=len(summaries),
                expected_model_physical_rows=len(positions), style_only_extra_blank_rows=len(extras),
                total_revenue_usd=oracle["total_revenue_usd"], all_source_statuses_match=True, all_accepted_fields_match=True,
                all_month_year_totals_match=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    print(json.dumps(generate(parser.parse_args().output), ensure_ascii=False, indent=2))
