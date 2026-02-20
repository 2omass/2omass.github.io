#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
import shutil
import time
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

DIVISION_RULES: dict[str, list[str]] = {
    "GovernmentRelations": ["الجوازات", "مقيم", "تأشيرة", "تاشيرة", "وزارة", "منصة", "قوى", "العمل", "government", "visa", "iqama", "gosi", "muqeem"],
    "Finance": ["مالية", "مالي", "finance", "invoice", "فاتورة", "سداد", "تحويل", "budget", "ميزانية", "expense", "cost", "tax", "vat", "ضريبة"],
    "HumanResources": ["hr", "موارد", "الموظف", "employee", "توظيف", "عقد", "راتب", "إجازة", "اجازة"],
    "Legal": ["legal", "قانوني", "قضية", "litigation", "عقوبة", "disciplinary"],
    "Operations": ["operations", "تشغيل", "sop", "إجراء", "process", "workflow"],
    "GeneralAdmin": ["admin", "إداري", "مراسلة", "letter", "خطاب"],
}

CATEGORY_RULES: dict[str, list[str]] = {
    "Recruitment": ["cv", "resume", "سيرة", "توظيف", "مرشح", "interview", "مقابلة", "offer"],
    "Contracts": ["contract", "عقد", "اتفاقية", "nda", "annex"],
    "Payroll": ["salary", "payroll", "راتب", "بدل", "خصم", "مسير", "كشف رواتب"],
    "Performance": ["kpi", "performance", "تقييم", "أداء", "objective", "هدف"],
    "Attendance": ["attendance", "حضور", "انصراف", "leave", "اجازة", "دوام", "timesheet"],
    "Training": ["training", "تدريب", "دورة", "workshop", "شهادة"],
    "Compliance": ["policy", "سياسة", "compliance", "امتثال", "تحقيق", "disciplinary"],
    "GovernmentDocuments": ["iqama", "اقامة", "جواز", "passport", "visa", "تأشيرة", "work permit", "رخصة"],
    "FinanceDocuments": ["invoice", "فاتورة", "payment", "دفعة", "receipt", "إيصال", "tax", "ضريبة", "balance", "ميزان"],
}

TEXT_EXTENSIONS = {".txt", ".csv", ".md", ".json", ".log", ".xml"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
DATE_PATTERNS = [
    re.compile(r"(20\d{2})[-/.](1[0-2]|0?[1-9])[-/.]([12]\d|3[01]|0?[1-9])"),
    re.compile(r"([12]\d|3[01]|0?[1-9])[-/.](1[0-2]|0?[1-9])[-/.](20\d{2})"),
]


@dataclass
class FilePlan:
    original_path: Path
    division: str
    category: str
    employee: str
    date: str
    new_name: str
    destination: Path


def sanitize(text: str) -> str:
    cleaned = unicodedata.normalize("NFKC", text or "").strip()
    cleaned = re.sub(r"[\s/\\]+", "_", cleaned)
    cleaned = re.sub(r"[^\w\u0600-\u06FF.-]", "", cleaned)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    return cleaned or "Unknown"


def unzip_xml_text(path: Path, members_pattern: str, max_chars: int = 12000) -> str:
    try:
        with zipfile.ZipFile(path) as zf:
            parts: list[str] = []
            for name in zf.namelist():
                if re.search(members_pattern, name):
                    xml = zf.read(name).decode("utf-8", errors="ignore")
                    parts.append(re.sub(r"<[^>]+>", " ", xml))
            return " ".join(parts)[:max_chars]
    except Exception:
        return ""


def read_text_content(path: Path, max_chars: int = 12000) -> str:
    ext = path.suffix.lower()
    if ext in TEXT_EXTENSIONS:
        try:
            return path.read_text(encoding="utf-8", errors="ignore")[:max_chars]
        except OSError:
            return ""
    if ext == ".docx":
        return unzip_xml_text(path, r"^word/.*\.xml$", max_chars=max_chars)
    if ext in EXCEL_EXTENSIONS:
        return unzip_xml_text(path, r"^xl/.*\.xml$", max_chars=max_chars)
    return ""


def detect_by_rules(text: str, rules: dict[str, list[str]], default: str) -> str:
    lowered = text.lower()
    best_label = default
    best_score = 0
    for label, keywords in rules.items():
        score = sum(1 for k in keywords if k.lower() in lowered)
        if score > best_score:
            best_score = score
            best_label = label
    return best_label


def detect_division(text: str) -> str:
    return detect_by_rules(text, DIVISION_RULES, "GeneralDivision")


def detect_category(text: str) -> str:
    return detect_by_rules(text, CATEGORY_RULES, "General")


def detect_date(text: str) -> str:
    for patt in DATE_PATTERNS:
        m = patt.search(text)
        if not m:
            continue
        parts = [p.zfill(2) for p in m.groups()]
        if len(parts[0]) == 4:
            yyyy, mm, dd = parts
        else:
            dd, mm, yyyy = parts
        try:
            return dt.date(int(yyyy), int(mm), int(dd)).isoformat()
        except ValueError:
            continue
    return "NoDate"


def cleanup_employee_name(name: str) -> str:
    blocked = {"date", "تاريخ", "employee", "name", "الموظف", "الاسم"}
    tokens = [t for t in re.split(r"[_\s]+", sanitize(name).lower()) if t and t not in blocked]
    if not tokens:
        return "UnknownEmployee"
    return sanitize("_".join(tokens[:3]))


def detect_employee(text: str, filename: str) -> str:
    patterns = [
        r"(?:employee|name|الموظف|الاسم)\s*[:\-]?\s*([A-Za-z\u0600-\u06FF]{3,}(?:[ ]+[A-Za-z\u0600-\u06FF]{2,}){0,2})",
    ]
    for patt in patterns:
        m = re.search(patt, text, flags=re.IGNORECASE)
        if m:
            return cleanup_employee_name(m.group(1))
    stem_tokens = [t for t in re.split(r"[_\-\s]+", Path(filename).stem) if t]
    if stem_tokens:
        return cleanup_employee_name("_".join(stem_tokens[:2]))
    return "UnknownEmployee"


def generate_name(department: str, division: str, category: str, employee: str, date: str, seq: int, ext: str) -> str:
    return (
        f"{sanitize(department).upper()}_"
        f"{sanitize(division).upper()}_"
        f"{sanitize(category).upper()}_"
        f"{sanitize(employee)}_"
        f"{sanitize(date)}_"
        f"{seq:04d}{ext}"
    )


def iter_files(root: Path, excluded_roots: Sequence[Path] | None = None) -> Iterable[Path]:
    excluded = [p.resolve() for p in (excluded_roots or [])]
    for path in root.rglob("*"):
        if not path.is_file() or path.name == ".DS_Store":
            continue
        resolved = path.resolve()
        if any(ex in resolved.parents or resolved == ex for ex in excluded):
            continue
        yield path


def build_plan(source_dir: Path, output_dir: Path, department: str) -> list[FilePlan]:
    plans: list[FilePlan] = []
    files = sorted(iter_files(source_dir, excluded_roots=[output_dir]))
    for i, path in enumerate(files, start=1):
        combined_text = (path.name + " " + read_text_content(path)).lower()
        division = detect_division(combined_text)
        category = detect_category(combined_text)
        date = detect_date(combined_text)
        employee = detect_employee(combined_text, path.name)
        new_name = generate_name(department, division, category, employee, date, i, path.suffix.lower() or ".file")
        destination = output_dir / division / category / employee / new_name
        plans.append(FilePlan(path, division, category, employee, date, new_name, destination))
    return plans


def apply_plan(plans: list[FilePlan], dry_run: bool) -> None:
    seen: set[Path] = set()
    for plan in plans:
        target = plan.destination
        while target in seen or target.exists():
            target = target.with_name(f"{target.stem}_DUP{target.suffix}")
        seen.add(target)
        if dry_run:
            print(f"[DRY-RUN] {plan.original_path} -> {target}")
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(plan.original_path), str(target))
        print(f"[MOVED] {plan.original_path} -> {target}")


def export_report(plans: list[FilePlan], report_path: Path) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["original_path", "division", "category", "employee", "date", "new_name", "destination"])
        for p in plans:
            writer.writerow([str(p.original_path), p.division, p.category, p.employee, p.date, p.new_name, str(p.destination)])


def process_once(source: Path, output: Path, report: Path, department: str, dry_run: bool) -> dict[str, object]:
    plans = build_plan(source, output, department)
    if not plans:
        return {
            "files_processed": 0,
            "divisions": [],
            "categories": [],
            "report": str(report),
            "mode": "dry-run" if dry_run else "apply",
        }

    export_report(plans, report)
    apply_plan(plans, dry_run=dry_run)
    return {
        "files_processed": len(plans),
        "divisions": sorted({p.division for p in plans}),
        "categories": sorted({p.category for p in plans}),
        "report": str(report),
        "mode": "dry-run" if dry_run else "apply",
    }


def run_loop(source: Path, output: Path, report: Path, department: str, dry_run: bool, interval_s: int, max_cycles: int) -> None:
    cycle = 0
    while True:
        cycle += 1
        summary = process_once(source, output, report, department, dry_run)
        summary["cycle"] = cycle
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        if max_cycles and cycle >= max_cycles:
            break
        time.sleep(max(5, interval_s))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Professional HR file organizer: analyzes content (including Excel/XML-inside), "
            "classifies by division/category, renames, and moves files directly on your computer."
        )
    )
    parser.add_argument("source", type=Path, help="Source folder that contains files")
    parser.add_argument("--department", default="HR", help="Department/company prefix in the filename")
    parser.add_argument("--output", type=Path, default=None, help="Output root folder (default: <source>/organized)")
    parser.add_argument("--dry-run", action="store_true", help="Preview without moving files")
    parser.add_argument("--report", type=Path, default=None, help="CSV report path")
    parser.add_argument("--watch", action="store_true", help="Run continuously and process new files periodically")
    parser.add_argument("--interval", type=int, default=30, help="Watch interval in seconds (default: 30)")
    parser.add_argument("--max-cycles", type=int, default=0, help="For testing watch mode only: stop after N cycles (0 = infinite)")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source = args.source.resolve()
    if not source.exists() or not source.is_dir():
        raise SystemExit("Source folder not found or is not a directory")

    output = args.output.resolve() if args.output else source / "organized"
    report = args.report.resolve() if args.report else output / "hr_organizer_report.csv"

    if args.watch:
        run_loop(source, output, report, args.department, args.dry_run, args.interval, args.max_cycles)
        return

    summary = process_once(source, output, report, args.department, args.dry_run)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
