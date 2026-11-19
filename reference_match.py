"""Conservative matching of PDF timetable entries against a weekly XLSX reference."""
from __future__ import annotations

import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from xlsx_parser import parse_xlsx

SUBGROUP_RE = re.compile(r"\(\s*([12])\s*п\s*/\s*г\s*\)", re.I)
OPEN_SUBGROUP_RE = re.compile(r"\(\s*[12]\s*(?:п\s*/\s*г?)?\s*$", re.I)


def _norm(value: Any) -> str:
    text = str(value or "").casefold().replace("ё", "е")
    text = SUBGROUP_RE.sub("", text)
    text = re.sub(r"\s*(?:с\s*/\s*р|самостоятельная\s+работа)\s*$", "", text, flags=re.I)
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    return " ".join(text.split())


def _display_subject(value: Any) -> str:
    text = str(value or "").strip()
    text = SUBGROUP_RE.sub("", text)
    text = re.sub(r"\s*(?:с\s*/\s*р|самостоятельная\s+работа)\s*$", "", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip(" ()")


def _subgroup(value: Any) -> int | None:
    match = SUBGROUP_RE.search(str(value or ""))
    return int(match.group(1)) if match else None


def _normalize_room(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").casefold())


def _looks_like_teacher_name(value: str) -> bool:
    # A few XLSX layouts put the teacher row under the subgroup-2 pair marker;
    # the legacy timetable parser can otherwise misread that surname as a subject.
    return bool(re.fullmatch(r"[А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.[А-ЯЁ]\.?", value.strip()))


def _normalize_teacher(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").casefold().replace("ё", "е")).strip()


def load_reference(path: str | Path) -> dict[str, Any]:
    """Parse the weekly XLSX into a conservative reference catalogue and slot map."""
    path = Path(path).expanduser().resolve()
    parsed = parse_xlsx(path)
    records: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    slots: dict[tuple[str, str, int], list[dict[str, Any]]] = defaultdict(list)
    subjects: set[str] = set()
    teachers: set[str] = set()
    rooms: set[str] = set()

    for group, group_data in parsed.get("groups", {}).items():
        for date_text, lessons in group_data.get("days", {}).items():
            for lesson in lessons:
                subject_raw = str(lesson.get("subject") or "").strip()
                subject = _display_subject(subject_raw)
                if not subject or _looks_like_teacher_name(subject):
                    continue
                subgroup = _subgroup(subject_raw)
                try:
                    pair = int(lesson.get("pair") or 0)
                except (TypeError, ValueError):
                    pair = 0
                teacher = str(lesson.get("teacher") or "").strip() or None
                room = str(lesson.get("room") or "").strip() or None
                item = {
                    "group": str(group),
                    "date": str(date_text),
                    "pair": pair,
                    "subject": subject,
                    "subject_raw": subject_raw,
                    "subgroup": subgroup,
                    "room": room,
                    "teacher": teacher,
                }
                key = (item["group"], item["date"], item["pair"], _norm(subject), subgroup,
                       _normalize_room(room), _normalize_teacher(teacher))
                if key in seen:
                    continue
                seen.add(key)
                records.append(item)
                slots[(item["group"], item["date"], pair)].append(item)
                subjects.add(subject)
                if teacher:
                    teachers.add(teacher)
                if room:
                    rooms.add(room)

    return {
        "source_file": path.name,
        "date_from": parsed.get("date_from"),
        "date_to": parsed.get("date_to"),
        "dates": parsed.get("dates", []),
        "group_count": len(parsed.get("groups", {})),
        "table_count": parsed.get("table_count", 0),
        "groups": sorted(parsed.get("groups", {}).keys()),
        "records": records,
        "slots": slots,
        "subjects": sorted(subjects),
        "teachers": sorted(teachers),
        "rooms": sorted(rooms),
    }


def _candidate_records(entry: dict[str, Any], reference: dict[str, Any]) -> tuple[list[dict[str, Any]], str]:
    raw_subject = str(entry.get("subject_original") or entry.get("subject_raw") or entry.get("subject") or "")
    raw_key = _norm(raw_subject)
    teacher_key = _normalize_teacher(entry.get("teacher"))
    room_value = entry.get("room")
    room_key = _normalize_room(room_value)
    if not raw_key or not teacher_key:
        return [], "none"

    allow_teacher_only = bool(entry.get("self_study")) or not room_value

    def collect(scope_group: bool) -> list[dict[str, Any]]:
        found: dict[tuple[str, str], dict[str, Any]] = {}
        for item in reference["records"]:
            if scope_group and str(item.get("group")) != str(entry.get("group")):
                continue
            if _normalize_teacher(item.get("teacher")) != teacher_key:
                continue
            if not allow_teacher_only and _normalize_room(item.get("room")) != room_key:
                continue
            title = _display_subject(item.get("subject"))
            title_key = _norm(title)
            # A dangling '(' is a clipped subgroup marker: the base subject can
            # match exactly. Otherwise require a longer canonical title.
            is_dangling_subgroup_open = raw_subject.rstrip().endswith("(")
            if (title_key.startswith(raw_key) and title_key != raw_key) or (is_dangling_subgroup_open and title_key == raw_key):
                found[(title_key, title)] = item
        return list(found.values())

    # Prefer evidence from the same group; only fall back to the whole XLSX
    # catalogue when that group has no candidate for the same teacher/room.
    same_group = collect(True)
    if same_group:
        return same_group, "same_group"
    global_matches = collect(False)
    return global_matches, "global_catalog" if global_matches else "none"


def _same_slot_reference(entry: dict[str, Any], reference: dict[str, Any]) -> list[dict[str, Any]]:
    try:
        pair = int(entry.get("pair") or 0)
    except (TypeError, ValueError):
        pair = 0
    return reference["slots"].get((str(entry.get("group")), str(entry.get("date")), pair), [])


def apply_reference(data: dict[str, Any], reference_path: str | Path) -> dict[str, Any]:
    """Enrich PDF records from XLSX only on unique teacher+room+prefix matches.

    The XLSX never overwrites a replacement merely because it differs from the
    baseline. It is used to complete an obviously clipped PDF title and to add
    an auditable comparison status.
    """
    reference = load_reference(reference_path)
    recovered = 0
    ambiguous = 0
    exact_slot_matches = 0
    different_slots = 0
    missing_slots = 0
    for entry in data.get("entries", []):
        raw_before = str(entry.get("subject_original") or entry.get("subject_raw") or entry.get("subject") or "")
        raw_key = _norm(raw_before)
        # A candidate is considered only for a row already marked uncertain,
        # a clipped marker, or an unfinished parenthesis. No broad expansion of
        # complete subject names is performed.
        needs_reference_help = bool(
            entry.get("needs_review")
            or entry.get("possibly_truncated")
            or entry.get("source_text_truncated")
            or entry.get("subgroup_marker_repaired")
            or OPEN_SUBGROUP_RE.search(raw_before)
            or raw_before.rstrip().endswith("(")
        )

        candidates, candidate_scope = _candidate_records(entry, reference) if needs_reference_help else ([], "not_needed")
        entry["reference_candidate_scope"] = candidate_scope
        distinct_titles = sorted({_norm(x["subject"]): x["subject"] for x in candidates}.values(), key=str.casefold)
        entry["reference_match"] = "not_needed" if not needs_reference_help else "no_candidate"
        entry["reference_candidates"] = distinct_titles
        entry["reference_source"] = reference["source_file"]

        if len(distinct_titles) == 1:
            restored_title = distinct_titles[0]
            is_dangling_open = raw_before.rstrip().endswith("(")
            if _norm(restored_title) != raw_key or is_dangling_open:
                entry["subject"] = restored_title
                entry["subject_recovered"] = True
                entry["subject_reference_match"] = restored_title
                entry["reference_recovered"] = True
                entry["reference_match"] = "recovered_unique"
                # A dangling opening parenthesis usually means the PDF clipped
                # the subgroup marker. Restore it only if the same teacher and
                # room map to exactly one subgroup in the XLSX reference.
                if is_dangling_open and entry.get("subgroup") is None:
                    candidate_subgroups = {x.get("subgroup") for x in candidates if x.get("subgroup") in (1, 2)}
                    if len(candidate_subgroups) == 1:
                        entry["subgroup"] = next(iter(candidate_subgroups))
                        entry["subgroup_marker_repaired"] = True
                flags = list(entry.get("quality_flags") or [])
                for flag in ("ambiguous_subject_prefix", "possible_subject_truncation"):
                    flags = [x for x in flags if x != flag]
                if "subject_recovered_from_reference_xlsx" not in flags:
                    flags.append("subject_recovered_from_reference_xlsx")
                if entry.get("subgroup_marker_repaired") and "subgroup_marker_repaired_from_reference_xlsx" not in flags:
                    flags.append("subgroup_marker_repaired_from_reference_xlsx")
                entry["quality_flags"] = flags
                entry["needs_review"] = bool(
                    "missing_subject" in flags or "ambiguous_reference_candidate" in flags
                )
                entry["quality_status"] = "recovered_by_reference" if not entry["needs_review"] else "review"
                recovered += 1
            else:
                entry["reference_match"] = "exact_catalog_match"
        elif len(distinct_titles) > 1:
            entry["reference_match"] = "ambiguous"
            entry["needs_review"] = True
            entry["quality_status"] = "review"
            flags = list(entry.get("quality_flags") or [])
            if "ambiguous_reference_candidate" not in flags:
                flags.append("ambiguous_reference_candidate")
            entry["quality_flags"] = flags
            ambiguous += 1

        baseline = _same_slot_reference(entry, reference)
        entry["reference_slot_entries"] = [
            {"subject": x["subject"], "subgroup": x.get("subgroup"), "room": x.get("room"), "teacher": x.get("teacher")}
            for x in baseline
        ]
        if not baseline:
            entry["reference_slot_status"] = "not_in_reference"
            missing_slots += 1
        else:
            subject_key = _norm(entry.get("subject"))
            teacher_key = _normalize_teacher(entry.get("teacher"))
            room_key = _normalize_room(entry.get("room"))
            same = any(
                _norm(x.get("subject")) == subject_key
                and (not teacher_key or _normalize_teacher(x.get("teacher")) == teacher_key)
                and (not room_key or _normalize_room(x.get("room")) == room_key)
                for x in baseline
            )
            entry["reference_slot_status"] = "same_as_reference" if same else "differs_from_reference"
            if same:
                exact_slot_matches += 1
            else:
                different_slots += 1

    summary = {
        "source_file": reference["source_file"],
        "date_from": reference["date_from"],
        "date_to": reference["date_to"],
        "group_count": reference["group_count"],
        "table_count": reference["table_count"],
        "record_count": len(reference["records"]),
        "subject_count": len(reference["subjects"]),
        "teacher_count": len(reference["teachers"]),
        "room_count": len(reference["rooms"]),
        "recovered_from_reference": recovered,
        "ambiguous_reference_matches": ambiguous,
        "same_as_reference_slot": exact_slot_matches,
        "differs_from_reference_slot": different_slots,
        "not_in_reference_slot": missing_slots,
    }
    data["reference_xlsx"] = summary
    # Quality totals are refreshed after reference recovery.
    data["quality_summary"] = {
        **(data.get("quality_summary") or {}),
        "needs_review": sum(1 for x in data.get("entries", []) if x.get("needs_review")),
        "subject_recovered": sum(1 for x in data.get("entries", []) if x.get("subject_recovered")),
        "reference_recovered": recovered,
    }
    data.setdefault("limitations", []).append(
        f"Сопоставление с XLSX «{reference['source_file']}»: названия восстанавливаются только при единственном совпадении по фрагменту, преподавателю и аудитории (для самостоятельной работы — преподавателю). Эталон не заменяет отличающиеся занятия в PDF."
    )
    return summary


def select_reference_xlsx(pdf_path: str | Path, reference_dir: str | Path) -> Path | None:
    """Select the newest XLSX whose date range covers all dates in a PDF filename."""
    pdf_path = Path(pdf_path)
    reference_dir = Path(reference_dir)
    if not reference_dir.exists():
        return None
    date_matches = re.findall(r"\d{2}\.\d{2}\.\d{4}", pdf_path.name)
    target_dates: list[str] = []
    for item in date_matches:
        try:
            target_dates.append(datetime.strptime(item, "%d.%m.%Y").date().isoformat())
        except ValueError:
            continue
    candidates: list[tuple[str, Path]] = []
    for candidate in reference_dir.glob("*.xlsx"):
        try:
            parsed = parse_xlsx(candidate)
            start = str(parsed.get("date_from") or "")
            end = str(parsed.get("date_to") or "")
            if not start or not end:
                continue
            if target_dates and not all(start <= d <= end for d in target_dates):
                continue
            candidates.append((end, candidate))
        except Exception:
            continue
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1].stat().st_mtime), reverse=True)
    return candidates[0][1]
