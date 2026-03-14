from __future__ import annotations

import argparse
import json
import os
import random
import sys
import time
from dataclasses import dataclass
from copy import deepcopy
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Dict, List, Optional, Set, Tuple

DAY_ORDER = ["Mon", "Tue", "Wed", "Thu", "Fri"]
DAY_JP = {"Mon": "月", "Tue": "火", "Wed": "水", "Thu": "木", "Fri": "金"}
DAY_START_COL = {"Mon": 2, "Tue": 8, "Wed": 15, "Thu": 21, "Fri": 28}  # B, H, O, U, AB
CLASS_ORDER_DEFAULT = ["1-1", "1-2", "1-3", "1-4", "2-1", "2-2", "2-3", "2-4", "3-1", "3-2", "3-3"]
DEFAULT_SKILL_SUBJECTS = {"音楽", "美術", "音美", "技術", "家庭", "技家", "保体", "学活", "道徳", "総合", "個別"}
ONBI_FILL_COLOR = "F0E68C"
TECHKA_FILL_COLOR = "FFB6C1"


@dataclass(frozen=True)
class Slot:
    day: str
    period: int

    @property
    def key(self) -> str:
        return f"{self.day}-{self.period}"


# orjson があれば高速シリアライズを使い、なければ標準 json へフォールバックする。
try:
    import orjson  # type: ignore
except Exception:
    orjson = None


def read_json(path: Path) -> dict:
    raw = path.read_bytes()
    if orjson is not None:
        return orjson.loads(raw)
    return json.loads(raw.decode("utf-8"))


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if orjson is not None:
        raw = orjson.dumps(data, option=orjson.OPT_INDENT_2)
    else:
        raw = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")

    # 途中書き込みを避けるため、同一ディレクトリへ一時保存してから置換する。
    with NamedTemporaryFile("wb", delete=False, dir=path.parent) as tmp:
        tmp.write(raw)
        tmp_path = Path(tmp.name)
    os.replace(tmp_path, path)


def build_all_slots(day_periods: Dict[str, int]) -> List[Slot]:
    out: List[Slot] = []
    for d in DAY_ORDER:
        for p in range(1, int(day_periods[d]) + 1):
            out.append(Slot(d, p))
    return out


def parse_slot_key(slot_key: str) -> Slot:
    day, p = slot_key.split("-")
    return Slot(day, int(p))


def grade_of(class_name: str) -> str:
    return class_name.split("-")[0]


def _class_sort_key(class_name: str) -> tuple:
    parts = class_name.split("-")
    if len(parts) != 2:
        return (class_name, 0)
    try:
        return (int(parts[0]), int(parts[1]))
    except ValueError:
        return (parts[0], parts[1])


def _normalize_subject_name(subject: str) -> str:
    return str(subject or "").replace("　", "").replace(" ", "").strip()


def _build_skill_subject_keywords(skill_subjects: Set[str] | None = None) -> Set[str]:
    keywords = {_normalize_subject_name(s) for s in (skill_subjects or DEFAULT_SKILL_SUBJECTS) if _normalize_subject_name(s)}
    if "保体" in keywords:
        keywords.update({"体育", "保健体育"})
    return keywords


def _is_skill_subject_name(subject: str, skill_subject_keywords: Set[str]) -> bool:
    normalized = _normalize_subject_name(subject)
    if not normalized:
        return False
    if normalized in skill_subject_keywords:
        return True
    return any(keyword in normalized for keyword in skill_subject_keywords)


def _subject_group_key(subject: str) -> str | None:
    normalized = _normalize_subject_name(subject)
    if normalized in {"音美", "音楽", "美術"}:
        return "onbi"
    if normalized in {"技家", "技術", "家庭"}:
        return "techka"
    return None


def _origin_group_key(origin_subject: str | None) -> str | None:
    normalized = _normalize_subject_name(origin_subject or "")
    if normalized == "音美":
        return "onbi"
    if normalized == "技家":
        return "techka"
    return None


class CSPTimetableSolver:
    def __init__(self, config: dict, scenario: dict):
        self.config = config
        self.scenario = scenario
        self.random = random.Random(config.get("seed", 42))
        self.start_time = 0.0
        self.time_limit_sec = int(config.get("solver", {}).get("time_limit_sec", 60))

        self.classes: List[str] = config["classes"]
        self.day_periods: Dict[str, int] = config["day_periods"]
        self.slots: List[Slot] = build_all_slots(self.day_periods)
        self.slot_keys: List[str] = [s.key for s in self.slots]

        self.skill_subject_keywords: Set[str] = _build_skill_subject_keywords(set(config.get("skill_subjects", [])))
        self.cross_grade_exempt: Set[str] = set(config.get("cross_grade_duplicate_exempt_subjects", []))

        self.class_subject_teacher: Dict[str, Dict[str, str]] = config.get("class_subject_teacher", {})
        self.teacher_availability: Dict[str, Set[str]] = self._build_teacher_availability(config.get("teachers", []))
        self.class_subject_unavailable: Set[Tuple[str, str, str]] = self._build_class_subject_unavailable(
            config.get("class_subject_unavailable", [])
        )

        self.weekly_requirements: Dict[str, Dict[str, int]] = self._normalize_requirements(
            scenario.get("weekly_requirements", {})
        )
        self.fixed_assignments = scenario.get("fixed_assignments", [])
        self.manual_skill_assignments = scenario.get("manual_skill_assignments", [])

        self.schedule: Dict[str, Dict[str, str]] = {c: {} for c in self.classes}
        self.teacher_slot_used: Dict[str, Set[str]] = {}

        self._apply_fixed_assignments(self.fixed_assignments + self.manual_skill_assignments)
        self.remaining = self._compute_remaining_requirements()
        self.variable_slots = self._build_variable_slots()

    def _build_teacher_availability(self, teachers: List[dict]) -> Dict[str, Set[str]]:
        all_slot_keys = set(self.slot_keys)
        result: Dict[str, Set[str]] = {}
        for t in teachers:
            name = t.get("name", "").strip()
            if not name:
                continue
            unavailable = set(t.get("unavailable_slots", []))
            result[name] = all_slot_keys - unavailable
        return result

    def _build_class_subject_unavailable(self, items: List[dict]) -> Set[Tuple[str, str, str]]:
        result: Set[Tuple[str, str, str]] = set()
        for x in items:
            c = x.get("class", "").strip()
            s = x.get("subject", "").strip()
            k = x.get("slot", "").strip()
            if c and s and k:
                result.add((c, s, k))
        return result

    def _normalize_requirements(self, req: Dict[str, Dict[str, float]]) -> Dict[str, Dict[str, int]]:
        normalized: Dict[str, Dict[str, int]] = {}
        for c, subj_map in req.items():
            normalized[c] = {}
            for subj, val in subj_map.items():
                if _is_skill_subject_name(subj, self.skill_subject_keywords):
                    continue
                f = float(val)
                if abs(f - round(f)) > 1e-6:
                    raise ValueError(
                        f"週時数に小数があります: class={c}, subject={subj}, value={val}. "
                        "このシナリオでは整数時数を入力してください（小数は別シナリオ化）。"
                    )
                hours = int(round(f))
                if hours > 0:
                    normalized[c][subj] = hours
        return normalized

    def _apply_fixed_assignments(self, assignments: List[dict]) -> None:
        for a in assignments:
            c = a["class"]
            slot = a["slot"]
            subj = a["subject"]
            if slot in self.schedule[c] and self.schedule[c][slot] != subj:
                raise ValueError(f"固定入力が衝突しています: class={c}, slot={slot}")
            self.schedule[c][slot] = subj
            teacher = self.class_subject_teacher.get(c, {}).get(subj)
            if teacher:
                self.teacher_slot_used.setdefault(teacher, set()).add(slot)

    def _compute_remaining_requirements(self) -> Dict[str, Dict[str, int]]:
        remaining = deepcopy(self.weekly_requirements)
        for c in self.classes:
            for subj in self.schedule[c].values():
                if subj in remaining.get(c, {}):
                    remaining[c][subj] -= 1
        for c in self.classes:
            for subj, cnt in remaining.get(c, {}).items():
                if cnt < 0:
                    raise ValueError(f"固定入力が週時数を超過: class={c}, subject={subj}, remaining={cnt}")
        return remaining

    def _build_variable_slots(self) -> List[Tuple[str, str]]:
        out: List[Tuple[str, str]] = []
        for c in self.classes:
            for s in self.slot_keys:
                if s not in self.schedule[c]:
                    out.append((c, s))
        return out

    def _subject_used_same_day(self, c: str, day: str, subject: str) -> bool:
        for slot_key, subj in self.schedule[c].items():
            if subj != subject:
                continue
            sl = parse_slot_key(slot_key)
            if sl.day == day:
                return True
        return False

    def _cross_grade_conflict(self, c: str, slot: str, subject: str) -> bool:
        if subject in self.cross_grade_exempt or _is_skill_subject_name(subject, self.skill_subject_keywords):
            return False
        g = grade_of(c)
        for other in self.classes:
            if other == c:
                continue
            if grade_of(other) == g:
                continue
            if self.schedule[other].get(slot) == subject:
                return True
        return False

    def _teacher_conflict(self, c: str, slot: str, subject: str) -> bool:
        teacher = self.class_subject_teacher.get(c, {}).get(subject)
        if not teacher:
            return False
        if slot not in self.teacher_availability.get(teacher, set()):
            return True
        if slot in self.teacher_slot_used.get(teacher, set()):
            return True
        return False

    def _class_subject_slot_forbidden(self, c: str, slot: str, subject: str) -> bool:
        return (c, subject, slot) in self.class_subject_unavailable

    def _domain(self, c: str, slot: str) -> List[str]:
        day = parse_slot_key(slot).day
        candidates: List[str] = []
        for subj, rem in self.remaining.get(c, {}).items():
            if rem <= 0:
                continue
            if self._subject_used_same_day(c, day, subj):
                continue
            if self._class_subject_slot_forbidden(c, slot, subj):
                continue
            if self._cross_grade_conflict(c, slot, subj):
                continue
            if self._teacher_conflict(c, slot, subj):
                continue
            candidates.append(subj)
        self.random.shuffle(candidates)
        return candidates

    def _forward_check(self, c: str) -> bool:
        unassigned = [s for s in self.slot_keys if s not in self.schedule[c]]
        for subj, need in self.remaining.get(c, {}).items():
            if need <= 0:
                continue
            possible = 0
            for slot in unassigned:
                day = parse_slot_key(slot).day
                if self._subject_used_same_day(c, day, subj):
                    continue
                if self._class_subject_slot_forbidden(c, slot, subj):
                    continue
                if self._cross_grade_conflict(c, slot, subj):
                    continue
                if self._teacher_conflict(c, slot, subj):
                    continue
                possible += 1
            if possible < need:
                return False
        return True

    def _class_has_remaining_requirements(self, c: str) -> bool:
        return any(cnt > 0 for cnt in self.remaining.get(c, {}).values())

    def _pick_next_variable(self) -> Optional[Tuple[str, str, List[str]]]:
        best = None
        best_domain = None
        for c, slot in self.variable_slots:
            if slot in self.schedule[c]:
                continue
            if not self._class_has_remaining_requirements(c):
                continue
            dom = self._domain(c, slot)
            if not dom:
                return (c, slot, [])
            if best is None or len(dom) < len(best_domain):  # type: ignore[arg-type]
                best = (c, slot)
                best_domain = dom
                if len(dom) == 1:
                    break
        if best is None:
            return None
        return (best[0], best[1], best_domain or [])

    def _assign(self, c: str, slot: str, subj: str) -> None:
        self.schedule[c][slot] = subj
        self.remaining[c][subj] -= 1
        teacher = self.class_subject_teacher.get(c, {}).get(subj)
        if teacher:
            self.teacher_slot_used.setdefault(teacher, set()).add(slot)

    def _unassign(self, c: str, slot: str, subj: str) -> None:
        del self.schedule[c][slot]
        self.remaining[c][subj] += 1
        teacher = self.class_subject_teacher.get(c, {}).get(subj)
        if teacher and slot in self.teacher_slot_used.get(teacher, set()):
            self.teacher_slot_used[teacher].remove(slot)

    def _deadline_exceeded(self) -> bool:
        return (time.time() - self.start_time) >= self.time_limit_sec

    def _dfs(self) -> bool:
        if self._deadline_exceeded():
            return False

        picked = self._pick_next_variable()
        if picked is None:
            return True
        c, slot, domain = picked
        if not domain:
            return False

        for subj in domain:
            self._assign(c, slot, subj)
            if self._forward_check(c) and self._dfs():
                return True
            self._unassign(c, slot, subj)
        return False

    def solve(self) -> Dict[str, Dict[str, str]]:
        self.start_time = time.time()
        retries = int(self.config.get("solver", {}).get("random_restarts", 6))
        best = None
        for _ in range(retries):
            if self._dfs():
                return self.schedule
            if best is None or sum(len(v) for v in self.schedule.values()) > sum(len(v) for v in best.values()):
                best = deepcopy(self.schedule)
            self.schedule = {c: {k: v for k, v in self.schedule[c].items() if k in self._fixed_slots(c)} for c in self.classes}
            self.teacher_slot_used = {}
            self._apply_fixed_assignments(self.fixed_assignments + self.manual_skill_assignments)
            self.remaining = self._compute_remaining_requirements()
            self.random.seed(self.random.randint(0, 10_000_000))
            self.start_time = time.time()
        if best is not None:
            self.schedule = best
        raise RuntimeError("解を見つけられませんでした。制約が厳しすぎる可能性があります。")

    def _fixed_slots(self, c: str) -> Set[str]:
        fixed = {a["slot"] for a in (self.fixed_assignments + self.manual_skill_assignments) if a["class"] == c}
        return set(self.schedule[c].keys()) & fixed


# Excel generation helpers were moved to timetable_excel.py


def validate_config(config: dict) -> None:
    required_root = ["classes", "day_periods", "teachers", "class_subject_teacher", "scenarios"]
    for k in required_root:
        if k not in config:
            raise ValueError(f"config missing key: {k}")
    for d in DAY_ORDER:
        if d not in config["day_periods"]:
            raise ValueError(f"day_periods missing {d}")


def make_template(config_path: Path) -> None:
    template = {
        "year": 6,
        "output_path": "./output/timetable.xlsx",
        "seed": 42,
        "classes": CLASS_ORDER_DEFAULT,
        "day_periods": {"Mon": 5, "Tue": 6, "Wed": 5, "Thu": 6, "Fri": 6},
        "skill_subjects": ["音楽", "美術", "技術", "家庭", "技家", "保体", "音美", "総合"],
        "cross_grade_duplicate_exempt_subjects": ["保体", "音楽", "美術", "技術", "家庭", "技家", "音美", "総合"],
        "teachers": [],
        "class_subject_teacher": {},
        "class_subject_unavailable": [],
        "solver": {"time_limit_sec": 90, "random_restarts": 8},
        "scenarios": [
            {
                "id": "1grade-music-art",
                "target_block": "upper",
                "weekly_requirements": {},
                "fixed_assignments": [],
                "manual_skill_assignments": []
            },
            {
                "id": "1grade-general",
                "target_block": "lower",
                "weekly_requirements": {},
                "fixed_assignments": [],
                "manual_skill_assignments": []
            }
        ]
    }
    write_json(config_path, template)


def extract_pdf_text(pdf_path: Path, output_txt: Path) -> None:
    try:
        from pypdf import PdfReader
    except ImportError as e:
        raise RuntimeError("PDF抽出には pypdf が必要です。`pip install pypdf` を実行してください。") from e

    reader = PdfReader(str(pdf_path))
    lines = []
    for i, page in enumerate(reader.pages, start=1):
        txt = page.extract_text() or ""
        lines.append(f"===== PAGE {i} =====\n{txt}")
    output_txt.parent.mkdir(parents=True, exist_ok=True)
    output_txt.write_text("\n\n".join(lines), encoding="utf-8")


def run_solve(config_path: Path) -> None:
    from timetable_excel import create_workbook_by_structure

    config = read_json(config_path)
    if "output_path" not in config:
        config["output_path"] = "./output/timetable.xlsx"
    validate_config(config)

    solved: Dict[str, Dict[str, Dict[str, str]]] = {}
    for sc in config["scenarios"]:
        solver = CSPTimetableSolver(config, sc)
        solved[sc["id"]] = solver.solve()

    output_path = Path(config["output_path"])
    create_workbook_by_structure(config, solved, output_path)
    print(f"生成完了: {output_path}")


def main() -> int:
    parser = argparse.ArgumentParser(description="時間割自動生成ツール")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_init = sub.add_parser("init", help="設定テンプレートJSONを作成")
    p_init.add_argument("--config", default="./config.json")

    p_solve = sub.add_parser("solve", help="JSON設定から時間割を作成してExcelへ出力")
    p_solve.add_argument("--config", default="./config.json")

    p_pdf = sub.add_parser("extract-pdf", help="PDFテキスト抽出（入力補助）")
    p_pdf.add_argument("--pdf", required=True)
    p_pdf.add_argument("--out", default="./output/pdf_text.txt")

    args = parser.parse_args()

    if args.cmd == "init":
        cfg = Path(args.config)
        make_template(cfg)
        print(f"テンプレート作成: {cfg}")
        return 0

    if args.cmd == "solve":
        run_solve(Path(args.config))
        return 0

    if args.cmd == "extract-pdf":
        extract_pdf_text(Path(args.pdf), Path(args.out))
        print(f"PDF抽出完了: {args.out}")
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())



