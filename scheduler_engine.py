from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

DAY_ORDER = ["Mon", "Tue", "Wed", "Thu", "Fri"]
IMMUTABLE_SKILLS = {"技術", "家庭", "技家"}
PE_SUBJECT = "保体"


def _normalize_subject_name(subject: str) -> str:
    return str(subject or "").replace("　", "").replace(" ", "").strip()


def _build_excluded_subject_keywords(excluded_subjects: set[str] | None) -> set[str]:
    keywords = {_normalize_subject_name(s) for s in (excluded_subjects or set()) if _normalize_subject_name(s)}
    if PE_SUBJECT in keywords:
        keywords.update({"体育", "保健体育"})
    return keywords


def _is_excluded_subject(subject: str, excluded_subjects: set[str] | None) -> bool:
    normalized = _normalize_subject_name(subject)
    if not normalized:
        return False
    keywords = _build_excluded_subject_keywords(excluded_subjects)
    if normalized in keywords:
        return True
    return any(keyword in normalized for keyword in keywords)


class SchedulerError(RuntimeError):
    pass


def grade_of(class_name: str) -> str:
    return class_name.split("-")[0]


def parse_slot(slot: str) -> tuple[str, int]:
    d, p = slot.split("-")
    return d, int(p)


def build_slots(day_periods: dict[str, int]) -> list[str]:
    slots: list[str] = []
    for d in DAY_ORDER:
        for p in range(1, int(day_periods[d]) + 1):
            slots.append(f"{d}-{p}")
    return slots


@dataclass(frozen=True)
class Unit:
    teacher: str
    class_name: str
    subject: str
    tt: bool


class ScenarioScheduler:
    def __init__(
        self,
        classes: list[str],
        day_periods: dict[str, int],
        non_tt_demands: dict[tuple[str, str, str], int],
        tt_demands: dict[tuple[str, str, str], int],
        teacher_unavailable: dict[str, set[str]],
        class_subject_forbidden: set[tuple[str, str, str]],
        fixed_subjects: list[dict],
        exempt_subjects: set[str],
        progress_callback: Optional[Callable[[int, str], None]] = None,
    ):
        self.classes = classes
        self.day_periods = day_periods
        self.slots = build_slots(day_periods)
        self.exempt_subjects = exempt_subjects
        self.teacher_unavailable = teacher_unavailable
        self.class_subject_forbidden = class_subject_forbidden
        self.progress_callback = progress_callback
        self.total_cells = max(1, len(self.classes) * len(self.slots))
        self.total_search_slots = self.total_cells
        self.max_assigned_seen = 0
        self._last_progress = -1
        self.nodes_visited = 0
        self.backtracks = 0

        self.fixed_subject: dict[tuple[str, str], str] = {}
        for a in fixed_subjects:
            c = str(a.get("class", "")).strip()
            s = str(a.get("slot", "")).strip()
            subj = str(a.get("subject", "")).strip()
            if not c or not s or not subj:
                continue
            key = (c, s)
            if key in self.fixed_subject and self.fixed_subject[key] != subj:
                raise SchedulerError(f"固定入力衝突: {c} {s}")
            self.fixed_subject[key] = subj

        self.remaining: dict[tuple[str, str, str], int] = {
            k: int(v) for k, v in non_tt_demands.items() if int(v) > 0
        }
        self.tt_remaining: dict[tuple[str, str, str], int] = {
            k: int(v) for k, v in tt_demands.items() if int(v) > 0
        }
        self.assigned: dict[tuple[str, str], tuple[str, str]] = {}
        self.prefilled_slots: set[tuple[str, str]] = set()
        self.teacher_busy: dict[str, set[str]] = {}
        self.class_day_subjects: dict[str, dict[str, set[str]]] = {
            c: {d: set() for d in DAY_ORDER} for c in classes
        }
        self.slot_subject_by_class: dict[str, dict[str, str]] = {s: {} for s in self.slots}
        self.last_dead_end: str = ""
        self.last_dead_end_key: Optional[tuple[str, str]] = None

        self._validate_supply_and_fixed()
        self._apply_fixed_subjects()
        self._emit_progress(5, "入力検証が完了しました")

    def _emit_progress(self, percent: int, message: str, force: bool = False) -> None:
        if not self.progress_callback:
            return
        p = max(0, min(100, int(percent)))
        if (not force) and p == self._last_progress and p < 100:
            return
        self._last_progress = p
        self.progress_callback(p, message)

    def _validate_supply_and_fixed(self) -> None:
        slot_count = len(self.slots)
        demand_by_class: dict[str, int] = {c: 0 for c in self.classes}
        demand_by_class_subject: dict[tuple[str, str], int] = {}
        for (_t, c, subj), n in self.remaining.items():
            demand_by_class[c] = demand_by_class.get(c, 0) + n
            key = (c, subj)
            demand_by_class_subject[key] = demand_by_class_subject.get(key, 0) + n

        fixed_count_by_class: dict[str, int] = {c: 0 for c in self.classes}
        fixed_count: dict[tuple[str, str], int] = {}
        for (c, slot), subj in self.fixed_subject.items():
            if c not in self.classes:
                raise SchedulerError(f"固定入力のクラス不正: {c}")
            if slot not in self.slots:
                raise SchedulerError(f"固定入力のコマ不正: {slot}")
            fixed_count_by_class[c] = fixed_count_by_class.get(c, 0) + 1
            key = (c, subj)
            fixed_count[key] = fixed_count.get(key, 0) + 1

        for c in self.classes:
            remain_slots = slot_count - fixed_count_by_class.get(c, 0)
            if remain_slots < 0:
                raise SchedulerError(
                    f"固定入力数がコマ数を超過: class={c} fixed={fixed_count_by_class.get(c, 0)} slots={slot_count}"
                )
            if demand_by_class.get(c, 0) != remain_slots:
                raise SchedulerError(
                    f"クラス{c}の担当時数合計が不一致 demand={demand_by_class.get(c, 0)} slots={remain_slots}"
                )

        for key, cnt in fixed_count.items():
            if key in demand_by_class_subject and cnt > demand_by_class_subject.get(key, 0):
                raise SchedulerError(
                    f"固定入力が時数を超過: class={key[0]} subject={key[1]} fixed={cnt}"
                )

    def _apply_fixed_subjects(self) -> None:
        for (c, slot), subj in self.fixed_subject.items():
            self.prefilled_slots.add((c, slot))
            day, _ = parse_slot(slot)
            self.class_day_subjects[c][day].add(subj)
            self.slot_subject_by_class[slot][c] = subj
        self.total_search_slots = max(1, self.total_cells - len(self.prefilled_slots))

    def _current_progress_percent(self) -> int:
        return 10 + int((70 * len(self.assigned)) / self.total_search_slots)

    def _teacher_available(self, teacher: str, slot: str) -> bool:
        if slot in self.teacher_unavailable.get(teacher, set()):
            return False
        if slot in self.teacher_busy.get(teacher, set()):
            return False
        return True

    def _same_grade_conflict(self, class_name: str, slot: str, subject: str) -> bool:
        if subject in self.exempt_subjects:
            return False
        g = grade_of(class_name)
        for other_class, other_subj in self.slot_subject_by_class[slot].items():
            if grade_of(other_class) == g and other_subj == subject:
                return True
        return False

    def _domain(self, class_name: str, slot: str) -> list[tuple[str, str]]:
        day, _ = parse_slot(slot)
        out: list[tuple[str, str]] = []
        for (teacher, c, subj), remain in self.remaining.items():
            if c != class_name or remain <= 0:
                continue
            if (class_name, slot, subj) in self.class_subject_forbidden:
                continue
            if subj in self.class_day_subjects[class_name][day]:
                continue
            if not self._teacher_available(teacher, slot):
                continue
            if self._same_grade_conflict(class_name, slot, subj):
                continue
            out.append((subj, teacher))
        return out

    def _diagnose_empty_domain(self, class_name: str, slot: str) -> str:
        day, _ = parse_slot(slot)
        total_candidates = 0
        ng_forbidden = 0
        ng_daily_dup = 0
        ng_unavailable_or_busy = 0
        ng_same_grade = 0

        for (teacher, c, subj), remain in self.remaining.items():
            if c != class_name or remain <= 0:
                continue
            total_candidates += 1
            if (class_name, slot, subj) in self.class_subject_forbidden:
                ng_forbidden += 1
                continue
            if subj in self.class_day_subjects[class_name][day]:
                ng_daily_dup += 1
                continue
            if not self._teacher_available(teacher, slot):
                ng_unavailable_or_busy += 1
                continue
            if self._same_grade_conflict(class_name, slot, subj):
                ng_same_grade += 1
                continue

        if total_candidates == 0:
            return (
                f"候補ゼロ: class={class_name} slot={slot} "
                "（このクラスに割り当て可能な残時数がありません）"
            )
        return (
            f"候補ゼロ: class={class_name} slot={slot} "
            f"candidates={total_candidates} "
            f"forbidden={ng_forbidden} "
            f"daily_dup={ng_daily_dup} "
            f"teacher_ng={ng_unavailable_or_busy} "
            f"same_grade={ng_same_grade}"
        )

    def _pick_next_var(self) -> Optional[tuple[str, str, list[tuple[str, str]]]]:
        best: Optional[tuple[str, str]] = None
        best_domain: Optional[list[tuple[str, str]]] = None
        best_peak = -1
        for c in self.classes:
            for s in self.slots:
                if (c, s) in self.prefilled_slots:
                    continue
                if (c, s) in self.assigned:
                    continue
                dom = self._domain(c, s)
                if not dom:
                    self.last_dead_end = self._diagnose_empty_domain(c, s)
                    self.last_dead_end_key = (c, s)
                    return (c, s, [])
                # 週4時間など残時数が多い教科を先に処理するため、
                # ドメイン内の最大残時数を優先し、同点ならMRV(候補最少)。
                peak = max(self.remaining.get((t, c, subj), 0) for subj, t in dom)
                if (
                    best is None
                    or peak > best_peak
                    or (peak == best_peak and len(dom) < len(best_domain or []))
                ):
                    best = (c, s)
                    best_domain = dom
                    best_peak = peak
                    if len(dom) == 1:
                        return (best[0], best[1], best_domain)
        if best is None:
            return None
        return (best[0], best[1], best_domain or [])

    def _assign(self, class_name: str, slot: str, subject: str, teacher: str) -> None:
        self.assigned[(class_name, slot)] = (subject, teacher)
        self.remaining[(teacher, class_name, subject)] -= 1
        self.teacher_busy.setdefault(teacher, set()).add(slot)
        day, _ = parse_slot(slot)
        self.class_day_subjects[class_name][day].add(subject)
        self.slot_subject_by_class[slot][class_name] = subject

        assigned_count = len(self.assigned)
        if assigned_count > self.max_assigned_seen:
            self.max_assigned_seen = assigned_count
            pct = self._current_progress_percent()
            self._emit_progress(pct, f"探索中: {assigned_count}/{self.total_search_slots} コマ確定")

    def _unassign(self, class_name: str, slot: str, subject: str, teacher: str) -> None:
        del self.assigned[(class_name, slot)]
        self.remaining[(teacher, class_name, subject)] += 1
        self.teacher_busy.get(teacher, set()).discard(slot)
        day, _ = parse_slot(slot)
        if subject in self.class_day_subjects[class_name][day]:
            self.class_day_subjects[class_name][day].remove(subject)
        self.slot_subject_by_class[slot].pop(class_name, None)

    def _rank_candidates(
        self,
        class_name: str,
        domain: list[tuple[str, str]],
    ) -> list[tuple[str, str]]:
        # 高時数教科を先に入れ、同点時は教師負荷が小さい候補を優先
        return sorted(
            domain,
            key=lambda st: (
                -self.remaining.get((st[1], class_name, st[0]), 0),
                len(self.teacher_busy.get(st[1], set())),
                st[0],
                st[1],
            ),
        )

    def _dfs(self) -> bool:
        # 高時数優先で進め、詰まったら直前手を入れ替える（局所バックトラック）
        picked = self._pick_next_var()
        if picked is None:
            return True
        class_name, slot, domain = picked
        if not domain:
            return False

        ranked = self._rank_candidates(class_name, domain)
        for subject, teacher in ranked:
            self.nodes_visited += 1
            if self.nodes_visited % 5000 == 0:
                pct = self._current_progress_percent()
                self._emit_progress(
                    pct,
                    f"優先割当+入替中: {len(self.assigned)}/{self.total_search_slots} コマ確定 | "
                    f"試行ノード={self.nodes_visited} バックトラック={self.backtracks}",
                    force=True,
                )
            self._assign(class_name, slot, subject, teacher)
            if self._dfs():
                return True
            self._unassign(class_name, slot, subject, teacher)
            self.backtracks += 1
        return False

    def _assign_tt(self) -> dict[tuple[str, str], list[str]]:
        tt_map: dict[tuple[str, str], list[str]] = {}
        units: list[Unit] = []
        for (teacher, c, subj), n in self.tt_remaining.items():
            for _ in range(n):
                units.append(Unit(teacher=teacher, class_name=c, subject=subj, tt=True))

        def candidate_count(u: Unit) -> int:
            cnt = 0
            for s in self.slots:
                subj = self.slot_subject_by_class.get(s, {}).get(u.class_name, "")
                if not subj:
                    continue
                if subj != u.subject:
                    continue
                if s in self.teacher_unavailable.get(u.teacher, set()):
                    continue
                if s in self.teacher_busy.get(u.teacher, set()):
                    continue
                cnt += 1
            return cnt

        units.sort(key=candidate_count)

        def rec(i: int) -> bool:
            if i >= len(units):
                return True
            u = units[i]
            for s in self.slots:
                subj = self.slot_subject_by_class.get(s, {}).get(u.class_name, "")
                if not subj:
                    continue
                if subj != u.subject:
                    continue
                if s in self.teacher_unavailable.get(u.teacher, set()):
                    continue
                if s in self.teacher_busy.get(u.teacher, set()):
                    continue
                self.teacher_busy.setdefault(u.teacher, set()).add(s)
                tt_map.setdefault((u.class_name, s), []).append(u.teacher)
                if rec(i + 1):
                    return True
                tt_map[(u.class_name, s)].pop()
                if not tt_map[(u.class_name, s)]:
                    del tt_map[(u.class_name, s)]
                self.teacher_busy[u.teacher].remove(s)
            return False

        if not rec(0):
            raise SchedulerError("TT割当が成立しませんでした。TT時数または不可コマを見直してください。")
        return tt_map

    def solve(self) -> tuple[dict[tuple[str, str], str], dict[tuple[str, str], str], dict[tuple[str, str], list[str]]]:
        self._emit_progress(10, "制約探索を開始します")
        ok = self._dfs()
        if not ok:
            detail = self.last_dead_end.strip()
            if detail:
                raise SchedulerError(
                    "制約を満たす時間割を見つけられませんでした。"
                    f"\n診断: {detail}\n"
                    "入力条件（不可コマ・時数・同日重複禁止・学年跨ぎ同時刻重複禁止）を見直してください。"
                )
            raise SchedulerError("制約を満たす時間割を見つけられませんでした。入力条件を見直してください。")

        remain_sum = sum(v for v in self.remaining.values() if v > 0)
        if remain_sum != 0:
            raise SchedulerError(f"未割当時数が残っています: {remain_sum}")

        self._emit_progress(82, "TT割当を実行しています")
        tt_map = self._assign_tt()
        self._emit_progress(98, "出力データを整形しています")
        subject_map = dict(self.fixed_subject)
        for (c, s), (subj, _t) in self.assigned.items():
            subject_map[(c, s)] = subj
        teacher_map = {(c, s): t for (c, s), (_subj, t) in self.assigned.items()}
        self._emit_progress(100, "シナリオ完了")
        return subject_map, teacher_map, tt_map


def build_teacher_demands(
    teachers: list[dict],
    excluded_subjects: set[str] | None = None,
) -> tuple[
    dict[tuple[str, str, str], int],
    dict[tuple[str, str, str], int],
    dict[str, set[str]],
    set[tuple[str, str, str]],
]:
    non_tt: dict[tuple[str, str, str], int] = {}
    tt: dict[tuple[str, str, str], int] = {}
    unavail: dict[str, set[str]] = {}
    class_subject_forbidden: set[tuple[str, str, str]] = set()

    for t in teachers:
        name = str(t.get("name", "")).strip()
        if not name:
            continue
        subj_list = t.get("subjects", [])
        subject = str(subj_list[0]).strip() if subj_list else ""
        if not subject:
            continue
        excluded = _is_excluded_subject(subject, excluded_subjects)
        unavail[name] = set(t.get("unavailable_slots", []))
        target_classes: set[str] = set()
        for ca in t.get("class_assignments", []):
            cls = str(ca.get("class", "")).strip()
            if not cls:
                continue
            target_classes.add(cls)
            try:
                hours = int(ca.get("hours", 0))
            except Exception:
                hours = 0
            if hours <= 0:
                continue
            if excluded:
                continue
            key = (name, cls, subject)
            if bool(ca.get("tt", False)):
                tt[key] = tt.get(key, 0) + hours
            else:
                non_tt[key] = non_tt.get(key, 0) + hours
        for cls in t.get("assigned_classes", []):
            c = str(cls).strip()
            if c:
                target_classes.add(c)
        for cls in target_classes:
            for slot in unavail[name]:
                class_subject_forbidden.add((cls, slot, subject))
    return non_tt, tt, unavail, class_subject_forbidden


def _fixed_list_to_map(fixed_subjects: list[dict]) -> dict[tuple[str, str], str]:
    out: dict[tuple[str, str], str] = {}
    for a in fixed_subjects:
        c = str(a.get("class", "")).strip()
        s = str(a.get("slot", "")).strip()
        subj = str(a.get("subject", "")).strip()
        if c and s and subj:
            out[(c, s)] = subj
    return out


def _fixed_map_to_list(fixed_map: dict[tuple[str, str], str]) -> list[dict]:
    return [{"class": c, "slot": s, "subject": subj} for (c, s), subj in fixed_map.items()]


def _try_solve_with_fixed(
    classes: list[str],
    day_periods: dict[str, int],
    non_tt: dict[tuple[str, str, str], int],
    tt: dict[tuple[str, str, str], int],
    unavail: dict[str, set[str]],
    class_subject_forbidden: set[tuple[str, str, str]],
    fixed_map: dict[tuple[str, str], str],
    exempt_subjects: set[str],
) -> bool:
    try:
        scheduler = ScenarioScheduler(
            classes=classes,
            day_periods=day_periods,
            non_tt_demands=non_tt,
            tt_demands=tt,
            teacher_unavailable=unavail,
            class_subject_forbidden=class_subject_forbidden,
            fixed_subjects=_fixed_map_to_list(fixed_map),
            exempt_subjects=exempt_subjects,
            progress_callback=None,
        )
        scheduler.solve()
        return True
    except SchedulerError:
        return False


def _build_move_specs(
    day_periods: dict[str, int],
    fixed_map: dict[tuple[str, str], str],
) -> list[dict]:
    slots = build_slots(day_periods)
    specs: list[dict] = []

    for (cls, from_slot), subj in fixed_map.items():
        if subj in IMMUTABLE_SKILLS or subj == PE_SUBJECT:
            continue
        for to_slot in slots:
            if to_slot == from_slot:
                continue
            if (cls, to_slot) in fixed_map:
                continue
            specs.append(
                {
                    "remove": [((cls, from_slot), subj)],
                    "add": [((cls, to_slot), subj)],
                    "desc": f"{cls} {from_slot} {subj} -> {to_slot}",
                }
            )

    pe_groups: dict[str, list[str]] = {}
    for (cls, slot), subj in fixed_map.items():
        if subj == PE_SUBJECT:
            pe_groups.setdefault(slot, []).append(cls)

    for from_slot, pe_classes in pe_groups.items():
        if len(pe_classes) < 2:
            continue
        for to_slot in slots:
            if to_slot == from_slot:
                continue
            if any((cls, to_slot) in fixed_map for cls in pe_classes):
                continue
            specs.append(
                {
                    "remove": [((cls, from_slot), PE_SUBJECT) for cls in pe_classes],
                    "add": [((cls, to_slot), PE_SUBJECT) for cls in pe_classes],
                    "desc": f"保体同時移動 {','.join(pe_classes)} {from_slot} -> {to_slot}",
                }
            )
    return specs


def _apply_specs(
    base_map: dict[tuple[str, str], str],
    specs: list[dict],
) -> Optional[dict[tuple[str, str], str]]:
    cur = dict(base_map)
    for sp in specs:
        for (k, subj) in sp["remove"]:
            if cur.get(k) != subj:
                return None
            del cur[k]
        for (k, subj) in sp["add"]:
            if k in cur:
                return None
            cur[k] = subj
    return cur


def suggest_skill_changes(
    classes: list[str],
    day_periods: dict[str, int],
    non_tt: dict[tuple[str, str, str], int],
    tt: dict[tuple[str, str, str], int],
    unavail: dict[str, set[str]],
    class_subject_forbidden: set[tuple[str, str, str]],
    fixed_subjects: list[dict],
    exempt_subjects: set[str],
    max_one_step: int = 5,
    max_two_step: int = 3,
    spec_limit: int = 240,
    two_pool_limit: int = 40,
    exhaustive_two_step: bool = False,
    dead_end_key: Optional[tuple[str, str]] = None,
) -> list[str]:
    base_map = _fixed_list_to_map(fixed_subjects)
    if not base_map:
        return []

    specs = _build_move_specs(day_periods, base_map)
    if not specs:
        return []

    if dead_end_key:
        dead_cls, dead_slot = dead_end_key
        def _priority(sp: dict) -> tuple[int, int]:
            score = 0
            for (k, _subj) in sp["remove"] + sp["add"]:
                cls, slot = k
                if cls == dead_cls:
                    score += 3
                if slot == dead_slot:
                    score += 2
            # 高優先を先頭へ
            return (-score, len(sp["remove"]) + len(sp["add"]))
        specs.sort(key=_priority)

    if spec_limit > 0:
        specs = specs[:spec_limit]
    suggestions: list[str] = []

    for sp in specs:
        moved = _apply_specs(base_map, [sp])
        if not moved:
            continue
        if _try_solve_with_fixed(
            classes, day_periods, non_tt, tt, unavail, class_subject_forbidden, moved, exempt_subjects
        ):
            suggestions.append(f"1手: {sp['desc']}")
            if len(suggestions) >= max_one_step:
                return suggestions

    limit_specs = specs if exhaustive_two_step else specs[: max(1, two_pool_limit)]
    two_count = 0
    for i in range(len(limit_specs)):
        for j in range(i + 1, len(limit_specs)):
            sp1 = limit_specs[i]
            sp2 = limit_specs[j]
            moved = _apply_specs(base_map, [sp1, sp2])
            if not moved:
                continue
            if _try_solve_with_fixed(
                classes, day_periods, non_tt, tt, unavail, class_subject_forbidden, moved, exempt_subjects
            ):
                suggestions.append(f"2手: {sp1['desc']} / {sp2['desc']}")
                two_count += 1
                if two_count >= max_two_step:
                    return suggestions
    return suggestions


def solve_all_scenarios(
    config_data: dict,
    exempt_subjects: set[str],
    progress_callback: Optional[Callable[[int, str], None]] = None,
    suggestion_options: Optional[dict] = None,
) -> dict[str, dict]:
    solver_cfg = dict(config_data.get("solver", {}) or {})
    engine = str(solver_cfg.get("engine", "")).strip().lower()
    if engine == "cp_sat":
        try:
            from scheduler_engine_cp import solve_all_scenarios_cp
        except Exception as e:
            raise SchedulerError(
                "CP-SATエンジンを利用するには `ortools` が必要です。"
                f" import error: {e}"
            )
        return solve_all_scenarios_cp(
            config_data=config_data,
            exempt_subjects=exempt_subjects,
            progress_callback=progress_callback,
            time_limit_sec=float(solver_cfg.get("time_limit_sec", 30)),
        )

    classes = config_data.get("classes", [])
    day_periods = config_data.get("day_periods", {})
    teachers = config_data.get("teachers", [])
    scenarios = config_data.get("scenarios", [])

    non_tt, tt, unavail, class_subject_forbidden = build_teacher_demands(teachers, excluded_subjects=exempt_subjects)
    if not non_tt:
        raise SchedulerError("教員タブに非TTの担当時数がありません。")
    suggestion_options = dict(suggestion_options or {})

    solved: dict[str, dict] = {}
    total = max(1, len(scenarios))
    for idx, sc in enumerate(scenarios):
        sid = sc.get("id", "")
        fixed = sc.get("fixed_assignments", [])

        def _scenario_progress(p: int, msg: str, i: int = idx, sid_now: str = sid) -> None:
            if not progress_callback:
                return
            base = int((i * 100) / total)
            span = int(100 / total)
            global_p = min(100, base + int((span * p) / 100))
            progress_callback(global_p, f"{sid_now}: {msg}")

        scheduler = ScenarioScheduler(
            classes=classes,
            day_periods=day_periods,
            non_tt_demands=non_tt,
            tt_demands=tt,
            teacher_unavailable=unavail,
            class_subject_forbidden=class_subject_forbidden,
            fixed_subjects=fixed,
            exempt_subjects=exempt_subjects,
            progress_callback=_scenario_progress,
        )
        try:
            subject_map, teacher_map, tt_map = scheduler.solve()
        except SchedulerError as e:
            if progress_callback:
                progress_callback(95, f"{sid}: 代替提案を探索中...")
            suggestions = suggest_skill_changes(
                classes=classes,
                day_periods=day_periods,
                non_tt=non_tt,
                tt=tt,
                unavail=unavail,
                class_subject_forbidden=class_subject_forbidden,
                fixed_subjects=fixed,
                exempt_subjects=exempt_subjects,
                max_one_step=int(suggestion_options.get("max_one_step", 5)),
                max_two_step=int(suggestion_options.get("max_two_step", 3)),
                spec_limit=int(suggestion_options.get("spec_limit", 240)),
                two_pool_limit=int(suggestion_options.get("two_pool_limit", 40)),
                exhaustive_two_step=bool(suggestion_options.get("exhaustive_two_step", False)),
                dead_end_key=scheduler.last_dead_end_key,
            )
            if suggestions:
                lines = [str(e), "", "技能科目の変更提案（制約は変更せず技能配置のみ変更）:"]
                lines.extend([f"- {s}" for s in suggestions])
                raise SchedulerError("\n".join(lines))
            raise

        solved[sid] = {
            "subject_map": subject_map,
            "teacher_map": teacher_map,
            "tt_map": tt_map,
        }

    if progress_callback:
        progress_callback(100, "全シナリオ完了")
    return solved
