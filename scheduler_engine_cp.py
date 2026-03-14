from __future__ import annotations
from dataclasses import dataclass
from typing import Callable, Optional
from ortools.sat.python import cp_model

DAY_ORDER = ["Mon","Tue","Wed","Thu","Fri"]
PE_SUBJECT = "保体"


def _normalize_subject_name(subject: str) -> str:
    return str(subject or "").replace("　", "").replace(" ", "").strip()


def _build_excluded_subject_keywords(excluded_subjects: set[str]) -> set[str]:
    keywords = {_normalize_subject_name(s) for s in excluded_subjects if _normalize_subject_name(s)}
    if PE_SUBJECT in keywords:
        keywords.update({"体育", "保健体育"})
    return keywords


def _is_excluded_subject(subject: str, excluded_keywords: set[str]) -> bool:
    normalized = _normalize_subject_name(subject)
    if not normalized:
        return False
    if normalized in excluded_keywords:
        return True
    return any(keyword in normalized for keyword in excluded_keywords)


class SchedulerError(RuntimeError):
    pass


def grade_of(class_name:str)->str:
    return class_name.split("-")[0]


def parse_slot(slot:str)->tuple[str,int]:
    d,p=slot.split("-")
    return d,int(p)


def build_slots(day_periods:dict[str,int])->list[str]:
    slots=[]
    for d in DAY_ORDER:
        for p in range(1,int(day_periods[d])+1):
            slots.append(f"{d}-{p}")
    return slots


def _build_subject_requirements(non_tt: dict[tuple[str, str, str], int]) -> dict[tuple[str, str], int]:
    req: dict[tuple[str, str], int] = {}
    for (_k, c, s), h in non_tt.items():
        req[(c, s)] = req.get((c, s), 0) + int(h)
    return req


def _find_class_open_slot_mismatches(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed: dict[tuple[str, str], str],
) -> list[tuple[str, int, int]]:
    demand_by_class: dict[str, int] = {c: 0 for c in classes}
    for (_teacher, cls, _subject), h in non_tt.items():
        if cls in demand_by_class:
            demand_by_class[cls] += int(h)

    fixed_count_by_class: dict[str, int] = {c: 0 for c in classes}
    for (cls, _slot) in fixed.keys():
        if cls in fixed_count_by_class:
            fixed_count_by_class[cls] += 1

    total_slots = len(slots)
    mismatches: list[tuple[str, int, int]] = []
    for c in classes:
        demand = demand_by_class.get(c, 0)
        free_slots = total_slots - fixed_count_by_class.get(c, 0)
        if demand != free_slots:
            mismatches.append((c, demand, free_slots))
    return mismatches


def _is_scenario_feasible_with_fixed(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_subject: dict[str, str],
    exempt_subjects: set[str],
    teacher_unavailable: dict[str, set[str]],
    time_limit_sec: float,
) -> bool:
    req = _build_subject_requirements(non_tt)
    fixed: dict[tuple[str, str], str] = {}
    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        t = str(a.get("slot", "")).strip()
        s = str(a.get("subject", "")).strip()
        if c and t and s:
            fixed[(c, t)] = s

    auto_exempt = {a.get("subject", "") for a in fixed_raw if a.get("subject")}
    exempt = set(exempt_subjects) | auto_exempt

    class_subjects = {c: set() for c in classes}
    for (c, s), _ in req.items():
        class_subjects[c].add(s)
    for (c, _), s in fixed.items():
        class_subjects[c].add(s)

    model = cp_model.CpModel()
    x: dict[tuple[str, str, str], cp_model.IntVar] = {}

    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                x[(c, t, s)] = model.NewBoolVar(f"x_{c}_{t}_{s}")

    for c in classes:
        for t in slots:
            if (c, t) in fixed:
                fs = fixed[(c, t)]
                for s in class_subjects[c]:
                    model.Add(x[(c, t, s)] == (1 if s == fs else 0))
            else:
                model.Add(sum(x[(c, t, s)] for s in class_subjects[c]) == 1)

    for (c, s), h in req.items():
        fixed_count = sum(1 for (cc, _t), ss in fixed.items() if cc == c and ss == s)
        target = h - fixed_count
        if target < 0:
            return False
        non_fixed = [t for t in slots if (c, t) not in fixed]
        model.Add(sum(x[(c, t, s)] for t in non_fixed) == target)

    for c in classes:
        for s in class_subjects[c]:
            for d in DAY_ORDER:
                dslots = [t for t in slots if parse_slot(t)[0] == d and (c, t) not in fixed]
                if dslots:
                    model.Add(sum(x[(c, t, s)] for t in dslots) <= 1)

    by_grade: dict[str, list[str]] = {}
    for c in classes:
        by_grade.setdefault(grade_of(c), []).append(c)

    subjects = sorted({s for (_c, s) in req} | set(fixed.values()))
    for _g, g_classes in by_grade.items():
        for t in slots:
            for s in subjects:
                if s in exempt:
                    continue
                vars_auto = []
                for c in g_classes:
                    if (c, t) in fixed:
                        continue
                    if (c, t, s) in x:
                        vars_auto.append(x[(c, t, s)])
                if vars_auto:
                    model.Add(sum(vars_auto) <= 1)

    y: dict[tuple[str, str, str], cp_model.IntVar] = {}
    teachers_used = sorted({k for (k, _c, _s) in non_tt})
    for (k, c, _s), _ in non_tt.items():
        for t in slots:
            if (c, t) in fixed:
                continue
            y[(k, c, t)] = model.NewBoolVar(f"y_{k}_{c}_{t}")

    for c in classes:
        for t in slots:
            if (c, t) in fixed:
                continue
            cand = [k for k in teachers_used if (k, c, t) in y]
            if cand:
                model.Add(sum(y[(k, c, t)] for k in cand) == 1)

    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                teachers_for_subject = [k for k in teachers_used if teacher_subject.get(k) == s and (k, c, t) in y]
                if teachers_for_subject:
                    model.Add(x[(c, t, s)] == sum(y[(k, c, t)] for k in teachers_for_subject))

    for k in teachers_used:
        for t in slots:
            vars_t = [y[(k, c, t)] for c in classes if (k, c, t) in y]
            if vars_t:
                model.Add(sum(vars_t) <= 1)
            if t in teacher_unavailable.get(k, set()) and vars_t:
                model.Add(sum(vars_t) == 0)

    main_h: dict[tuple[str, str], int] = {}
    for (k, c, _s), h in non_tt.items():
        main_h[(k, c)] = main_h.get((k, c), 0) + int(h)
    for (k, c), h in main_h.items():
        vars_kc = [y[(k, c, t)] for t in slots if (k, c, t) in y]
        model.Add(sum(vars_kc) == h)

    # 可解診断ではTTは評価対象外（本体モデルでのみ最適化対象）。
    tt: dict[tuple[str, str, str], int] = {}

    # -----------------------------
    # TT VARIABLES / CONSTRAINTS
    # -----------------------------

    z={}

    for (k,c,s),h in tt.items():
        if h<=0:
            continue
        for t in slots:
            if (c,t,s) not in x:
                continue
            z[(k,c,t)] = model.NewBoolVar(f"z_{k}_{c}_{t}")
            # TT参加は当該クラスで自担当教科が実施されるコマに限定。
            model.Add(z[(k,c,t)] <= x[(c,t,s)])
            if t in teacher_unavailable.get(k,set()):
                model.Add(z[(k,c,t)] == 0)

    # 非TT(y) と TT(z) を同時に担当させない。
    all_teachers = sorted(set(teachers_used) | {k for (k,_c,_s) in tt})
    for k in all_teachers:
        for t in slots:
            vars_t=[]
            vars_t.extend(
                y[(k,c,t)]
                for c in classes
                if (k,c,t) in y
            )
            vars_t.extend(
                z[(k,c,t)]
                for c in classes
                if (k,c,t) in z
            )
            if vars_t:
                model.Add(sum(vars_t) <= 1)

    # TT時間は要求時間を上限とし、過剰に割り当てない。
    for (k,c,_s),h in tt.items():
        vars_kc=[
            z[(k,c,t)]
            for t in slots
            if (k,c,t) in z
        ]
        if vars_kc:
            model.Add(sum(vars_kc) <= int(h))

    # 可能な範囲でTT割当を最大化する（不可な場合は不足を許容）。
    if z:
        model.Maximize(sum(z.values()))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(3.0, float(time_limit_sec))
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    return status in (cp_model.OPTIMAL, cp_model.FEASIBLE)


def _suggest_relax_fixed_candidates(
    *,
    scenario_id: str,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_subject: dict[str, str],
    exempt_subjects: set[str],
    teacher_unavailable: dict[str, set[str]],
    time_limit_sec: float,
    max_candidates: int = 3,
    trial_limit: int = 30,
) -> list[str]:
    fixed_list = [
        {
            "class": str(a.get("class", "")).strip(),
            "slot": str(a.get("slot", "")).strip(),
            "subject": str(a.get("subject", "")).strip(),
        }
        for a in fixed_raw
        if str(a.get("class", "")).strip() and str(a.get("slot", "")).strip() and str(a.get("subject", "")).strip()
    ]
    if not fixed_list:
        return []

    # 固定入力順を尊重しつつ、診断時間の上限を設ける。
    candidates: list[str] = []
    for i, victim in enumerate(fixed_list[: max(1, int(trial_limit))]):
        test_fixed = [a for j, a in enumerate(fixed_list) if j != i]
        ok = _is_scenario_feasible_with_fixed(
            classes=classes,
            slots=slots,
            non_tt=non_tt,
            fixed_raw=test_fixed,
            teacher_subject=teacher_subject,
            exempt_subjects=exempt_subjects,
            teacher_unavailable=teacher_unavailable,
            time_limit_sec=max(5.0, min(20.0, float(time_limit_sec) / 2.0)),
        )
        if ok:
            candidates.append(f"{victim['class']} {victim['slot']} {victim['subject']} を一旦外す")
            if len(candidates) >= max(1, int(max_candidates)):
                break
    return candidates


def _diagnose_infeasible_fixed_core(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_subject: dict[str, str],
    exempt_subjects: set[str],
    teacher_unavailable: dict[str, set[str]],
    time_limit_sec: float,
    max_candidates: int = 3,
) -> list[str]:
    # assumption API がない環境では診断をスキップ。
    if not fixed_raw:
        return []
    if not hasattr(cp_model.CpModel, "AddAssumptions"):
        return []

    req = _build_subject_requirements(non_tt)
    fixed_list = [
        {
            "class": str(a.get("class", "")).strip(),
            "slot": str(a.get("slot", "")).strip(),
            "subject": str(a.get("subject", "")).strip(),
        }
        for a in fixed_raw
        if str(a.get("class", "")).strip() and str(a.get("slot", "")).strip() and str(a.get("subject", "")).strip()
    ]
    if not fixed_list:
        return []

    model = cp_model.CpModel()
    class_subjects = {c: set() for c in classes}
    for (c, s), _ in req.items():
        class_subjects[c].add(s)
    for a in fixed_list:
        class_subjects[a["class"]].add(a["subject"])

    x: dict[tuple[str, str, str], cp_model.IntVar] = {}
    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                x[(c, t, s)] = model.NewBoolVar(f"x_{c}_{t}_{s}")

    for c in classes:
        for t in slots:
            model.Add(sum(x[(c, t, s)] for s in class_subjects[c]) == 1)

    for (c, s), h in req.items():
        model.Add(sum(x[(c, t, s)] for t in slots) == h)

    for c in classes:
        for s in class_subjects[c]:
            for d in DAY_ORDER:
                dslots = [t for t in slots if parse_slot(t)[0] == d]
                if dslots:
                    model.Add(sum(x[(c, t, s)] for t in dslots) <= 1)

    by_grade: dict[str, list[str]] = {}
    for c in classes:
        by_grade.setdefault(grade_of(c), []).append(c)
    subjects = sorted({s for (_c, s) in req} | {a["subject"] for a in fixed_list})
    auto_exempt = {a["subject"] for a in fixed_list}
    exempt = set(exempt_subjects) | auto_exempt

    for _g, g_classes in by_grade.items():
        for t in slots:
            for s in subjects:
                if s in exempt:
                    continue
                vars_auto = [x[(c, t, s)] for c in g_classes if (c, t, s) in x]
                if vars_auto:
                    model.Add(sum(vars_auto) <= 1)

    y: dict[tuple[str, str, str], cp_model.IntVar] = {}
    teachers_used = sorted({k for (k, _c, _s) in non_tt})
    for (k, c, _s), _ in non_tt.items():
        for t in slots:
            y[(k, c, t)] = model.NewBoolVar(f"y_{k}_{c}_{t}")

    for c in classes:
        for t in slots:
            cand = [k for k in teachers_used if (k, c, t) in y]
            if cand:
                model.Add(sum(y[(k, c, t)] for k in cand) == 1)

    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                teachers_for_subject = [k for k in teachers_used if teacher_subject.get(k) == s and (k, c, t) in y]
                if teachers_for_subject:
                    model.Add(x[(c, t, s)] == sum(y[(k, c, t)] for k in teachers_for_subject))

    for k in teachers_used:
        for t in slots:
            vars_t = [y[(k, c, t)] for c in classes if (k, c, t) in y]
            if vars_t:
                model.Add(sum(vars_t) <= 1)
            if t in teacher_unavailable.get(k, set()) and vars_t:
                model.Add(sum(vars_t) == 0)

    assumptions: list[cp_model.IntVar] = []
    lit_to_fixed: dict[int, dict] = {}
    for i, a in enumerate(fixed_list):
        c = a["class"]
        t = a["slot"]
        s = a["subject"]
        if (c, t, s) not in x:
            continue
        lit = model.NewBoolVar(f"assume_fix_{i}")
        model.Add(x[(c, t, s)] == 1).OnlyEnforceIf(lit)
        assumptions.append(lit)
        lit_to_fixed[lit.Index()] = a

    if not assumptions:
        return []

    solver0 = cp_model.CpSolver()
    solver0.parameters.max_time_in_seconds = max(5.0, min(15.0, float(time_limit_sec)))
    solver0.parameters.num_search_workers = 8
    base_status = solver0.Solve(model)
    if base_status == cp_model.INFEASIBLE:
        return ["固定配置を外しても不可解（教員時数・不可コマ・同日重複などを確認）"]

    model.AddAssumptions(assumptions)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(8.0, min(30.0, float(time_limit_sec)))
    solver.parameters.num_search_workers = 8

    status = solver.Solve(model)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return []
    if status != cp_model.INFEASIBLE:
        return []

    core = list(solver.SufficientAssumptionsForInfeasibility())
    out: list[str] = []
    for lit_idx in core:
        fixed = lit_to_fixed.get(int(lit_idx))
        if not fixed:
            continue
        out.append(f"{fixed['class']} {fixed['slot']} {fixed['subject']} を見直す")
        if len(out) >= max(1, int(max_candidates)):
            break
    return out


def _diagnose_teacher_capacity_shortage(
    *,
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    teacher_unavailable: dict[str, set[str]],
    fixed_raw: list[dict],
    max_lines: int = 12,
) -> list[str]:
    teacher_required: dict[str, int] = {}
    teacher_class_required: dict[tuple[str, str], int] = {}
    for (teacher, cls, _subject), h in non_tt.items():
        teacher_required[teacher] = teacher_required.get(teacher, 0) + int(h)
        key = (teacher, cls)
        teacher_class_required[key] = teacher_class_required.get(key, 0) + int(h)

    # クラス固定コマはそのクラスの担当可能枠を減らす（y変数は fixed 上で作られないため）。
    fixed_by_class: dict[str, int] = {}
    for a in fixed_raw:
        cls = str(a.get("class", "")).strip()
        slot = str(a.get("slot", "")).strip()
        if cls and slot:
            fixed_by_class[cls] = fixed_by_class.get(cls, 0) + 1

    lines: list[str] = []
    total_slots = len(slots)
    for teacher, need in sorted(teacher_required.items()):
        unavail = len(teacher_unavailable.get(teacher, set()))
        avail = total_slots - unavail
        if avail < need:
            lines.append(
                f"教員不足: {teacher} 可用{avail} < 担当{need}（不可コマ{unavail}）"
            )

    for (teacher, cls), need in sorted(teacher_class_required.items()):
        unavail = len(teacher_unavailable.get(teacher, set()))
        fixed = fixed_by_class.get(cls, 0)
        # 教員不可とクラス固定は単純加算の上界評価（不足検知のため保守的に使用）。
        avail_for_class = total_slots - unavail - fixed
        if avail_for_class < need:
            lines.append(
                f"担当枠不足: {teacher}/{cls} 可用見積{avail_for_class} < 担当{need} "
                f"（不可{unavail}, 固定{fixed}）"
            )

    if not lines:
        return ["可用コマ不足は検出されませんでした（他制約の衝突の可能性）。"]
    return lines[: max(1, int(max_lines))]


def _diagnose_constraint_collision_candidates(
    *,
    classes: list[str],
    slots: list[str],
    req: dict[tuple[str, str], int],
    fixed_raw: list[dict],
    exempt_subjects: set[str],
    max_examples: int = 3,
) -> list[str]:
    fixed: dict[tuple[str, str], str] = {}
    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        t = str(a.get("slot", "")).strip()
        s = str(a.get("subject", "")).strip()
        if c and t and s:
            fixed[(c, t)] = s

    auto_exempt = {str(a.get("subject", "")).strip() for a in fixed_raw if str(a.get("subject", "")).strip()}
    exempt = set(exempt_subjects) | auto_exempt

    lines: list[str] = []

    # 同日重複: 1クラス同一教科は1日1回まで。
    daily_hits: list[tuple[str, str, int, int]] = []
    for c in classes:
        fixed_count_by_subj: dict[str, int] = {}
        fixed_days_by_subj: dict[str, set[str]] = {}
        for (cc, slot), subj in fixed.items():
            if cc != c:
                continue
            fixed_count_by_subj[subj] = fixed_count_by_subj.get(subj, 0) + 1
            day = parse_slot(slot)[0]
            fixed_days_by_subj.setdefault(subj, set()).add(day)

        for (cc, subj), h in req.items():
            if cc != c:
                continue
            target = int(h) - fixed_count_by_subj.get(subj, 0)
            if target <= 0:
                continue
            used_days = len(fixed_days_by_subj.get(subj, set()))
            available_days = max(0, len(DAY_ORDER) - used_days)
            if target > available_days:
                daily_hits.append((c, subj, target, available_days))

    if daily_hits:
        examples = ", ".join(
            [f"{c}/{s} 要{need}>日{cap}" for c, s, need, cap in daily_hits[: max(1, int(max_examples))]]
        )
        lines.append(f"同日重複候補: {len(daily_hits)}件（例: {examples}）")
    else:
        lines.append("同日重複候補: 0件")

    # 学年重複: 同学年同一教科は同一コマで重複不可（exempt除外）。
    grade_map: dict[str, list[str]] = {}
    for c in classes:
        grade_map.setdefault(grade_of(c), []).append(c)

    class_subjects: dict[str, set[str]] = {c: set() for c in classes}
    for (c, s), _h in req.items():
        class_subjects[c].add(s)
    for (c, _slot), s in fixed.items():
        class_subjects[c].add(s)

    grade_hits: list[tuple[str, str, int, int]] = []
    for g, g_classes in grade_map.items():
        all_subjects = sorted({s for c in g_classes for s in class_subjects.get(c, set())})
        for subj in all_subjects:
            if subj in exempt:
                continue

            need_total = 0
            for c in g_classes:
                h = int(req.get((c, subj), 0))
                if h <= 0:
                    continue
                fixed_count = sum(1 for (cc, _t), ss in fixed.items() if cc == c and ss == subj)
                need_total += max(0, h - fixed_count)

            if need_total <= 0:
                continue

            cap = 0
            for t in slots:
                exists = False
                for c in g_classes:
                    if (c, t) in fixed:
                        continue
                    if subj in class_subjects.get(c, set()):
                        exists = True
                        break
                if exists:
                    cap += 1

            if need_total > cap:
                grade_hits.append((g, subj, need_total, cap))

    if grade_hits:
        examples = ", ".join(
            [f"{g}年/{s} 要{need}>枠{cap}" for g, s, need, cap in grade_hits[: max(1, int(max_examples))]]
        )
        lines.append(f"学年重複候補: {len(grade_hits)}件（例: {examples}）")
    else:
        lines.append("学年重複候補: 0件")

    return lines


def _diagnose_teacher_slot_conflicts_top3(
    *,
    non_tt: dict[tuple[str, str, str], int],
    teacher_unavailable: dict[str, set[str]],
    fixed_raw: list[dict],
    max_items: int = 3,
) -> list[str]:
    fixed_slots = {
        (str(a.get("class", "")).strip(), str(a.get("slot", "")).strip())
        for a in fixed_raw
        if str(a.get("class", "")).strip() and str(a.get("slot", "")).strip()
    }

    teacher_classes: dict[str, set[str]] = {}
    for (teacher, cls, _subj), h in non_tt.items():
        if int(h) <= 0:
            continue
        teacher_classes.setdefault(teacher, set()).add(cls)

    hits: list[tuple[int, str, str, list[str]]] = []
    for teacher, slots_ng in teacher_unavailable.items():
        cls_set = teacher_classes.get(teacher, set())
        if not cls_set:
            continue
        for slot in sorted(slots_ng):
            candidates = [c for c in sorted(cls_set) if (c, slot) not in fixed_slots]
            if not candidates:
                continue
            hits.append((len(candidates), teacher, slot, candidates))

    if not hits:
        return ["不可コマx同時担当候補: 0件"]

    hits.sort(key=lambda x: (-x[0], x[1], x[2]))
    out: list[str] = []
    for cnt, teacher, slot, classes in hits[: max(1, int(max_items))]:
        out.append(
            f"不可コマx同時担当候補: {teacher} {slot} で{cnt}クラス候補（{', '.join(classes[:6])}）"
        )
    return out


def _diagnose_subject_shortage_from_req_fixed(
    *,
    req: dict[tuple[str, str], int],
    fixed_raw: list[dict],
    max_lines: int = 8,
) -> list[str]:
    fixed_count: dict[tuple[str, str], int] = {}
    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        s = str(a.get("subject", "")).strip()
        if c and s:
            key = (c, s)
            fixed_count[key] = fixed_count.get(key, 0) + 1

    by_subject: dict[str, tuple[int, int, int]] = {}
    for (c, s), need in req.items():
        req_h = int(need)
        fixed_h = fixed_count.get((c, s), 0)
        remain = max(0, req_h - fixed_h)
        old_req, old_fixed, old_remain = by_subject.get(s, (0, 0, 0))
        by_subject[s] = (old_req + req_h, old_fixed + fixed_h, old_remain + remain)

    shortage = [
        (s, req_h, fixed_h, remain_h)
        for s, (req_h, fixed_h, remain_h) in by_subject.items()
        if remain_h > 0
    ]
    shortage.sort(key=lambda x: (-x[3], x[0]))

    if not shortage:
        return ["教科別不足内訳(REQ-固定): 不足0"]

    out: list[str] = []
    for s, req_h, fixed_h, remain_h in shortage[: max(1, int(max_lines))]:
        out.append(f"教科別不足: {s} 残{remain_h} (REQ{req_h} / 固定{fixed_h})")
    return out


def _diagnose_structural_conflicts(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_unavailable: dict[str, set[str]],
    exempt_subjects: set[str],
    max_lines: int = 12,
) -> list[str]:
    lines: list[str] = []

    fixed_by_cell: dict[tuple[str, str], list[str]] = {}
    fixed_by_class_subject_day: dict[tuple[str, str, str], int] = {}
    fixed_grade_slot_subject: dict[tuple[str, str, str], int] = {}

    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        t = str(a.get("slot", "")).strip()
        s = str(a.get("subject", "")).strip()
        if not c or not t or not s:
            continue
        fixed_by_cell.setdefault((c, t), []).append(s)
        if "-" in t:
            d = parse_slot(t)[0]
            fixed_by_class_subject_day[(c, s, d)] = fixed_by_class_subject_day.get((c, s, d), 0) + 1
        if s not in exempt_subjects:
            g = grade_of(c)
            fixed_grade_slot_subject[(g, t, s)] = fixed_grade_slot_subject.get((g, t, s), 0) + 1

    cell_conflicts = [(c, t, sorted(set(subjs))) for (c, t), subjs in fixed_by_cell.items() if len(set(subjs)) > 1]
    for c, t, subjs in cell_conflicts[:3]:
        lines.append(f"固定衝突: {c} {t} に複数教科固定 ({'/'.join(subjs)})")

    daily_conflicts = [(c, s, d, cnt) for (c, s, d), cnt in fixed_by_class_subject_day.items() if cnt >= 2]
    for c, s, d, cnt in sorted(daily_conflicts)[:3]:
        lines.append(f"同日重複確定: {c} {d} {s} が{cnt}コマ固定")

    grade_conflicts = [(g, t, s, cnt) for (g, t, s), cnt in fixed_grade_slot_subject.items() if cnt >= 2]
    for g, t, s, cnt in sorted(grade_conflicts)[:3]:
        lines.append(f"学年同時重複確定: {g}年 {t} {s} が{cnt}クラス固定")

    fixed_cells = set(fixed_by_cell.keys())
    req_teacher_class: dict[tuple[str, str], int] = {}
    for (teacher, c, _s), h in non_tt.items():
        req_teacher_class[(teacher, c)] = req_teacher_class.get((teacher, c), 0) + int(h)

    cap_shortages: list[tuple[str, str, int, int]] = []
    for (teacher, c), need in sorted(req_teacher_class.items()):
        cap = 0
        for t in slots:
            if (c, t) in fixed_cells:
                continue
            if t in teacher_unavailable.get(teacher, set()):
                continue
            cap += 1
        if cap < need:
            cap_shortages.append((teacher, c, cap, need))

    for teacher, c, cap, need in cap_shortages[:3]:
        lines.append(f"教師担当枠不足: {teacher}/{c} 可用{cap} < 必要{need}")

    if not lines:
        return ["構造矛盾チェック: 直接矛盾は未検出"]
    return lines[: max(1, int(max_lines))]


def solve_all_scenarios_cp(
        config_data:dict,
        exempt_subjects:set[str],
        progress_callback:Optional[Callable[[int,str],None]]=None,
        time_limit_sec:float=60.0
)->dict:

    log_file = open("debug_log.txt", "w", encoding='utf-8')

    log_file.write("\n==============================\n")
    log_file.write("CP-SAT TIMETABLE DEBUG START\n")
    log_file.write("==============================\n")

    classes=list(config_data["classes"])
    teachers=list(config_data["teachers"])
    scenarios=list(config_data["scenarios"])
    day_periods=config_data["day_periods"]

    slots=build_slots(day_periods)
    excluded_subject_keywords = _build_excluded_subject_keywords(set(exempt_subjects))

    log_file.write("CLASSES:" + str(classes) + "\n")
    log_file.write("SLOTS:" + str(len(slots)) + "\n")
    log_file.write("TEACHERS:" + str(len(teachers)) + "\n")
    log_file.write("SCENARIOS:" + str(len(scenarios)) + "\n")

    # -----------------------------
    # teacher demand
    # -----------------------------

    teacher_subject={}
    teacher_unavailable={}
    non_tt={}
    tt={}

    for t in teachers:

        name=t["name"]
        subject=t["subjects"][0]
        excluded_subject = _is_excluded_subject(subject, excluded_subject_keywords)

        teacher_subject[name]=subject
        teacher_unavailable[name]=set(t.get("unavailable_slots",[]))

        for ca in t["class_assignments"]:

            cls=ca["class"]
            hours=int(ca["hours"])

            if excluded_subject:
                continue

            if ca.get("tt",False):

                key=(name,cls,subject)

                tt[key]=tt.get(key,0)+hours

            else:

                key=(name,cls,subject)

                non_tt[key]=non_tt.get(key,0)+hours

    # -----------------------------
    # class subject hours
    # -----------------------------

    req={}

    for (_k,c,s),h in non_tt.items():
        req[(c,s)]=req.get((c,s),0)+h

    log_file.write("\nCLASS SUBJECT HOURS\n")

    for k,v in sorted(req.items()):
        log_file.write(str(k) + " " + str(v) + "\n")

    solved={}

    for sc in scenarios:

        sid=sc["id"]

        log_file.write("\n==============================\n")
        log_file.write("SCENARIO:" + str(sid) + "\n")
        log_file.write("==============================\n")

        fixed_raw=sc.get("fixed_assignments",[])
        fixed={}

        for a in fixed_raw:
            fixed[(a["class"],a["slot"])]=a["subject"]

        log_file.write("FIXED LESSONS:" + str(len(fixed)) + "\n")

        # 固定入力後の空きコマ数と、TTを除く担当コマ数がクラスごとに一致するかを先に検証。
        mismatches = _find_class_open_slot_mismatches(
            classes=classes,
            slots=slots,
            non_tt=non_tt,
            fixed=fixed,
        )
        if mismatches:
            log_file.write("\nCLASS FREE SLOT CHECK (NG)\n")
            for c, demand, free in mismatches:
                log_file.write(
                    "MISMATCH " + str(c) + " demand=" + str(demand) + " free=" + str(free) + "\n"
                )
            log_file.close()
            lines = [
                f"{sid}: クラス別コマ数の整合チェックで不一致があります。",
                "教員タブのTTを含まない担当コマ数と、固定入力後の空きコマ数を一致させてください。",
            ]
            lines.extend(
                [f"- {c}において、教員タブでの授業の指定数は{demand}ですが、空きコマが{free}あります。" for c, demand, free in mismatches]
            )
            raise SchedulerError("\n".join(lines))

        # -----------------------------
        # auto exempt
        # -----------------------------

        auto_exempt={a["subject"] for a in fixed_raw}
        exempt=set(exempt_subjects)|auto_exempt

        log_file.write("EXEMPT SUBJECTS:" + str(sorted(exempt)) + "\n")

        # -----------------------------
        # subject set
        # -----------------------------

        subj_set={s for (_c,s) in req}
        subj_set.update(fixed.values())

        subjects=sorted(subj_set)

        # -----------------------------
        # class subjects
        # -----------------------------

        class_subjects={c:set() for c in classes}

        for (c,s),_ in req.items():
            class_subjects[c].add(s)

        for (c,_),s in fixed.items():
            class_subjects[c].add(s)

        # -----------------------------
        # fixed vs req debug
        # -----------------------------

        log_file.write("\nREQ vs FIXED CHECK\n")

        for (c,s),h in req.items():

            fixed_count=sum(
                1 for (cc,t),ss in fixed.items()
                if cc==c and ss==s
            )

            if fixed_count>h:

                log_file.write("ERROR FIXED>REQ " + str(c) + " " + str(s) + " " + str(fixed_count) + " " + str(h) + "\n")

        # -----------------------------
        # MODEL
        # -----------------------------

        model=cp_model.CpModel()

        # -----------------------------
        # VARIABLES
        # -----------------------------

        x={}

        for c in classes:
            for t in slots:
                for s in class_subjects[c]:
                    x[(c,t,s)]=model.NewBoolVar(f"x_{c}_{t}_{s}")

        # -----------------------------
        # SLOT FILL
        # -----------------------------

        for c in classes:
            for t in slots:

                if (c,t) in fixed:

                    fs=fixed[(c,t)]

                    for s in class_subjects[c]:

                        model.Add(
                            x[(c,t,s)]==(1 if s==fs else 0)
                        )

                else:

                    model.Add(
                        sum(
                            x[(c,t,s)]
                            for s in class_subjects[c]
                        )==1
                    )

        # -----------------------------
        # WEEKLY HOURS
        # -----------------------------

        for (c,s),h in req.items():

            fixed_count=sum(
                1 for (cc,t),ss in fixed.items()
                if cc==c and ss==s
            )

            target=h-fixed_count

            if target<0:

                raise SchedulerError(
                    f"FIXED>{c} {s}"
                )

            non_fixed=[
                t for t in slots
                if (c,t) not in fixed
            ]

            model.Add(
                sum(x[(c,t,s)] for t in non_fixed)==target
            )

        # -----------------------------
        # DAILY SUBJECT
        # -----------------------------

        for c in classes:
            for s in class_subjects[c]:

                for d in DAY_ORDER:

                    dslots=[
                        t for t in slots
                        if parse_slot(t)[0]==d
                           and (c,t) not in fixed
                    ]

                    if dslots:

                        model.Add(
                            sum(x[(c,t,s)] for t in dslots)<=1
                        )

        # -----------------------------
        # GRADE UNIQUE
        # -----------------------------

        by_grade={}

        for c in classes:
            by_grade.setdefault(grade_of(c),[]).append(c)

        for g,g_classes in by_grade.items():
            for t in slots:
                for s in subjects:

                    if s in exempt:
                        continue

                    vars_auto=[]

                    for c in g_classes:

                        if (c,t) in fixed:
                            continue

                        if (c,t,s) in x:
                            vars_auto.append(x[(c,t,s)])

                    if vars_auto:

                        model.Add(sum(vars_auto)<=1)

        # -----------------------------
        # TEACHER VARIABLES
        # -----------------------------

        y={}
        teachers_used=sorted({k for (k,_c,_s) in non_tt})

        for (k,c,s),_ in non_tt.items():

            for t in slots:

                if (c,t) in fixed:
                    continue

                y[(k,c,t)]=model.NewBoolVar(
                    f"y_{k}_{c}_{t}"
                )

        # -----------------------------
        # TEACHER ASSIGN
        # -----------------------------

        for c in classes:
            for t in slots:

                if (c,t) in fixed:
                    continue

                cand=[
                    k for k in teachers_used
                    if (k,c,t) in y
                ]

                if cand:

                    model.Add(
                        sum(y[(k,c,t)] for k in cand)==1
                    )

        # -----------------------------
        # TEACHER LINK
        # -----------------------------

        for c in classes:
            for t in slots:

                if (c,t) in fixed:
                    continue

                for s in class_subjects[c]:

                    teachers_for_subject = [
                        k for k in teachers_used
                        if teacher_subject[k] == s and (k,c,t) in y
                    ]

                    if teachers_for_subject:

                        model.Add(
                            x[(c,t,s)] ==
                            sum(y[(k,c,t)] for k in teachers_for_subject)
                        )

        # -----------------------------
        # TEACHER OVERLAP
        # -----------------------------

        for k in teachers_used:
            for t in slots:

                vars_t=[
                    y[(k,c,t)]
                    for c in classes
                    if (k,c,t) in y
                ]

                if vars_t:

                    model.Add(sum(vars_t)<=1)

        # -----------------------------
        # TEACHER HOURS
        # -----------------------------

        main_h={}

        for (k,c,_s),h in non_tt.items():
            main_h[(k,c)]=main_h.get((k,c),0)+h

        log_file.write("\nTEACHER HOURS DEMAND\n")

        for k,v in main_h.items():
            log_file.write(str(k) + " " + str(v) + "\n")

        log_file.write("\nTT HOURS DEMAND\n")
        for k,v in sorted(tt.items()):
            log_file.write(str(k) + " " + str(v) + "\n")

        for (k,c),h in main_h.items():

            vars_kc=[
                y[(k,c,t)]
                for t in slots
                if (k,c,t) in y
            ]

            model.Add(sum(vars_kc)==h)

        # -----------------------------
        # TT VARIABLES / CONSTRAINTS
        # -----------------------------

        z={}

        for (k,c,s),h in tt.items():
            if h<=0:
                continue
            for t in slots:
                if (c,t,s) not in x:
                    continue
                z[(k,c,t)] = model.NewBoolVar(f"z_{k}_{c}_{t}")
                # TT参加は当該クラスで自担当教科が実施されるコマに限定。
                model.Add(z[(k,c,t)] <= x[(c,t,s)])
                if t in teacher_unavailable.get(k,set()):
                    model.Add(z[(k,c,t)] == 0)

        # 非TT(y) と TT(z) を同時に担当させない。
        all_teachers = sorted(set(teachers_used) | {k for (k,_c,_s) in tt})
        for k in all_teachers:
            for t in slots:
                vars_t=[]
                vars_t.extend(
                    y[(k,c,t)]
                    for c in classes
                    if (k,c,t) in y
                )
                vars_t.extend(
                    z[(k,c,t)]
                    for c in classes
                    if (k,c,t) in z
                )
                if vars_t:
                    model.Add(sum(vars_t) <= 1)

        # TT時間は要求時間を上限とし、過剰に割り当てない。
        for (k,c,_s),h in tt.items():
            vars_kc=[
                z[(k,c,t)]
                for t in slots
                if (k,c,t) in z
            ]
            if vars_kc:
                model.Add(sum(vars_kc) <= int(h))

        # 可能な範囲でTT割当を最大化する（不可な場合は不足を許容）。
        if z:
            model.Maximize(sum(z.values()))

        # -----------------------------
        # SOLVER
        # -----------------------------

        solver=cp_model.CpSolver()

        solver.parameters.max_time_in_seconds=time_limit_sec
        solver.parameters.num_search_workers=8
        solver.parameters.log_search_progress=True
        solver.parameters.cp_model_presolve=True

        log_file.write("\nMODEL STATS\n")
        log_file.write(model.ModelStats() + "\n")

        log_file.write("\nSTART SOLVING...\n\n")

        status=solver.Solve(model)

        log_file.write("SOLVER STATUS:" + solver.StatusName(status) + "\n")

        if status not in (cp_model.OPTIMAL,cp_model.FEASIBLE):
            log_file.write("\nDEBUG: MODEL UNSAT\n")
            diagnostics = _diagnose_structural_conflicts(
                classes=classes,
                slots=slots,
                non_tt=non_tt,
                fixed_raw=fixed_raw,
                teacher_unavailable=teacher_unavailable,
                exempt_subjects=exempt_subjects,
                max_lines=12,
            )
            diagnostics.extend(
                _diagnose_teacher_capacity_shortage(
                    slots=slots,
                    non_tt=non_tt,
                    teacher_unavailable=teacher_unavailable,
                    fixed_raw=fixed_raw,
                    max_lines=12,
                )
            )
            diagnostics.extend(
                _diagnose_teacher_slot_conflicts_top3(
                    non_tt=non_tt,
                    teacher_unavailable=teacher_unavailable,
                    fixed_raw=fixed_raw,
                    max_items=3,
                )
            )
            diagnostics.extend(
                _diagnose_subject_shortage_from_req_fixed(
                    req=req,
                    fixed_raw=fixed_raw,
                    max_lines=8,
                )
            )
            diagnostics.extend(
                _diagnose_constraint_collision_candidates(
                    classes=classes,
                    slots=slots,
                    req=req,
                    fixed_raw=fixed_raw,
                    exempt_subjects=exempt_subjects,
                    max_examples=3,
                )
            )
            log_file.write("\nAUTO DIAGNOSTICS\n")
            for dline in diagnostics:
                log_file.write("- " + dline + "\n")
            base_ok = _is_scenario_feasible_with_fixed(
                classes=classes,
                slots=slots,
                non_tt=non_tt,
                fixed_raw=[],
                teacher_subject=teacher_subject,
                exempt_subjects=exempt_subjects,
                teacher_unavailable=teacher_unavailable,
                time_limit_sec=max(5.0, min(20.0, float(time_limit_sec) / 2.0)),
            )
            if not base_ok:
                log_file.close()
                lines = [
                    f"{sid}: 固定配置を全解除しても解が見つかりません。",
                    "教員時数・不可コマ・同日重複制約のいずれかが矛盾しています。",
                    "",
                    "自動診断:",
                ]
                lines.extend([f"- {d}" for d in diagnostics])
                raise SchedulerError("\n".join(lines))
            # 解なし時は、固定配置を1件ずつ緩和する最小診断を実施して候補を返す。
            suggestions = _suggest_relax_fixed_candidates(
                scenario_id=sid,
                classes=classes,
                slots=slots,
                non_tt=non_tt,
                fixed_raw=fixed_raw,
                teacher_subject=teacher_subject,
                exempt_subjects=exempt_subjects,
                teacher_unavailable=teacher_unavailable,
                time_limit_sec=time_limit_sec,
                max_candidates=3,
                trial_limit=30,
            )
            if not suggestions:
                suggestions = _diagnose_infeasible_fixed_core(
                    classes=classes,
                    slots=slots,
                    non_tt=non_tt,
                    fixed_raw=fixed_raw,
                    teacher_subject=teacher_subject,
                    exempt_subjects=exempt_subjects,
                    teacher_unavailable=teacher_unavailable,
                    time_limit_sec=time_limit_sec,
                    max_candidates=3,
                )
            log_file.close()
            if suggestions:
                lines = [
                    f"{sid}: 解が見つかりません",
                    "固定配置の異動候補（まずは1件のみ変更）:",
                ]
                lines.extend([f"- {s}" for s in suggestions])
                lines.extend(["", "自動診断:"])
                lines.extend([f"- {d}" for d in diagnostics])
                raise SchedulerError("\n".join(lines))
            lines = [
                f"{sid}: 解が見つかりません（固定配置1件の緩和では可解化できません）",
                "",
                "自動診断:",
            ]
            lines.extend([f"- {d}" for d in diagnostics])
            raise SchedulerError("\n".join(lines))

        log_file.write("SOLUTION FOUND\n")

        # デバッグ: 変数値の確認
        log_file.write("\nDEBUG: VARIABLE VALUES\n")
        for (c, t, s), var in list(x.items())[:10]:  # 最初の10個を確認
            val = solver.Value(var)
            log_file.write(f"x[{c},{t},{s}] = {val} (type: {type(val)})\n")

        # 教科割り当ての抽出
        assignments = {}
        for (c, t, s), var in x.items():
            val = solver.Value(var)
            if val == 1 or val >= 0.99:  # 浮動小数点対応
                assignments[(c, t)] = s
                log_file.write(f"ASSIGN: {c} {t} {s}\n")

        # 教師割り当ての抽出
        teacher_assignments = {}
        for (k, c, t), var in y.items():
            val = solver.Value(var)
            if val == 1 or val >= 0.99:
                teacher_assignments[(c, t)] = k
                log_file.write(f"TEACHER: {c} {t} {k}\n")

        # TT割り当ての抽出
        tt_assignments = {}
        for (k, c, t), var in z.items():
            val = solver.Value(var)
            if val == 1 or val >= 0.99:
                tt_assignments.setdefault((c, t), []).append(k)
                log_file.write(f"TT: {c} {t} {k}\n")

        log_file.write(f"TOTAL ASSIGNMENTS: {len(assignments)}\n")


        solved[sid] = {
            "assignments": assignments,
            "teacher_assignments": teacher_assignments,
            "tt_assignments": tt_assignments,
        }


    log_file.close()
    return solved

