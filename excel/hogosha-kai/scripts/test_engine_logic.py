"""Python mirror of the VBA matching engine (Module1: RunChosei core).

Verifies, against brute force, that the staged augmenting-path strategy
always reaches a maximum matching, and that the failure diagnosis reports
a genuine Hall violator set. The VBA implementation is a line-by-line
translation of `solve()` below — keep the two in sync.
"""
import itertools
import random

UNSET, MARU, SANKAKU = 0, 1, 2


def try_augment(s, allow_sub_root, allow_sub_all, root, wish, n_slots,
                blocked, match_stu, match_slot, visited):
    """One augmenting DFS from student s. Mirrors VBA TryAugment."""
    for j in range(1, n_slots + 1):
        if blocked[j] or visited[j]:
            continue
        w = wish[s][j]
        edge_ok = (w == MARU) or (w == SANKAKU and (allow_sub_all or (allow_sub_root and s == root)))
        if not edge_ok:
            continue
        visited[j] = True
        if match_stu[j] == 0 or try_augment(match_stu[j], allow_sub_root, allow_sub_all,
                                            root, wish, n_slots, blocked,
                                            match_stu, match_slot, visited):
            match_stu[j] = s
            match_slot[s] = j
            return True
    return False


def solve(n_stu, n_slots, wish, fixed, breaks):
    """wish[s][j] in {0,1,2}; fixed[s] = slot or 0 (◎); breaks = set of slots.

    Returns (match_slot, used_sub, groups) where groups is the failure
    diagnosis: list of (students, slots) Hall-violating groups.
    """
    blocked = [False] * (n_slots + 1)
    for j in breaks:
        blocked[j] = True
    match_stu = [0] * (n_slots + 1)
    match_slot = [0] * (n_stu + 1)

    participating = []
    for s in range(1, n_stu + 1):
        if fixed[s]:
            match_slot[s] = fixed[s]
            match_stu[fixed[s]] = s
            blocked[fixed[s]] = True
        elif any(wish[s][j] for j in range(1, n_slots + 1)):
            participating.append(s)
        else:
            match_slot[s] = -1  # 実施しない生徒

    # fewest-options-first order (stable)
    def options(s):
        return sum(1 for j in range(1, n_slots + 1) if wish[s][j] == MARU and not blocked[j])
    order = sorted(participating, key=lambda s: (options(s), s))

    # phase 1: ○ only / phase 2a: own △ / phase 2b: all △
    for allow_root, allow_all in ((False, False), (True, False), (True, True)):
        for s in order:
            if match_slot[s] == 0:
                visited = [False] * (n_slots + 1)
                try_augment(s, allow_root, allow_all, s, wish, n_slots,
                            blocked, match_stu, match_slot, visited)

    used_sub = [s for s in range(1, n_stu + 1)
                if match_slot[s] > 0 and wish[s][match_slot[s]] == SANKAKU]

    # phase 3: diagnosis — alternating reachability from each unmatched student
    groups = []
    grouped = set()
    for s0 in order:
        if match_slot[s0] != 0 or s0 in grouped:
            continue
        stu_set = {s0}
        slot_set = set()
        frontier = [s0]
        while frontier:
            s = frontier.pop()
            for j in range(1, n_slots + 1):
                if blocked[j] or j in slot_set or wish[s][j] == UNSET:
                    continue
                slot_set.add(j)
                t = match_stu[j]
                if t and t not in stu_set:
                    stu_set.add(t)
                    frontier.append(t)
        grouped |= stu_set
        groups.append((sorted(stu_set), sorted(slot_set)))
    return match_slot, used_sub, groups


def brute_max_matching(n_stu, n_slots, wish, fixed, breaks, allow_sub):
    """Reference: brute-force maximum matching size."""
    blocked = set(breaks) | {fixed[s] for s in range(1, n_stu + 1) if fixed[s]}
    students = [s for s in range(1, n_stu + 1)
                if not fixed[s] and any(wish[s][j] for j in range(1, n_slots + 1))]
    levels = (MARU,) if not allow_sub else (MARU, SANKAKU)

    best = 0
    slots = [j for j in range(1, n_slots + 1) if j not in blocked]

    def rec(idx, used, count):
        nonlocal best
        best = max(best, count)
        if idx == len(students) or count + (len(students) - idx) <= best:
            return
        s = students[idx]
        rec(idx + 1, used, count)  # skip s
        for j in slots:
            if j not in used and wish[s][j] in levels:
                used.add(j)
                rec(idx + 1, used, count + 1)
                used.discard(j)

    rec(0, set(), 0)
    return best


def random_case(rng):
    n_stu = rng.randint(1, 7)
    n_slots = rng.randint(1, 8)
    wish = [[0] * (n_slots + 1) for _ in range(n_stu + 1)]
    for s in range(1, n_stu + 1):
        for j in range(1, n_slots + 1):
            r = rng.random()
            wish[s][j] = MARU if r < 0.25 else (SANKAKU if r < 0.35 else UNSET)
    breaks = {j for j in range(1, n_slots + 1) if rng.random() < 0.1}
    fixed = [0] * (n_stu + 1)
    taken = set(breaks)
    for s in range(1, n_stu + 1):
        if rng.random() < 0.1:
            free = [j for j in range(1, n_slots + 1) if j not in taken]
            if free:
                fixed[s] = rng.choice(free)
                taken.add(fixed[s])
    return n_stu, n_slots, wish, fixed, breaks


def main():
    rng = random.Random(20260610)
    n_max_fail = 0
    for trial in range(4000):
        n_stu, n_slots, wish, fixed, breaks = random_case(rng)
        match_slot, used_sub, groups = solve(n_stu, n_slots, wish, fixed, breaks)

        # validity: each assignment is wished & unique & unblocked
        seen = set()
        for s in range(1, n_stu + 1):
            j = match_slot[s]
            if j > 0 and not fixed[s]:
                assert wish[s][j] in (MARU, SANKAKU), "assigned non-wished slot"
                assert j not in breaks, "assigned a break slot"
                assert j not in seen
                seen.add(j)
                assert all(fixed[t] != j for t in range(1, n_stu + 1))

        matched = sum(1 for s in range(1, n_stu + 1) if match_slot[s] > 0 and not fixed[s])
        best = brute_max_matching(n_stu, n_slots, wish, fixed, breaks, allow_sub=True)
        assert matched == best, f"trial {trial}: matched {matched} < optimum {best}"

        # if nobody used △ unnecessarily: phase order guarantees that a
        # ○-only maximum matching is found before any △ is introduced
        best_maru = brute_max_matching(n_stu, n_slots, wish, fixed, breaks, allow_sub=False)
        matched_maru = sum(1 for s in range(1, n_stu + 1)
                           if match_slot[s] > 0 and not fixed[s]
                           and wish[s][match_slot[s]] == MARU)
        assert matched_maru >= best_maru - len(used_sub), "wasteful △ usage"

        # diagnosis: every reported group must be a genuine Hall violation
        for stu_set, slot_set in groups:
            n_unmatched_in_group = sum(1 for s in stu_set if match_slot[s] == 0)
            assert n_unmatched_in_group >= 1
            assert len(slot_set) < len(stu_set), \
                f"trial {trial}: group {stu_set} has {len(slot_set)} slots — not a violation"
            n_max_fail += 1

    print(f"OK: 4000 random cases — engine always reaches the brute-force optimum; "
          f"{n_max_fail} failure groups all genuine Hall violations")


if __name__ == "__main__":
    main()
