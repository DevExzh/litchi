# Independent pre-measurement protocol review

The schedule-only repair preserves every non-native field and all qualification
and Heaptrack rows. The native row multiset is unchanged, while each of the
32 cases now has six matched blocks with three before-first and three
after-first pairs. Counts are 64 qualification reports, 384 native reports
containing 11,520 samples, and 64 one-sample Heaptrack reports: 512 reports /
11,648 samples. Two observer compatibility preflights remain separate.

Capture, qualification and analysis use `measurement-plan.json`; immutable
build freezes retain the original `plan.json`. Admission checks their exact
row multiset and non-native equality, then freezes both plans and the repair
record. Qualification rows and actual command lines are unchanged. Every
report is bound to a retained build/binary/command, and analysis checks exact
native/heap order and excludes overlapping Cargo work from those populations.

A packet-local Python cache found during review was removed before admission.
The admission and closure scripts reject any remaining `__pycache__` directory.
No production or measurement workload was changed by this review.

The two capture-qualification manifests intentionally retain the original
plan identity used when those captures ran. They are pre-repair evidence, not
rewritten manifests. Admission verifies their original plan hash and exact
receipt lists; original/final qualification rows are identical. Formal native
and heap capture manifests must instead bind the final measurement plan, and
the analyzer verifies their exact receipt lists and order.
