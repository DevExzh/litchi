# Addendum: truthful Callgrind repair and replay

The original profile capture used the admitted 0844 reader. Three profile
rows completed and the fourth produced valid raw Callgrind files and a zero
workload exit, but the parent `profile-capture` command returned exit 1 while
the old reader parsed the fourth trace. The failure is retained at its exact
receipt, log, and source line. No `capture-profile.json` is synthesized from
that failed parent command.

The Callgrind grammar permits `fe` to switch the current source file and
`cfl` as the alias of the callee-file record. `reader_v2.py` adds those two
bounded cases while retaining function identity, owner, terminal-dump, and
cost-conservation checks. `reader_v2_tests.py` reruns the inherited archive
checks, exercises the eight grammar regressions, and validates the four raw
profile traces before any reduction. The installed Valgrind format reference
and the repair scope remain bound in `reader-repair-spec.json`.

`profile_resume.py` records the four-row prefix in
`profile-resume-freeze.json`, preserves its raw files, and records the
post-replay custody in `profile-resume.json`. The v2 analysis keeps raw
Callgrind descriptors, owner totals, and concrete inclusive edge rows. It is
diagnostic instruction-count evidence inside `profile_publish`; it makes no
native-cycle, elapsed-phase, allocation, or pooled-main-matrix claim. The
16-row profile population remains separate from the 512-report native and
Heaptrack result.

`close_v2.py verify` is read-only. It accepts exactly the bound original
`profile-capture` exit 1 and requires every other recorded command to exit 0.
It validates the repaired reader, freeze, replay manifest, v2 analysis, main
analysis, tables, crosscheck, quality receipts, and owner-test totals without
capturing workloads, running cleanup, or rewriting an admitted file.

The explicit `close_v2.py clean` and `seal` modes retain the original owned-root marker, binary witness, closure equality, exact file inventory and post-commit parent/path guards. Full profile analysis and the resume manifest are independently reconstructed during verification.

The first full closure rejected relative paths in the draft decision descriptors. `decision-draft.json` and `closure-draft-repair.md` retain that bookkeeping failure; the final decision uses absolute paths with identical sizes and hashes. No workload was rerun.
