# 0424 evidence-script review

Review scope: static read-only review of `measure.py`, `summarize.py`,
`portable-replay.py`, `profile-candidate.py`, and `analyze.py`, against the
current 0424 bundle. No Cargo, workload, or profiler command was run for this
review.

## Blocking finding and disposition

At the initial review snapshot, `portable-replay.py:404` and
`portable-replay.py:406` passed an undefined `bundle` name to
`copy_pinned_tools` and `modified_tool_guard`. A portable replay would have
raised `NameError` before writing its receipt. The current frozen edit now
passes the copied `export` at both call sites, so the guard mutates and restores
only the temporary export's pinned tool and the replay commands use that same
export. This closes the blocker statically; a portable replay receipt still
needs to be retained as the execution proof.

## Binding review

The matched measurement path binds both roles to distinct full source
identities and copied binary hashes. `measure.py` rechecks source files and
binary hashes before, during, and after each child, and `summarize.py` checks
all 16 journal/artifact/verifier chains. The control receipt is explicitly
required to be an exact prior clean build and its origin receipt, source
identity, binary hashes, and historical protocol hash are cross-checked in
`summarize.py:410-437`. The candidate receipt is bound to the current
measurement protocol in `summarize.py:393-395`.

The new candidate profile has the expected isolated layout: its two captures
record the candidate revision and the candidate-profile protocol hash, while
the normal candidate build receipt remains at the bundle root. During analysis,
`analyze.py:108-126` falls back to that root
`measurement-build-candidate.json`; this works for the complete exported
bundle, and `binding_for_capture` still checks the capture build hash, source
revision, candidate ancestry, report binary, and report/catalog hashes. A
candidate-profile directory copied by itself is therefore not a self-contained
replay unit; publication must retain the bundle-root candidate receipt (as the
current portable export does).

The profile provenance hardening is now present: `profile-candidate.py` only
accepts a byte-identical copy of the frozen root protocol and adds the three
authorized role/hash fields, while `analyze.py` rechecks those fields,
common flags, scope, and the parent hash during replay. The current
`candidate-profile/protocol.json` satisfies that binding.

No additional measured-data or source-custody blocker was found in the other
four scripts. The portable replay NameError is resolved in the current edit;
the retained portable replay receipt is the remaining execution check.

## Final execution disposition

Root's retained standalone export replay now passes all 16 report validations,
four control/candidate trace replays and 104 report mutation probes, with the
modified pinned tool rejected. `portable-replay.json` and
`standalone-replay.json` are the execution receipts. The later candidate
report/catalog base-path defect was reproduced, retained and corrected before
this pass; root-only protocol/build and role-specific report paths stay distinct.
Original worktrees and binaries were absent for this replay.
