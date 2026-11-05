# 0362ul152a — signed archive verification memory return

After the live observations in [152](0362ul152-bound-policy-test-integration.md),
the signed v4 generation service fails at 22:12:18 UTC on 2026-10-25.
Systemd reports signal 9, a 4,294,967,296-byte memory peak and a kernel OOM
kill in the unit. No generation receipt or source assembly is written.
This is a failed source qualification, not a passing archive capture.

Root preserves the terminal service output and journal in the external
`signed-native-Cf8cd398a-P4ab42bb-generation-control-v4/terminal-oom.json`:
3,074 bytes, SHA-256
`158bbcc485206c87b990d1c3be74f22a867a2cd33cd9e9248aaf3918beaeef96`.
The constructed tar is 6,646,016,000 bytes and remains unverified. Its partial
packet and historical protocol pin are preserved.

Code inspection identifies an unbounded archive read in
`prepare_signed_native_inputs_v2.py::_artifact_identity`, reached through
`_checked_artifact` during archive validation. `_archive_tree_identity` also
reads each entire member. The journal proves the OOM; it does not provide a
Python stack proving the precise allocation site. The repair must stream
artifact SHA-256/byte counts and member Git-blob/ignored-lock digests while
retaining exact tree, mode, size, link and duplicate checks. Bounded-read
regressions are required; raising the memory cap is not the repair.

A new external recipe may reuse the constructed tar only through fresh,
complete verification against the frozen source and repaired protocol. It
must preserve the failed packet and write a new receipt with the reuse
provenance. No previous failure may be relabelled, and no new candidate
binding may be emitted from the tar's existence or size alone.

The current-native emitter and full65 gate remain live at this observation.
P stays at `4ab42bb9678769930d7c17173827fe5693370ab6` until the emitter
finishes. The installation sequence now allows a first P commit containing
the current-native and DrawingML bindings, their consumer migrations and the
signed streaming repair. The signed family remains pending in that commit;
its later binding and consumer migration must land together after verified
generation. The final common P pin comes after both commits. This preserves
same-commit consumer migration for each change without invalidating a live
emitter's explicitly bound prebinding HEAD.

The baseline materializer is also returned before execution. Its default
plan wrongly assumes C has a sparse file; C is a full checkout. A tiny Git
reproduction shows that new worktrees can already contain sparse/config
files, contradicting its absent-file assumption. It also treats unrelated
ref advances as its own mutation and writes its completion artifact before
final checks. Its successor must handle actual Git checkout metadata,
protect the relevant sources without blocking unrelated refs, and publish
completion only after final validation. No real baseline root was created or
changed by this return.

All final gates, native fidelity/performance, census review, trunk integration
and cleanup remain required. Nothing is pushed.
