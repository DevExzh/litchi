# 0493 bundle integration review

Static review completed 2026-09-10 against the current `measure.py` and
`cleanup.py` contracts. No Cargo build or benchmark capture was run. The
following notes record the integration fixes and the live-gate sequencing
precondition.

1. **The bundle verifier now invokes canonical cleanup verification.**
   The previous `verify_bundle.py:61-73` checks covered only
   `cleanup.json.status`, target absence, and the retained file/directory set.
   The verifier now calls `cleanup.verify(root=ROOT, temp=TEMP,
   target=TARGET_DIR)` before sealing, binding the cleanup schema, roots,
   retained-attempt scope, retained-binary inventory, removal inventory, and
   process audits.

2. **A live gate is a pending precondition for cleanup and sealing.**
   During this review `validation/final-docx-tests.started.json` was live and
   had no terminal receipt yet. Both `cleanup.py:403-420` and
   `verify_bundle.py:57-59` require a terminal receipt for every
   `*.started.json`, including development attempts. The active gate must be
   allowed to finish and produce its terminal receipt; do not remove or replace
   that start receipt while it is running. Root subsequently reported the gate
   terminal with exit 0 and unchanged source.

3. **The final-gate contract must remain strict and source-bound.**
   The original `verify_bundle.py:34-40` required only ten entries and compared
   each receipt with the command self-described by `final-gates.json`; it did
   not validate a schema, uniqueness, or the required final command set. In
   addition, `measure._gate_binding()` (`measure.py:370-409`) does not bind the
   receipt's `label` to the JSON filename. The verifier now requires the exact
   twelve-command set, allows attempt-suffixed labels, and requires
   `receipt["label"] == path.stem`.

4. **Cleanup now authenticates build custody before target deletion.**
   The original `cleanup.py:346-379` validation ignored each build receipt's
   gate descriptor, command, environment, `git_revision`, and retainer hash.
   Its production preflight now calls `measure.load_builds()` at
   `cleanup.py:602-617`, and the execute path repeats that check immediately
   before removal.

The evidence inventory now rejects non-symlink special files under the evidence
root instead of silently omitting FIFOs, sockets, devices, and other
non-regular paths.
