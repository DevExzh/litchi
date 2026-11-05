# 0362ul80a — tooling-row capture ownership correction

This source/tooling-only follow-up records the capture ownership correction for
0362ul80. No Cargo, rustdoc, native qualification, allocator, timing, or
performance run was performed.

## Capture distinction

Three outcomes are kept separate:

1. The first shell launch attempted to redirect to
   `(deleted-coord-artifact)`
   before its parent directory existed. It exited before starting
   `run_safe_gate.py`; no runner state, test process, or source capture was
   created. This is a setup failure, not a source-run receipt.
2. The retry used mutable integration at
   `317f6482421e8a50a0361d1590d19077518c7fb8`. Its child completed 1,746
   tests, but integration advanced to
   `11048ed3cbcd664308d5ad77c8361f21802578a4` during execution. The wrapper
   correctly rejected that receipt for provenance change; it remains failed
   historical evidence and is not acceptance evidence.
3. The accepted run used the independently hydrated detached worktree
   `(deleted-coord-artifact)` at fixed
   `11048ed3cbcd664308d5ad77c8361f21802578a4`. Root verified its top-level,
   detached HEAD, clean status, full tracked hydration, runner root, and lock
   identities before launch. Integration was not used by that command.

The earlier failed 0dc integration receipt from 0362ul80 remains preserved;
none of these failed receipts is relabelled as a pass.

## Detached source and command

The preflight manifest is external at
`(deleted-coord-artifact)/tooling-11048-source-manifest.json`
(3,811 bytes, SHA-256
`70b60cd846c0d0cd1840b7f4f003f3f010689895ea874f3aebf308ac2b03583f`). It
records 461,107 tracked files, zero skip-worktree entries, and 39 modules.
The source runner is `e88893d497f7e7ef70bde4314c3e577d47f20997f2ea9d82b4d4595edbbe8458`
(52,787 bytes). Exact lock identities:

```text
Cargo.lock                      efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811
tools/perf-baseline/Cargo.lock  de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea
tools/native-resave/Cargo.lock  989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae
```

The root `Cargo.lock` is ignored by Git and was copied into the detached
worktree before verification; its bytes match the recorded `efff87af...`
identity. The exact launched command was:

```text
env -u PYTHONPATH -u CARGO_TARGET_DIR -u LITCHI_GATE_TARGET_ROOT PYTHONDONTWRITEBYTECODE=1 \
  python3 (deleted-coord-artifact) \
  --root (deleted-coord-artifact) \
  --target (deleted-coord-artifact) \
  --state (deleted-coord-artifact) \
  --only tooling-unit
```

Immediately after launch, `plan.json` recorded the detached source path and
exact 11048 commit; the live worktree matched those identities.

## Accepted detached receipt

The external wrapper is
`(deleted-coord-artifact)/tooling-row-11048-detached-20261025T-wrapper.log`
(4,121 bytes, SHA-256
`d5bb90124ab842f33923edbaff95a687cc7dc56affd73b6e4fdad2485ff438b3`). The
corresponding overall state receipt is
`(deleted-coord-artifact)/state.json`
(10,603 bytes, SHA-256
`07ceb0ca5851dd1479b6a6e01fc0adc11d846cd95b84c8bf8d199b6e295e6880`). The
overall state status is `partial-passed` because this capture selected only
row 48 of 53; its row 48 status is `passed` with effective exit 0. The
child log is
`(deleted-coord-artifact)/48-tooling-unit-attempt-1792903033125826266-1634645.log`
(635,222 bytes, SHA-256
`d3bb1741180463c5883ab9e8644192dcb6446e52eaec3ad6ad2843fe57391950`).

All 39 modules passed:

```text
Ran 1752 tests in 308.544s
OK
```

The wrapper recorded `status=passed`, `effective_exit=0`, detached 11048,
clean status, and unchanged locks. Target usage was zero bytes throughout
(`max_sampled_bytes=0`); the Cargo-target trace was empty (0 bytes, SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`). No
actual Cargo/native/rustdoc process ran; cargo-like lines in the child log
come only from synthetic Python fixture tests. After exit, runner and unittest PIDs
1634645 and 1635276 were gone, owned/builders lists were empty, and the
source worktree remained detached, clean, and at 11048.
