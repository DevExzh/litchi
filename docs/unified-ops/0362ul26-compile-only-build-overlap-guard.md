# 0362ul26 — compile-only native build overlap guard

Date: 2026-10-24. This leaf adds the approved protocol boundary for
overlapping compile-only work. It changes only the generic native runner and
its unit tests; the OLE-owner and ODF shared-triad paired runners remain
globally quiet. Existing manifests omit `build_execution` and therefore retain
the old exclusive guard. No frozen source packet, manifest, receipt, or
qualification result was rebound.

## Contract

A rendered manifest may opt into overlap only with:

```json
"build_execution": {
  "class": "compile_only",
  "permit_path": "/external/coordinator/native-build-permit.json"
}
```

The external `litchi.native-build-permit.v1` file binds a grant and job group,
an expiry, `max_jobs` from one through four, exact admitted lanes and manifest
SHA-256 values, the admitted Cargo command classes, and an external lease
directory. Each participating job atomically creates an external lease file
under that directory. The lease binds its group, controller PID/start time,
manifest and exact normalized source/target lists, and (once spawned) the
Cargo child PID/start time. A controller first publishes an exact pending
phase (phase ID, normalized `cargo build` argv, cwd and target); only a direct
child matching that phase may be admitted during the spawn-to-bind transition.
The bound record then carries that child's identity and exact phase, and the
child must remain directly parented by the live controller. The phase argv is
also matched to the admitted manifest build entry and checked by the runner's
Cargo path/config restrictions before process allowance. The runner rejects
malformed, stale, expired, changed, reused, over-capacity, or phase-changed
leases. Peer roots must still be live Cargo processes whose command class and
argv/cwd match the lease; Cargo docs, rustdoc, nested Cargo, profilers, and
other disallowed descendants fail closed. An
unregistered process remains an offender, and the Python controller itself is
never an allowed process, so the permit is not a blanket process allowlist.

Only the generic runner's own rendered `cargo build` commands can select this
class. A build-time permit expiry terminates and reaps the owned child and
writes a failed build receipt. Build events and the build receipt retain the
permit/lease identity and start/end peer snapshots. Existing 20,000,000,000
byte target and address-space limits, 19,200,000,000 byte preflight, one Cargo
job, fresh output/targets, source identity checks, output/log hashing, and
periodic descendant checks are unchanged. ABBA measurements and all controls
and validators always use the exclusive quiet guard.

## Checks

The source was based on integration `c204999a444b799658ed56b5e92609986b21fbd4`.
No Cargo or native qualification command was run. The exact reduced receipt is
[results/0362ul26/compile-only-build-overlap.json](results/0362ul26/compile-only-build-overlap.json).

| Check | Result |
| --- | --- |
| `python3 -m py_compile tools/native_qualification_runner.py tools/test_native_qualification_runner.py` | exit 0 |
| `python3 -m unittest tools.test_native_qualification_runner -v` | 38 tests passed, exit 0 |
| `python3 -m unittest tools.test_drawingml_raw_qualification tools.test_legacy_native_qualification -v` | 31 tests passed, exit 0 |

External logs are under
`(deleted-coord-artifact)`.
The final runner source is 176,110 bytes (SHA-256
`f550adf27ff6559a2318cc456c6f8cc4f3d338e74119e21c470e2c767452bf03`) and its
focused test source is 68,468 bytes (SHA-256
`a469731e5947342bbc97574447ddbc15d816d002693ac2822666dc09bfa411d4`). The
final logs are: `pycompile.log`, 0 bytes, SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`;
`native-runner-tests.log`, 7,663 bytes, SHA-256
`e6acb516b9eefeadacdf7ac90054cceb27a08f4c991e44a34c9b81bd4cad5906`; and
`paired-legacy-tests.log`, 5,941 bytes, SHA-256
`c0e3c58176d3dd1e23218ca30eafa81358b1715a5afd7d78b9caa1151c57c3c9`.
The focused integration creates temporary committed fake-Cargo trees and
asserts overlapping SafeRunner event intervals; it invokes no Cargo. This
leaf makes no performance, ABBA, RSS, or native correctness acceptance claim.
