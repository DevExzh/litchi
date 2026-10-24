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
The bound record then carries that child's identity and exact phase. The
runner rejects malformed, stale, expired, changed, reused, over-capacity, or
phase-changed leases. Peer roots must still be live Cargo processes whose
command class and argv/cwd match the lease; Cargo docs, rustdoc, nested
Cargo, profilers, and other disallowed descendants fail closed. An
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
| `python3 -m unittest tools.test_native_qualification_runner -v` | 36 tests passed, exit 0 |
| `python3 -m unittest tools.test_drawingml_raw_qualification tools.test_legacy_native_qualification -v` | 31 tests passed, exit 0 |

External logs are under
`/home/zhuhe/code/litchi-worktrees/.coord/cfb-compile-overlap-protocol-20261024-final/`.
The final runner source is 171,733 bytes (SHA-256
`96979f6f16bb9ebf2fe89c2c098ad074ef0f6797f78b88347f5a88633aac714e`) and its
focused test source is 65,566 bytes (SHA-256
`c58d4e890d53cb980bc74cb04f8adefaff6419877ac7df45efe028dea19b1587`). The
final logs are: `pycompile.log`, 0 bytes, SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`;
`native-runner-tests.log`, 7,259 bytes, SHA-256
`8e147d150a75b88389024f77085e5c8f7573f2bb9a4250cbaf80864e4bd0f242`; and
`paired-legacy-tests.log`, 5,941 bytes, SHA-256
`88f4282257eb8791d9e6ea1ab438414f6641422911a881cb8423f0f3dc49596f`.
The focused integration creates temporary committed fake-Cargo trees and
asserts overlapping SafeRunner event intervals; it invokes no Cargo. This
leaf makes no performance, ABBA, RSS, or native correctness acceptance claim.
