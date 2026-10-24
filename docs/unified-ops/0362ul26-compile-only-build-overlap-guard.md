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
under that directory. The lease binds its group, root PID and `/proc` start
time, manifest and source/target paths, exact command class, and `phase=build`.
The runner rejects malformed, stale, expired, changed, reused, over-capacity,
or phase-changed leases. Peer roots must still be live Cargo processes whose
command class matches the lease; Cargo docs, rustdoc, nested Cargo, profilers,
and other disallowed descendants fail closed. An unregistered process remains
an offender, so the permit is not a blanket process allowlist.

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
| `python3 -m unittest tools.test_native_qualification_runner -v` | 31 tests passed, exit 0 |
| `python3 -m unittest tools.test_drawingml_raw_qualification tools.test_legacy_native_qualification -v` | 31 tests passed, exit 0 |

External logs are under
`/home/zhuhe/code/litchi-worktrees/.coord/cfb-compile-overlap-protocol-20261024/`.
The runner source is 152,389 bytes (SHA-256
`597c2e46027b11c7f882d7ae28bc3c8b3a8abff9d92e015fbb25e6be32a8752c`) and its
focused test source is 50,216 bytes (SHA-256
`e767c6943d50a70417a68aef27b01371ab03fb5a0b35882f940ecc53fbd5d947`). This
leaf makes no performance, ABBA, RSS, or native correctness acceptance claim.
