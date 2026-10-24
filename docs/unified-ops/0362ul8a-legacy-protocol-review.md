# 0362ul8a — legacy protocol review returns and composition

Date: 2026-10-24. Parent: [0362ul8](0362ul8-legacy-native-protocol.md).

Independent review accepts source preparation packet
`e8d51bb07a00dd0b9feeb6ed5e134c521c89be4d` after the following returns.
This is preparation; no native measurement or final performance acceptance
is claimed.

The validators now bind actual Git identities and allowed consumer overlays,
reviewed input hashes, lock records, build receipts and journals, and exact
RTF/OLE/XLS workloads. The RTF probe checks the fixed refusal kind/reason and
unchanged/replayed bytes. Its tracked source hash is enforced by a real
repository consistency test. Raw symlink components are rejected before
resolution. The adapter helpers require an explicit owned root, a contained
fresh destination and exclusive creation; real temporary-directory tests
exercise source, destination, parent and root symlinks and path escapes.

The OLE property baseline remains
`604822b44107999aaddb0bffb96b705e9464f546`: its object owner is already
migrated while the property API is still legacy. The frozen fdb probe cannot
use its legacy-owner branch at this revision. A reviewed consumer-only
feature split selects the actual public current-owner/legacy-property
boundary, preserving the property workload and production source. The
candidate adds the required CFB limit argument at the single parser call.
The exact adapter hashes are:

- Before: `6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed`.
- After: `e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782`.

The distinct original object-owner baseline `0ce4bb8b61` and its eleven
routes, clone and retained-history qualification are separate required work.
This property pair does not replace them.

Coordinator composition retains the XLSX raw protocol and all three new
legacy lanes in the shared runner. The conflict resolution unions the lane
registry, both workload/preparation validators and both initialization paths.
The resolved source index is `29f2b28a0bcbed33c534c91e6dedd493ae066de8`.
A single mutually exclusive selection keeps the shared XLSX receipt and the
top-level legacy receipt available to the same external/disjoint path checks.
Four XLSX source/target overlap cases guard this composition. All 67 composed
Python tests pass in 9.981 seconds:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_prepare_xlsx_raw_inputs tools.test_xlsx_raw_native_qualification tools.test_native_qualification_runner tools.test_legacy_native_qualification tools.test_ppt_lazy_native_qualification tools.test_no_3rdparty_references
```

The previous `384740be13` integration was returned before commit despite
64 passing tests. Actual handoff review found obsolete OLE feature flags,
a preparer/runner digest-framing mismatch, and missing legacy receipt path
separation. Its 150,269-byte resolved patch remains external, SHA-256
`8759ae84c441b128720e195e640c2629662cb3768971d7246fe88c7c3096093a`.
The corrected leaf adds actual manifest-loading and real Git identity
round-trip tests. The extra XLSX overlap regression initially expected the
wrong error wording; the guards correctly refused both source overlaps.
Only that assertion wording changed before the passing run above.

Independent review also executes the preparer CLI on the real XLS baseline
`509bf28e7531ad75f278eaec3a732714a243915a` and candidate `dcc7cfbdef`.
The runner consumes the resulting receipt and correctly stops because no
build receipt exists. This proves the preparation handoff, not a build or
native run. Reduced evidence is in
[`results/0362ul8a/composed-review.json`](results/0362ul8a/composed-review.json).
The immutable CLI files remain under
`/tmp/legacy-cli-review-xls-staging-2984703/`.

The source-only receipt does not claim a Cargo build of the RTF probe or
legacy endpoints. Those builds, execution, native/allocation review and the
assembled acceptance gate remain required. The earlier accidental feature
matrix attempts remain recorded in 0362ul8 as 31 failures before compilation;
they are not validation passes. Only redaction audit 22 is deferred.
