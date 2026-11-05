# 0362d — ODraw picture package publication

- Date: 2026-10-21.
- Parent: 0351, 0355a, 0362c.
- Status: picture slice accepted after the retained-memory repair; final workspace integration remains open.

`litchi_odraw::prop::picture::Package` owns one admitted Opt-family record.
`open` and `open_with_options` borrow its input bytes; `view()` exposes a
read-only `View`. The former `Snapshot`, `Owned`, free `parse` and snapshot
publication methods are removed. Every consumer migrates in this change.
The old consumers were the picture codec tests; DOC/PPT/XLS use the broader
property table rather than this focused picture editor.

`Package::edit` and `edit_with_options` stage an independent `Edit`. Commit
returns the shared `Commit<Package, Patch>` and `Diagnostics`; dropping the
edit leaves the source unchanged. `Package::apply` returns the same triad.
`Patch::inverse` restores exact original record bytes, including an absent
property slot. Replay compares the complete recorded base content, including
the record kind, ordered descriptors, flags and opaque complex payloads.
An independently opened matching record is accepted; a different recipient
returns `PatchConflict`. Recipient limits govern the replay target.

The owner and patches share immutable states. Edit start and no-op commit
are constant time. A changed commit shares its original property table and retains only the
name/flags delta. It borrows or shares every untouched complex payload.
Chained edits keep a single original base, preserve append order and retain
a cleared descriptor slot; they do not form a replay chain. It produces a deferred owner;
`write_to` and `save` stream the record. `metadata()` reads admitted name/flags
without rendering. `view()` creates a read-only descriptor projection of a contiguous raw record
and materializes an edited record on its first call; `to_bytes()` explicitly
allocates a complete output vector. No-op publication retains the original
state. Chained edits share previously authored complex bytes.
`Package::into_owned` detaches borrowed payloads; detaching the source before
editing gives a whole commit that can outlive its input. `Patch::into_owned`
detaches both endpoints while copying each common borrowed payload only once.
This retains the former owned publication capability without forcing opaque
payload copies on ordinary commits. A one-KiB stack buffer batches small
output fields; large retained payloads pass directly to the writer.

Flags preserve the distinction between an absent descriptor and an explicit
zero. A name-only edit does not invent a flags change. Clearing a name preserves
its existing descriptor slot and producer bits, as the prior codec did;
applying the inverse of an append removes the appended slot exactly. Unknown
properties, ordering, BLIP bits, producer flag bits, secondary/tertiary Opt
record kinds and opaque payloads survive publication.

Picture `Limits` name three dimensions: record body bytes (excluding the
eight-byte header), descriptor count, and name UTF-16 units (excluding its
terminating NUL). Defaults retain 16 MiB, 4095 descriptors and 4096 name units.
Validated noun setters and `*_value` getters replace exported cap constants.
Owner limits enter through `OpenOptions` and optional tighter `EditOptions`;
default edits inherit ingress policy. Raw `View::parse(record, limits)` and
`Props::picture(limits)` require explicit policy. The latter checks its
projection's count and name budget; body admission belongs to `View`/`Package`.

The family keeps its own error type. Structural failures retain their ODraw
cause; limits carry typed kind and observed/maximum values; invalid options
carry value/maximum. Invalid authored names and semantic flags use a fixed,
content-free `Refused(Refusal)` with `InvalidValue`. Errors retain their typed causes. This focused codec has no current facade
operation; a facade consumer must classify its conflict/refusal/I/O variants
directly rather than flatten them into a string. UTF-16 validation walks the borrowed
payload without a temporary vector/string. Public owner, view, patch, name and change Debug output omits authored data;
I/O formatting exposes only its kind while retaining the source error.

Patch value getters borrow their before/after name payloads and return a
`Change<&[u8]>`; flags return `Change<u32>`. They do not retain duplicate copies
of the authored values. Public `Limits` remain machine-sized values while the
private retained policy uses checked wire-sized fields. The contiguous-view
cache is out of line and allocated only when requested.

The initial candidate failed memory acceptance: opening added 80 bytes and
128 small retained commits peaked at 62,464 bytes against 21,864. Its timings
were faster, but that did not justify retaining the memory regressions. The
repair replaces copied descriptor vectors and duplicate patch values with the
shared base/delta representation. Tables of at most 16 properties use bounded
linear lookup; larger tables retain the hash index. Duplicate rejection and
wire order remain identical. Small package ingress validates then releases
the temporary descriptor projection; its admitted raw iterator performs no
allocation. Public raw `View` parsing retains its own read-only projection.

The glossary profile registers the family methods and immutable owner/view,
and bans the removed API. The ODraw→core dependency is canonical because the
commit/diagnostics/refusal vocabulary is shared. The local inventory is empty for the rules it can detect; this does not
stand in for the remaining manual raw-policy and builder review.

This closes the picture publication slice. Raw record/iterator limits,
image-file scanner policies, the remaining global ODraw error vocabulary,
write-builder namespaces and retained host policies remain required work.
It does not close W4/W4b/W8 integration or the final workspace gate.

## Validation

The final source passes 5,156 tests across 232 targets in ODraw, imgconv, DOC,
PPT and XLS, with six existing ignored tests. All five crates pass Clippy with
all targets/features and `-D warnings`. The public picture example passes its
doctest. The ODraw glossary checks 15 bans, one family and ten immutable owners;
`results/0362d/api-debt-inventory.tsv` contains only its header. The independent
review found no blocking ownership or fidelity issue. Replay deliberately
shares the validated patch target while retaining recipient limits.

Coverage includes exact independent-base replay/inverse and conflicts; opaque
bytes, array payloads, fBid bits, Opt-family record kinds and descriptor order;
absent flags versus explicit zero; cleared appended name slots; chained edits;
no-op/drop behavior; narrowed no-op policy inheritance; UTF-16 limits and invalid
inputs; retained input independence; writer failures; and content-free Debug.
The private large-payload checks verify base/name sharing and that commit,
metadata and streaming leave the contiguous-view cache unmaterialized.

Earlier checks in this slice also passed the full crate-boundary check (68
packages, 289 internal edges, 11 existing boundary-debt entries), the locked
perf harness, the benchmark matrix and 1,167 glossary/boundary tooling tests.
One initial tooling command named the nonexistent `tools.test_api_debt_inventory`
module; the corrected command above passed. No result from that failed command
is claimed as a successful suite. Final warnings-denied rustdoc and the locked
native-resave check also pass (the latter emits four existing tool warnings).
The final matrix command was corrected from a nonexistent script path to
`python3 tools/unified_bench_matrix.py --check`; that check and final formatting
pass.

## Protected performance comparison

The source baseline is `610e3c1afe`. Both binaries use the same release profile,
compiler and probe source. Each route has three serial ABBA blocks pinned to
CPU 24, calibrated to approximately 500 ms per leg; the receipt retains every
leg and allocation observation. All workers acknowledged the heavy hold and
the process table was clear before each run. Output checksums match across
all timed pairs. Allocation observations use a separate instrumented binary;
their clock values are not used for timing claims. History retains 128 commits. Heap values are the additional tracked live
bytes above fixture/source setup, not process RSS.
Fixtures contain four lossless descriptors and either 24 bytes or 2 MiB of
opaque data; all authored fixture strings are synthetic.

The rejected initial and repaired final receipts are both committed as
`results/0362d/measurements-v1.json` and `measurements-v2.json`. Only v2 is the
accepted implementation. Ratios below are the median of the paired block
ratios, with candidate/base values below one indicating a reduction.

| Route | Time ratio | Block range | Peak allocated bytes, base → final |
|---|---:|---:|---:|
| open-small | 0.461870 | 0.456276–0.466238 | 360 → 208 |
| open-large | 0.460566 | 0.448183–0.484530 | 360 → 208 |
| view-small | 0.398962 | 0.398863–0.399290 | 360 → 208 |
| view-large | 0.392786 | 0.386076–0.401094 | 360 → 208 |
| metadata-small | 0.719930 | 0.712503–0.720116 | 11 → 0 |
| metadata-large | 0.712774 | 0.700164–0.718245 | 11 → 0 |
| noop-small | 0.168493 | 0.166559–0.168990 | 432 → 0 |
| noop-large | 0.000116 | 0.000116–0.000127 | 4,194,488 → 0 |
| commit-small | 0.340811 | 0.339876–0.346963 | 440 → 112 |
| commit-large | 0.000267 | 0.000265–0.000272 | 4,194,496 → 112 |
| write-small | 0.541633 | 0.539129–0.543668 | 440 → 112 |
| write-large | 0.721155 | 0.720723–0.721171 | 4,194,496 → 112 |
| props-open-4 | 0.388095 | 0.367425–0.395742 | 360 → 208 |
| props-open-16 | 0.398905 | 0.397001–0.403719 | 1,392 → 832 |
| props-open-17 | 0.991817 | 0.991712–1.020773 | 1,444 → 1,444 |
| props-open-256 | 0.980195 | 0.976353–0.982471 | 22,032 → 22,032 |
| props-get-4 | 0.225293 | 0.222336–0.225890 | 0 → 0 |
| props-get-16 | 0.433206 | 0.432899–0.433966 | 0 → 0 |
| props-get-17 | 0.742432 | 0.726655–0.747756 | 0 → 0 |
| props-get-256 | 0.744079 | 0.741670–0.747590 | 0 → 0 |

The small retained history falls from 21,864 to 21,504 peak bytes; retained
bytes are exactly 21,504 on both sides. The large retained history falls from
270,551,176 to 21,504 peak bytes (268,453,888 to 21,504 retained). Total allocation
for either repaired history is 29,472 bytes; the corresponding baselines are
153,760 and 537,018,528 bytes. No-op commits allocate nothing. Changed commits
peak at 112 bytes for both opaque payload sizes, versus 440 and 4,194,496.
These are the measured shapes and routes, not a claim about every possible
record, allocator or deployment.

The general property lookup change is covered separately on both sides of
the 16-descriptor threshold and on a 256-descriptor table. Hash-indexed tables
retain their allocation footprint, while the small tables avoid the hash
allocation. All final timing medians are at or below the recorded baseline.

## Reproduction and artifact provenance

From the commit containing this record, use the capped cargo environment and
an external output directory:

```sh
source (deleted-coord-artifact)
python3 docs/unified-ops/results/0362d/picture-probe.py build /tmp/picture-0362d-repro
# Run only in a coordinator-approved quiet window.
python3 docs/unified-ops/results/0362d/picture-probe.py run /tmp/picture-0362d-repro --blocks 3 --target-ms 500
```

The build command creates isolated before/after source copies and separate
timed/allocation binaries. The receipt records their sizes and SHA-256 values,
the compiler version, lockfile hashes and every copied source hash. It includes
the driver and runner SHA-256 values. No compiled binary is committed. Frozen
sources, logs, binaries and per-route checkpoints remain under the external
`(deleted-coord-artifact)` and `probe-v2`.
The two source archives in that external parent directory are:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| picture-sources-v1.tar.gz | 609,386 | `42ef2849dfc9f416ad6eb41d8c58ca4c87fb0cc607f4b4eea3c46cb9a9f5b1f2` |
| picture-sources-v2.tar.gz | 611,813 | `9b31c728c2bddef1ad0ae5a4126482c76deffbc2009de91bb4d3f08e5f5ea7db` |

The source-copy directories can be regenerated with the build command; the
initial implementation is frozen in the v1 archive. The archives contain only
those sources, manifests, lockfiles and driver files. Final timed binaries:
baseline 447,728 bytes, SHA-256
`462073bad50e76d3db3fd7409e1e97d37fc2ffcf28cbbffb0fec13bba464b954`;
candidate 458,696 bytes, SHA-256
`ea8ad80aa7890c6754da69b806b189c8ac75f737e3decf3451d3c953a654e76e`.
