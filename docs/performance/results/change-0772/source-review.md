# Current OPC mutation publication boundary

Reviewed at `6074e10e57`, before a production candidate. This is source evidence,
not measured attribution.

`tools/perf-baseline/src/lib.rs::run_opc_mutated_save` opens the generated owned
archive, flips the first target byte, calls `set_blob`, then constructs an
expected archive with `PackageWriter::to_bytes` **before** its sample loop.
Each sample times only publication to a pre-reserved bounded `CountingSink`.
The sink comparison and deterministic summary checks are outside the clock.
The eight corpora are four shapes times two payload types; the large shape
has four 4 MiB parts. No filesystem save or sync is measured.

An initial source review identified the original target's lazy decode in
`pkgwriter.rs::source_blob_retained`. Root checked its position against the
actual harness: expected-output construction has already called this function,
and `payload.rs::DeferredPayload` caches the result in a `OnceLock`. Thus that
first decode is outside this selector's timed repetitions. More precisely,
`get_part_mut` already forces the target before `set_blob`, and provenance
shares that deferred cell; expected serialization is an additional pre-loop
writer pass. It is **not** an
explanation of this selector's measured publication cost. The replacement's
first byte differs, so the later equality test can stop immediately.

Untouched deferred parts retain source spans. The replacement is regenerated
with Deflate, while 0742's compressed transfer applies only to bytes proved
equal to a verified compressed representation. Parallel member compression
cannot remove a single dominant changed member's serial compression cost.
These observations do not establish the actual cycle distribution; it still
requires a qualified profile.

The current harness oracle also needs strengthening before candidate acceptance:
it checks against the current writer's own expected output and omits the output
digest from its report. An untimed oracle should expose the exact output digest,
decode the changed target against the independently constructed replacement,
and compare untouched source local records and central records (with only their
relocatable offsets normalized). Existing `PreservationIndex` spans provide
the raw boundaries. This should supplement, not remove, the existing byte check.

Any profile must distinguish the timed publication from corpus construction and
the untimed expected serialization. Merely multiplying a large sample count
does not prove that startup work is negligible. A dedicated call boundary or
verified call-site addresses are needed for strict timed attribution.

Do not skip payload equality merely because `set_blob` was called: equal-byte
replacement must still retain the source's physical representation. Caller-
defined parts also retain their existing observation and fallback rules.

Other current work was checked before selecting this investigation: 0761 save
durability and 0766 XLS registration fixes already exist on separate branches;
0770/0771 also have worktrees. They are not absent implementations merely
because their final records are absent on the integration branch. This batch
does not change or merge those worktrees.
