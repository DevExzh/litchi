# DOCX `stylesWithEffects` optimization attribution plan

Status: **plan only**. This document authorizes no code change, timing run, or
optimization. It defines the smallest additional evidence needed to decide
whether a library change is justified while retaining the complete CRUD and
preservation goal.

## Evidence boundary

The retained profile in
[`results/profile-clean-96958498f`](results/profile-clean-96958498f/README.md)
is a bounded absolute observation from production source
`d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119`. Its 1 MiB replacement publication
medians, including roughly 122 ms for main and 125 ms for glossary, are
historical measurements. They are not an attribution of the current OPC
implementation and are not a before/after result.

The current source has since changed in the physical package owner, including
`83e60e8c62d52d46d75e06f3ddda121692775dc3` (reuse of source content types for
unchanged membership) and
`aab225305e98f90b60daf67441b888741fed2724` (source relationship preflight).
Those changes can move work between publication, validation, and serialization.
Any optimization decision therefore starts with a new clean current baseline
whose complete local Cargo source closure, lockfile, generated inputs, and
toolchain are sealed. The old profile remains a reproducible historical
control only.

The decision follows the order in `docs/GOAL.md` and ADR 0005: first remove
unnecessary work, I/O, parsing, allocation, copying, and recompression; then
consider layout or algorithms. A proposed change must preserve source-backed
opaque XML, exact no-ops, owner independence, typed limits and refusals,
atomic publication, exact inverse, signatures, and all add/replace/remove
CRUD scenarios. SIMD, unsafe code, hidden parallelism, and a broad profiling
framework are outside this plan.

## What the historical `publish_ns` actually contains

The profile harness prepares the package, baseline bytes, member map, and
replacement `Resource` before the allocator and elapsed window. For a changed
`replace_*` lane, `snapshot_ns`, `stage_ns`, and `commit_ns` are separate. The
`publish_ns` timer then covers this sequence:

| measured segment | Work inside the historical `publish_ns` | What it means for attribution |
| --- | --- | --- |
| Public patch application | `Package::apply_styles_with_effects_patch`, including its current owner load and source precondition check | Includes the DOCX facade and effects owner path, not only the final byte splice. |
| Candidate publication | `edit_semantic_opc` checks the current source and signature policy, clones the `OpcPackage`, invokes `effects::apply_patch`, validates the candidate, reloads the main/custom property state, and assigns the candidate | Includes candidate cloning, graph inspection, package-limit preflight, source relationship/content-type work, changed XML ownership, candidate owner readback, and facade publication checks. |
| Effects operation | `effects::apply_patch` loads the current owner again, validates the patch graph/resource, calls `publish_resource`, reloads the staged owner, and compares resource and graph state before returning | May include XML validation/projection and graph/relationship parsing more than once. Exact call counts must be measured before describing this as duplicate work. |
| Package output | `package_bytes` calls `Package::to_stream`, its `write_plain` wrapper, `OpcPackage::to_stream`, and `PackageWriter::write_to_stream` into an in-memory `Cursor<Vec<u8>>` | Includes `write_plain` rollback/property/custom-property staging, source preservation-index work, the selected writer mode's member copying or compression, central-directory output, and output-buffer allocation. It is not a disk save, fsync, or source-file I/O measurement. |

The prepared fixtures do not stage unrelated property edits, but `write_plain`
still remains inside the measured call. Resource XML construction belongs to
the apply phase. Copying and compression depend on the selected writer mode;
the plan does not assume every publication performs both.

The timer stops after the output bytes are returned. It excludes package open,
baseline serialization, fixture/XML generation, replacement construction,
external `Package::from_reader` reopen, selected-owner semantic checks,
unchanged-member hashing, graph metrics, and receipt serialization. The
historical `reopen_ns` is therefore a separate external reopen, while any
candidate owner readback performed inside `apply_patch` remains part of
`publish_ns`. Process RSS from `/usr/bin/time -v` is a maximum for the complete
process invocation and cannot be assigned to this one clock.

No disk or remote source attribution can be inferred from this harness: the
operation uses an already prepared in-memory package and an in-memory output
cursor. A future disk, sequential sink, or caller supplied range-source
scenario must be named separately.

## Minimum attribution capture

Keep the existing bounded matrix and 3 fresh processes × 2 warmups × 20
samples. Add only two disjoint operation clocks to the successful prepared
publication lanes currently covered by the scaffold: `noop_main`,
`noop_glossary`, `replace_main`, `replace_glossary`, `remove_main`,
`remove_glossary`, `add_main_absent`, `inverse_replace_main`,
`inverse_remove_main`, `independent_main`, and `independent_glossary`.
Capture, projection, refusal, and signed lanes retain null split fields.
The two clocks are:

1. `apply_ns`: from immediately before the public mutation closure until it
   returns, before output serialization. In source no-op lanes this closure
   includes `apply_styles_with_effects_patch` followed by
   `put_styles_with_effects`; in inverse lanes it also constructs the inverse
   patch inside this clock; and
2. `serialize_ns`: from immediately before `Package::to_stream` until the
   output cursor is complete.

Retain the existing outer elapsed clock and allocator windows. Record, for
each subphase, elapsed time, requested/live/peak allocator bytes, output byte
count, and failure status. Do not put hashing, opaque checks, graph metrics, or
receipt construction into either clock.

Add one narrow physical-publication attribution record, using an existing OPC
accounting seam if it provides these fields or a small opt-in diagnostic seam
owned by `litchi-opc`:

* writer mode: exact-source copy, targeted source-preserving write, or full
  physical writer fallback;
* source-preservation index construction and source-member scan bytes;
* unchanged, changed, omitted, and appended member counts;
* changed-member uncompressed/compressed bytes and compression work, where
  an existing accounting seam exposes them;
* output bytes and write-call count/size summary; and
* relationship/content-type XML bytes considered during publication.

For `apply_ns`, a short call-count/provenance record is sufficient. Count or
sample call stacks for `effects::load`, graph capture, XML limit validation,
style projection construction, source relationship capture, content-type
planning, candidate cloning, and candidate owner readback. Use `perf`/callgrind
or a narrowly scoped internal counter if call counts alone cannot distinguish
the dominant stack. Do not add counters to the public API or record document
content. Any internal counters or writer diagnostics must be test/feature-gated
with zero default production overhead, or collected externally through tools
such as `perf` or callgrind.

The current seam inspection leaves the OPC writer record unimplemented. The
public `OpcOperationAccounting` report covers source-backed cold Part reads,
exact source copies, and the single-Part overlay publisher; ordinary
`OpcPackage::to_stream` and `PackageWriter::to_bytes` do not accept that report,
and the preservation writer keeps its ZIP counters internal. The attribution
scaffold therefore records only the DOCX apply/serialize split and allocator
subpeaks. If the next evidence gate assigns the dominant work to OPC, add one
opt-in diagnostic writer entry point or feature-gated callback that returns
writer mode, member counts, and accepted-byte counters without recording
document content. The normal writer path must retain zero accounting overhead;
this seam is not part of the current profile or an optimization claim.

The minimum controls are:

* a source-backed exact no-op that reaches output serialization, to measure the
  source-copy/output floor;
* one present-owner replacement with the same deterministic 64 KiB and 1 MiB
  resource inputs used for both main and glossary; and
* the existing remove, add-absent, inverse, source-no-op, and owner-
  independence lanes as correctness and regression controls.

Compare main and glossary separately. A large `serialize_ns` share with a
targeted-preservation mode points to the OPC writer; a large `apply_ns` share
requires the call-count/call-stack record before choosing between DOCX owner
logic and OPC candidate publication. If the no-op floor accounts for the
observed time, report the floor rather than optimizing a guessed hot loop.

## Current-baseline closure

Before attribution, capture one immutable baseline and retain its manifest with
the receipts. The baseline must include:

| binding | required evidence |
| --- | --- |
| Production source | exact Git `HEAD`; clean status; every local path dependency and production file blob; `Cargo.lock`; `cargo metadata --locked --offline` digest; target triple/linker; and the exact feature set. The source manifest must fail closed on a moving or dirty input. |
| Generated and native inputs | byte-exact hashes for all four native packages and selected effects members; deterministic XML bytes, complete generated package bytes, seed, conformance, XML event/depth/style counts, and replacement marker hash. Baseline and candidate must consume identical generated inputs. |
| Build | `rustc -vV`, `cargo -V`, rustup toolchain, release profile, `--locked --offline`, `CARGO_INCREMENTAL`, `RUSTFLAGS` and related variables, allocator observer, binary hash, build log, and exact commands. |
| Host | OS/kernel, CPU model, physical and logical core counts when available, affinity, memory, filesystem/mount identity, `/usr/bin/time -v` version, load, and unrelated-process census. If these are incomplete or the host is shared, state that no idle or uncontended claim is made. |
| Run | fresh external results and target paths, process indexes, warmups/samples, output and target cleanup receipts, raw phase/allocation/RSS sidecars, and no simultaneous competing profiler or timing run. |

The source closure must include the current OPC commits above and all their
transitive local inputs. A baseline built from the old captured `d1` closure
cannot be compared with the current workspace, and a candidate may change only
the explicitly selected owner files after the baseline is sealed.

## Optimization decision gate

An optimization branch may be authorized only when all of these conditions
hold:

1. The current baseline passes the complete correctness smoke, including all
   add/replace/remove/no-op/inverse lanes, main/glossary independence, exact
   opaque-member checks, stale-patch rejection, failed publication/readback
   atomicity, signed-change refusals, signature policy, malformed-input checks,
   and cap refusals. Include the current OPC focused tests and exact/one-under
   caller-limit and source-preservation refusal checks. The 1 MiB success lanes
   alone are insufficient.
2. The split clocks and writer record pass phase-sum and allocator validity
   checks, and the dominant component repeats across main/glossary and the two
   initial scales with uncertainty reported.
3. The dominant work is assigned to an owning crate from evidence: OPC
   physical publication/writing, DOCX effects graph/resource validation, or
   another named owner. A `publish_ns` total by itself is not an attribution.
4. The proposed change follows GOAL steps 1 or 2, has a concrete preservation
   and limit invariant, and has a small falsifiable correctness test plan.
5. The candidate can be rebuilt with the same generated inputs, lock/toolchain,
   host protocol, and process settings. The matched rerun must include the
   complete CRUD smoke and the same attribution receipts before any speed or
   memory claim is written.

If no component clears this gate, retain the baseline and record that the
profile did not justify an optimization. If OPC serialization dominates,
investigate only measured source-member reuse, compression, or output-copy
work under ADR 0011. If effects publication dominates, investigate measured
duplicate graph/XML work or candidate copies under the DOCX owner and ADR 0003.
In either case, preserve the full CRUD matrix and stop before speculative
low-level tuning.
