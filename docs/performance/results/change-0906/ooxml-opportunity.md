# OOXML opportunity after 0905

This is a source audit, not a timing result. The shared checkout is at
`59f75c7f79026a4d124b61e97c178c7b45de19ab` and contains unrelated iWork
conflicts. The Office source context receipt says that the audited files match
the accepted Office branch `2b974e3541517ccf032021dc5497243b315fd313`; the
0905 measurement remains frozen to its `4f39f5a88d` source and must not be used
as current OOXML timing. The current harness source is
`tools/perf-baseline/src/lib.rs` SHA-256
`b62775d903ccb2490c28af61d6c5c95d8ed826afad4f0e7dc3c83bd41b877903`.

## Finding and priority

The next high-ROI capture should be a fresh paired current-source run of
`pptx_eager_multi_slide_batch_edit_save` versus
`pptx_source_backed_multi_slide_batch_edit_save`. This is a requalification of
a historically expensive path, not a current speedup claim. The strongest
strictly unmeasured current edit/save route is
`docx_source_backed_existing_section_layout_edit_save`; it should be the first
optimization target if the PPTX requalification does not preserve its old
cost separation. The DOCX route already has phase vectors and complete
preservation gates, but no accepted performance result or matched eager
control.

The current edit/save matrix is small and explicit: DOCX has the media-rich
one-paragraph source-backed save plus this section-layout save
(`lib.rs:3801-3807`); PPTX has one-edit, same-slide batch, and multi-slide
batch eager/source saves (`lib.rs:3830-3838`). Cross-copy and image-query
selectors are separate workflows, not edit/save coverage. The section-layout
selector is opt-in; change 0273 records that it is outside the default 36-case
baseline.

The 0900 record says an old ranking becomes invalid after publication auditing
and runtime changes, and that historical times are not a before leg
(`docs/performance/0900-xlsx-edit-save-phase-refresh.md:10-18`, SHA-256
`fc083ee99f27e412c8a2acd98611782353bdc49aeacfec206e71766d147a0b8d`). The
same evidence rule applies to the older PPTX results, which is why PPTX is the
first capture rather than an adopted optimization.

This ordering follows the goal's requirement to rank measured end-to-end work,
name the scenario/corpus/build/metric, and avoid unsupported claims
(`docs/GOAL.md:12-16,45-47,799-817`, SHA-256
`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`). The
CRUD checklist keeps opened-document editing separate from authoring, requires
exact no-op, atomic, reversible, and preservation evidence, and says to omit
performance claims without reproducible measurement
(`docs/CRUD_Scenario_Checklist.md:9-12,46-69,71-81`, SHA-256
`d3a63f3aa001ad6e3dda7627b6294230ae87b3b5aaa595c45c3327e026dbf02a`).

## Highest-ROI requalification: PPTX multi-slide batch

The current enum and dispatch expose the paired cases at
`tools/perf-baseline/src/lib.rs:1177-1185,3830-3838,11031-11068`. Both cases
use `build_pptx_source_edit_corpus` (`:16928-16964,17016-17065`), whose fixed
shape is `media-rich`: 200 slides, eight text boxes per slide, eight
deterministic incompressible 2 MiB PNG Parts, 229 ordinary Parts, and 445 ZIP
members. The retained archive identity is 17,017,139 bytes, 17,568,429
logical Part bytes, SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
The edit changes all eight text boxes on slide positions
`[0,28,57,85,114,142,171,199]`; the retained output identity is 17,017,145
bytes, SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.

The exact runner is `run_pptx_multi_slide_batch_edit_save`
(`lib.rs:51039-51213`). Its guards require the source generator and digest,
then require source-backed materializations `PPTX_MULTI_SLIDE_BATCH_COUNT + 1`
(`9`) versus eager materialization of every manifest Part (`229`)
(`:51122-51146`). `verify_pptx_multi_slide_edit_output`
(`:50705-50750`) reopens the package, checks semantic text, topology,
relationships, content types, the exact selected slide set, untouched Part
identity, and all eight exact media payloads. Each sample also requires a
nonempty ordinary source read, a stable output digest, and a sequential sink
whose largest write is at most 64 KiB (`:51158-51213`).

The timer boundary needs to be reported precisely. Elapsed time starts around
the call at `:51166-51169`. The source-backed helper performs patch replay and
inverse checks before publication (`:51060-51067`) and inverse/forward checks
after publication (`:51068-51075`) inside that interval; eager performs the
analogous replay checks (`:51093-51113`). The paired result is useful only as a
guarded operation interval unless a future harness change records a separate
pure operation interval. Full semantic and package verification after the
timer remains a correctness gate, not latency.

Change 0077 is the strongest retained cost signal: on production base
`a82a5d9f3f3e34706a9ee890fd23f48b92718e6f`, it reports p50 331.362 ms eager
versus 13.997 ms source-backed (23.67x) and 229 versus 9 semantic
materializations. Its report SHA-256 is
`6142438ea026f38660b7a87f532e87b775c07ec4f91b5b1108a26888bfc92159`. Those
times, old binary hashes, and the related one-edit and same-slide results in
changes 0044 and 0063 are historical evidence only. They do not describe the
current `lib.rs` or the post-0905 executable.

## Strongest genuinely unmeasured route: DOCX section layout

The current selector is dispatched separately at
`tools/perf-baseline/src/lib.rs:10962-10973` and is implemented by
`run_docx_source_backed_existing_section_layout_edit_save`
(`:50319-50652`). Its fixed corpus is
`docx-source-backed-existing-section-layout`, generator
`litchi-docx-section-layout-source-v1`, shape `existing-section-layout`:
three logical sections, 257 main-story paragraphs, a table-cell section
sentinel, headers/footers and opaque fragments, four deterministic 512 KiB
media Parts, seven OPC entries and ten physical ZIP members. The retained
archive is 2,101,427 bytes with SHA-256
`183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8`; the
main-document payload is 18,855 bytes with SHA-256
`67cb9066f112bfee6905a23f3eabdc76b65c86cd0c6613dd60a5a18895ceac81`; the
retained changed output is 2,101,446 bytes with SHA-256
`5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec`.
These identities and the `performance_claim: none` status are recorded in
change 0273 (SHA-256
`4f634da9ec2631b404f7f8dbd6c461c68dc0b230e83b311d16ec2e5cbbc050b5`).

The operation queries the existing section owned by paragraph 64 and changes
columns from 2 to 3 with 240 twips of space and a separator
(`:50330-50354`). Its timed phases are preparation, source-backed open,
section query, edit, commit, and sequential publication
(`:50410-50461`). Correctness after the timer checks semantic reopen, exact
main-Part-only mutation, untouched ZIP members and order, headers/footers,
media, opaque fragments, source immutability, forward/inverse/stale patches,
and publication inverse restoration (`:49859-49991,50463-50505`). The route
requires exactly one retained main-Part load and deterministic cache counters
(`:50506-50531`), nonempty ordinary source reads, stable output digests, and a
64 KiB sink bound (`:50532-50650`). No-op, foreign-source, stale-source,
signed-source, limit, and partial-sink refusals are explicit in
`:50184-50316`.

Change 0273's only performance observations were a dirty CPU-2 smoke
(one warmup, five samples: median total 4.826203 ms, query 1.256769 ms,
publication 2.932747 ms) and a contaminated whole-process five-sample profile.
It explicitly makes no latency, allocator, RSS, I/O, decompression, copy, or
broad DOCX claim. There is no same-corpus eager control, so a future run must
report this as a source-backed phase baseline unless an eager control is first
added and measured.

## Evidence limits and corpus coverage

The current DOCX one-edit runner uses a separate fixed media-rich corpus built
by `build_docx_source_edit_corpus` (`tools/perf-baseline/src/lib.rs:16966-17014`):
200 paragraphs, eight incompressible 2 MiB media Parts, and a retained source
archive SHA-256 `a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.
It is the current one-Part edit/save path, not a section-layout or multi-Part
semantic edit. The current PPTX one-edit and same-slide batch cases share the
200-slide media-rich generator described above; their retained one-Part and
same-slide results are historical controls for the multi-slide requalification.

The old PPTX records are media-heavy and exercise same-topology existing-slide
text replacement, which is a good high-cost preservation workload. They do not
cover real-producer PPTX, notes/layout/chart/media topology edits, cold
filesystem input, high-latency range input, native Office save, or concurrency.
The old DOCX record 0039 similarly covers only one main-document paragraph edit
on a different 200-paragraph/eight-2 MiB-media corpus; its p50 223.183 to
5.732 ms and 17 to 1 materialization result is not evidence for section-layout
editing. Its report SHA-256 is
`a5f8e73791b418543ce0757965ea82b5089ad78177257ce37993efca82fe8c0e`.

For retained PPTX context, change 0044 reports p50 296.590 to 8.545 ms and
229 to 2 materializations for one shape; change 0063 reports p50 322.306 to
8.206 ms and 229 to 2 for a same-slide eight-shape batch. Their report
SHA-256 values are respectively
`bd153d442132c3a5856547f665f3d763a40b059486f086bfc9c80c1b879653ba` and
`994b190eb07eb8f374309f63317b7484eaf0eccfa62d6af1e65202878d7c1a4f`.
They use the old PPTX source archive and binaries, so they rank cost but do not
establish current timings.

The current source-backed edit loops use an instrumented in-memory positional
`ReadAt` and a bounded counting sink, not physical filesystem, range-latency,
remote, cold-cache, or native-producer inputs. The current PPTX edit results
also expose no operation metrics or allocator/RSS measurement. Logical source
read counters must not be described as device I/O. The synthetic generators
prove deterministic semantic and preservation contracts, but not producer
representativeness. Selector registration or a completed correctness test does
not equal measured CRUD coverage; the separate 0906 registry census makes the
same distinction.

## Reproducible next capture

Use a clean checkout of the accepted Office source after the pending target
integration, bind the source manifest and resulting release binary to a fresh
receipt, and do not run from the dirty shared iWork checkout. For the first ROI
packet, run four legs on one fixed CPU in eager/source/source/eager order, each
with 30 untimed warmups and 200 retained samples:

```text
cargo run --release --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 30 --samples 200 \
  --case pptx_eager_multi_slide_batch_edit_save \
  --output pptx-eager-a.json

cargo run --release --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 30 --samples 200 \
  --case pptx_source_backed_multi_slide_batch_edit_save \
  --output pptx-source-a.json
```

Repeat the two commands as source B and eager B, preserving raw JSON, binary
SHA-256, source SHA-256, machine/toolchain, CPU placement, and corpus manifest.
Any fresh source/output identity difference from the retained fixture must be
investigated and explicitly re-pinned before comparison; it must not be hidden
by reusing the old digest. Require the 9-versus-229 materialization contract,
selected/unselected semantic oracle, source/sink counters, and bounded writes.
Report the guarded timer scope separately from all untimed verification.

Then capture the unmeasured DOCX route with the same warmup/sample policy using
`--case docx_source_backed_existing_section_layout_edit_save` and no
`--shape`/`--payload` (the validator requires its fixed corpus,
`lib.rs:10051-10065`). Retain p50/p95/p99 for each phase, source reads and
cache counters, materialization count, output bytes/hash, sink distribution,
and all refusal/preservation gates. Do not call that result a speedup without a
same-corpus eager baseline. Only after these in-memory captures should the
winning route be extended to filesystem, cold, and caller-supplied range
sources required by the goal.

Audit provenance: `docs/performance/results/change-0906/source-context.json`
SHA-256 `710a84f9a5ec6d4e8693da2c8fd5d6ba55917348df4b7642c8e34ee7b3603b02`;
`docs/performance/0906-office-coverage-qualification.md` SHA-256
`3a5f74bcdaff449b8c9eab805dda6b9cc67641f749e3cd84163eb90e8c009176`.
