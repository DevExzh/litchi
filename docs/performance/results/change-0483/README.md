# Change 0483: bounded DOCX tail append evidence

This bundle is the reproducibility boundary for the focused DOCX main-story
tail-append comparison. It is being prepared for a same-executable comparison
between the existing materialized paragraph-copy route and the bounded
caller-supplied plain-text tail-append route.

The [source checkpoint](docx-source-checkpoint.md) has passed the accepted
DOCX feature configurations, lints, documentation, non-iWork workspace and
boundary checks, five focused benchmark tests, independent LibreOffice
readback, and two 10,000-iteration sanitizer smoke runs. The evidence helpers
have separate corruption tests. Formal captures and performance claims remain
pending; the full authored-stream and broader CRUD objective remains open.

The comparison uses deterministic DOCX/OOXML/OPC/ZIP sources containing 64,
8,192, or 131,072 plain body paragraphs and one opaque 32 KiB member. The
first paragraph's text is the explicit append text. The materialized route
copies that paragraph through the existing paragraph-copy transaction; the
bounded route receives the same UTF-8 text. Independent checks authenticate
the source archive, output main XML, semantic order and text, untouched ZIP
members, source immutability, and each route's source-bound proof.

The initial matrix measures only the complete total lifecycle. It has two
fresh instrumentation processes, three paragraph counts, and four ordered
arms: A1 materialized, B1 bounded, B2 bounded, and A2 materialized. Each case
has three warmups and 30 retained samples on CPU 2, for 24 processes and 720
samples. A1/B1 use forward count order; B2/A2 use reverse order. Normal
processes provide elapsed time and `/usr/bin/time -v` RSS. Allocator processes
provide operation-scoped allocation counters. The source adapter and scalar
short-write sink report logical read and output-write observations.

The timed region includes source/package admission, route preparation, commit,
sequential publication, sink digest finalization, and owner destruction. Corpus
construction, independent output reopening, ZIP member inspection, semantic
oracles, and report serialization remain outside the timed region. Allocator
and process counters do not establish a complete process RSS or arbitrary
DOCX memory bound; caller storage, package metadata and ZIP codec state remain
separate costs. The route's typed admission/refusal contract limits the scope
of any performance result.

## Reproduction

The coordinator must first confirm the harness CLI and report schema described
in [harness-contract.md](harness-contract.md), complete the pilot and required
gate receipts, and write `validation-plan.json` plus `fuzz-plan.json` from
those actual retained labels and paths. The freeze command refuses to invent
labels when either plan is absent. Then run each heavy command serially under
the retained CPU lock:

`write-plans.py` performs that receipt-only derivation. It requires the
accepted normal/allocator build custody, explicit final gate labels, explicit
pilot specifications, a classification and reason for every other completed
validation receipt, and one actual accepted prepared/build/smoke fuzz receipt
set. It refuses missing, failed, source-different, or synthetic inputs and
does not execute any command. The pilot argument convention is
`LABEL:INSTRUMENTATION:ROUTE:COUNT:SAMPLES:WARMUPS:REPORT_PATH`; the retained
pilot gate may use the direct harness form or the exact
`/usr/bin/time -v -o PILOT.resource /usr/bin/taskset -c 2` wrapper. Fuzz
receipts use `LABEL=KIND:PATH`.

For example, after the accepted receipts exist, the coordinator supplies the
actual labels and report paths (the values below are placeholders for the
coordinator's real retained records):

```sh
python3 -B docs/performance/results/change-0483/write-plans.py \
  --attempt accepted \
  --required-label build-normal-accepted \
  --required-label build-allocator-accepted \
  --required-label <accepted-required-gate-label> \
  --pilot <label>:normal:materialized:64:1:1:<pilots/report.json> \
  --classify <developmental-label>=developmental:<recorded reason> \
  --seed-manifest fuzz/seed-manifest-v2.json \
  --generator fuzz/generator-v2.json \
  --fuzz-receipt <prepared-label>=prepared:fuzz/accepted/prepared.json \
  --fuzz-receipt <build-label>=build:fuzz/accepted/build.json \
  --fuzz-receipt <smoke-label>=smoke:fuzz/accepted/smoke.json
```

Repeat `--pilot`, `--required-label`, and `--classify` for the complete
retained set. For a long receipt directory, `--classification-file` accepts a
JSON object mapping each non-final receipt label to
`{"classification":"developmental"|"historical","reason":"..."}`. The
planner writes both plans only after all checks pass.

`validation-plan.json` contains `required_labels`, `pilot_labels`, an `argv`
object whose keys are exactly their union, a `pilot_reports` object, and a
`developmental` object. Each argv list is the command recorded by the gate;
each pilot report reference includes its path/metadata, report spec, and
sample/warmup counts. A pilot argv is exactly the harness command
`BINARY --route ROUTE --counts COUNT --samples SAMPLES --warmups WARMUPS
--json ABSOLUTE_REPORT_PATH`; its executable, route, count, sample shape, and
report suffix must agree with the pilot spec. The report is checked against the
same schema, generated corpus, XML, member, semantic, and route oracles used
for formal captures. Every retained non-final gate is classified in
`developmental` with a reason, its source-before/source-after hashes, and the
recorded difference from the accepted source. `fuzz-plan.json` contains
`required_labels`, `seed_manifest`, `generator`, and `receipts`; each receipt
carries its retained path, byte count, SHA-256, label, and kind. Freeze runs
the read-only build, gate, pilot, and fuzz custody checks before writing the
immutable protocol. The verifier also requires every retained build,
validation, and fuzz receipt to finish before the freeze timestamp, and every
formal capture to start after it. Final gates must be source-stable and bind to
the accepted source manifest; historical/developmental receipts remain
retained with an explicit classification and may have either outcome.

```sh
python3 -B docs/performance/results/change-0483/build.py --attempt accepted
python3 -B docs/performance/results/change-0483/capture.py --attempt accepted --freeze
python3 -B docs/performance/results/change-0483/capture.py --attempt accepted --arm a1
python3 -B docs/performance/results/change-0483/capture.py --attempt accepted --arm b1
python3 -B docs/performance/results/change-0483/capture.py --attempt accepted --arm b2
python3 -B docs/performance/results/change-0483/capture.py --attempt accepted --arm a2
python3 -B docs/performance/results/change-0483/analyze.py
PYTHONDONTWRITEBYTECODE=1 python3 -B docs/performance/results/change-0483/test_evidence.py
python3 -B docs/performance/results/change-0483/verify.py
```

Every attempt uses exclusive JSON, text, report, and binary paths. Development
attempts must use a distinct `--attempt` token; their failures remain in the
bundle and cannot replace accepted receipts. The frozen protocol binds every
owned helper hash, exact capture and gate argv, environment, source manifest,
binary metadata, route labels, pilot and required validation labels, fuzz
seed/receipt custody, and sample matrix.

The bundle is portable after the temporary executables and lock are removed.
The final accepted bundle must retain `SHA256SUMS`, the final validation
receipts, independently generated measurements, and copied-bundle corruption
refusal evidence before it is sealed.

The settings/MCE fuzz extension is generated separately with
`python3 -B docs/performance/results/change-0483/fuzz-seeds.py
--extend-settings`. It authenticates the original 30 main-story seeds and
leaves `fuzz/seed-manifest.json` and `fuzz/generator.json` unchanged, then
writes the 62-seed `fuzz/seed-manifest-v2.json` and
`fuzz/generator-v2.json` records. The fuzz runner inventories every file in `fuzz/seeds`; the accepted custody
plan binds that inventory to the v2 manifest and generator explicitly.

No performance claim is authorized until the accepted source, binary builds,
formal captures, independent analysis, focused tests, required Rust gates and
final portability checks are all present and reviewed.

## Settings fuzz corpus and temporary storage

The expanded fuzz corpus contains 62 seeds. The original 30 main-story seeds
and their original manifest remain unchanged; 32 additional settings seeds
cover both Word dialects and Store/Deflate entries. They exercise empty
settings, protection flags, duplicates, valid MCE fallback, unsupported
`MustUnderstand`, inherited namespaces, and directive amplification. The
combined records are `fuzz/seed-manifest-v2.json` and `fuzz/generator-v2.json`.
Pass those paths through `write-plans.py --seed-manifest` and `--generator`,
with `--seed-root fuzz/seeds`. The runner inventories the complete seed
directory. Generator history and intermediate attempts remain retained.

The final-32 all-feature test run exhausted the user quota on `/tmp` after
successful compilation. `temporary-storage-migration.json` records the move
of this batch's scratch to `/home/zhuhe/.cache/litchi-goal-0483`; earlier
helper versions remain content-addressed in `driver-history`. Subsequent
file-writing test commands explicitly set `TMPDIR` to that directory's
`test-tmp` child. No unrelated files or shared build caches were removed.
The source-stable final-41 rerun passed all 1,378 tests with the same 31
ignored documentation examples.
