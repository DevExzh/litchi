# 0556 prospective measurement plan: XLSX provenance merge

status: prospective performance plan; no performance capture, claim, or adoption

scope: source-backed XLSX cell-value edit/save; OLE2/OOXML priority; ODF deferred; iWork excluded

This plan evaluates one private boundary:
litchi_xlsx::cell::Store::merge_omitted_cells in
crates/litchi-xlsx/src/cell.rs:805, called by try_rewritten_value_cells in
crates/litchi-xlsx/src/cell_values/snapshot.rs:736. The candidate source
review and this measurement plan are separate measurement prerequisites; neither may be replaced by a lower instruction count or by a
missing symbol.

## Question and current evidence

The existing path performs a complete rewritten-output validation, reduced
readback, and source/parsed eligibility checks. merge_omitted_cells then counts
the source records selected by ordered omission rectangles, reserves a combined
vector, moves parsed records, clones omitted source records, copies merge ranges,
and calls Store::from_unsorted. The latter sorts and duplicate-checks cells and
rows, rebuilds the cell-row index and stored, content, and styled bounds, and
constructs the merge index.

The candidate hypothesis is a checked private linear merge of the two
address-ordered cell sequences. It may remove a second cell sort and duplicate
scan while retaining the complete validator, reduced parser, omission producer,
source identity check, fallback route, merge-index construction, publication
checks, and all resource boundaries. An uncertain precondition must return the
existing optional refusal and use the complete parser.

The retained 0550 owner profile reports provenance merge at 4.6405--6.0308%
of inclusive MultiSourceEdit::commit instructions across the four shapes. That
is a nested instruction attribution, not a removable fraction, elapsed share,
or expected speedup. The same profile reports Store::from_unsorted inclusive
work, Cell::clone, and memcpy underneath the merge; those rows must remain
separate and must never be added together. The current source identity is HEAD
3210f07ac7a9daa2686e781645f6539a156a26ce; the source and lock manifests in
this bundle must be frozen again immediately before the campaign's first build.

## Preconditions and unresolved plan issues

The source-proof report must establish all of the following before a candidate
is measured:

* parsed.cells and source.cells are strictly address-ordered at this call
  site, including every parser and fallback route that can construct a Store.
  A current from_unsorted call is useful evidence, but it is not by itself a
  proof for a new constructor or future call path.
* Omission rectangles are ordered, non-overlapping, single-row rectangles with
  the same half-open semantics used by omitted_entries. The lower-column address
  on a later row must remain valid. Equal addresses, out-of-order input,
  omitted-source cardinality mismatches, and parsed collisions must decline
  before a fast result is published.
* Rows remain unique and in the same semantic order. The existing constructor
  sorts and duplicate-checks rows; a fast constructor must preserve that
  validation or decline. It must not silently rely on a row invariant that the
  source proof did not establish.
* Every Stored field is preserved, including style, shared-string lineage,
  inline-rich state, formula ranges, shared-formula storage, metadata, and the
  Cell variant. cell_rows, all three calculated bounds, the declared dimension,
  rows, columns, defaults, merge ranges, and merge-index behavior must be
  identical to the complete parse.
* Exact checked reservation, overflow, allocation-failure, execution,
  cancellation, and error-precedence behavior remains intact. A failed fast
  attempt must leave the complete fallback available and must not publish
  partially moved state.

The candidate must also document how it handles an inline or out-of-line
merge_omitted_cells symbol. If the symbol is inlined or absent after the
change, positive assembly and caller mapping are required; absence is never a
zero-cost substitution. The rejected 0552 and 0553 compact collectors and the
0555 OLE2 physical-marker candidate are closed experiments and are not
enablers for this work.

This preparation packet deliberately leaves four choices open for the later
measurement batch. Before any measured candidate run, the coordinator must
record the decision and its source/hash bindings in the frozen plan: (1)
whether every case/shape gets its own child or only local-attribution jobs do,
(2) the baseline noise result and resulting primary floor, (3) whether the
harness schema adds the two sequential allocation intervals, and (4) whether
the target remains a stable Callgrind/heaptrack frame after code generation.
These are real evidence dependencies, not reasons to invent a gate in this
preparation document. Preparing this plan involved no build or performance
capture; the separate root-run correctness builds and tests are retained under
quality-attempts and summarized in this bundle's README.

## Frozen corpus and scenario matrix

Use the existing deterministic generator
litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1 through
build_xlsx_cell_crud_corpus. Freeze archive, member, worksheet, media,
vendor-extension, update-list, and expected-output identities before the first
build. The four required shapes are:

| Shape | Existing corpus characteristics | Required purpose |
| --- | --- | --- |
| medium | Four 48x48 sheets with the fixed media members | ordinary source-backed edit |
| dense-sparse | Dense 128x128 sheet, stride-4 and stride-8 sparse sheets, and a diagonal sheet | different omission density and row/column locality |
| noncompact | Medium cells with alternating qualified XML forms | namespace and lexical preservation control |
| vendor-extension | Medium cells plus opaque extension XML/binary parts and a relationship | unknown-member and relationship preservation control |

The primary timing matrix is eight rows: each primary case below is run on all
four shapes. Controls are also run on all four shapes unless a receipt records
that a case is inapplicable; an inapplicable case cannot be silently dropped.

| Class | Existing case | Updates | Reason |
| --- | --- | ---: | --- |
| Primary | xlsx_source_backed_cell_values_one_edit_save | 1 | smallest changed-cell path with omitted source records |
| Primary | xlsx_source_backed_cell_values_one_percent_edit_save | ceil(total stored cells / 100) | proportional changed-cell path and the retained 0550 owner profile case |
| Control | xlsx_source_backed_cell_values_batch_edit_save | litchi_xlsx::cell_values::MAX_BATCH_EDITS | larger value-only commit and reservation control |
| Control | xlsx_source_backed_cell_values_multi_sheet_edit_save | 2 on distinct sheets | multi-sheet atomic commit and calculation invalidation control |
| Control | xlsx_source_backed_managed_cell_values_one_edit_save | 1 | managed budget and publication planning control |
| Control | xlsx_source_backed_managed_cell_values_one_percent_edit_save | 1% | managed budget under the larger source path |
| Control | xlsx_eager_cell_values_one_edit_save | 1 | negative-path control; it does not use provenance merging |
| Control | xlsx_eager_cell_values_one_percent_edit_save | 1% | negative-path scale control |

The timed workflow is source-backed open, selector planning, edit staging,
commit, and sequential publication. Corpus construction, expected eager output,
source-layout ranges, correctness oracles, output hashing, semantic reopen,
untouched-member comparison, and report construction are outside the elapsed
clock, as in the existing cell-values runner. The runner must still execute
those checks on every measured iteration or on an explicitly equivalent
independent validation sample. Instrumented elapsed time is never compared to
normal elapsed time.

Correctness-only controls must cover exact no-op commit/save, clear versus
remove, inverse/recommit, source-version and foreign-source conflicts, managed
output-budget refusal, cancellation/execution interruption, partial
sequential-sink failure, and malformed or unsupported source records. They are
not substitutes for the eight primary rows. A source formula, shared-formula
closure, rich inline string, unknown cell payload, unsupported metadata, or
uncertain omission proof must use the complete writer/fallback and preserve its
existing diagnostic.

## Source, binary, and target custody

The campaign driver should be the small serial driver derived from the existing
0555 run.py interface. It must not copy the large benchmark implementation.
Use only these owned paths for generated output and temporary storage:

    /home/zhuhe/litchi-goal-0556-target
    /home/zhuhe/litchi-goal-0556-target/tmp

Before the first build, freeze plan.json, the full source manifest for the
workspace crates and tools/perf-baseline, the ignored workspace Cargo.lock
copy, accepted-ADR manifest, corpus manifest, driver, and analyzer inputs.
Every receipt must bind the output stage and live execution stage, source
manifest, lock hashes, plan and script hashes, binary hash, command, relevant
environment, and artifact inventory. The baseline and candidate source trees
must be independently hashed. A candidate patch and every failed preparation
or quality attempt remain retained.

The normal binary is built with the existing command:

    env TMPDIR=/home/zhuhe/litchi-goal-0556-target/tmp CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline --target-dir /home/zhuhe/litchi-goal-0556-target

The allocator binary is built separately:

    env TMPDIR=/home/zhuhe/litchi-goal-0556-target/tmp CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline-alloc --features allocator-metrics --target-dir /home/zhuhe/litchi-goal-0556-target

Before each child, verify the selected source manifest and ignored lock. The
normal and allocator binaries are retained by stage and checked by hash before
cleanup. No shared /tmp output, unbound binary, or ambient compiler process
may enter the evidence.

## ABBA native and allocator execution

Use CPU 2, fresh child processes, two repeats, and the existing high-sample
configuration: 20 warmups and 1,000 measured samples for the normal binary;
three warmups and 30 measured samples for the allocator binary. A child should
select one case and one shape when local attribution or RSS is required. The
underlying normal invocation for one such job is:

    taskset -c 2 /home/zhuhe/litchi-goal-0556-target/retained/baseline/normal --case xlsx_source_backed_cell_values_one_percent_edit_save --xlsx-cell-crud-shape dense-sparse --warmup 20 --samples 1000 --json docs/performance/results/change-0556/baseline/native-r1-dense-sparse-one-percent.json --corpus-manifest docs/performance/results/change-0556/baseline/native-r1-dense-sparse-one-percent.catalog.json

The allocator invocation has the same selectors and uses the retained allocator
binary with warmup 3 and samples 30. The driver should record GNU-time
max_rss_kib, user/system time, and all case result vectors in separate
receipt-bound sidecars. Report p50, mean, p95, p99, min, and max for every
row; do not merge shapes or hide an adverse row in a geometric mean.

The exact stage order for both native and allocator lanes is:

    baseline R1 (A1)
    candidate R1 (B1)
    candidate R2 (B2)
    baseline R2 (A2), using the retained baseline binary with candidate execution-stage identity

The command form for the retained final leg is, for example:

    python3 -B docs/performance/results/change-0556/run.py native --stage baseline --execution-stage candidate --repeat 2

The other legs use the same driver with the corresponding stage and repeat. The
driver must reject an execution-stage/source-manifest mismatch and must retain
every failed child. A baseline-only noise pilot using the same normal binary,
matrix, child isolation, CPU, warmups, and samples must be completed before the
candidate gate is frozen; its result is descriptive and cannot be replaced with
a retained 0550 repeat-drift number.

## Practical admission thresholds

The retained 4.64--6.03% owner attribution does not justify importing the
0555 OLE2 3% primary threshold unchanged. Before any candidate output is seen,
freeze a short threshold record from the baseline-only noise pilot:

* Let N be the largest absolute relative p50 change between the two matched
  baseline-only repeats among the eight primary rows, with the row-level value
  retained. If the pilot is too unstable to establish N, stop with no adoption
  decision.
* The prospective primary floor is
  max(1.0%, 3*N, 50,000 ns / baseline p50) for each row and repeat. The 1%
  term avoids treating a sub-noise change as useful; the 50 microsecond term
  prevents a tiny absolute change in a short operation from passing. These are
  a proposed practical floor and must be recorded in the frozen plan.json
  before candidate capture; this document itself makes no result claim.
* Every one of the eight primary p50 rows must improve by that floor in both
  paired repeats. Each primary mean must change by no more than +5%.
* Every control p50 and mean, every process-RSS row, and every allocator row
  must change by no more than +5% in each paired repeat. Allocation rows include
  calls, reallocations, allocated/deallocated bytes, live endpoints, and the
  checked region peak where measured.

The exact floor is a preregistered gate, not a post-result relaxation. Every
absolute adverse or same-build drift above 5% remains in the machine-readable
report and receives individual review. No selective rerun, mean-only rescue,
geometric-mean masking, or profile-only adoption is allowed. A passing numeric
matrix still requires source proof, correctness, preservation, quality, and
positive mechanism evidence.

## Operation-local allocation evidence

The existing commit_allocation_metrics interval in
tools/perf-baseline/src/lib.rs:41924 starts before the edit.set loop. It
therefore combines edit staging with the commit and cannot answer how much
allocation work belongs to the provenance merge. Its values remain useful
end-to-end diagnostics but are not merge-local evidence.

Make one harness-only, receipt-bound amendment before the first candidate build.
Keep the commit_ns timer unchanged, so native scope remains the real end-to-end
operation. Finish a non-nested allocator region immediately after the edit.set
loop, then begin a second region immediately before edit.commit() and finish
it immediately after the returned commit. Report staging_allocation_metrics and
commit_core_allocation_metrics separately; never nest allocation_metrics.begin
regions. This removes the known staging contamination without changing
production code or public APIs. The commit-core region still includes all
commit work and its returned object's temporary destruction before return, so
it is not by itself a merge-only counter.

For operation-local target attribution, use the installed heaptrack tool on a
fresh normal release child with one case, one shape, zero warmups, and one
measured sample. The command form is:

    taskset -c 2 /usr/bin/heaptrack --record-only -o /home/zhuhe/litchi-goal-0556-target/heaptrack/baseline-r1-dense-sparse-one-percent.data.gz /home/zhuhe/litchi-goal-0556-target/retained/baseline/normal --case xlsx_source_backed_cell_values_one_percent_edit_save --xlsx-cell-crud-shape dense-sparse --warmup 0 --samples 1 --json /home/zhuhe/litchi-goal-0556-target/heaptrack/baseline-r1-dense-sparse-one-percent.json --corpus-manifest /home/zhuhe/litchi-goal-0556-target/heaptrack/baseline-r1-dense-sparse-one-percent.catalog.json
    heaptrack_print -f /home/zhuhe/litchi-goal-0556-target/heaptrack/baseline-r1-dense-sparse-one-percent.data.gz --filter-bt-function 'litchi_xlsx::cell::Store::merge_omitted_cells' --print-allocators --print-temporary

Run the same target-filtered fresh-child jobs for both primary cases, all four
shapes, both repeats, and both ABBA stages when the pilot permits profiling.
Retain the unfiltered heaptrack data, filtered text, command output, and
receipt. The filtered report supplies inclusive allocation-call attribution for stacks
containing the target; a second filter for Store::from_unsorted may explain
descendants. Retain complete allocation-cost folded stacks or prove that text
limits did not truncate the selected rows before aggregating counts. Target
requested-byte totals require a separately verified raw-event aggregation;
printed peak consumption must not be relabeled as cumulative allocated bytes.
Until that aggregation is available, report target allocated bytes as unavailable
and retain the measured commit-core byte counters with their broader scope. Caller ancestry and symbol
identity must be retained so setup or an unrelated call cannot be counted as
the merge.

Heaptrack process peak is not a merge peak. Do not use --merge-backtraces for
a peak conclusion (the tool explicitly warns that merged peak consumption is
incorrect), and do not turn filtered process peak into an operation-local
peak. The commit-core allocator interval supplies a bounded operation-level
peak with its scope stated; any target-only peak remains unavailable unless a
separate test-only allocator census is added and independently bound. A
missing target frame, zero filtered row, or unavailable stack is indeterminate,
never evidence of zero allocation.

## Callgrind attribution and static mapping

If the native/allocation pilot passes, collect two profile views for each
primary case and shape: the exact owner
litchi_xlsx::cell_values::source::MultiSourceEdit::commit and the exact target
litchi_xlsx::cell::Store::merge_omitted_cells. Use zero warmups and five
positive timed samples per fresh child, with setup and termination dumps
retained and classified separately. The target view's command form is:

    taskset -c 2 /usr/bin/valgrind --vgdb=no --tool=callgrind --collect-atstart=no --toggle-collect='litchi_xlsx::cell::Store::merge_omitted_cells' --zero-before='litchi_xlsx::cell::Store::merge_omitted_cells' --dump-after='litchi_xlsx::cell::Store::merge_omitted_cells' --callgrind-out-file=/home/zhuhe/litchi-goal-0556-target/profile/target-r1-dense-sparse-one-percent.callgrind --dump-instr=yes --dump-line=no --compress-pos=no --collect-jumps=yes /home/zhuhe/litchi-goal-0556-target/retained/baseline/normal --case xlsx_source_backed_cell_values_one_percent_edit_save --xlsx-cell-crud-shape dense-sparse --warmup 0 --samples 5 --json /home/zhuhe/litchi-goal-0556-target/profile/target-r1-dense-sparse-one-percent.json --corpus-manifest /home/zhuhe/litchi-goal-0556-target/profile/target-r1-dense-sparse-one-percent.catalog.json

The owner view substitutes the owner string in all three Callgrind collection
options. Use the same ABBA stage order as the native lane; the final baseline
profile uses stage baseline and execution-stage candidate. Before capture,
record nm -C output and require a unique baseline target symbol. After each
stage, inspect emitted assembly with the existing bounded assembly tooling and
retain the requested source-to-instruction mapping.

The profile consumer must verify positive target ancestry and expected changed
worksheet calls for every positive timed dump. It must report self and
inclusive Ir separately, preserve every dump, and keep from_unsorted, clone,
copy, and owner rows nested rather than summing them. A target that was inlined,
renamed, or split is accepted for mechanism evidence only with an exact mapped
caller/assembly proof. A symbol disappearance without that proof is
indeterminate. The owner profile is a mechanism and attribution gate, not an
elapsed-time claim.

## Correctness, preservation, and quality gates

The candidate source proof and targeted private tests must pass before matched
capture. The direct differential suite must compare the current complete route
and the candidate route for valid gaps, same-row gaps, lower-column next-row
addresses, empty records, inferred references, and all supported Stored
facets. Refusal tests must assert the optional fallback and then compare the
complete output and exact public error. The public XLSX suite must additionally
verify:

* exact worksheet/package bytes where the existing contract promises them,
  semantic reopened cells and formulas, output determinism, and untouched
  member hashes;
* vendor extension parts, relationships, styles, media, namespace choices,
  lexical forms, declared dimensions, row/column defaults, merge ranges, and
  calculation-chain invalidation;
* source sharing and exact no-op behavior, inverse patch restoration,
  stale/foreign source conflicts, atomic failure, cancellation, execution
  limits, output-budget refusal, allocation failure, and sequential non-seek
  publication;
* duplicate addresses, unsorted rows/cells, out-of-order omission rectangles,
  unsupported cells, shared formulas, rich inline values, malformed XML, and
  all existing first-error precedence cases.

At minimum, retain receipt-bound runs of the existing checks below for both
candidate and restored final source as applicable:

    cargo test -p litchi-xlsx --all-features
    cargo clippy -p litchi-xlsx --all-features --all-targets -- -D warnings
    cargo fmt --all -- --check
    cargo check --workspace --all-features
    RUSTDOCFLAGS='-D warnings' cargo doc --workspace --all-features --no-deps
    cargo clippy --workspace --all-targets --all-features -- -D warnings
    python3 -B tools/check_crate_boundaries.py
    python3 -B tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --evidence-root . --mode strict

Run the relevant native Office fixtures and available fuzz or sanitizer checks.
If a tool is unavailable, retain the availability result and make no claim from
it. The quality lane must cover the exact candidate source and the restored
baseline source; prior 0555 quality does not silently substitute for fresh
candidate correctness.

## Disposition and terminal custody

The numerical analyzer must produce machine-readable before/after rows with
source, corpus, binary, environment, uncertainty, and stage identities. It must
distinguish normal elapsed, allocator vectors, process RSS, commit-core
allocation, heaptrack filtered attribution, and Callgrind Ir. No ODF workload
or claim enters this campaign.

Any failed mandatory native, memory, source-proof, correctness, preservation,
profile-attribution, or quality gate rejects the candidate. Restore the exact
baseline source, rerun the applicable final quality and strict-claim checks, and
retain the decision, failure rows, and restoration evidence. A passing candidate
still requires an independent review of every adverse row and a fresh recursive
evidence seal.

After the complete precleanup verifier passes, check all retained binary
hashes, accessible process references, owned Python caches, and the exact owned
target path. Remove only /home/zhuhe/litchi-goal-0556-target, retain the cleanup
receipt, and verify the final bundle recursively. No production Rust change,
performance claim, or adoption decision is made by this plan.
