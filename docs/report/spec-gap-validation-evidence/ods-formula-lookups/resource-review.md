# ODS lookup and reference-function resource and cache review

This review covers the resolver-aware and scalar implementations of
`ADDRESS`, `CHOOSE`, `HLOOKUP`, `INDEX`, `INDIRECT`, `LOOKUP`, `MATCH`,
`OFFSET`, and `VLOOKUP`. It applies ADR 0005's bounded work and storage rules
and ADR 0006's typed-failure, source-fence, and publication rules to the
lookup contract. This review changed no production or test source.

The disposition is **PASS** for the resource and demand-cache contract on the
current staged source. Table lookup control validation now checks known index
domain and table geometry before resolving a reference-valued mode or key, and
lifted reference arguments remain scalar-cell descriptors until that check has
completed. The focused semantic, limits, and oracle receipts pass on the
current staged source. The final seven-gate receipt also passes 1,700 package
tests with zero failures and zero ignored tests.

## Frozen source identity

The current staged freeze baseline is commit
`635fd2e1348b621426b50909cbd5765c91837306`. Complete source custody is
recorded in [`gates/freeze.json`](gates/freeze.json), which lists 62 selected
inputs. The relevant implementation, contract, and focused-test inputs have
these SHA-256 values:

| Frozen file | SHA-256 |
| --- | --- |
| `Cargo.lock` | `58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3` |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `4e2f383e9fa6a9aa80457b8b7b440957a01a45de625bfc765fb50038a4b18882` |
| `crates/litchi-ods/src/codec/formula/evaluation/lookup.rs` | `25c24f585b9d4ff74d57fa605d99caac033455297852eb7e8ffcae31fd4c2c67` |
| `crates/litchi-ods/src/codec/formula/evaluation/lookup/address.rs` | `7da925466d57320d4727f9ff992edc92ab64d7276fe32442464edaacee9771d8` |
| `crates/litchi-ods/src/codec/formula/evaluation/lookup/indirect.rs` | `10f5629a90463bbd5dd7f1fb41b298e6a2a800713af57c35e3c5c7241ac8b213` |
| `crates/litchi-ods/src/codec/formula/evaluation/lookup/search.rs` | `2b0920dfc73a1282a9db4272499f96be3a0a3811d08f4774f8464d5c95a384de` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `8c0fb5ade784a02bb46268dba93c4141ff787cae2dd70cc286abae9cfc32f12c` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup.rs` | `58bf26633bed3c443a2cadbee10a1b47eb6ec5b9683a2d56284932af7459c0be` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup/indirect.rs` | `e42c71c76b34200de355cc8cd293b8a3d2d2bf2da9bce140b6fc1b551f2b0c01` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup/lifting.rs` | `d202a46463fe91173bd42027f8a383012a43a6b5ce158db231b1faa2a03d0571` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup/search.rs` | `3e2f001f91dd9d690730798c0a22ecae4867e427ae6b5f333ef07e6fa0b8b225` |
| `crates/litchi-ods/docs/FEATURE_MATRIX.md` | `48b5c5cc25e51bd216dc08220b4da862d6a86675923f400f339bf8cfa59c23c5` |
| `crates/litchi-ods/tests/ods_formula_lookup_evaluation.rs` | `5537e8820144cc587c6f13c794b3d02edc12fdf0b0c988c26e3919f77faa0965` |
| `crates/litchi-ods/tests/ods_formula_lookup_limits.rs` | `16a9ff3fdf0bb3c65148f68e91ca1ae2455c26c5ba6cbda03377554efccf9899` |
| `crates/litchi-ods/tests/ods_formula_lookup_oracle.rs` | `19dbe697d9064fe80a3236055d5f3950b447ac08fc48bb58d704d44c0441ca37` |
| `docs/report/spec-gap-validation-evidence/ods-formula-lookups/contract.md` | `b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf` |
| `docs/report/spec-gap-validation-evidence/ods-formula-lookups/lookup_oracle.py` | `a9e33442aec45cce4ee361f2bd54daaeb22d7a0778e39b7669cc3ab725f8054b` |
| `docs/report/spec-gap-validation-evidence/ods-formula-lookups/lookup-goldens.json` | `ad2a9bd6c92823b1a40d012d2c64bea66a7b9a0b87c998bb16fc4fcd89887758` |
| `docs/report/spec-gap-validation-evidence/ods-formula-lookups/coverage-scope.json` | `07e01ded006916789c442cfdc1d22c2a522f955a8a141c61fcd743eb9c4308b7` |

The isolated gate checkout uses the frozen `Cargo.lock` hash shown above. The
ambient workspace lockfile is a different copy and is not the dependency
input for this custody claim. Compared with diagnostic attempt 4, the current
freeze includes the ADDRESS character-loop normalization, lookup preflight
formula-error preservation, owned H/V key transfer, the focused formula-error
key test, and the immutable coverage-scope input. The changed implementation
hashes are `lookup/address.rs` `7da92546…`, `value/lookup.rs` `58bf2663…`,
and `value/lookup/search.rs` `3e2f001f…`; the focused evaluation test is
`5537e882…`. Other selected hashes are unchanged.

## Resource and ownership checks

* **Admission and zero-read refusal:** Lookup dispatch performs the metadata
  preflight before search scanning or projected argument lifting. ForceArray
  data rejects `ReferenceList`, 3-D and invalid scalar shapes before cell
  reads; MATCH and LOOKUP validate the searched/result descriptor shapes and
  vector requirements before scanning. LOOKUP's conditional result-reference
  extension is intentionally decided only after the selected match, then its
  extension geometry and bounds are checked before the selected result read.
  ADDRESS, INDEX, OFFSET, and INDIRECT perform descriptor construction and
  geometry checks without reading cell values. Metadata work, source policy,
  sheet extent/name calls, area geometry, and cancellation remain charged. A
  computed child may read while producing the descriptor whose type or shape
  must be discovered; the zero-read guarantee begins once that refused
  descriptor is known. If a scalar lookup key already contains a formula
  error, preflight preserves that key error over a later metadata shape
  refusal without projecting or reading the key.
* **Lazy CHOOSE:** The index is visited first and only selected branches are
  evaluated or shape-probed. Unselected branches are not inspected for shape,
  provider errors, source references, or resolver work. Matrix indices select
  branches per demanded coordinate, while selected scalar references, arrays,
  lists, and errors retain their runtime identity for the caller.
* **Streaming search:** Search vectors and tables are traversed through the
  existing cell resolver in logical row/column order. Work is charged before
  each read; cumulative reference limits and cancellation are checked before
  and after the provider call, and successful reads are counted only after a
  successful return. Exact matching keeps fixed state and retains a formula
  error while continuing the required suffix scan; a later match cannot erase
  that error. Approximate matching uses bounded search state and applies the
  final-candidate type barrier. No range-sized candidate or text vector is
  materialized.
* **Selected results and derived references:** The selected lookup result is
  the only result cell read after a search unless the candidate is already the
  selected cell. LOOKUP result-reference extension checks checked geometry,
  bounds, work, and cancellation before constructing the derived descriptor.
  INDEX, OFFSET, and INDIRECT retain bounded area/record metadata and borrow
  stable resolver sheet names; they do not retain parser-owned strings or
  read cells during construction.
* **Text and conversion:** `read_to_element` enforces text limits and keeps
  provider text borrowed. Comparison and address parsing charge text work
  without cloning a provider string for every candidate. Formula errors retain
  their original conversion channel; malformed or non-finite text follows the
  contract's formula-error profile. The owned H/V path transfers a computed
  key into the shared finish step; borrowed control conversion may clone an
  owned computed Text value, bounded by the normal text and storage limits.
* **Bounded state and drop order:** Runtime area descriptors, AST shape and
  reference-kind stacks, projected output arrays, scalar-argument scratch, and
  CHOOSE lookup probes use checked capacities and storage reservations. Search
  data remains borrowed. The corrected lookup lifting code declares scratch
  reservations before their vectors, so vectors drop before their budget
  tokens. Computed CHOOSE shape probes retain only selected complete results,
  bound probe contexts and entries by evaluator limits, and release their
  depth reservation after restoring the suspended planner; sequential probes
  therefore cannot leak depth budget.
* **Typed precedence and fences:** Resolver `Unsupported`, resource,
  cancellation, source-version, and allocation failures remain
  `EvaluationFailure` values. They are not converted into formula errors,
  search fallback success, or catchable `IFERROR`/`IFNA` values. Formula
  errors are retained only where the contract requires a formula result. The
  outer value evaluator checks source identity and cancellation before and
  after evaluation and immediately before publication.

## Selector read-order clearance

`search::preflight` validates data and vector descriptors before projected
lifting. `table_controls` first consumes selector values that are already
available and applies the one-based index and table-dimension checks before it
resolves reference-valued controls. Consequently a reference mode such as
`D1` is not read for `=VLOOKUP(B1;A1:C4;0;D1)`; the known index refusal wins.

`lifting::select_search_argument` maps a shaped reference key, index, or mode
to a bounded `ScalarCell` descriptor without reading it. The borrowed search
entry point resolves that descriptor only after control validation, so failed
output coordinates remain read-free while valid coordinates retain their
normal key/search/result reads. The focused suite covers direct reference
selectors, reference-valued range flags, lifted references, literal arrays,
and mixed valid/invalid output coordinates.

## Demand-cache checks

Lookup functions are classified before projected-branch cache lookup and use
the same conservative classifier at apply-time. Complete ForceArray data and
result descriptors remain complete under projected lazy `IF`; a data range is
not implicitly intersected for one output coordinate. Direct local descriptors
may be cacheable within one source-fenced evaluation, while computed projected
source expressions remain uncached unless their invariance is proven.

Scalar search keys, row/column/index selectors, and range flags remain
position-sensitive unless they are direct coordinate-independent scalar
arguments. Nested conditional criterion propagation marks only complete data
and result slots; it does not make scalar selector descendants complete.
`MUNIT` remains excluded from full-argument propagation. `CHOOSE` retains
selected scalar branches only when the index and branch are invariant, and its
unselected branches never enter the cache walk. Descriptor results from
`OFFSET` and `INDIRECT` are not scalar-cache payloads.

The demand cache stores only bounded scalar values and formula-error payloads
with evaluator/source context. It does not retain arrays, references, borrowed
text, or typed failures. Cache hits occur before source argument scheduling,
but remain inside the outer source and cancellation fences.

## Validation evidence

The current candidate test inventory is:

* `ods_formula_lookup_evaluation`: 34 tests, including direct and lifted
  selector refusal cases;
* `ods_formula_lookup_limits`: 9 tests;
* independent lookup oracle: 127 observations across all nine functions; and
* native corroboration: 30 function observations, with 17 parity rows and 13
  documented divergences against the independent contract.

The focused execution passes 34 semantic tests, 9 limits tests, and the
independent 127-observation lookup oracle. The selector cases include
reference-valued range flags and shaped-reference lifted controls; mixed
coordinates read only the valid output's key, search, and selected result.
Formula-error keys also retain their original error over data shape refusal in
both evaluator modes without resolver reads.

The focused limits cover list/3-D/vector/table zero-read refusal, ordered
exact and approximate reads, selected-result reads, derived-reference
construction, borrowed text, cumulative work/read/storage limits, typed
provider failure, formula-error continuation, cancellation, source fencing,
projected cache reuse, position-sensitive scalar parameters, and sequential
computed-CHOOSE probe depth.

Diagnostic attempts 2 and 3 each completed all seven gates with 1,697 package
tests and zero failures or ignored tests. Attempt 4 completed all seven gates
with 1,699 tests before the owned-key transfer fix and is retained as a
superseded receipt under `gates/diagnostic-attempt-04/`. The current seven-gate
receipt passes with 1,700 tests and zero failures or ignored tests; standalone
source/receipt verification also reports all required checks passed with stable
sources. Attempts 5 and 6 are retained as superseded seven-gate receipts, and
attempts 1 through 4 remain available under their diagnostic directories. This
review makes no timing, throughput, RSS, or allocator-performance claim.
