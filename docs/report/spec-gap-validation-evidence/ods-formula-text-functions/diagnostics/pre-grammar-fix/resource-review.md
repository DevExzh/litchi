# ODS ordinary text-functions resource and cache review

This review covers the 26 ordinary OpenFormula 1.4 §6.20 text functions in
the scalar evaluator and the resolver-aware value bridge. It applies ADR 0005
bounded work and storage rules and ADR 0006 typed-failure, source-fence, and
publication rules to the frozen local text contract. The review changed no
production or test source.

The disposition is **PASS** for the frozen resource and cache boundaries. The
implementation streams reference text through matrix positions, retains no
input cell vector, preserves typed failures, and has no position-insensitive
text-result demand-cache path.

## Frozen source identity

The frozen baseline is commit
`8f09231e36982248eface4d599432143a67f6e49`. Complete source custody is
recorded in [`gates/freeze.json`](gates/freeze.json). The selected inputs have
these SHA-256 values:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `971f7589f5abd32930e039e7cb07a1b6f20dcab776836ade37327648ebc85847` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `3f59c37f883cd1a91278c224d3d97edcb86d5a6a3d7ea7c1cc4adfc7848cc55a` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `cd882fa914bffa26ede20c51a4483c53fc36f4070e37cfbfef2bfec6e2db69f3` |
| `crates/litchi-ods/src/codec/formula/evaluation/text.rs` | `7959904110633b298cc2fcfb3affbe205b91c8f408a864cb57dc7818648f696e` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/core.rs` | `e2832703658582c516a6605fc9197de6af8d636266777bca8e66ab5b025809ba` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/format.rs` | `23c0df92d8659ce9f5630b676fa1e1a1ebae471f110a2597ae92802d59d9619f` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/fraction.rs` | `a2196dd849e2d85faaacd49f876aba8920f9aff00b50e1fad474ae6c278f276e` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/search.rs` | `523b93ce563976bc617c626133e3dd90760d0580b63ef677a7eb73bda0f78d0f` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/unicode.rs` | `336584ece56ee00e9e68a0c2165b5396ac1f007d5ad671f1fe8eba30aac877b6` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/width.rs` | `69a4fae2bb52744e6fa041e3b5de2ffcb9ffa17f6fa7ed2139e433ffe9136b7c` |
| `crates/litchi-ods/tests/ods_formula_text_evaluation.rs` | `6b8a12108e27b609f5c90b52136173a22a5c79916d4a0e86314eb5b9fdec7620` |
| `crates/litchi-ods/tests/ods_formula_text_limits.rs` | `51e69b69947a404322329fab296cdfa352fc1d52bc6bb9c9e9b8028c2a3a1b1b` |
| `crates/litchi-ods/tests/ods_formula_text_native.rs` | `f480e81b3bc2a3a6ab9eabb7fbb254072e3f792068dd5da1719e961857794473` |
| `crates/litchi-ods/tests/ods_formula_text_oracle.rs` | `0e935d7154a804abe738685edc0f00d9aa7c545560fa3606b2e05798404b8f87` |
| `crates/litchi-ods/docs/FEATURE_MATRIX.md` | `89236557575bb3a51d858a5606e30a30a3cb8946c85d1d48b5b5bf93c03c84ef` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/contract.md` | `b0b7f6ffd8a33c93f98c9728c938eb2d55476680e11797a9b87dcf49223312e7` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/text_oracle.py` | `e00cc04f6b263000c3a3e5836968b340739a50a3adad99008074c2f4c77ba57f` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/text-goldens.json` | `67a358b923a5311a260c8dcb977a5bff49d6ef11f3b1504712e6f49e8265e857` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/native/cached-results.json` | `f28fbe4afecd3e92ba9c320113a84ab5a926346d05cdd5e42f114846b792fbc3` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/native/provenance.json` | `6b47bb8c77679e708bed8e83950ca9b409ef27cf3c0ec573548c0cd05f6ac3a4` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/unicode-data/UNICODE-LICENSE.txt` | `e7a93b009565cfce55919a381437ac4db883e9da2126fa28b91d12732bc53d96` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/unicode-data/generate.py` | `ecdd04283e0f01ed61e7af5aa3a9b11ddfc645db4d9570684f439116fca6e6d6` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/unicode-data/provenance.json` | `79feeef3c69b84875dda504cf309a914625b63de2e5f65d0b25818a6d19d0c6a` |
| `docs/report/spec-gap-validation-evidence/ods-formula-text-functions/unicode-data/source-receipts.json` | `4d8741f3fbaaddd6f39bf7f4cf7721273c4650e066d01418e66bcebaf3797096` |
| `Cargo.lock` (isolated gate copy) | `58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3` |

The isolated gate checkout uses the frozen `Cargo.lock` above. The ambient
workspace lockfile has SHA-256
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02` and is
not the dependency input for this review.

## Resource and ownership checks

* **Shape and read admission:** The text value bridge computes broadcast shape
  from runtime descriptors and checks cell-count and array limits before
  allocating its output. A known `ReferenceList` or multi-area reference is
  rejected by the shape/area selector before the text function scans it, so
  the refused descriptor causes zero resolver cell reads. A computed
  expression may perform the upstream work needed to reveal that descriptor;
  the zero-read guarantee begins once the refusal is known. Scalar projection
  reads only the selected cell, while matrix mode selects each required
  position according to the ordinary broadcast geometry.
* **Streaming matrix evaluation:** `map_text_function` keeps reference areas
  as descriptors and selects one element per output position. It does not use
  the generic `materialize_for_array` path and does not build an input cell
  vector. The reusable scalar-argument buffer is reserved once and drained
  for each output cell. Output cell work is charged before selection, and
  reference cell work is charged before every provider call.
* **Resolver and text boundaries:** `read_reference_cell` checks the
  cumulative reference-cell limit and cancellation before the resolver call,
  checks cancellation afterward, and increments the successful-read count
  only after the provider returns. `read_to_element` enforces the text-byte
  ceiling and retains provider Text as `TextValue::borrowed`; reference text
  is not cloned into a per-cell buffer. Inline conversion and every inspected
  input or output byte charge the evaluator's hierarchical work budget.
* **Unicode and search work:** Explicit Unicode mapping, case, width, search,
  replacement, trimming, and repeated-output loops use periodic execution
  checkpoints and checked output growth. SEARCH's KMP state is bounded by the
  admitted pattern, while FIND/SEARCH scans charge their inspected text. A
  skipped JIS pair and UTF-8 offset path still reaches the same checkpoints.
  Standard bounded string primitives have execution checks around their use.
* **Owned output and reservation order:** Transformations allocate only when
  a changed or owned result is required. `new_output` returns the storage
  reservation before the String, and callers retain that order so a failure
  drops the String before refunding its lease. CONCATENATE retains the full
  left `TextValue` while growing and replaces its lease only after a checked
  growth succeeds; chunked copies have post-copy execution checks. Mapped
  output, scalar argument scratch, formatter output, and all temporary
  vectors use fallible checked reservations with the buffer declared after
  its lease.
* **Formatter bounds:** Format-code bytes are charged before parsing. Grammar
  scans and every previously un-emitting positional scan checkpoint through
  `FormatChecks`; the formatter also checks at render boundaries. CountSink
  charges and admits each emitted byte before the second pass, rejects an
  output over the text ceiling before publication, and StringSink checks both
  sides of large copies. The exact fraction helper uses fixed `u128` state,
  no allocation, and a callback charge for each continued-fraction iteration;
  it replaces the million-candidate scan. Borrowed fraction pattern spans
  avoid render-time substring allocation. Exact arithmetic overflow remains a
  formula numeric error, while work, cancellation, allocation, and storage
  failures remain typed evaluator failures.
* **Formula and typed errors:** Formula errors are retained through the
  required argument and matrix-position evaluation. A later resolver,
  unsupported, resource, allocation, cancellation, or source-version
  failure remains an `EvaluationFailure`; it is not converted into a formula
  error and is not caught by `IFERROR`/`IFNA`. The text bridge continues its
  admitted output scan after formula errors, so a later typed provider or
  budget failure supersedes the retained formula result.
* **Source fence and publication:** The scalar and value evaluators check the
  source version and cancellation before evaluation, after selected matrix
  positions and formatting work, and immediately before publishing the
  result. Every temporary reservation is released on formula, typed failure,
  cancellation, formatter, and source-fence paths.

## Demand-cache checks

The frozen value classifier does not classify the text family as a sequence
reducer, and `apply_function` has no text demand-cache result branch. Text
results are therefore not reused as scalar cache payloads that could suppress
resolver reads or position-dependent argument evaluation. Matrix text calls
retain runtime shape and reference descriptors and perform ordinary
coordinate selection for each output position.

The generic scalar demand cache explicitly excludes Text payloads. If text
result caching is enabled in a later integration, the local contract requires
the source snapshot, conversion/Unicode/formatter profile, and complete
argument dependency identity; a projected scalar argument or a nested MUNIT
size criterion must remain position-sensitive. The current conservative
classification leaves no such cache path in the frozen implementation.

## Validation evidence

The frozen focused targets cover:

* 13 text semantic evaluation tests;
* 12 resource, cancellation, typed-failure, and shape-limit tests;
* 2 native-profile tests; and
* 26 independent-oracle tests covering all 26 function names.

The resource cases cover known-list zero-read refusal, scalar and matrix
broadcast selection, cumulative reference and text-byte limits, output
reservation failure and refund paths, work and cancellation boundaries,
borrowed resolver Text, formula-error continuation, typed provider failure,
formatter output admission, exact fraction work failure, and source fences.
The independent oracle contains 219 typed observations over all 26 functions;
the retained native evidence contains 91 corroborating observations. The
independent exact-fraction verifier covers 5,474 deterministic cases and is
pinned by `fraction-reproduction.json` to the frozen `fraction.rs` hash.

The isolated gate receipts under [`gates/`](gates/) record all seven required
checks. Each command exited zero, and the retained log hashes are:

| Gate | Command | Log SHA-256 |
| --- | --- | --- |
| `ods-tests` | `cargo test --locked --offline -p litchi-ods` | `0ec03135bafc35ea5ae5e04766da88efeeb0f1bd09364c814531f57340d2e011` |
| `clippy` | `cargo clippy --locked --offline -p litchi-ods --all-targets -- -D warnings` | `c400583cbada29002ada4317352eacb32a0741389802ab62cca40142d3dc1c1c` |
| `rustdoc` | `cargo doc --locked --offline -p litchi-ods --no-deps` | `b125e0f4c8442f57fc4415c91cdd1f03e17d655ec7911b1bac7f37eab80e3925` |
| `format` | `cargo fmt -p litchi-ods -- --check` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `batch-format` | frozen selected-file `rustfmt --check` batch | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `boundaries` | `python3 tools/check_crate_boundaries.py` | `cdea9514ba7c77c1608a6e73b2eeaf093accc70aefd1c3591af13d73e8b871fa` |
| `diff-check` | `git diff --check` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |

The package receipt passes 1,592 tests with zero failures and zero ignored
tests. Its focused summaries are evaluation 13, limits 12, native 2, and
oracle 26. `verification.json` records stable source-before/source-after
custody and all required checks passed. These receipts make no timing,
throughput, RSS, or allocator-performance claim.

The §6.7 byte-position family remains outside this review. Native evidence is
corroborating behavior only; the local contract and independent oracle define
the semantic boundary.
