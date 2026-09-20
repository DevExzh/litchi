# ODF byte text function resource and cache review

This review covers the seven OpenFormula 1.4 §6.7 functions: `FINDB`,
`LEFTB`, `LENB`, `MIDB`, `REPLACEB`, `RIGHTB`, and `SEARCHB`. It applies
ADR0005's bounded work, storage, source, and cancellation rules and ADR0006's
typed-failure and publication rules to the selected UTF-8-octet profile in
[`contract.md`](contract.md). The review made no production edits.

The disposition is **PASS** for the resource and demand-cache boundaries. The
frozen implementation has no unresolved resource blocker.

## Frozen source identity

The freeze manifest is [`gates/freeze.json`](gates/freeze.json), based on
commit `3844f235bac545ff0ae1580b97612883c1fd9f89`. Recomputing the selected
files after the temporary cache probe was removed matched every selected
source, test, contract, oracle, and native receipt hash. The selected hashes
are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `971f7589f5abd32930e039e7cb07a1b6f20dcab776836ade37327648ebc85847` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `e9e266ad59658d378fb758a436fd7a4a12e66b9924c225e074a03a82c8ceb51f` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `cd882fa914bffa26ede20c51a4483c53fc36f4070e37cfbfef2bfec6e2db69f3` |
| `crates/litchi-ods/src/codec/formula/evaluation/text.rs` | `11d7e945eeb00b733941c34026ae670d997d58a76be16808c2d10b0a9221ac41` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/bytes.rs` | `3a625110a77697bba2e7414f9c898bedf61a70bed8517cc9d67db6036add3d61` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/core.rs` | `17d1a1d1f9bbd4a30eec6a48d1aa1d049e29040b489779c0ca867a6f27f962f0` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/format.rs` | `129b4323c1f577f048e6bc3431653d4c2d979cc04f2fa60305d3bcf84d923c06` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/fraction.rs` | `a2196dd849e2d85faaacd49f876aba8920f9aff00b50e1fad474ae6c278f276e` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/search.rs` | `523b93ce563976bc617c626133e3dd90760d0580b63ef677a7eb73bda0f78d0f` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/unicode.rs` | `336584ece56ee00e9e68a0c2165b5396ac1f007d5ad671f1fe8eba30aac877b6` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/width.rs` | `69a4fae2bb52744e6fa041e3b5de2ffcb9ffa17f6fa7ed2139e433ffe9136b7c` |
| `crates/litchi-ods/tests/ods_formula_byte_text_evaluation.rs` | `57be711b33521d02197552bba33a09e105e4c927e8cc7e1fc5a6451e68ca8007` |
| `crates/litchi-ods/tests/ods_formula_byte_text_limits.rs` | `ad4723fe4e3ee6a31a82e59a63e036fcabbff01a25b52ec06778f09f534981d2` |
| `crates/litchi-ods/tests/ods_formula_byte_text_oracle.rs` | `be7897323c2dbb937a19acc3546be18da2f4739ec843326c74e7325c5261783b` |
| `crates/litchi-ods/docs/FEATURE_MATRIX.md` | `80494d5da9129ac1957fc8206e9b9b780144d1bea95e75d7074941069493d1d8` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/contract.md` | `88dafc6f0d111672e724b4238289afc0a17d879b856059ac3886ca60bbc99db5` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/byte_oracle.py` | `0b677d1e3b67d2597fdfa292d656abd98debf308f3a7be4387a8e37341888859` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/byte-goldens.json` | `9f30f244b3f98d1f6ddbdc481165efbf153c80205ad60a85cab1c7561f47746c` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/native/README.md` | `634d549abc98a9155a157fe30217ec56bff9e6cc8650490b3923064e6fdae24c` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/native/byte-functions-native.fods` | `bfe3de382ea70223509e99b4709b9b25bbb4f7eb73bea77c461f8f07ba6c9011` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/native/native-results.json` | `a9a376ef554ada95ef0d9cb864be4f4d195c46fb3d9a14fd1148e1813e3c8091` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/native/provenance.json` | `ca084d339cbfd49c86cfc75fc015e6eac092bc454c46375fdf340d106f4fb03a` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/native/recalculated.ods` | `25551e366e07853532d4afd958c02645b08d308b0ef5a1cbe4ff61f4d0d0a0b5` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/native/reproduce.py` | `61c025979febe643e10219f876b840dc15d7d4855cd3628c3c00c219cc4d9d7f` |

The isolated gate checkout uses the frozen `gates/Cargo.lock`, SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`. The
ambient root `Cargo.lock` is
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`; this
expected workspace mismatch does not change the frozen gate receipt.

## Resource and ownership checks

* **Shape and reference admission:** The value evaluator routes byte functions
  through the scalar-result shape planner and the dedicated text mapper. A
  mapped reference selects and reads one cell per output position; it does not
  materialize a range. `ReferenceList` and multi-area/3-D descriptors are
  refused before cell reads. The focused shape test verifies zero reads for a
  `ReferenceList`, zero reads for a typed 3-D reference refusal, and the
  documented formula-error distinction for the list case.
* **Bounded matrix state:** Matrix output cells, function argument buffers,
  evaluator frames, and shape/cache scratch vectors use checked dimensions and
  fallible capacity reservations. The seven functions return scalar values per
  position, so only the bounded result array is retained by matrix evaluation.
* **Work, reads, and cancellation:** Each mapped output charges cell work
  before selecting a reference element. `read_reference_cell` checks the
  cumulative read limit and cancellation before the resolver call, increments
  successful reads only after the call, and checks cancellation again before
  retaining the returned cell. Text byte work, boundary decoding, Unicode
  folding, search comparisons, and output bytes are charged; long scans check
  execution at bounded intervals.
* **Streaming and borrowed text:** `LENB`, literal slicing, and search inspect
  borrowed semantic Text. `read_to_element` applies the text limit and creates
  a borrowed `TextValue`; it does not clone provider strings. Prefix/suffix
  boundary mapping needs at most three UTF-8 boundary adjustments. `SEARCHB`
  retains only bounded folded-pattern, failure-table, and ring-origin vectors;
  its matcher scans the borrowed haystack as a stream without materializing the
  whole value. The checked scalar-to-byte and byte-to-scalar mapping helpers may
  make bounded additional scans to normalize the start or report the original
  UTF-8 position.
* **Output reservation and drop order:** `REPLACEB` checks output-length
  arithmetic, charges output bytes, reserves checked storage, checks execution
  between source/replacement copies, and drops source and replacement values
  before publishing the owned result. Slices preserve the source reservation
  when ownership is already available. Array and argument leases are declared
  so buffers are dropped before their reservation tokens on success and typed
  failure paths.
* **Typed-failure precedence and fences:** Generated formula errors remain
  scalar values. Resolver `Unsupported`, resource, cancellation, source, and
  allocation failures return as `EvaluationFailure`; they are not converted to
  formula errors or caught by `IFERROR`. A retained formula error in an earlier
  reference cell does not hide a later typed provider failure. Evaluation checks
  cancellation and the resolver source version before reads, after evaluation,
  and immediately before publication.
* **Shape/type refusal:** Direct scalar byte consumers reject a list before
  reading its cells. Matrix mapping rejects a list or unsupported multi-area
  descriptor before selecting a cell. Computed expressions may have performed
  their own admitted reads before their resulting descriptor is refused, which
  is the contract's per-argument rule rather than a whole-expression zero-read
  promise.

## Demand-cache checks

Byte functions remain position-sensitive scalar consumers and are not admitted
as direct sequence reducers. Their scalar results therefore do not enter the
sequence demand cache merely because they occur in a projected `IF` branch.
The value cache classifier now rejects text-function nodes during a computed
matrix/reducer walk. This prevents a computed `LENB` or `LEN` reference from
being reused at another projected coordinate while retaining cacheability for
fixed, complete matrix arguments.

The permanent projected-reducer check records the resulting tradeoff:

* `AVERAGE(LENB([.A1:.A2]))` returns `[10, 3]` with two reads.
* `AVERAGE(LEN([.A1:.A2]))` returns `[4, 2]` with two reads.
* `SUM(LENB([.A1:.A2]))` runs its computed argument in the existing complete
  matrix aggregate context and returns `[13, 13]` with four reads. The
  repeated scan is conservative and is intentional; no unsound
  cross-coordinate cache reuse is introduced.

Nested conditional-criterion propagation remains conservative and keeps
`MUNIT`'s scalar size argument position-sensitive. A frozen-source probe of a
computed criterion (`COUNTIF` over `AVERAGE(LENB(reference))`) returned
`[1, 0]` with six reads, confirming that criterion projection follows the
selected output coordinate. The temporary probe was removed before the source
hash handoff. Demand-cache entries continue to hold only scalar values or
formula-error payloads; text, references, arrays, and typed failures are not
cached.

## Validation evidence

The focused frozen targets passed:

| Target | Result |
| --- | --- |
| `ods_formula_byte_text_evaluation` | 6/6 |
| `ods_formula_byte_text_limits` | 5/5 |
| `ods_formula_byte_text_oracle` | 2/2 |

The isolated package gate recorded 1,605 passing tests. Package Clippy with
`-D warnings`, rustdoc with warnings denied, package format, and the batch
rustfmt gate all passed. The independent UTF-8 oracle covers 1,376 semantic
observations. The native receipt contains 19 observations: all 7 ASCII rows
match the selected profile, while 10 of 12 non-ASCII rows intentionally expose
LibreOffice's native-width behavior. That is a documented profile divergence,
not a resource failure. This review makes no timing, allocation, RSS, or
throughput claim.
