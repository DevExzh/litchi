# ODS OpenFormula reference grammar

This batch follows `b7a66574a` and addresses the audit's missing inert external
workbook references and incomplete Part 4 section 5.8 reference representation.
The function catalog remains the complete 393-name standard catalog from the
previous batch.

The `codec::formula::reference` reader represents source IRIs, cell/whole-column/whole-row
ranges, explicit and inherited sheet locators, absolute markers, nested table
locations, and reference errors. Source locations remain data: no URI resolution,
filesystem access, networking, dependency evaluation, or workbook refresh occurs.
The wider expression grammar, arrays, host-defined functions, and evaluation
remain distinct work.

## Public API and bounds

`reference::Reference::parse` reads a complete bracketed reference;
`FormulaParser` emits `Token::Reference` for metadata that legacy cell/range
values cannot represent. `extract_references` borrows all reference occurrences,
and `Reference::address()` borrows the address without cloning nested locators.
`SheetSelector::Inherited` retains the omitted endpoint locator without copying
its predecessor. Formula text is retained exactly.

The legacy `extract_cell_refs` query now returns `Cow<CellRef>` elements: existing
legacy values are borrowed, while rich local cells and cell-range starts without
subtables yield owned projections. Callers previously using `.cloned()` should
use `.map(Cow::into_owned)`. `Cell::formula_cell_refs()` still returns
`Vec<CellRef>`. External sources, whole-axis references, subtables, and reference
errors are excluded from that local projection; the complete query retains them.

Default `FormulaLimits` admit 1 MiB and 65,536 tokens. Nested `reference::Limits`
admit a 64 KiB bracket body, 256 components, and 16 KiB per decoded name/source
or lexical coordinate component. All are configurable with finite builders;
formula parsing forwards its nested limits. Limit failures retain the typed
`Error::ResourceLimit` cause. Rows use checked `u32` representation; larger
syntactically numeric rows are refused. No host sheet lookup or axis expansion
occurs. Reference-body whitespace is preserved and checked, while outer parser
whitespace retains the existing compatibility behavior.

This batch covers the reference family within those representation/resource
bounds. It does not close the broader specification audit.

## Architecture and verification contract

- ADRs 0001/0006: reference correctness and exact original formula retention take
  precedence over parser speed. Decoded IRIs and locators remain inert metadata.
- ADRs 0002/0023/0024: spreadsheet grammar stays in `litchi-ods`; the MathML/StarMath
  `litchi-odf-formula` owner and shared package topology are unchanged.
- ADRs 0003/0004: the new owner is a typed reader value, not a mutable package or
  evaluator. Public variants distinguish address families and invalidated
  references; source qualification must never be silently projected as local.
- ADR 0005: admit bounded input/components before growth, avoid recursive locator
  parsing and repeated prefix scans, and measure accepted existing formulas
  separately from additional syntax coverage. Peak-live memory is a maximum,
  not an allocation total divided by operation count.
- ADR 0008: final source must pass the ODS all-target, warning-denied lint/docs,
  doctest, and formatting gates, independent grammar/resource tests, and isolated
  candidate checks. Benchmark manifests and patch replay must match those files.

The [normative review](spec-review.md) distinguishes complete reference-family
representation from evaluation and the remaining expression grammar. The
[performance evidence](performance/report.md) compares the previous committed
parser with this batch and documents material regressions as well as improvements.

## Validation

The final [gate receipt](gates/results.json) records 810 passing tests across
47 targets and successful warning-denied Clippy, rustdoc, doctest, and formatting
checks. The crate currently has no executable doctests. The ten independent
[reference integration tests](../../../../crates/litchi-ods/tests/ods_formula_references.rs)
cover all address alternatives, metadata, malformed input, and exact limits.
The existing 393-function catalog remains unchanged.

The standalone [IRI allocation probe](iri-probe.rs) performed 37,016 observations,
including 64 KiB inputs, with zero allocations in the IRI validator. This does
not claim that formula token construction is allocation-free.

A first gate attempt exposed two Clippy `manual_strip` findings and a quota
failure in the default temporary directory. The prefix checks now use
`strip_prefix`; the final successful run used a dedicated `/var/tmp` directory.
The initial performance capture also exposed extra token-vector reallocations;
the final parser restores a bounded four-slot initial capacity and checks token
count before parsing while reserving only after a token succeeds and the vector
needs to grow. ASCII coordinate
copying uses a fallible bounded copy followed by in-place case conversion. Final receipts
replace the superseded attempts. Owned build targets, probe
binaries, and temporary directories are removed after capture.

The isolated candidate checks cover 47 formula/reference unit tests, 10 new
reference integration tests, and seven catalog integration tests. The retained
patch is independently replayed from `b7a66574a`; all five changed or new source
files must match the gated digests.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-references/verify.py`
from the checkout to verify final source hashes, test receipts, isolated checks,
raw benchmark fields/outcomes, and artifact integrity. This validates retained
evidence; reproducing timings requires the standalone harness described in the
performance report.

The final paired microbenchmark records an 18.2–29.0% latency increase for the
three unbracketed control workloads and 26.6–48.4% for bracketed references.
Common-formula allocation counts and requested bytes are unchanged; three
metadata-rich cases add one allocation, and one range process records +5.1%
maximum RSS. These are material performance costs, retained explicitly under
the ADR correctness and bounded-resource priority. The capture does not isolate
which new checks account for each timing difference, and it supports no general
CRUD speedup claim. Further tokenizer optimization remains performance work.
