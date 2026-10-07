# 0946 XML publication source audit

Reviewed 2026-10-07 from checkout HEAD `deccd87f3efdfd88ea5474550406ed65e7e5670c`
(`perf(xlsx): share owned replacement payloads during publication`). This is a
read-only source audit. No source, build, or benchmark change was made. The
user excludes iWork.

Applicable accepted ADRs read before this review were 0003, 0005, 0006, 0011,
0030, 0031, and 0032, together with `docs/adr/README.md`. The relevant
requirements are: immutable source/version freshness, full original-first XML
validation, exact typed error/refusal ordering, finite aggregate limits,
namespace/MCE preservation, and no unbounded parsed-XML memo or cross-source
cache.

## Current source map

`crates/xml-minifier/src/audit.rs` (SHA-256
`eddc437156c976c55ce36b0bf8f0ef47cc4943d296735a1ac8d6ef120b8754e7`):

* `verify_source_replacement` at lines 1046-1076 compares the common prefix and
  suffix, always audits the original through `verify_observed`, then either
  proves one `Window` or completely audits the replacement. The proof variants
  are `Identical`, one `Window`, or `Complete` at lines 938-962.
* `verify_observed` and `scan` at lines 1763-1928 perform byte/UTF-8 and
  character checks, quick-XML event parsing, namespace/name/attribute checks,
  depth and aggregate counters, and source-policy error ordering.
* `WindowSearch` and `Window::prove` at lines 2188-2391 capture one balanced
  original element, inherited namespace state, removed counters, and the
  replacement window. `Window::prove` re-totals aggregate limits at lines
  2321-2349; a doubt falls back to the complete replacement audit.
* `common_prefix_len` and `common_suffix_len` at lines 2425-2464 can compare
  most of both payloads before the parser work. `inspect_attributes` and
  `check_start` at lines 2563-2605 and 2734-2845 retain lexical,
  namespace, duplicate-attribute, `xml:space`, and limit contracts. The
  AttributeSummary change is the already-adopted 0905 optimization.

`crates/litchi-opc/src/source_backed.rs` (SHA-256
`5e0cf3451e70784ff76896b0b2dc93b4409b15db9ef7c88535cc6a20885f8c43`):

* Ordinary topology replacements are audited in the changed loop at lines
  9611-9645; exact byte no-ops continue before XML auditing, and source-proof
  payloads remain a separate path.
* The single-part, relationship, and multi-overlay routes call the same helper
  at lines 10201-10218, 10318-10343, 10544-10566, and 10797-10815.
* `validate_source_part_replacement_xml` at lines 12952-12966 calls the pair
  audit with default limits and maps either side to the existing typed
  `OpcError::XmlPublication`.

The XLSX value-only writer already exposes the useful provenance boundary but
not enough for publication proof:

* `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` (SHA-256
  `c6d5c29fecac0b67415014cad92e7c4fa08bfe88eeee8145d9ab59777543a202`) defines
  `OmittedCells` and `ValueOnlyRewrite` at lines 21-55. The metadata records
  output spans copied verbatim from unchanged cells and a pointer to the exact
  source slice; it does not record corresponding source spans for changed
  cells.
* `codec/snapshot/write/sheet_data.rs` (SHA-256
  `b462dc62df4cd24f6964464a4166eab3c1a78715e9a2d0004ef21b182a807da4`) records
  unchanged runs at lines 84-187 and 189-273 while writing changed cell
  records. Generated cell lengths can shift output offsets, so omitted output
  spans cannot be treated as source/replacement range pairs.
* `crates/litchi-xlsx/src/cell_values/snapshot.rs` (SHA-256
  `fff0da0fd3fe86824ba6a84531ad81d8499ebff0e73c824e4c16bf53d77129a5`) uses
  the provenance only for reduced semantic readback at lines 784-845. The
  topology plan at lines 1338-1358 and worksheet replacement at lines
  1651-1675 pass the finished payload into ordinary OPC replacement, so the
  audit currently sees only bytes and loses writer span knowledge.

## Deferred hypothesis, not an implementation claim

The 0944 profile reports nested XML audit instructions of 49,735,942 (53.48%)
for medium and 95,298,865 (56.88%) for dense-sparse publication. Five calls
cover four changed worksheets and the workbook. These aggregate columns do not
identify original versus replacement cost or prove that any particular helper
is removable.

The bounded hypothesis is the multi-window extension explicitly deferred in
0747: the value-only writer could provide exact `(original range,
replacement range)` pairs for changed, balanced cell elements, while the
existing full original audit remains mandatory. The audit could then reuse the
existing inherited namespace/counter state for each replacement window and
skip only unchanged replacement bytes. It must sum all removed/added counters
against the whole replacement, require disjoint balanced ranges, and fall back
to the current complete replacement audit for any unsupported namespace/MCE
shape, mapping ambiguity, allocation refusal, or failed window check. A
replacement-window failure must use the complete audit to retain canonical
replacement error and offset behavior.

This does not require the larger 0747 prerequisite of fusing the worksheet
planner's traversal with the XML audit to eliminate the original scan. That
would require exposing the audit state machine and an unforgeable allocation /
policy / limits token; 0747 bounded its one-edit prize to tokenizer work and
deferred it. The multi-window experiment keeps the original scan and needs only
private source/replacement span metadata, if the per-call census shows enough
replacement bytes can be skipped. It must not become a public trust token or a
parsed-XML cache under ADR 0032.

## Required discriminator before code

Use a diagnostic-only per-call census on the exact 0944 medium and
dense-sparse routes. For each XML Part record URI, original/replacement lengths,
common prefix/suffix, changed-span count and byte coverage, `Identical` /
`Window` / `Complete` outcome, original-audit work, replacement-window or full
replacement work, aggregate event/attribute/text counters, and namespace/MCE
cases. On the writer side record source/output span pairs, omitted-run count,
changed-cell coverage, and whether the metadata reservation succeeded. The
useful comparison is estimated replacement events/bytes that a valid
multi-window proof would omit, not the 53–57% aggregate XML share.

The current raw profile's quick-XML cost is not itself a safe shortcut. In
quick-xml 0.41, `reader/state.rs:221-240` compares end names and
`:343-348` stores opened names even when `check_end_names` is false. The audit
also performs its own end/depth/namespace checks. Disabling end-name checks,
turning off the observer after it stops, or trusting the XLSX planning parser
would either repeat rejected 0902 work or weaken the source-policy contract.

Historical boundaries remain active: 0902's stopped-observer branch failed its
native gates and RSS check; 0905's AttributeSummary is already adopted; 0899's
duplicate-check/QName shortcuts remain deferred; and 0747's one-window proof
preserves complete original-first validation. No aggregate raw graph currently
provides dynamic original/replacement attribution, so no production change is
justified until the per-call proof/coverage census exists.
