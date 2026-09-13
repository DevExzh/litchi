# Current ODS audit disposition

This companion checks the audit’s older ODS inventory against `5fae21a34`
and the implementation evidence below. Complete expression grammar remains
unfinished.

- **Full OpenFormula semantics** (read 🟡/write ❌, Part 4 ch. 5–8): the ODS
  `codec::formula` tokenizer recognizes all 393 normative chapter-6 function names
  (`b7a66574a`) and bounded §5.8 references (`b88341bf4`), including inert source
  IRIs, whole-axis ranges, nested/inherited sheet locators, and reference errors.
  Function-name recognition does not validate arity or evaluate volatile/external
  functions. Complete expression grammar, array expressions (§5.13), named and
  host-defined expressions, type/evaluation semantics, and recalculation remain
  incomplete. Formula-string transactions preserve or replace inert text without
  requiring full tokenization. [Catalog evidence](../ods-formula-functions/README.md)
  and [reference evidence](../ods-formula-references/README.md)
  distinguish lexical coverage from evaluation.

- **Sheet metadata**: bounded public `sheet_metadata` transactions cover consolidation, ordered
  label ranges, and cell detective metadata, including source-backed publication,
  exact no-ops/inverses, and checked repeated-cell selectors (`d5b5afdb4`,
  `5347801fc`). Consolidation calculation, dependency evaluation, and arrow
  rendering remain incomplete. [Lifecycle evidence](../ods-sheet-metadata/README.md).
- **DDE** (§9.8): bounded `dde::{Snapshot, Edit, Commit, Patch}` now supports
  worksheet-source and formula-link CRUD, scalar cache authoring, source-only
  replacement with exact cached-XML preservation, and ordinary/source-backed
  publication. Unknown cache content remains opaque or refuses destructive
  replacement. DDE sessions, source access, refresh, and evaluation remain inert.
  [Transaction evidence](../ods-dde-transactions/README.md).

A separate memory follow-up remains in `FormulaParser::parse_string`: each
string currently reserves the entire remaining input before decoding. Multiple
short literals can therefore retain substantially more capacity than their
combined decoded length. This scan batch does not change that allocation policy;
it needs a dedicated allocation baseline and bounded decoding change.
