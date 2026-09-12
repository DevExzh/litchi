# Next OLE2/OOXML priority after 0522

`scope: read-only queue review after the terminal 0522 comparison`

`performance_claim: none`

Keep OLE2 and OOXML ahead of ODF until the OLE2/OOXML optimization goal is
complete. The 0522 scanner candidate was rejected and reverted. Its
codec tests, explicit `noncompact` guard, source-bound records, and adverse
flags remain useful evidence. This review selects the next measured work; it
does not authorize a production rewrite.

## Ranking

1. **CFB/OLE2 residual operation attribution, followed by one bounded
   work-elimination gate.** This is the next family-level investigation after
   0522. The accepted 0511 FAT batching change must not be repeated. Its
   retained source-open profile leaves the larger mandatory CFB walks as the
   strongest unresolved OLE2 owners: `SectorChainScratch::collect_exact`
   37.03% exclusive Ir, `OleFile::claim_sector` 16.46%,
   `validate_physical_sector_layout` 13.17%, and
   `validate_stream_allocations` 13.08%; `load_fat` is 9.30%. These are
   diagnostic instruction shares, not proof that any validation is redundant.

2. **OOXML semantic cell-reference ownership, only if an allocator/profile
   lane confirms value.** The current semantic parser stores every cell
   reference as `Option<String>` at
   [`codec.rs#L176`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L176)
   and forces `decoded_and_normalized_value(...).into_owned()` through
   [`codec.rs#L215`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L215).
   A narrow `Cow` reference candidate could borrow the usual unchanged XML
   lexical value until `parse_a1` and error formatting finish, while leaving
   style, metadata, type, formula, and value ownership unchanged. The current
   0522 profile shows `scan_cell_attributes` at 19,992,989 dense-sparse Ir and
   10,258,905 medium Ir in its baseline annotations, with
   `decode_cell_attribute` at 7,561,889 and 3,857,241 Ir respectively. Those
   rows motivate measurement; raw annotation call metadata does not establish
   allocation counts. This candidate is lower risk than parser fusion but
   likely smaller than the CFB and full reconstruction opportunities, so it
   needs an operation-local allocator result before implementation.

3. **OOXML reconstruction/layout owner attribution.** The 0522 baseline
   profiles put `Snapshot::from_rewritten_source` at 702,367,967 aggregate Ir
   (61.147% of selected commit Ir) and worksheet `rewrite` at 443,626,538
   (38.621%). The source owner at
   [`snapshot.rs#L686`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L686)
   validates, reparses, clones, stores the rewritten bytes, and rechecks
   execution. The semantic parser at
   [`codec.rs#L736`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L736)
   and finalization at [`codec.rs#L989`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L989)
   retain required cell checks. The 0522 scanner reduction lowered rewrite Ir
   11.603% but raised reconstruction Ir 0.640% and still failed native total
   admission. Therefore this is a profile-and-proof target, not permission to
   pass events between rewrite and semantic parse.

If the first CFB attribution cannot separate required validation, setup,
source service, and verification, close it without a rewrite and move to the
OOXML reference-ownership probe. If the reference probe is too small or its
borrowed lifetime changes error/resource behavior, close it and profile the
full reconstruction/layout boundary. A later DOCX provider/publication
measurement remains part of OOXML, but it does not outrank these current CFB
and XLSX owner leads.

## First investigation: CFB/OLE2

Use the current production CFB owner and a fresh plan. The open sequence at
[`file.rs#L539`](../../../../crates/litchi-cfb/src/file.rs#L539) installs the
physical-sector roles, then calls `load_fat`, `load_directory`, optional
`load_minifat`, `validate_stream_allocations`, and
`validate_physical_sector_layout` at
[`file.rs#L699`](../../../../crates/litchi-cfb/src/file.rs#L699). Keep these
phases separately attributable. Directory validation and public parsing meet
at [`file.rs#L934`](../../../../crates/litchi-cfb/src/file.rs#L934),
stream-chain validation uses reusable scratch at
[`file.rs#L1058`](../../../../crates/litchi-cfb/src/file.rs#L1058), and the
physical reconciliation loop is at
[`file.rs#L1157`](../../../../crates/litchi-cfb/src/file.rs#L1157).

The smallest defensible CFB experiment is therefore an evidence-only
operation-local profile and allocator bracket over `OleFile::open_with_limits`
on the existing FAT-heavy/opaque XLS corpus, with tiny MiniFAT and few-large
regular-FAT guards. Record source calls and bytes, operation allocations and
incremental peak, and valid hardware counters when available. Only after a
single repeated operation is named should a candidate remove or reuse work.
Retain fallible reservations, sector ownership and physical-layout checks,
cycle and marker errors, source-version and cancellation fences, FILEPASS
refusal, limits, and all existing error precedence. A generic CFB cache,
ReadAt API change, broad concurrent chain walk, or a repeat of 0511 FAT
batching lacks a current measured justification.

## OOXML follow-up contract

The source-backed XLSX path remains a valid second investigation. Use a new
before/after plan on the restored production scanner with medium and
dense-sparse primary shapes, retaining the `noncompact` prefixed and
multi-attribute guard. Bracket the candidate parser or reconstruction owner
with operation-local allocations and profile it separately from publication,
reopen, and post-timing oracles. Require useful repeatable **total** improvement
and unchanged semantic/output/resource identities before retaining any code.

For the narrow reference-ownership probe, preserve the complete checked
attribute walk, duplicate and malformed-attribute errors, source-order decode
precedence, `parse_a1` row and column checks, cell limits, namespace/MCE
handling, final semantic readback, original-source preservation, and bounded
temporary ownership. Keep an authoritative owned fallback for escaped or
normalized values. Do not pass emitted events into the parser or combine
validation and rewrite state: 0514 and 0516 fusion remain rejected.

For the larger reconstruction/layout owner, first prove which representation
or pass is actually removable. `Snapshot::from_rewritten_source` must continue
to validate and fully reparse rewritten XML before installing the new source
state, and all source/resource checks before and after it must remain. A
profile share alone cannot justify skipping that work.

No ODF optimization is scheduled in this queue. OLE2/OOXML remains active,
and the broader provider, native-producer, physical-I/O, cold-cache, scaling,
and coverage requirements remain open until separately measured.
