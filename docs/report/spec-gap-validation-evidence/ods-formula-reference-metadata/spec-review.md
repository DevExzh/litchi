# Reference and worksheet metadata specification review

Status: independent pre-implementation review for baseline
`049c09cdde3978593149079c4257df047a3fa419`. This review covers the eight
functions in the companion contract: `AREAS`, `COLUMN`, `COLUMNS`, `ISREF`,
`ROW`, `ROWS`, `SHEET`, and `SHEETS`. It is a specification and integration
review; it is not a production-support or passing-test receipt.

## Evidence and decision basis

The normative source is the repository-local OpenDocument 1.4 distribution:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`;
* `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The function clauses are §§6.13.2, 6.13.4, 6.13.5, 6.13.24, and 6.13.29–
6.13.32. Their input and evaluation behavior also depends on §§3.2.3, 3.3,
4.8, 4.9, 4.11.2, 4.11.13, 5.6, 5.8, 5.9, and 6.3. The accepted design
constraints come from ADRs 0001, 0002, 0003, 0004, 0005, 0006, 0008, 0010,
0011, 0023, and 0024.

The repository-local native fixture is useful compatibility evidence. Its
receipt records 32 formula rows and 53 independent-oracle observations over a
four-sheet ordered workbook; hidden sheets are included. LibreOffice reports
`Err:504` for the four multi-reference-list dimension/sheet cases. That
observation does not replace the contract’s formula `#VALUE!` spelling or its
typed-failure rules.

## Normative findings

The function syntax and operation must remain separate from scalar result
publication. `COLUMN` and `ROW` accept one Reference, require
`AREAS(R)=1`, and return all columns or rows when the admitted cuboid spans
more than one axis coordinate. A 3-D reference is one logical reference and
uses one common row/column rectangle; its sheet planes do not duplicate the
axis result. A formula at `C5` with `COLUMN([.B2:.D2])` therefore computes the
complete row array `{2,3,4}` before a scalar-output boundary may select an
element. Input intersection to `C2` would change the specified operation. A
ReferenceList, including one with one retained Reference, is rejected for both
functions before the `AREAS(R)=1` check.

`COLUMNS` and `ROWS` accept a Reference or Array and report complete shape.
They do not dereference a reference or inspect Array elements. A 3-D reference
reports its row/column extent once. Every ReferenceList, including one with a
single retained Reference, is refused by this profile at the Reference
pseudotype boundary. Sections 4.9 and 5.9 make admission conditional on
identical computation for an arbitrary sequence of single references occupying
the same cells; they do not define a one-entry list-to-Reference conversion.
The normative COLUMNS counterexample shows why the condition fails for these
shape functions: a rectangular Reference reports its column extent, while its
single-cell decomposition reports a sequence length. A list therefore remains
a list at each function boundary and is rejected before cell reads.

`AREAS` is the list consumer. It counts retained logical reference records in
operator order, including duplicates, while a 3-D cuboid remains one record.
The count is not the number of expanded physical sheet planes. `ISREF` is the
runtime-kind consumer: it returns `TRUE` for Reference and ReferenceList,
including a source descriptor if that descriptor reaches classification, and
`FALSE` for formula Errors and all ordinary scalar/Array values. Neither
function reads a cell.

`SHEET` has two overloads. A Reference is metadata and is never converted to
the content of its current intersection. A Text argument is looked up by exact
sheet name. Number and Logical arguments use the §6.3.14 Text conversion
before lookup; a Text array is iterated elementwise in matrix mode. A 3-D
reference returns the first sheet in normalized workbook order. `SHEETS()`
reports the complete workbook count, including hidden sheets, and
`SHEETS(R)` reports the inclusive sheet span of one admitted cuboid. A
ReferenceList, including a one-entry list, is refused by the profile.

The empty argument list is distinct from an explicit Empty value or empty
parameter slot. `COLUMN()`, `ROW()`, `SHEET()`, and `SHEETS()` use their
documented current-position/document defaults. Required arguments and
explicit empty slots are validated before metadata execution. Formula Errors
remain values and propagate through ordinary conversion, except that `ISREF`
classifies an Error as `FALSE`.

## Runtime and resolver feasibility

The existing value evaluator exposes enough local metadata for this batch:

* `Context::position` supplies the fixed current sheet, row, and column;
* `Resolver::sheet_extent` supplies checked row/column bounds;
* `Resolver::sheet_index` supplies stable workbook order;
* `Resolver::sheet_name_at` supports ordered name lookup and inspection; and
* `Resolver::sheet_count` supplies the document count.

`RuntimeAreaSet` already distinguishes logical reference records from expanded
physical areas. Direct references use one record; a 3-D cuboid may retain
multiple sheet planes within that record. A concatenation retains one record
per operand and sets `is_list`. Metadata functions should consume those fields
directly. Calling `select_area_element`, `materialize_for_array`, or
`read_reference_cell` would erase the distinction that these functions report
and would violate the zero-cell-read rule.

Logical admission still has a resource cost. The established
`max_reference_cells` check applies to a reference’s admitted geometry even
when no cell is read. `max_reference_areas`, checked arithmetic, cancellation,
and fallible descriptor/output capacity must remain active. Hidden sheets need
no separate visibility field in the current host because the resolver’s
ordered set includes every supplied sheet; a host that filters hidden sheets
would violate `SHEET`/`SHEETS` semantics.

The current resolver has no external-workbook metadata provider, but that does
not erase source kind semantics. The accepted limited profile is explicit:
`ISREF` on a direct source leaf returns `TRUE`; `SHEET` and `SHEETS` on a
direct source leaf return formula `#VALUE!` from their source-location
constraint; and the same results hold when `IF(TRUE(); source; local)` selects
the source leaf. `IF(FALSE(); source; local)` remains lazy and uses the local
descriptor. These paths call no external provider and read no cell.

The same preservation rule applies through `IFERROR` and `IFNA`. A source leaf
is a successful Reference value, so
`ISREF(IFERROR(source;"Main"))` and `ISREF(IFNA(source;"Main"))` return
`TRUE`; `SHEET(IFERROR(source;"Main"))` and `SHEET(IFNA(source;"Main"))`
return formula `#VALUE!` at the final metadata consumer. If the source is used
by arithmetic inside the handler, such as `IFERROR(source+1;"Main")`, the
operator produces typed `Unsupported(Reference)`. The handler does not catch
that typed failure or substitute its fallback text.

For `AREAS`, `COLUMN`, `COLUMNS`, `ROW`, and `ROWS`, a source leaf is a typed
`Unsupported(Reference)` capability refusal in the limited profile because
their current implementation has no external geometry provider. A source
operand entering generic reference arithmetic (`:`, `!`, or `~`) is also typed
`Unsupported(Reference)` before a metadata function can inspect the result.
This separates unsupported source arithmetic from provider-free descriptor
selection. Source IRIs remain inert lexical values and are never fetched.

The specification also makes evaluation outside a table cell an Error for
`SHEET`. The current value API has an explicit worksheet `Position` but no
outside-table bit. It can faithfully implement the ordinary cell-context
profile; adding a host API flag is required before claiming the outside-table
case. The scalar evaluator has neither Position nor Resolver, so its metadata
entry points must return typed `Unsupported(Reference)` for no-argument
current-context operations, workbook lookups, and reference descriptors. It
must not invent process-global workbook state.

## Evaluation scheduling and cache consequences

Metadata functions require complete argument demand. The value scheduler must
use a matrix/complete argument path for Reference, ReferenceList, and Array
operands before scalar publication. A projected `IF` branch must preserve its
complete descriptor through the selected branch. `ISREF` must classify the
runtime kind before scalar coercion; coercing a reference or list to its
intersected cell makes `ISREF` wrong. `COLUMNS` and `ROWS` likewise need full
Array shape before any element selection.

`COLUMN` and `ROW` are the special result-shape cases. Their operation emits a
complete row or column array, then the ordinary outer scalar-demand boundary
may select the profile’s first element for an origin-less scalar publication.
The function must not intersect its Reference input first. No-argument
`COLUMN()` and `ROW()` are position-sensitive; `SHEET()` without an argument
is position/workbook-sensitive. Explicit invariant descriptors can be cached
only after complete argument propagation proves that the descriptor and result
do not depend on the projected output coordinate. `SHEET(Text)` and
`SHEETS()` have the same complete-argument requirement; Text-array results are
coordinate-sensitive unless the complete array is retained. Keep existing
MUNIT scalar descendants position-sensitive and out of complete-reference
propagation.

Known list, pseudotype, source-constraint, and shape refusals must occur before
cell reads. If a computed child must be evaluated to discover its runtime
descriptor, that child may perform its own work. Once a descriptor is known to
be inadmissible, the metadata reducer contributes no cell reads and cannot
catch typed provider, cancellation, resource, or source-version failures.

## Gap and implementation matrix

| Area | Existing capability | Required implementation decision | Gap status |
| --- | --- | --- | --- |
| Catalog and arity | Shared scalar catalog pattern exists | Add exact arity/default classification for eight names | Needed |
| Local geometry | `Area`, `RuntimeAreaSet`, and resolver extents exist | Add metadata reducer over complete descriptors | Needed |
| Reference lists | `is_list` and retained records exist | AREAS/ISREF admit all lists; COLUMN/COLUMNS/ROW/ROWS/SHEET/SHEETS refuse every list, including one-entry lists, under the arbitrary-decomposition rule in §§4.9/5.9 | Needed |
| 3-D references | Physical planes and record boundaries exist | Count records for AREAS; use common 2-D bounds; normalize sheet span | Needed |
| Arrays | Runtime arrays retain shape | COLUMNS/ROWS consume shape; SHEET Text iterates in matrix mode | Needed |
| Current position | `Context::position` exists | Implement no-argument COLUMN/ROW/SHEET profiles | Needed |
| Workbook metadata | `sheet_index`, `sheet_name_at`, `sheet_count` exist | Use stable ordered resolver set, including hidden sheets | Needed |
| External Source | Generic resolver boundary is typed Unsupported | Direct/selected source leaves: ISREF=`TRUE`, SHEET/SHEETS=`#VALUE!`; source arithmetic and other geometry functions remain typed Unsupported | Profile split |
| Outside-table SHEET | No context bit | Add host metadata before claiming that normative branch | Host API gap |
| Scalar evaluator | No resolver/position | Typed Unsupported for contextual metadata | Intentional limit |
| Resource/cancel fences | Existing evaluator budget and source fences | Charge metadata/output work and preserve final fences | Needed |
| Demand cache | Existing projected-branch machinery | Add complete-argument and position-sensitive classifiers | Needed |

## Validation required before handoff

The focused batch should cover the following observable contracts:

1. exact arity, omitted defaults, explicit Empty slots, and wrong pseudotypes;
2. single-cell, rectangle, whole-row, whole-column, reversed endpoint, and
   3-D references;
3. direct references, one-entry and multi-entry list refusals for the six
   Reference functions, plus AREAS/ISREF list admission, duplicate records,
   and nested `~` order;
4. `AREAS` record counts versus physical sheet-plane counts;
5. complete `COLUMN`/`ROW` axis arrays before scalar publication;
6. `COLUMNS`/`ROWS` over complete Arrays containing formula Errors;
7. `ISREF` over every scalar kind, formula Error, Reference, ReferenceList,
   source descriptor, Array, and deferred scalar-cell token;
8. `SHEET()` current position, exact Text, Number/Logical conversion, unknown
   names, local and 3-D references, matrix Text arrays, direct source refusal,
   `IF(TRUE(); source; local)` selected-source refusal, and
   `IFERROR`/`IFNA` source-preserving rejection;
9. `SHEETS()` document count, hidden sheets, local and 3-D references, list
   refusal, direct/selected source refusal, IFERROR/IFNA source-preserving
   refusal, and source-arithmetic typed boundary;
10. zero resolver cell reads for every metadata success and known refusal,
    including lists, source constraints, shape failures, and geometry-limit
    failures;
11. metadata work, reference/array capacity, cancellation, allocation drop
    order, provider failure, source-version, and final-cancellation behavior;
12. lazy `IF`, projected matrix branches, direct and selected source leaves,
    IFERROR/IFNA source preservation, source arithmetic refusal, nested
    metadata calls, and demand cache hits/misses across different current
    positions.

The independent oracle should compare scalar values, array shape and order,
ReferenceList classification, sheet order, cell-read count, and typed failure
kind. Native spreadsheet output may be retained as a compatibility appendix;
the local normative contract and resolver profile decide acceptance.
