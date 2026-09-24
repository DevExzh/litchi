# ODS sheet metadata review corrections

Review covered public metadata snapshots, staging, exact patches, owned and
source-backed facade integration, and publication boundaries. The final gate
receipt identifies the tested source; this document records the reasons for
changes rather than a separate approval claim.

## Resource and transaction boundaries

- Name and position selectors originally produced different staged keys for the
  same physical/logical cell. Keys now use resolved worksheet index and logical
  coordinates. Mixed-selector update and revert tests cover this identity.
- Equal budget limits, diagnostic names, and current memory usage cannot prove
  budget ownership. A zero-byte reservation merge now checks actual budget-node
  lineage. Both the supplied and retained source cancellation contexts are checked.
- Selector searches originally performed comparisons without charging work.
  Each candidate comparison now charges eight work units and checks cancellation.
  These are linear searches; charged work does not establish a faster algorithm.
- Exact source XML and even the same physical source-backed package owner do not
  imply equal destination policies. Two root-authored regressions failed before
  the fix: ordinary patch application and source-bound application under a second
  smaller output profile. Both paths now reopen the target under the receiving
  snapshot’s context and limits before acceptance.
- Local ceilings must produce structured resource errors, including the resource,
  observed amount, limit, and scope, rather than an unsupported-feature message.
- Changed signed packages are refused by the ordinary, mutable, and source-backed
  facades. Exact no-ops retain their source bytes.

## XML and native evidence

Review additionally required qualified ODF address endpoints, acceptance of
explicit paired-empty owner elements, namespace bindings at the insertion site,
prelude placement of label ranges when no table exists, and restricted admission
of known document-content root families. Generated known-owner markup now uses
short, locally declared `table` and `xlink` prefixes. It neither searches ancestor
bindings nor expands long inherited aliases. Unchanged XML remains exact; the
canonical admission gate recognizes this generated output for subsequent edits
without treating arbitrary lexical differences as safe to rewrite. The
conformance regressions and final receipts establish the resulting coverage.

Address validation follows the bundled ODF 1.4 RNG for these schema-typed
attributes. It admits quoted sheet names, absolute markers, and row/column range
alternatives; consolidation targets require `cellAddress`. It is lexical:
`.A0` matches the RNG, and no worksheet or dimension resolution is performed.
Part 3 §9.2.1 discusses subtable references, while the RNG for these attributes
permits one sheet separator; this API does not add subtable resolution.

A native program accepting and saving a package does not establish semantic
preservation. Shorthand second address endpoints in an earlier native harness
were invalid ODF and could not support conclusions about consumer loss. Final
native evidence must identify the corrected input, preserve the original API
expectations across saves, and distinguish owner schema checks from whole-package
validation. Any observed loss remains a limitation, even if the native process
exits successfully.

## Remaining boundaries

Owned package replacement/rehydration uses existing package policy, outside the
metadata context’s accounting. Source-backed publication uses the caller’s
publication options. Exact source-bound patches are in-memory lineage artifacts;
they are not a portable patch exchange format. The implementation still indexes
metadata through substantial owned XML structures and scans detective metadata
separately. Performance receipts report those costs and do not establish the
wider GOAL.md speed or memory targets.

## Final bounded disposition

Independent review of the frozen namespace renderer, exact canonical admission,
and eleven XML conformance regressions found no remaining original blockers
within the documented canonical/refusal scope. The focused XML conformance target
passed eleven tests. This disposition does not certify wider ODF coverage,
whole-package schema conformance, native semantic preservation, or the overall
performance program; those remain separately scoped in the receipts.
