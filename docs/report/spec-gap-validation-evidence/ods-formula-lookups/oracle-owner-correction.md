# Lookup oracle correction handoff

This note records the owner-side correction pass after the independent HOLD
review in `oracle-review.md`. It is a handoff for the independent second pass;
it is not a semantic approval.

The retained Python oracle now has 127 contract-bound observations. The prior
descending duplicate model keeps the last qualifying duplicate, exact formula
error cases expect the suffix scan, the LOOKUP short-reference duplicate case
uses an in-range query, and approximate resolver reads use envelopes. INDEX
array literals now use the evaluator grammar (`;` within a row and `|` between
rows), with `{1;2|3;4}` for the 2x2 cases. The full-area, duplicate-record,
and 3-D INDEX descriptor rows remain in the corpus. Reference expectations
carry `owner=direct` or `owner=derived`, and the Rust consumer checks lexical
owner retention through `ReferenceView::reference()`.

The added `match.ascending_mixed_type_midpoint` and
`match.descending_mixed_type_midpoint` rows exercise the Number/Text ordering
barriers in approximate mode. They are model-generated from the pinned
comparison owner and are included so the prior coverage gap is reviewable.

Current independent checks:

```text
python3 lookup_oracle.py --self-check  -> 127 observations
python3 lookup_oracle.py --check       -> verified=true
```

Current identities are:

| Input | SHA-256 |
| --- | --- |
| `lookup_oracle.py` | `a9e33442aec45cce4ee361f2bd54daaeb22d7a0778e39b7669cc3ab725f8054b` |
| `lookup-goldens.json` | `ad2a9bd6c92823b1a40d012d2c64bea66a7b9a0b87c998bb16fc4fcd89887758` |
| `contract.md` | `b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf` |

The native receipt was captured from the same contract hash and independently
links these identities in `native/provenance.json`. It retains 32 formula rows
(30 function observations and two Data-sheet seam rows), with 17 native
parities and 13 explicitly documented LibreOffice divergences. The native
reproduction and root `verify.py --allow-pending` both validate the retained
native fixture; the overall bundle remains pending the unrelated source/gate/
performance receipts.

The sheet-bound `lookup.reference_extension_sheet_bound_is_na` case now uses
the implementation-independent read envelope 2–4: binary search may reject
the out-of-extent result vector after two key probes and perform no result-cell
read.
