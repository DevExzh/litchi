# 0549 next OOXML target: rewritten worksheet readback

This is a bounded source and evidence review. It selects one OOXML/XLSX
attribution target after the two rejected CFB per-sector rewrites. It does not
claim a measured hotspot or a speedup, and it starts no new capture campaign
in this turn.

OLE2 remains at its sealed baseline while this OOXML target is qualified. ODF
remains deferred until the OLE2/OOXML optimization goal is complete.

## Selected target

The next target is the changed worksheet commit readback boundary:

```text
SourceEdit::commit
  -> rewrite_value_only_with_provenance
  -> Snapshot::from_rewritten_value_source
  -> try_rewritten_value_cells
```

The single-sheet caller is
`crates/litchi-xlsx/src/cell_values/source.rs:817-872`. For a non-empty
single-sheet edit it rewrites the worksheet, constructs a snapshot through
`Snapshot::from_rewritten_value_source` at lines 836-842, invalidates the
calculation part, and checks every staged value again at lines 843-867. The
multi-sheet caller is the same file at lines 1125-1230. It correctly skips
untouched sheets and empty action sets at lines 1153-1173, then repeats the
rewrite/readback path for each changed sheet at lines 1177-1214.

The concrete duplicate-work boundary is
`crates/litchi-xlsx/src/cell_values/snapshot.rs:715-759`:

1. `from_rewritten_value_source` validates the complete rewritten byte stream
   with `validation::worksheet_xml` at line 723.
2. `try_rewritten_value_cells` constructs another byte buffer with
   `raw::worksheet::edit::reduced_readback` at lines 740-741, then parses that
   reduced stream with `raw::worksheet::parse` at line 742.
3. Provenance and supported-entry checks then merge the parsed changed records
   with source records through `Store::merge_omitted_cells` at lines 751-758.
4. Any reduced-readback refusal, parse failure, unsupported entry, invalid
   rectangle, or merge collision returns `None`; the caller performs a full
   `raw::worksheet::parse` of the complete output at line 727.

This is a candidate for a private proof-bearing handoff between rewrite,
validation, and semantic readback. The source shows that the full output
validation and reduced parse both traverse XML-related state, while the
rewrite already records omitted spans. It does not show that either traversal
is removable: the complete output remains the validation authority, and the
reduced route is what verifies changed records before source records are
merged. A source-level fusion or an assumption that unchanged bytes are safe
to skip would therefore be premature.

The preceding accepted 0546 change optimized the source-backed worksheet load
path and retained its shared parser/validator traversal. Its profiled owner is
`SourceBackedEditor::edit_sheets`, not either commit method. The next target
must therefore be attributed in an isolated commit path; the retained planning
profile cannot be used as commit attribution.

## Current source and evidence custody

The current checkout has no changes under `crates/litchi-xlsx`. The latest
XLSX source commit is `2118c6fb1d023005e224aaa938e84d6ed0588b70`
(`perf(xlsx): retain shared worksheet traversal with exact event admission`);
the workspace `HEAD` is
`6d9fbb7401729aacf2afc5d6b6c681a9e7384056`. The provenance readback region
was introduced by `67028ab6037ae6eef15af92a0d540285c3c5362c`
(`perf(xlsx): reuse unchanged cells during rewrite readback`). Current source
SHA-256 values for the relevant files are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `b41af5d8c91a5c1e82f9030b8798a354ca4943072373846b8874c73996c03035` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `6db3c616e0e639b578b9aea2493cf717083f4d06f072a201a787002a9daa2178` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` |

The authoritative retained 0546 records are bound as follows:

| Record | SHA-256 |
| --- | --- |
| `change-0546/integration/plan.json` | `6835a4122a16bf6ec8a5f3513e27be50749d3a8b3a471261976de6f3c4302574` |
| `change-0546/integration/decision.json` | `6ca7e5042cb6305b0ed1f83ca32899efe36983b37520da76e8889e4e1d55623e` |
| `change-0546/integration/planning-profile-analysis.json` | `929f920bcc12f6de7594a8909892a96a8ef1bbe05ffe9a7a0354a6f443750f08` |
| `change-0546/integration/results-review.md` | `5ae5c0c541ff0f2279bef32017185c1876f4dd7ac9a09e947cec68eade37a276` |
| `change-0546/integration/source-review.md` | `a50d44294bc9cb482c01c376024dd835c8b113d2d4e51db753d81b520eee8701` |
| `change-0546/integration/candidate-source-hashes.json` | `889cbc45e5a74767b029bf70d5577f8deff8678073bb5ce8273b305238b4bba0` |
| `change-0546/integration/SHA256SUMS` | `4e2be2147eb83b77a01d7d4eb4a5b587e9e5ac125705b7389e2c9f9d579341eb` |

The retained profile is a planning diagnostic with
`performance_claim: none`, exact owner
`litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets`, and
explicit limitations against converting Ir to latency. The selected final
owner dumps and their exact receipts are retained here for custody; they are
listed to delimit the evidence, not to claim commit cost:

| Shape/repeat | Stage | Selected owner dump SHA-256 | Receipt SHA-256 |
| --- | --- | --- | --- |
| medium / 1 | baseline | `465bb7b342496a0b0809e7ea6f7cbc8de6b9469c2928c80711b03fc5cb606176` | `67806ed08f104c529381e382c439ba626142f2ceb925de3f14ca46c7badb10b2` |
| medium / 1 | candidate | `b46f7a2f4d8d1aff2c05a4a583712496fc20c93b97d1319435cea5946bc68c4b` | `0d1e81a449e6554f373a83abe185ecd78b06be3638523a17e32842d78b9191e6` |
| dense-sparse / 1 | baseline | `db3b5661fa467e0d1b4905e88e68bff8b0f1f83a23bb5a509cac42001c7a8901` | `d79745db1fa215197902927972c70917d51df4fd92a79f6090a0c1ea22f8ae12` |
| dense-sparse / 1 | candidate | `929bf6b27a932212b9ea010c80c6b6073e40e814095b6cb8e6a7df6592c33e14` | `3c10ab95d3d38cd848239ad47b43bd7036f07816e2da5dd78cbf395eb9f7c3fc` |

Those receipt commands toggle only the exact `edit_sheets` owner for one
planning operation. They do not isolate `SourceEdit::commit`,
`MultiSourceEdit::commit`, `from_rewritten_value_source`, or the reduced
readback route. The retained native comparison includes a commit phase inside
the whole workflow, but it does not supply a function-level commit
attribution. No prior numeric hotspot or speedup figure is used here.

## Measure before coding

The next capture, owned by the root agent, should bind the current source
checkout, exact binary hashes, source manifests, input identities, CPU and
repeat order before collecting any result. It should separate setup/open,
planning, commit, publication, and reopen/verification. The commit lane must
attribute both callers independently:

```text
SourceEdit::commit
MultiSourceEdit::commit
```

The exact symbol/caller edge must be recorded even if optimization inlines the
private helpers. Fresh line/callgraph and, where needed, disassembly evidence
should identify actual work in these regions before a candidate is designed:

* `rewrite_value_only_with_provenance`, including `scan` and
  `write_sheet_data_with_provenance`;
* `Snapshot::from_rewritten_value_source` and
  `try_rewritten_value_cells`;
* `validation::worksheet_xml`/`validate_xml`;
* `reduced_readback` and its allocation/copy bytes;
* complete and reduced `raw::worksheet::parse`, including
  `Parser::parse`, `Parser::finish_parse`, semantic materialization, and
  `Store::merge_omitted_cells`;
* the complete-parse fallback and workbook calculation invalidation; and
* staged readback comparisons and operation-local allocations.

For every route, the measure should record output bytes, omitted-span bytes
and count, reduced bytes, whether the reduced route succeeds or falls back,
parsed-entry counts, and allocation calls/bytes/peak. These counters explain
what work occurred; they must not be converted into a latency or hotspot
percentage claim. Native commit p50/p95/p99 and mean should be retained beside
the whole workflow, with the same ABBA and source/output/cache identities as
the primary verifier.

The input matrix must include at least:

* exact no-op and empty staged actions;
* one changed eligible worksheet and multi-sheet edits with untouched sheets;
* inserted, cleared, removed, and shared-formula actions;
* shared-formula source rows, noncompact layout, vendor/MCE and x14ac content,
  and any route that makes provenance ineligible;
* malformed or error-producing output, invalid omission spans, unsupported
  stored entries, and reduced-parse/merge refusal; and
* ordinary changed output whose bytes, parsed cells, source lineage, and
  publication readback must remain exact.

The fresh assembly review must answer whether the complete validator, reduced
parser, and fallback parser actually repeat XML reader/namespace/attribute
loads after optimization. Source-level repetition alone is insufficient,
just as in the preceding CFB assembly review.

## Constraints for any future candidate

Any candidate must preserve the current contracts and error order:

* `plan.is_empty()` and empty action sets remain exact no-op commits with no
  rewrite, publication, calculation invalidation, or changed count;
* full rewritten output validation remains authoritative before publication;
  validator-first and fallback error precedence remains unchanged;
* omission spans remain sorted, disjoint, in-bounds, checked for arithmetic
  overflow, and within existing worksheet, multi-worksheet, and allocation
  limits;
* unsupported provenance, structural collisions, reduced parser refusal, and
  merge failure continue to reach the complete parser fallback;
* UTF-8, XML depth/event, namespace/grammar, MCE/x14ac, scalar/formula/style,
  source identity, execution/version, source-lineage, and resource checks
  remain in force;
* staged value readback, exact output/source retention, workbook calculation
  invalidation, atomic multi-sheet publication, and untouched-sheet sharing
  remain unchanged; and
* no public API, dependency, or unsafe-code change is introduced.

If fresh attribution shows that layout rewriting dominates, that reduced
readback is small or required for correctness, or that no private handoff can
preserve the validation and readback contracts, close this target without a
candidate and return to the measured OOXML/OLE2 backlog. Do not infer a
commit optimization from the retained 0546 planning profile.
