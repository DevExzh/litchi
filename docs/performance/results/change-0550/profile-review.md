# 0550 source-backed XLSX commit profile review

This review records the baseline Callgrind attribution for the exact
`litchi_xlsx::cell_values::source::MultiSourceEdit::commit` owner. It is a
mechanism diagnostic for the selected source-backed one-percent edit/save
case. It does not claim a latency improvement, an allocation improvement, or
a removable fraction.

## Custody and replay

The frozen plan is bound by SHA-256
`3eeb52cdad1312b3953006e538490d6523a707e9858354dcb6688b4a615471a7` in the
retained receipts and machine-readable report. The report records the exact approved capture
amendment (`capture-amendment-inputs.json`, SHA-256
`672a57c7c1673666e834709d1f81712e3b44e3ec710ff620c548f8e114fb13f8`) and
its input hashes. The amendment adds `--vgdb=no` and an owned `--vgdb-prefix`
after the first failed profile child; the original failed artifacts remain in
`failed-attempts` and are not used as measured evidence.

The normal binary is bound to SHA-256
`6d2d1ccf3a47e0db1c7188d509385ae75dd80ad81c60a4e13cab93c251b75c5b`, and the
profile source manifest is bound to SHA-256
`fa85f76972a37f52644c60b955c9125ed80e4d78e2434ec54b77c3cb19e64163`.
Profile and native result identities match for every shape and repeat. The
profile analyzer is
[`analyze_profiles.py`](analyze_profiles.py), SHA-256
`ac23bbe997c8753557c141752ea2bb2c3e59d634a549d75b6ecd517aabcd0fbc`; its
retained output is [`profile-analysis.json`](profile-analysis.json), SHA-256
`64ee681a71674696e65f3d36922a03dec950e8af5d2e49ecf00b1c65c5dfd0ae`.

The analyzer imports the immutable 0521 parser and its 0519 dependencies by
path. Their hashes are recorded in `profile-analysis.json`:

| Imported helper | SHA-256 |
| --- | --- |
| `change-0521/analyze_profiles.py` | `5a5ba155c9eb775cd8653f41b10ef9b3e52e8fc871ea3df01166fba2e5c269a7` |
| `change-0519/analyze_profiles.py` | `9587f5a776b55d423b720a98cd228708c79af0956100fd164d81713e0c90488e` |
| `change-0519/compare_profile_lanes.py` | `a817bb19e1f6d1b5989e41a72dbcf3166b9d0178099781e5ed78c9e5f19f3c74` |

Replay with `python3 -B analyze_profiles.py --output profile-analysis.json`
is read-only when the output and annotation sidecars already exist. Output
uses exclusive-create or exact-identical semantics. The deterministic
`callgrind_annotate` environment is `PERL_HASH_SEED=0` and
`PERL_PERTURB_KEYS=0`.

## Scope classification

Each profile command disables collection at process start, toggles only the
exact owner, zeros before that owner, and dumps after that owner. The analyzer
classifies each numbered dump from its positive incoming owner parent or a
bounded positive ancestor. It does not assume that a dump number alone names
the selected operation.

| Shape | Lifecycle owner dumps | Measured dump | Termination dump |
| --- | ---: | ---: | --- |
| `medium` | `.1`–`.3`, lifecycle parent | `.4`, measured runner | `.callgrind`, part 5, `Program termination`, zero Ir |
| `dense-sparse` | `.1`–`.3`, lifecycle parent | `.4`, measured runner | `.callgrind`, part 5, `Program termination`, zero Ir |
| `noncompact` | `.1`–`.3`, lifecycle parent | `.4`, measured runner | `.callgrind`, part 5, `Program termination`, zero Ir |
| `vendor-extension` | `.1`–`.4`, lifecycle parent | `.5`, measured runner | `.callgrind`, part 6, `Program termination`, zero Ir |

All eight jobs have exactly one positive owner edge with `calls=1` in the
measured dump. The 26 lifecycle dumps are retained for custody and excluded
from measured attribution. Every numbered dump has the exact owner trigger,
`events: Ir`, and a raw owner self-plus-direct equation that balances its
summary. The unnumbered process dump is termination evidence only.

## Owner and immediate children

The table shows the measured owner inclusive Ir and owner self Ir for both
same-binary repeats. Immediate child rows in the report are disjoint raw
owner edge partitions keyed by Callgrind function ID; descendants are not
added a second time.

| Shape | Repeat 1 owner Ir | Repeat 2 owner Ir | Owner self Ir | Largest immediate children in repeat 1 |
| --- | ---: | ---: | ---: | --- |
| `medium` | 133,269,922 | 133,280,376 | 5,720 | `rewrite_value_only_with_provenance` 75,175,345; `Snapshot::from_rewritten_value_source` 57,589,036 |
| `dense-sparse` | 258,256,329 | 258,284,803 | 9,800 | `rewrite_value_only_with_provenance` 146,921,217; `Snapshot::from_rewritten_value_source` 110,596,196 |
| `noncompact` | 165,898,276 | 165,902,662 | 5,720 | `rewrite_value_only_with_provenance` 104,247,528; `Snapshot::from_rewritten_value_source` 61,151,127 |
| `vendor-extension` | 133,297,977 | 133,269,518 | 5,720 | `rewrite_value_only_with_provenance` 75,203,831; `Snapshot::from_rewritten_value_source` 57,589,176 |

Those two direct edges carry about 99.6%–99.7% of each owner direct partition.
The remaining direct rows include `append_actions`, effective action counting,
cell-store lookup, invalidation, and temporary destruction. Ir is a
Callgrind instruction-event count and is not a native elapsed measurement.

Repeat drift is small and is retained as a diagnostic rather than averaged
into a claim. Relative repeat-2 changes in owner inclusive Ir are +0.0078%
(`medium`), +0.0110% (`dense-sparse`), +0.0026% (`noncompact`), and −0.0213%
(`vendor-extension`).

## Nested attribution

The analyzer reports exact source-level targets when they have a positive
owner-reachable edge. Values below are repeat-1 inclusive Ir:

| Shape | `validate_xml` | rewrite target | `merge_omitted_cells` | reduced readback/parser |
| --- | ---: | ---: | ---: | --- |
| `medium` | 46,788,369 | 75,175,345 | 7,800,084 | absent from positive owner graph |
| `dense-sparse` | 89,942,800 | 146,921,217 | 15,574,920 | absent from positive owner graph |
| `noncompact` | 50,430,664 | 104,247,528 | 7,698,433 | absent from positive owner graph |
| `vendor-extension` | 46,790,484 | 75,203,831 | 7,800,137 | absent from positive owner graph |

`validate_xml`, rewrite, and merge rows are nested inclusive/self diagnostics;
they overlap their callers and descendants and must not be summed. The
reduced readback and reduced parser source-level names are recorded with an
explicit absence reason. Their absence can reflect inlining or codegen
omission and does not prove that the concepts were not executed.

Within the nested `validate_xml` target, the direct rows retain the visible
name and allocation costs. Across the four shapes, `QName::local_name` is
1,086,143–2,088,575 Ir, `BytesRef::decode` is 1,107,794–2,008,731 Ir,
`__rustc::__rust_alloc` is 673,656–1,296,147 Ir, and
`__rustc::__rust_dealloc` is 653,800–1,258,615 Ir. These are direct child
attribution rows from the measured dump. They do not establish allocation
counts, ownership, or removable work; the separate allocator lane remains the
source for allocation metrics.

The measured owner includes its internal validation, reduced-readback or
fallback decisions, rewrite, merge, invalidation, and temporary destruction.
External source open, selector planning, staged `edit.set` work, publication,
reopen, and later returned-commit destruction remain outside the exact owner
profile as established by the scope review. Any production optimization would
need fresh native, allocator, correctness, refusal, error-order, and
source-lineage evidence while preserving those boundaries.
