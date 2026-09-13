# 0555 OOXML next opportunity: a fallible sorted provenance merge

`status: read-only opportunity audit`

`performance_claim: none`

`scope: source-backed XLSX commit; OLE2/OOXML first, ODF deferred, iWork excluded`

The next independent OOXML measurement should target
`litchi_xlsx::cell::Store::merge_omitted_cells`. The current implementation
moves the already sorted cells from the reduced parse, clones the source cells
selected by the omission rectangles, and then calls `Store::from_unsorted`.
That constructor sorts and duplicate-checks the combined cells and sorts the
parsed rows again before rebuilding indexes and extents. A private, checked
linear merge of the two address-ordered sequences is therefore a concrete
work-elimination candidate after the reduced readback has already established
the changed-output authority.

This is a measurement target only. It does not establish a removable fraction,
latency improvement, allocation saving, or adoption decision. It leaves the
source layout scan, the complete output validator, the reduced parser, and the
omission producer unchanged. It consequently does not revive the rejected
source-layout or compact-collector designs from 0552 or 0553.

## Why this is the next independent boundary

The 0550 owner is the exact source-backed multi-sheet transaction:
`litchi_xlsx::cell_values::source::MultiSourceEdit::commit`, called directly by
`litchi_perf_baseline::run_xlsx_cell_values_edit_save`. Its nested attribution
ranked the rewrite/layout scan at 53.24–54.94% of owner inclusive Ir, complete
XML validation at 30.40–35.11%, and provenance merge at 4.64–6.03%. The larger
scan and validator are protected by distinct lossless and validation contracts;
the layout handoff and parser/layout fusion directions have already been
rejected or left conditional. The merge is smaller, but it has a directly
visible sorted-input algorithm boundary that can be changed without moving
work into planning or retaining new snapshot state.

The current path is [`Store::merge_omitted_cells`](../../../../crates/litchi-xlsx/src/cell.rs#L805),
called from [`try_rewritten_value_cells`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L736):

```text
complete rewritten-output validation
  -> reduced_readback
  -> raw worksheet parse of the reduced output
  -> source/parsed Stored eligibility checks
  -> Store::merge_omitted_cells
       -> validate ordered omission rectangles
       -> count source entries and reserve the combined vector
       -> move parsed cells and clone omitted source entries
       -> copy merge ranges
       -> Store::from_unsorted
            -> sort and duplicate-check cells
            -> sort and duplicate-check rows
            -> rebuild cell-row index and stored/content/styled bounds
            -> build merge index
```

Both `Store` values expose address-ordered cell slices: `Store::from_unsorted`
establishes that invariant for every parsed store, and `omitted_entries` walks
the source slice in that order while consuming ordered one-row rectangles.
The proposed helper can merge those sequences by address, detect an equal
address as a refusal, and retain the existing complete-parser fallback. The
rows, columns, defaults, declared extent, merge ranges, and every `Stored`
field must remain identical. The first candidate should keep merge-index
construction unchanged; moving or reusing that index is a separate proof
question.

## Retained owner attribution

The exact 0550 profile report is
[`profile-analysis.json`](../change-0550/profile-analysis.json), SHA-256
`64ee681a71674696e65f3d36922a03dec950e8af5d2e49ecf00b1c65c5dfd0ae`. The
following values are copied from its eight measured dumps. Each cell is
`inclusive Ir / self Ir`; nested inclusive values are not disjoint with their
callers and must not be summed.

| Shape | Owner R1 | Owner R2 | Merge R1 | Merge R2 | Merge share of owner R1 / R2 |
| --- | ---: | ---: | ---: | ---: | ---: |
| medium | 133,269,922 / 5,720 | 133,280,376 / 5,720 | 7,800,084 / 995,191 | 7,799,631 / 995,191 | 5.8528% / 5.8520% |
| dense-sparse | 258,256,329 / 9,800 | 258,284,803 / 9,800 | 15,574,920 / 1,917,350 | 15,572,293 / 1,917,350 | 6.0308% / 6.0291% |
| noncompact | 165,898,276 / 5,720 | 165,902,662 / 5,720 | 7,698,433 / 995,191 | 7,698,755 / 995,191 | 4.6405% / 4.6405% |
| vendor-extension | 133,297,977 / 5,720 | 133,269,518 / 5,720 | 7,800,137 / 995,191 | 7,799,049 / 995,191 | 5.8517% / 5.8521% |

Within the eight merge rows, the nested `Store::from_unsorted` inclusive
values range from 4,294,284 to 8,804,364 Ir, `Cell::clone` values range from
2,131,042 to 4,292,060 Ir, and `memcpy` values range from 275,532 to 559,016
Ir. Those rows explain why the sorted merge is worth a direct attribution;
they remain nested diagnostics rather than independent costs. The owner’s
self Ir is only 5,720 or 9,800, so owner self Ir must not be confused with the
merge inclusive work.

The owner profile uses one measured call per job and excludes lifecycle owner
dumps. The 0550 source review also records that reduced-readback and raw-parser
names are absent from the positive owner graph; that absence is code-generation
evidence, not proof that either concept is free. The merge target is present as
a positive owner-reachable function, which makes this a narrower attribution
question than trying to infer the cost of an inlined reduced reader.

## Raw profile and execution bindings

Every row below is bound to the same exact owner, measured parent, normal
binary, source manifest, and plan. The first hash is the measured raw
`callgrind.N` dump; the second is the canonical profile result; the third is
the execution receipt.

| Job | Raw measured dump SHA-256 | Profile result SHA-256 | Receipt SHA-256 |
| --- | --- | --- | --- |
| medium R1 (`profile-r1-medium-c0.callgrind.4`) | `969bf7f08a1f6d1a179e978408144d84822bde036560c42234db95be42da068b` | `f8ffab7a6044acc7dd30a5ae847c5dbac1c93d518748eb4096012ef9b24f96a8` | `5a4e0f36c8b381816b60c7ce1e7941ee50577d3373008cc83eaee498ec0f5898` |
| dense-sparse R1 (`profile-r1-dense-sparse-c0.callgrind.4`) | `3bfad4962266d4a345b0bba1c3ef83888129f5908218a2db344fd463d84c9787` | `36ee0ac34c9b790543a05e1d2a0e732a2189bb9a59fdafd3e3f699918acfc9e1` | `2610343eb149e040cdf1995b322a15cebd75ee1b6f4203f7efb59e93e1b1e6b5` |
| noncompact R1 (`profile-r1-noncompact-c0.callgrind.4`) | `713f7dfe8de3f007d25fc282b7d6738439329edf1e669e9d4c9a4d17c12ea431` | `321bb9f5cbd3591455ef4e99a4c614999d86652ff243c01ac6ecfe73e9d16639` | `141ccffe7f827a106b4743b190979f294b7f717071e1d45f9a7319250b9ad74e` |
| vendor-extension R1 (`profile-r1-vendor-extension-c0.callgrind.5`) | `9224bbb071b1ac0b003415a371dfe8fd7c0181568a3afb6bb0ad20ca9ae8aaca` | `0275f7159ead9f04fcd87275dbdefe509dd29c2b64830380d392aad3dcb6f932` | `814869605ae7a17cdcb8739a54f8de9854abb524bd523479ef7a5d4f5aa54903` |
| medium R2 (`profile-r2-medium-c0.callgrind.4`) | `c65c98852ab47bf7cdb78cede58315e548ae1b4d32ff92abd3b4b1064b500f2e` | `36eadc0887ca112bb723f1ab7d92ef5d56e9d99a4bfa4d1ec6538149e911869b` | `79b27a61259c853ad47c0f768cf01055f7f73bcbd165fd1e28d109b26d46dd87` |
| dense-sparse R2 (`profile-r2-dense-sparse-c0.callgrind.4`) | `ad12980733aceecf6f16867f019446edebd0f3410ece4e44e523a46a3cce56ba` | `4219098c02f2f6d9a3e944aeeba38f4564204ccab630898e3e89367f3b882374` | `dec16400a4e00b99098545bd3bc0a3ffb0a9233c500a4c5f175f0e0c6ceb68c3` |
| noncompact R2 (`profile-r2-noncompact-c0.callgrind.4`) | `916597a7bff31eb4bcb8471fcb9d02769a7d51a43f0b0ecc646db93aeb76b85e` | `7fa72df68f6ce3f20b1d0e493bb4490950ff5faed0fc3c8589b4499adc854980` | `419ae241b11395dd3b674c8a4fb3e389d864a34309889580ca164fbe0df0bdb5` |
| vendor-extension R2 (`profile-r2-vendor-extension-c0.callgrind.5`) | `8ea754b9bb31d2580e6a70107e2f87bdbce4f9a59a886fdb48056ebf68780bd3` | `77884ea36cff0de85c5baf4d5fb840c83bebc260a8d83f34e1791bc4d737ef4c` | `1ab80fc3651ca20ed6a938bdbc67e757dbd0945ad943579092b1b94c9c15df7a` |

Common bindings are:

| Binding | SHA-256 |
| --- | --- |
| 0550 baseline source manifest | `fa85f76972a37f52644c60b955c9125ed80e4d78e2434ec54b77c3cb19e64163` |
| 0550 normal profile binary | `6d2d1ccf3a47e0db1c7188d509385ae75dd80ad81c60a4e13cab93c251b75c5b` |
| 0550 plan | `3eeb52cdad1312b3953006e538490d6523a707e9858354dcb6688b4a615471a7` |
| 0550 run driver | `dcb20341f9558bd6de1befcfe0afe43d2045c90c6c7c5673529fe234d2e6d39f` |
| 0550 profile analyzer | `ac23bbe997c8753557c141752ea2bb2c3e59d634a549d75b6ecd517aabcd0fbc` |
| 0550 analysis inputs | `0878a2a8963c13d91fe1ff365f544eb5e3eb23350706fb5e6f21e9de3aaf8435` |
| 0550 metrics analysis | `49e7978f1dd41b21031abe45cc2394a51085ca701d17fca79eda4ab5a8de14e4` |
| current repository HEAD | `d3c62f19a7632114f9cce5d1ad350e237215582d` |

## Current source identity

The current tree has no production Rust change for this audit. These hashes
were calculated from the current files at the recorded HEAD and match the
0550 source review for the relevant path.

| Current source file | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `b41af5d8c91a5c1e82f9030b8798a354ca4943072373846b8874c73996c03035` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` |
| `tools/perf-baseline/src/lib.rs` | `b71922e53c507842e4d8fb4983f52872a690940e3f7e9bf1b0f2d03fdbcd525f` |

## Why the compact directions are closed

0551 identified a planning-time compact per-cell source proof but recorded no
candidate latency result. 0552 then measured a compact source-cell offset
candidate and recorded `adoption_allowed: false`; its decision and metrics
hashes are `df2fb73e62c0800898b585a04ac2a1eb2bc86f8b6f340f938bd0d557c3fc6a78`
and `c0dd9f49b7006c5d426211624e36217fc680af1ad37d1febaa5d82a6f10c2a3c`.
0553 measured a commit-local compact collector and also recorded
`adoption_allowed: false`; its decision and metrics hashes are
`3aba2611db5007ebbdf6c4a76d1bbc382192c8ad4492d1d816aa272349395154` and
`147abc506cde809dfe7f20368a5cfc567629b5f58feffeeba09d23ad1c9fb02a`.
Both final sources were restored. Their failed gates remain evidence about
those candidates, not permission to alter the merge consumer or infer a
benefit for this proposal.

The current hotspot inventory records the same 0550 attribution and the
0552/0553 rejections. Its current SHA-256 is
`12e7a588ca432d019c4e3ec3f1deb07f959c02cc58120a12c2badb55494c2a55`.
The 0550 source review, SHA-256
`0bfb5748fe0b7d8a528e43bd1ce1d4a32a2962b45bb4595fa8f636a7a3dd7430`,
explicitly names a specialized fallible sorted Store merge as the strongest
conditional implementation hypothesis.

## Feasible next change and proof boundary

Implement one private helper used only by `merge_omitted_cells` that:

1. reserves the exact combined cell count with the existing fallible error
   label;
2. walks the parsed cells and the ordered omitted source entries by address;
3. moves parsed entries, clones only the source entries that must remain owned
   by the new store, and refuses equal or out-of-order addresses;
4. rebuilds the existing cell-row index and stored/content/styled extents from
   the merged order; and
5. preserves the existing merge index, row, column, default, declared-extent,
   resource, and fallback behavior.

An uncertain precondition must return the same optional refusal route so
`try_rewritten_value_cells` falls back to the complete parse. No complete
output validation, reduced parser, source identity check, execution fence,
calculation invalidation, publication check, or patch behavior may move or be
removed. The helper must use safe Rust and must not introduce a row arena,
layout cache, compact proof, or public API.

The required evidence gap is direct attribution and operation-local memory:
0550 shows merge as a positive nested owner, but its allocation region includes
`edit.set` staging plus commit, and it does not report merge-only allocation,
copy, or retained-state counters. A fresh measurement must profile the exact
merge call and its `from_unsorted` descendants, then compare a current-source
baseline and candidate across the four existing shapes and both repeats. It
must retain native workflow, allocation calls/bytes and incremental peak,
profile self/inclusive rows, exact output and semantic readback, and all
fallback/refusal cases before any admission decision. If the direct merge work
is below measurement noise or the end-to-end gates do not pass, close this
candidate rather than weakening the validator, parser, or preservation
contracts.

This audit performed no build, test, capture, analyzer mutation, or Rust source
edit. The OLE2 physical-reconciliation campaign remains the active next
campaign; this document records the independent OOXML queue item for a later
turn. ODF remains deferred until the OLE2/OOXML optimization goal is complete.
