# Final native OOXML review: rejected candidate

This review is bound to `final-native-comparison.json` by SHA-256 `92f098ff42ceee6c59f30f12558add73dc98fb7d540a1da77a681501f6f0d055`. Evidence validation passed, while native admission is rejected: the final stage failed the dense-sparse primary repeat-2 total p50 and mean gates. The measurements identify the rebuilt candidate as the final stage for comparison arithmetic, with its true binary, stage, and source labels preserved. No production speedup claim is retained.

Every one of the 57 greater-than-five-percent adverse metrics and all 101 same-build drift metrics below is copied from the comparison and has an individual review and classification. The adverse rows remain material evidence. Same-build drift is reported as within-stage repeat variation; it does not erase adverse evidence. External Cargo observations are listed as context only, with no causal attribution.

## Binding and decision

| item | value |
| --- | --- |
| comparison report | `final-native-comparison.json` |
| comparison SHA-256 | `92f098ff42ceee6c59f30f12558add73dc98fb7d540a1da77a681501f6f0d055` |
| report status | `pass` (`final-native-compare`) |
| admission status | **`reject`** |
| production decision | **reject** |
| retained speedup claim | **none** |
| adverse flags retained | 57 |
| same-build drift flags retained | 101 |
| complete review | `true` |

| frozen input | SHA-256 |
| --- | --- |
| primary `plan.json` | `12de5c9a6d145e9d8ac6b905db0eb4df78614bf08041a8f99be84d5771067e15` |
| `final-native-plan.json` | `3655d564e42d79213374b16727f3709b9570b71d888e9d10e7fcca0a7a30139a` |
| `final_capture.py` | `e9a2337a615be2f459776c8d452f0d360ab304cd5dde8ff9ae17a003f1e23b13` |
| `final-native-frozen-inputs.json` | `c48fdba82628a2d125251401e7bdc300647c8b84b72ec28c8abe28dc39848ab8` |
| `final-source-binding.json` | `ec12d415cbb5501c6aa376be9ae04c37448e979787756ee2ec94e4dfd3baec8c` |
| `analyze_final_native.py` | `0c42938416b116587a3dccf37c867969b3e58dfc9560ff2653962207897d508f` |

| true stage | binary SHA-256 | source manifest SHA-256 |
| --- | --- | --- |
| baseline | `44b70af7bc9e613770fd79902c4f20167ee25100c9a177299612e07130482456` | `c5a7cfbd0d7da965c7d945aea3355d24e34685b91f9dbbb25baedbed97140a08` |
| final | `0da73c5ab5f152bfa535a5259efe41c952e027f9fe31d70190400bf507c87ef1` | `b749ceb1a1a379a3495ba9e949dd88edc93a522ec3323124dbdff2c1662d7570` |

The report’s `candidate` slot is used for compatible arithmetic, but it contains the true final stage; the table above is the authoritative binary/source labeling.

## Matrix and numerical method

| check | validated value |
| --- | --- |
| ABBA order | baseline r1 → final r1 → final r2 → baseline r2 |
| capture order status | `pass` |
| children | 44 / 44 |
| native rows per stage | baseline 22, final 22 |
| primary rows per stage | 4 |
| guard rows per stage | 18 across 7 guard IDs |
| guard shape jobs per repeat | 9 |
| native samples per stage | baseline 1340, final 1340 |
| matched timing comparisons | 22 |
| identity-equal timing rows | 22 |
| bootstrap | 2,000 iterations, seed 5,310,531; matched-child within-stage resampling |

The native elapsed metric is the measured workflow total (open, planning, commit, and publication). Whole-child RSS is a separate /usr/bin/time observation. Publication remains diagnostic and is excluded from native admission. Planning Ir is conditional and unmeasured; no value is manufactured.

Bootstrap metadata copied from the comparison:

```json
{
  "iterations": 2000,
  "scope": "matched-child within-stage resampling; final median / baseline median",
  "seed": 5310531
}
```

## Exact native admission gates

A row passes only when its measured reduction is at least the required threshold. The frozen plan requires every primary shape and repeat to pass total p50, total mean, and planning p50. The failed row is bolded in the table; its planning p50 still passes.

| repeat | shape | metric | baseline | final | reduction % | required % | passed |
| ---: | --- | --- | ---: | ---: | ---: | ---: | :---: |
| 1 | dense-sparse | total_p50 | 42690382 | 41397319 | 3.028932840188687 | 1.0 | pass |
| 1 | dense-sparse | total_mean | 42710745.99499999 | 41415760.500000015 | 3.031989877094572 | 1.0 | pass |
| 1 | dense-sparse | planning_p50 | 13813451 | 12834324 | 7.088214234082417 | 2.0 | pass |
| 1 | medium | total_p50 | 22445689 | 21562787 | 3.9335036674525785 | 1.0 | pass |
| 1 | medium | total_mean | 22261311.384999994 | 21570519.67500001 | 3.103104296296991 | 1.0 | pass |
| 1 | medium | planning_p50 | 7148282 | 6864161 | 3.9746753135928325 | 2.0 | pass |
| 2 | dense-sparse | total_p50 | 42524801 | 42144774 | 0.8936596787366506 | 1.0 | fail **FAIL** |
| 2 | dense-sparse | total_mean | 42561618.155000016 | 42303862.004999995 | 0.6056070261739811 | 1.0 | fail **FAIL** |
| 2 | dense-sparse | planning_p50 | 13498266 | 12828839 | 4.959355520183111 | 2.0 | pass |
| 2 | medium | total_p50 | 21992628 | 21461182 | 2.416473374623533 | 1.0 | pass |
| 2 | medium | total_mean | 21988398.804999996 | 21592675.984999992 | 1.7996891156532036 | 1.0 | pass |
| 2 | medium | planning_p50 | 7188207 | 6795686 | 5.460624603604209 | 2.0 | pass |

| frozen gate | required value |
| --- | ---: |
| total p50 reduction | 1.0% |
| total mean reduction | 1.0% |
| planning p50 reduction | 2.0% |
| planning Ir reduction | 1.0% (conditional, unmeasured, unused) |
| every shape/repeat | `true` |

The rejection is determined by primary dense-sparse repeat 2: total p50 reduction is `0.8936596787366506%` and total mean reduction is `0.6056070261739811%`, each below the required `1.0%`. The planning p50 reduction is `4.959355520183111%` and passes its `2.0%` gate, but that does not override the two failed total gates.

## Flag inventory

The count tables are descriptive indexes of the retained rows. They do not filter or downgrade any flag.

### Adverse flags: 57

| dimension | counts |
| --- | --- |
| scope | {"guard0":6,"guard1":1,"guard2":14,"guard3":2,"guard4":11,"guard5":4,"guard6":2,"primary":17} |
| format | {"DOCX":4,"PPTX":2,"XLSX":51} |
| phase | {"commit_ns":1,"elapsed_ns":7,"open_ns":27,"publication_ns":16,"reopen_ns":6} |
| shape | {"dense-sparse":20,"medium":10,"noncompact":13,"vendor-extension":14} |
| statistic | {"mean":9,"p50":9,"p95":16,"p99":23} |
| guard adverse metrics | 40 |

### Same-build drift flags: 101

| dimension | counts |
| --- | --- |
| scope | {"guard0":14,"guard1":22,"guard3":6,"guard4":27,"guard5":4,"guard6":6,"primary":22} |
| true stage | {"baseline":44,"final":57} |
| format | {"DOCX":4,"PPTX":6,"XLSX":91} |
| phase | {"commit_ns":5,"elapsed_ns":13,"open_ns":38,"publication_ns":25,"reopen_ns":20} |
| shape | {"dense-sparse":44,"medium":24,"noncompact":33} |
| statistic | {"mean":22,"p50":21,"p95":27,"p99":31} |

## External Cargo observations

Each host file is a pre-child process snapshot with the recorded scope `Pre-child observation; not proof of an idle host.`. A Cargo or rustc entry is retained verbatim where present. These observations are context only: they neither establish a causal explanation nor excuse, remove, or reclassify any metric.

| stage | child | observed UTC | processes | raw Cargo/rustc entries | assessment |
| --- | --- | --- | ---: | --- | --- |
| baseline | `final-native-r1-guard0-dense-sparse` | 2026-09-12T14:40:39.536185+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard0-medium` | 2026-09-12T14:40:36.786603+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard1-dense-sparse` | 2026-09-12T14:40:47.956838+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard1-medium` | 2026-09-12T14:40:44.472410+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard2-vendor-extension` | 2026-09-12T14:40:53.065543+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard3-noncompact` | 2026-09-12T14:40:56.510668+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard4-noncompact` | 2026-09-12T14:41:00.040035+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard5-medium` | 2026-09-12T14:41:03.616492+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-guard6-medium` | 2026-09-12T14:41:06.954169+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-primary-dense-sparse` | 2026-09-12T14:40:16.461651+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r1-primary-medium` | 2026-09-12T14:40:02.725208+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r2-guard0-dense-sparse` | 2026-09-12T14:44:11.354973+00:00 | 2 | 2834814  298594       00:00 cargo<br>2834829 2834814       00:00 rustc | Cargo/rustc present; context only; no causal attribution |
| baseline | `final-native-r2-guard0-medium` | 2026-09-12T14:44:08.585580+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r2-guard1-dense-sparse` | 2026-09-12T14:44:19.796875+00:00 | 2 | 2834814  298594       00:08 cargo<br>2834829 2834814       00:08 rustc | Cargo/rustc present; context only; no causal attribution |
| baseline | `final-native-r2-guard1-medium` | 2026-09-12T14:44:16.297126+00:00 | 2 | 2834814  298594       00:05 cargo<br>2834829 2834814       00:05 rustc | Cargo/rustc present; context only; no causal attribution |
| baseline | `final-native-r2-guard2-vendor-extension` | 2026-09-12T14:44:24.970700+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r2-guard3-noncompact` | 2026-09-12T14:44:28.410668+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r2-guard4-noncompact` | 2026-09-12T14:44:31.887283+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| baseline | `final-native-r2-guard5-medium` | 2026-09-12T14:44:35.365324+00:00 | 2 | 2835716  298594       00:02 cargo<br>2835727 2835716       00:02 rustc | Cargo/rustc present; context only; no causal attribution |
| baseline | `final-native-r2-guard6-medium` | 2026-09-12T14:44:38.687270+00:00 | 2 | 2835716  298594       00:06 cargo<br>2835727 2835716       00:05 rustc | Cargo/rustc present; context only; no causal attribution |
| baseline | `final-native-r2-primary-dense-sparse` | 2026-09-12T14:43:46.817301+00:00 | 1 | 2833095 2833078       00:17 cargo | Cargo/rustc present; context only; no causal attribution |
| baseline | `final-native-r2-primary-medium` | 2026-09-12T14:43:33.302463+00:00 | 3 | 2830548  298594       00:13 cargo<br>2833051 2830548       00:05 rustc<br>2833095 2833078       00:03 cargo | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r1-guard0-dense-sparse` | 2026-09-12T14:41:50.062453+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard0-medium` | 2026-09-12T14:41:47.378970+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard1-dense-sparse` | 2026-09-12T14:41:57.989785+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard1-medium` | 2026-09-12T14:41:54.602824+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard2-vendor-extension` | 2026-09-12T14:42:02.966894+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard3-noncompact` | 2026-09-12T14:42:06.343574+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard4-noncompact` | 2026-09-12T14:42:09.840545+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard5-medium` | 2026-09-12T14:42:13.404271+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-guard6-medium` | 2026-09-12T14:42:16.697918+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-primary-dense-sparse` | 2026-09-12T14:41:26.844225+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r1-primary-medium` | 2026-09-12T14:41:13.379359+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r2-guard0-dense-sparse` | 2026-09-12T14:42:59.745894+00:00 | 3 | 2826087 2826073       00:08 cargo<br>2829018 2826087       00:02 rustc<br>2829315 2826087       00:00 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard0-medium` | 2026-09-12T14:42:57.001972+00:00 | 2 | 2826087 2826073       00:05 cargo<br>2828327 2826087       00:01 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard1-dense-sparse` | 2026-09-12T14:43:08.042959+00:00 | 2 | 2826087 2826073       00:16 cargo<br>2830029 2826087       00:04 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard1-medium` | 2026-09-12T14:43:04.579607+00:00 | 2 | 2826087 2826073       00:12 cargo<br>2830029 2826087       00:01 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard2-vendor-extension` | 2026-09-12T14:43:12.923393+00:00 | 4 | 2826087 2826073       00:21 cargo<br>2830029 2826087       00:09 rustc<br>2830213  298594       00:03 cargo<br>2830253 2830213       00:03 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard3-noncompact` | 2026-09-12T14:43:16.336749+00:00 | 2 | 2826087 2826073       00:24 cargo<br>2830029 2826087       00:12 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard4-noncompact` | 2026-09-12T14:43:19.855180+00:00 | 2 | 2826087 2826073       00:28 cargo<br>2830029 2826087       00:16 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard5-medium` | 2026-09-12T14:43:23.459569+00:00 | 4 | 2826087 2826073       00:31 cargo<br>2830029 2826087       00:20 rustc<br>2830548  298594       00:03 cargo<br>2831963 2830548       00:00 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-guard6-medium` | 2026-09-12T14:43:27.025285+00:00 | 4 | 2826087 2826073       00:35 cargo<br>2830029 2826087       00:23 rustc<br>2830548  298594       00:06 cargo<br>2832251 2830548       00:03 rustc | Cargo/rustc present; context only; no causal attribution |
| final | `final-native-r2-primary-dense-sparse` | 2026-09-12T14:42:36.228748+00:00 | 0 | none listed | none listed; context only; no causal attribution |
| final | `final-native-r2-primary-medium` | 2026-09-12T14:42:22.885686+00:00 | 14 | 2823602  298594       00:00 cargo<br>2824244 2823602       00:00 rustc<br>2824300 2823602       00:00 rustc<br>2824457 2823602       00:00 rustc<br>2824493 2823602       00:00 rustc<br>2824511 2823602       00:00 rustc<br>2824701 2823602       00:00 rustc<br>2824718 2823602       00:00 rustc<br>2824780 2823602       00:00 rustc<br>2824781 2823602       00:00 rustc<br>2824792 2823602       00:00 rustc<br>2824793 2823602       00:00 rustc<br>2824802 2823602       00:00 rustc<br>2824823 2823602       00:00 rustc | Cargo/rustc present; context only; no causal attribution |

## Individual adverse reviews

Each source row below is copied from the report’s `adverse_flags_over_five_percent` array. The `candidate` value is the final-stage value. All changes are positive increases over the retained five-percent threshold.

### Adverse 001 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 open_ns p50

Source row (copied exactly): `{"baseline":100625,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":105675,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":5.018633540372663,"guard":0,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=p50: final candidate=105675 versus baseline=100625 gives change=5.018633540372663%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard0","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 002 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 open_ns p95

Source row (copied exactly): `{"baseline":113500,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":124301,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":9.516299559471374,"guard":0,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=p95: final candidate=124301 versus baseline=113500 gives change=9.516299559471374%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard0","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 003 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 open_ns p99

Source row (copied exactly): `{"baseline":124091,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":133990,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":7.977210273106028,"guard":0,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=p99: final candidate=133990 versus baseline=124091 gives change=7.977210273106028%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard0","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 004 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 open_ns mean

Source row (copied exactly): `{"baseline":102762.26666666668,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":108532.53333333334,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":5.615160947533271,"guard":0,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=mean: final candidate=108532.53333333334 versus baseline=102762.26666666668 gives change=5.615160947533271%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard0","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 005 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 publication_ns p95

Source row (copied exactly): `{"baseline":12974830,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":14113033,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":8.772392393580496,"guard":0,"lane":"native","phase":"publication_ns","repeat":2,"shape":"dense-sparse","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 phase=publication_ns stat=p95: final candidate=14113033 versus baseline=12974830 gives change=8.772392393580496%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard0","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 006 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 publication_ns p99

Source row (copied exactly): `{"baseline":13005030,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":14605184,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":12.304116176587065,"guard":0,"lane":"native","phase":"publication_ns","repeat":2,"shape":"dense-sparse","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse repeat=2 phase=publication_ns stat=p99: final candidate=14605184 versus baseline=13005030 gives change=12.304116176587065%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard0","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 007 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium repeat=2 open_ns p99

Source row (copied exactly): `{"baseline":106340,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":120750,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":13.550874553319536,"guard":1,"lane":"native","phase":"open_ns","repeat":2,"shape":"medium","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium repeat=2 phase=open_ns stat=p99: final candidate=120750 versus baseline=106340 gives change=13.550874553319536%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard1","shape":"medium","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 008 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 open_ns p50

Source row (copied exactly): `{"baseline":97765,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":103266,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.626758042244151,"guard":2,"lane":"native","phase":"open_ns","repeat":1,"shape":"vendor-extension","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=open_ns stat=p50: final candidate=103266 versus baseline=97765 gives change=5.626758042244151%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 009 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 open_ns p95

Source row (copied exactly): `{"baseline":105830,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":113980,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.701029953699323,"guard":2,"lane":"native","phase":"open_ns","repeat":1,"shape":"vendor-extension","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=open_ns stat=p95: final candidate=113980 versus baseline=105830 gives change=7.701029953699323%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 010 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 open_ns p99

Source row (copied exactly): `{"baseline":106490,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":114360,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.390365292515733,"guard":2,"lane":"native","phase":"open_ns","repeat":1,"shape":"vendor-extension","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=open_ns stat=p99: final candidate=114360 versus baseline=106490 gives change=7.390365292515733%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 011 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 open_ns mean

Source row (copied exactly): `{"baseline":98906.63333333336,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":104038.36666666665,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.188462250088333,"guard":2,"lane":"native","phase":"open_ns","repeat":1,"shape":"vendor-extension","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=open_ns stat=mean: final candidate=104038.36666666665 versus baseline=98906.63333333336 gives change=5.188462250088333%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 012 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 publication_ns p50

Source row (copied exactly): `{"baseline":6735590,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7263593,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.839001483166275,"guard":2,"lane":"native","phase":"publication_ns","repeat":1,"shape":"vendor-extension","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=publication_ns stat=p50: final candidate=7263593 versus baseline=6735590 gives change=7.839001483166275%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p50","phase":"publication_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 013 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 publication_ns p95

Source row (copied exactly): `{"baseline":6800756,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7349038,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":8.062074275271748,"guard":2,"lane":"native","phase":"publication_ns","repeat":1,"shape":"vendor-extension","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=publication_ns stat=p95: final candidate=7349038 versus baseline=6800756 gives change=8.062074275271748%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 014 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 publication_ns p99

Source row (copied exactly): `{"baseline":6956526,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7366148,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.8883126434085,"guard":2,"lane":"native","phase":"publication_ns","repeat":1,"shape":"vendor-extension","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=publication_ns stat=p99: final candidate=7366148 versus baseline=6956526 gives change=5.8883126434085%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 015 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 publication_ns mean

Source row (copied exactly): `{"baseline":6745253.200000001,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7089553.800000002,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.104339152161108,"guard":2,"lane":"native","phase":"publication_ns","repeat":1,"shape":"vendor-extension","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=1 phase=publication_ns stat=mean: final candidate=7089553.800000002 versus baseline=6745253.200000001 gives change=5.104339152161108%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.mean","phase":"publication_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 016 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 open_ns p95

Source row (copied exactly): `{"baseline":107740,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":114451,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":6.228884351215891,"guard":2,"lane":"native","phase":"open_ns","repeat":2,"shape":"vendor-extension","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 phase=open_ns stat=p95: final candidate=114451 versus baseline=107740 gives change=6.228884351215891%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 017 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 open_ns p99

Source row (copied exactly): `{"baseline":111101,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":117370,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.642613477826486,"guard":2,"lane":"native","phase":"open_ns","repeat":2,"shape":"vendor-extension","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 phase=open_ns stat=p99: final candidate=117370 versus baseline=111101 gives change=5.642613477826486%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 018 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 publication_ns p50

Source row (copied exactly): `{"baseline":6803616,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7263302,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":6.756495369521143,"guard":2,"lane":"native","phase":"publication_ns","repeat":2,"shape":"vendor-extension","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 phase=publication_ns stat=p50: final candidate=7263302 versus baseline=6803616 gives change=6.756495369521143%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p50","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 019 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 publication_ns p95

Source row (copied exactly): `{"baseline":6874976,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7368917,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.184621444496675,"guard":2,"lane":"native","phase":"publication_ns","repeat":2,"shape":"vendor-extension","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 phase=publication_ns stat=p95: final candidate=7368917 versus baseline=6874976 gives change=7.184621444496675%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 020 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 publication_ns p99

Source row (copied exactly): `{"baseline":6888496,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7385358,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.212924272584331,"guard":2,"lane":"native","phase":"publication_ns","repeat":2,"shape":"vendor-extension","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 phase=publication_ns stat=p99: final candidate=7385358 versus baseline=6888496 gives change=7.212924272584331%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 021 — XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 publication_ns mean

Source row (copied exactly): `{"baseline":6807886.933333333,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7272295.6,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":6.821627198195546,"guard":2,"lane":"native","phase":"publication_ns","repeat":2,"shape":"vendor-extension","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard2 xlsx_source_backed_cell_values_one_percent_edit_save shape=vendor-extension repeat=2 phase=publication_ns stat=mean: final candidate=7272295.6 versus baseline=6807886.933333333 gives change=6.821627198195546%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.mean","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard2","shape":"vendor-extension","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 022 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact repeat=1 open_ns p99

Source row (copied exactly): `{"baseline":97410,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":114290,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":17.328816343291244,"guard":3,"lane":"native","phase":"open_ns","repeat":1,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact repeat=1 phase=open_ns stat=p99: final candidate=114290 versus baseline=97410 gives change=17.328816343291244%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard3","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 023 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact repeat=2 open_ns p99

Source row (copied exactly): `{"baseline":90910,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":96051,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.655043449565511,"guard":3,"lane":"native","phase":"open_ns","repeat":2,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=open_ns stat=p99: final candidate=96051 versus baseline=90910 gives change=5.655043449565511%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard3","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 024 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=1 open_ns p99

Source row (copied exactly): `{"baseline":96611,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":103640,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":7.275569034581975,"guard":4,"lane":"native","phase":"open_ns","repeat":1,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=1 phase=open_ns stat=p99: final candidate=103640 versus baseline=96611 gives change=7.275569034581975%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 025 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 commit_ns p99

Source row (copied exactly): `{"baseline":9948108,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":10450240,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.047512552135536,"guard":4,"lane":"native","phase":"commit_ns","repeat":2,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=commit_ns stat=p99: final candidate=10450240 versus baseline=9948108 gives change=5.047512552135536%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"commit_ns.p99","phase":"commit_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 026 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 elapsed_ns p99

Source row (copied exactly): `{"baseline":24664514,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":26214299,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":6.283460521460094,"guard":4,"lane":"native","phase":"elapsed_ns","repeat":2,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=elapsed_ns stat=p99: final candidate=26214299 versus baseline=24664514 gives change=6.283460521460094%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"elapsed_ns.p99","phase":"elapsed_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 027 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 open_ns p50

Source row (copied exactly): `{"baseline":94390,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":114200,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":20.987392732280963,"guard":4,"lane":"native","phase":"open_ns","repeat":2,"shape":"noncompact","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=open_ns stat=p50: final candidate=114200 versus baseline=94390 gives change=20.987392732280963%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 028 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 open_ns p95

Source row (copied exactly): `{"baseline":105570,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":156271,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":48.02595434308989,"guard":4,"lane":"native","phase":"open_ns","repeat":2,"shape":"noncompact","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=open_ns stat=p95: final candidate=156271 versus baseline=105570 gives change=48.02595434308989%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 029 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 open_ns p99

Source row (copied exactly): `{"baseline":107650,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":156681,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":45.546679052484905,"guard":4,"lane":"native","phase":"open_ns","repeat":2,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=open_ns stat=p99: final candidate=156681 versus baseline=107650 gives change=45.546679052484905%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 030 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 open_ns mean

Source row (copied exactly): `{"baseline":95323.9,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":115858.0666666667,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":21.541467215112586,"guard":4,"lane":"native","phase":"open_ns","repeat":2,"shape":"noncompact","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=open_ns stat=mean: final candidate=115858.0666666667 versus baseline=95323.9 gives change=21.541467215112586%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 031 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 publication_ns p95

Source row (copied exactly): `{"baseline":6796326,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7294228,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":7.326046455099422,"guard":4,"lane":"native","phase":"publication_ns","repeat":2,"shape":"noncompact","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=publication_ns stat=p95: final candidate=7294228 versus baseline=6796326 gives change=7.326046455099422%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 032 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 publication_ns p99

Source row (copied exactly): `{"baseline":6814016,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7883250,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":15.691686077637623,"guard":4,"lane":"native","phase":"publication_ns","repeat":2,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=publication_ns stat=p99: final candidate=7883250 versus baseline=6814016 gives change=15.691686077637623%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 033 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 reopen_ns p95

Source row (copied exactly): `{"baseline":28551868,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":32225432,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":12.86628251433497,"guard":4,"lane":"native","phase":"reopen_ns","repeat":2,"shape":"noncompact","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=reopen_ns stat=p95: final candidate=32225432 versus baseline=28551868 gives change=12.86628251433497%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"reopen_ns.p95","phase":"reopen_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 034 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 reopen_ns p99

Source row (copied exactly): `{"baseline":28579299,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":32828885,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":14.869454985582387,"guard":4,"lane":"native","phase":"reopen_ns","repeat":2,"shape":"noncompact","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact repeat=2 phase=reopen_ns stat=p99: final candidate=32828885 versus baseline=28579299 gives change=14.869454985582387%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":true,"metric":"reopen_ns.p99","phase":"reopen_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard4","shape":"noncompact","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 035 — DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 elapsed_ns p50

Source row (copied exactly): `{"baseline":2102308,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":4361701,"candidate_label":"final","candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":107.47202598287218,"guard":5,"lane":"native","phase":"elapsed_ns","repeat":2,"shape":"medium","stat":"p50","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 phase=elapsed_ns stat=p50: final candidate=4361701 versus baseline=2102308 gives change=107.47202598287218%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":true,"metric":"elapsed_ns.p50","phase":"elapsed_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard5","shape":"medium","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 036 — DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 elapsed_ns p95

Source row (copied exactly): `{"baseline":2143258,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":5071849,"candidate_label":"final","candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":136.64201883301033,"guard":5,"lane":"native","phase":"elapsed_ns","repeat":2,"shape":"medium","stat":"p95","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 phase=elapsed_ns stat=p95: final candidate=5071849 versus baseline=2143258 gives change=136.64201883301033%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":true,"metric":"elapsed_ns.p95","phase":"elapsed_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard5","shape":"medium","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 037 — DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 elapsed_ns p99

Source row (copied exactly): `{"baseline":2158008,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":5716251,"candidate_label":"final","candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":164.88553332517765,"guard":5,"lane":"native","phase":"elapsed_ns","repeat":2,"shape":"medium","stat":"p99","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 phase=elapsed_ns stat=p99: final candidate=5716251 versus baseline=2158008 gives change=164.88553332517765%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":true,"metric":"elapsed_ns.p99","phase":"elapsed_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard5","shape":"medium","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 038 — DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 elapsed_ns mean

Source row (copied exactly): `{"baseline":2106580.333333333,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":4235577.7,"candidate_label":"final","candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":101.06414329321409,"guard":5,"lane":"native","phase":"elapsed_ns","repeat":2,"shape":"medium","stat":"mean","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium repeat=2 phase=elapsed_ns stat=mean: final candidate=4235577.7 versus baseline=2106580.333333333 gives change=101.06414329321409%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":true,"metric":"elapsed_ns.mean","phase":"elapsed_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard5","shape":"medium","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 039 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium repeat=1 elapsed_ns p99

Source row (copied exactly): `{"baseline":6931397,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":8372422,"candidate_label":"final","candidate_stage":"final","case":"pptx_source_backed_one_edit_save","change_percent":20.789820580180308,"guard":6,"lane":"native","phase":"elapsed_ns","repeat":1,"shape":"medium","stat":"p99","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium repeat=1 phase=elapsed_ns stat=p99: final candidate=8372422 versus baseline=6931397 gives change=20.789820580180308%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":true,"metric":"elapsed_ns.p99","phase":"elapsed_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard6","shape":"medium","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 040 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium repeat=2 elapsed_ns p95

Source row (copied exactly): `{"baseline":8010530,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":9034374,"candidate_label":"final","candidate_stage":"final","case":"pptx_source_backed_one_edit_save","change_percent":12.78122671034252,"guard":6,"lane":"native","phase":"elapsed_ns","repeat":2,"shape":"medium","stat":"p95","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium repeat=2 phase=elapsed_ns stat=p95: final candidate=9034374 versus baseline=8010530 gives change=12.78122671034252%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":true,"metric":"elapsed_ns.p95","phase":"elapsed_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"guard6","shape":"medium","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 041 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 open_ns p50

Source row (copied exactly): `{"baseline":90815,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":98715,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":8.699003468589982,"guard":null,"lane":"native","phase":"open_ns","repeat":1,"shape":"dense-sparse","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=open_ns stat=p50: final candidate=98715 versus baseline=90815 gives change=8.699003468589982%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 042 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 open_ns p95

Source row (copied exactly): `{"baseline":99920,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":107820,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.906325060048047,"guard":null,"lane":"native","phase":"open_ns","repeat":1,"shape":"dense-sparse","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=open_ns stat=p95: final candidate=107820 versus baseline=99920 gives change=7.906325060048047%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 043 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 open_ns p99

Source row (copied exactly): `{"baseline":107110,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":115721,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":8.039398748949678,"guard":null,"lane":"native","phase":"open_ns","repeat":1,"shape":"dense-sparse","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=open_ns stat=p99: final candidate=115721 versus baseline=107110 gives change=8.039398748949678%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 044 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 open_ns mean

Source row (copied exactly): `{"baseline":92163.855,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":99584.38500000001,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":8.051453576893032,"guard":null,"lane":"native","phase":"open_ns","repeat":1,"shape":"dense-sparse","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=open_ns stat=mean: final candidate=99584.38500000001 versus baseline=92163.855 gives change=8.051453576893032%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 045 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 reopen_ns p50

Source row (copied exactly): `{"baseline":40064278,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":42292898,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.562611161993236,"guard":null,"lane":"native","phase":"reopen_ns","repeat":1,"shape":"dense-sparse","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=reopen_ns stat=p50: final candidate=42292898 versus baseline=40064278 gives change=5.562611161993236%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p50","phase":"reopen_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 046 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 reopen_ns p95

Source row (copied exactly): `{"baseline":40612052,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":42922739,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.689658330980163,"guard":null,"lane":"native","phase":"reopen_ns","repeat":1,"shape":"dense-sparse","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=reopen_ns stat=p95: final candidate=42922739 versus baseline=40612052 gives change=5.689658330980163%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p95","phase":"reopen_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 047 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 reopen_ns p99

Source row (copied exactly): `{"baseline":40770150,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":43155881,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.851661080471859,"guard":null,"lane":"native","phase":"reopen_ns","repeat":1,"shape":"dense-sparse","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=reopen_ns stat=p99: final candidate=43155881 versus baseline=40770150 gives change=5.851661080471859%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p99","phase":"reopen_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 048 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 reopen_ns mean

Source row (copied exactly): `{"baseline":40086004.81500001,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":42348310.55999998,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.643629878908318,"guard":null,"lane":"native","phase":"reopen_ns","repeat":1,"shape":"dense-sparse","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=1 phase=reopen_ns stat=mean: final candidate=42348310.55999998 versus baseline=40086004.81500001 gives change=5.643629878908318%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.mean","phase":"reopen_ns","repeat_scope":"final repeat 1 versus baseline repeat 1","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 049 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 open_ns p50

Source row (copied exactly): `{"baseline":97820,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":104886,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.22347168268247,"guard":null,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=p50: final candidate=104886 versus baseline=97820 gives change=7.22347168268247%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p50"}`

### Adverse 050 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 open_ns p95

Source row (copied exactly): `{"baseline":108080,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":137351,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":27.08271650629164,"guard":null,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=p95: final candidate=137351 versus baseline=108080 gives change=27.08271650629164%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 051 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 open_ns p99

Source row (copied exactly): `{"baseline":116460,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":152410,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":30.868967885969422,"guard":null,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=p99: final candidate=152410 versus baseline=116460 gives change=30.868967885969422%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 052 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 open_ns mean

Source row (copied exactly): `{"baseline":99140.20000000004,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":110622.685,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":11.582067617374125,"guard":null,"lane":"native","phase":"open_ns","repeat":2,"shape":"dense-sparse","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 phase=open_ns stat=mean: final candidate=110622.685 versus baseline=99140.20000000004 gives change=11.582067617374125%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"mean"}`

### Adverse 053 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 publication_ns p95

Source row (copied exactly): `{"baseline":13798752,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":14555005,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.480589838849204,"guard":null,"lane":"native","phase":"publication_ns","repeat":2,"shape":"dense-sparse","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 phase=publication_ns stat=p95: final candidate=14555005 versus baseline=13798752 gives change=5.480589838849204%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 054 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 publication_ns p99

Source row (copied exactly): `{"baseline":13813512,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":14751676,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":6.791639953691719,"guard":null,"lane":"native","phase":"publication_ns","repeat":2,"shape":"dense-sparse","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse repeat=2 phase=publication_ns stat=p99: final candidate=14751676 versus baseline=13813512 gives change=6.791639953691719%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"dense-sparse","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 055 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium repeat=2 open_ns p99

Source row (copied exactly): `{"baseline":95271,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":104491,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":9.677656369724264,"guard":null,"lane":"native","phase":"open_ns","repeat":2,"shape":"medium","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium repeat=2 phase=open_ns stat=p99: final candidate=104491 versus baseline=95271 gives change=9.677656369724264%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"medium","stage_scope":"final-vs-baseline","stat":"p99"}`

### Adverse 056 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium repeat=2 publication_ns p95

Source row (copied exactly): `{"baseline":6829076,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7296518,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":6.844879160811801,"guard":null,"lane":"native","phase":"publication_ns","repeat":2,"shape":"medium","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium repeat=2 phase=publication_ns stat=p95: final candidate=7296518 versus baseline=6829076 gives change=6.844879160811801%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"medium","stage_scope":"final-vs-baseline","stat":"p95"}`

### Adverse 057 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium repeat=2 publication_ns p99

Source row (copied exactly): `{"baseline":6848736,"baseline_stage":"baseline","baseline_zero_adverse":false,"candidate":7446398,"candidate_label":"final","candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":8.726602981922493,"guard":null,"lane":"native","phase":"publication_ns","repeat":2,"shape":"medium","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium repeat=2 phase=publication_ns stat=p99: final candidate=7446398 versus baseline=6848736 gives change=8.726602981922493%, above the retained 5.0% threshold. Retain this adverse metric as material evidence; it does not support retaining a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 2 versus baseline repeat 2","review_outcome":"retained adverse evidence; no speedup claim","scope":"primary","shape":"medium","stage_scope":"final-vs-baseline","stat":"p99"}`

## Individual same-build drift reviews

Each source row below is copied from the report’s `same_build_drift_over_five_percent` array. These are repeat-1 to repeat-2 comparisons within the named true stage; they are not final-vs-baseline measurements.

### Drift 001 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 commit_ns p50

Source row (copied exactly): `{"baseline":12599861,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":13304200,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":5.59005373154513,"guard":0,"kind":"guard","lane":"native","phase":"commit_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=commit_ns stat=p50: within-stage repeat drift is 5.59005373154513% (value 12599861 to 13304200). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"commit_ns.p50","phase":"commit_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 002 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 commit_ns p99

Source row (copied exactly): `{"baseline":12800946,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":13588642,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":6.1534202237865765,"guard":0,"kind":"guard","lane":"native","phase":"commit_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=commit_ns stat=p99: within-stage repeat drift is 6.1534202237865765% (value 12800946 to 13588642). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"commit_ns.p99","phase":"commit_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 003 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 commit_ns mean

Source row (copied exactly): `{"baseline":12607766.9,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":13318801.099999998,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":5.639652173455056,"guard":0,"kind":"guard","lane":"native","phase":"commit_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=commit_ns stat=mean: within-stage repeat drift is 5.639652173455056% (value 12607766.9 to 13318801.099999998). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"commit_ns.mean","phase":"commit_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 004 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":114400,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":124091,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":8.471153846153845,"guard":0,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 8.471153846153845% (value 114400 to 124091). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 005 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":105025,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":117930,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":12.287550583194484,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is 12.287550583194484% (value 105025 to 117930). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 006 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":111580,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":130761,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":17.190356694748154,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is 17.190356694748154% (value 111580 to 130761). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 007 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":113731,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":131800,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":15.887488899244717,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 15.887488899244717% (value 113731 to 131800). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 008 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns mean

Source row (copied exactly): `{"baseline":105483.59999999999,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":115728.29999999999,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":9.712125866011402,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=mean: within-stage repeat drift is 9.712125866011402% (value 105483.59999999999 to 115728.29999999999). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 009 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 publication_ns p50

Source row (copied exactly): `{"baseline":13681025,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":14410055,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.328767398641543,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=publication_ns stat=p50: within-stage repeat drift is 5.328767398641543% (value 13681025 to 14410055). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p50","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 010 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":13778522,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":14659586,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":6.3944739501087255,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is 6.3944739501087255% (value 13778522 to 14659586). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 011 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":13782891,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":14726926,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":6.849325007358753,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is 6.849325007358753% (value 13782891 to 14726926). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 012 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 publication_ns mean

Source row (copied exactly): `{"baseline":13696335.399999999,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":14434270.63333333,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.38782974994414,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=publication_ns stat=mean: within-stage repeat drift is 5.38782974994414% (value 13696335.399999999 to 14434270.63333333). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.mean","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 013 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":97341,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":88681,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-8.896559517572245,"guard":3,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is -8.896559517572245% (value 97341 to 88681). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard3","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 014 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":97410,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":90910,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-6.672826198542248,"guard":3,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is -6.672826198542248% (value 97410 to 90910). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard3","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 015 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 reopen_ns p50

Source row (copied exactly): `{"baseline":32236481,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":30426111,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-5.615904539952732,"guard":3,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=reopen_ns stat=p50: within-stage repeat drift is -5.615904539952732% (value 32236481 to 30426111). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p50","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard3","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 016 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 reopen_ns mean

Source row (copied exactly): `{"baseline":32250649.93333333,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":30420843.799999997,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-5.6737031257224295,"guard":3,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=reopen_ns stat=mean: within-stage repeat drift is -5.6737031257224295% (value 32250649.93333333 to 30420843.799999997). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.mean","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard3","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 017 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":89505,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":94390,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.457795653874076,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is 5.457795653874076% (value 89505 to 94390). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 018 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":95940,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":105570,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":10.037523452157604,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is 10.037523452157604% (value 95940 to 105570). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 019 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":96611,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":107650,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":11.42623510780345,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 11.42623510780345% (value 96611 to 107650). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 020 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 open_ns mean

Source row (copied exactly): `{"baseline":90435.8,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":95323.9,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.405049770113157,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=open_ns stat=mean: within-stage repeat drift is 5.405049770113157% (value 90435.8 to 95323.9). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 021 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 publication_ns p50

Source row (copied exactly): `{"baseline":6137913,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":6767166,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":10.251904841270964,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=publication_ns stat=p50: within-stage repeat drift is 10.251904841270964% (value 6137913 to 6767166). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p50","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 022 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":6181293,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":6796326,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":9.94990853855333,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is 9.94990853855333% (value 6181293 to 6796326). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 023 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":6213483,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":6814016,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":9.664997876392345,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is 9.664997876392345% (value 6213483 to 6814016). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 024 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 publication_ns mean

Source row (copied exactly): `{"baseline":6141887.999999999,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":6767408.800000001,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":10.184503527254197,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=publication_ns stat=mean: within-stage repeat drift is 10.184503527254197% (value 6141887.999999999 to 6767408.800000001). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.mean","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 025 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 reopen_ns p50

Source row (copied exactly): `{"baseline":32274446,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":28279392,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-12.378381336119604,"guard":4,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=reopen_ns stat=p50: within-stage repeat drift is -12.378381336119604% (value 32274446 to 28279392). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p50","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 026 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 reopen_ns p95

Source row (copied exactly): `{"baseline":32482302,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":28551868,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-12.100232304964099,"guard":4,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=reopen_ns stat=p95: within-stage repeat drift is -12.100232304964099% (value 32482302 to 28551868). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p95","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 027 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 reopen_ns p99

Source row (copied exactly): `{"baseline":32977844,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":28579299,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-13.337879213692682,"guard":4,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=reopen_ns stat=p99: within-stage repeat drift is -13.337879213692682% (value 32977844 to 28579299). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p99","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 028 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 reopen_ns mean

Source row (copied exactly): `{"baseline":32234227.1,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":28300933.466666665,"candidate_stage":"baseline","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-12.202227219939566,"guard":4,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"noncompact","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact baseline repeat 1 to 2 phase=reopen_ns stat=mean: within-stage repeat drift is -12.202227219939566% (value 32234227.1 to 28300933.466666665). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.mean","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 029 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium baseline repeat 1 to 2 elapsed_ns p95

Source row (copied exactly): `{"baseline":6905036,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":8010530,"candidate_stage":"baseline","case":"pptx_source_backed_one_edit_save","change_percent":16.009967218128907,"guard":6,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium baseline repeat 1 to 2 phase=elapsed_ns stat=p95: within-stage repeat drift is 16.009967218128907% (value 6905036 to 8010530). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":false,"metric":"elapsed_ns.p95","phase":"elapsed_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard6","shape":"medium","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 030 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium baseline repeat 1 to 2 elapsed_ns p99

Source row (copied exactly): `{"baseline":6931397,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":9916537,"candidate_stage":"baseline","case":"pptx_source_backed_one_edit_save","change_percent":43.06693152909868,"guard":6,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium baseline repeat 1 to 2 phase=elapsed_ns stat=p99: within-stage repeat drift is 43.06693152909868% (value 6931397 to 9916537). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":false,"metric":"elapsed_ns.p99","phase":"elapsed_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard6","shape":"medium","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 031 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium baseline repeat 1 to 2 elapsed_ns mean

Source row (copied exactly): `{"baseline":6746540.6,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":7103817.299999999,"candidate_stage":"baseline","case":"pptx_source_backed_one_edit_save","change_percent":5.295702215147102,"guard":6,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium baseline repeat 1 to 2 phase=elapsed_ns stat=mean: within-stage repeat drift is 5.295702215147102% (value 6746540.6 to 7103817.299999999). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":false,"metric":"elapsed_ns.mean","phase":"elapsed_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard6","shape":"medium","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 032 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":90815,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":97820,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.713483455376324,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is 7.713483455376324% (value 90815 to 97820). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 033 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":99920,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":108080,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":8.166533226581274,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is 8.166533226581274% (value 99920 to 108080). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 034 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":107110,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":116460,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":8.72934366539071,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 8.72934366539071% (value 107110 to 116460). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 035 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 open_ns mean

Source row (copied exactly): `{"baseline":92163.855,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":99140.20000000004,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.569502165463948,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=open_ns stat=mean: within-stage repeat drift is 7.569502165463948% (value 92163.855 to 99140.20000000004). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 036 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 reopen_ns p50

Source row (copied exactly): `{"baseline":40064278,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":46617282,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":16.3562263620475,"guard":null,"kind":"primary","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=reopen_ns stat=p50: within-stage repeat drift is 16.3562263620475% (value 40064278 to 46617282). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p50","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 037 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 reopen_ns p95

Source row (copied exactly): `{"baseline":40612052,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":47568270,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":17.1284573357682,"guard":null,"kind":"primary","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=reopen_ns stat=p95: within-stage repeat drift is 17.1284573357682% (value 40612052 to 47568270). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p95","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 038 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 reopen_ns p99

Source row (copied exactly): `{"baseline":40770150,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":47896122,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":17.478405156713926,"guard":null,"kind":"primary","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=reopen_ns stat=p99: within-stage repeat drift is 17.478405156713926% (value 40770150 to 47896122). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p99","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 039 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 reopen_ns mean

Source row (copied exactly): `{"baseline":40086004.81500001,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":46696930.47500001,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":16.49185467723693,"guard":null,"kind":"primary","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"dense-sparse","stage":"baseline","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse baseline repeat 1 to 2 phase=reopen_ns stat=mean: within-stage repeat drift is 16.49185467723693% (value 40086004.81500001 to 46696930.47500001). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.mean","phase":"reopen_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-baseline same-build","stat":"mean"}`

### Drift 040 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":101240,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":92870,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-8.267483208218096,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is -8.267483208218096% (value 101240 to 92870). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 041 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":108271,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":95271,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-12.006908590481292,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is -12.006908590481292% (value 108271 to 95271). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 042 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 publication_ns p50

Source row (copied exactly): `{"baseline":7255642,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":6749686,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-6.973276796181505,"guard":null,"kind":"primary","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 phase=publication_ns stat=p50: within-stage repeat drift is -6.973276796181505% (value 7255642 to 6749686). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p50","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-baseline same-build","stat":"p50"}`

### Drift 043 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":7373438,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":6829076,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-7.382743301021854,"guard":null,"kind":"primary","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is -7.382743301021854% (value 7373438 to 6829076). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-baseline same-build","stat":"p95"}`

### Drift 044 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":7389918,"baseline_stage":"baseline","baseline_zero_drift":false,"candidate":6848736,"candidate_stage":"baseline","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-7.323247700448099,"guard":null,"kind":"primary","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"baseline","shape":"medium","stage":"baseline","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium baseline repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is -7.323247700448099% (value 7389918 to 6848736). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"baseline repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-baseline same-build","stat":"p99"}`

### Drift 045 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":92375,"baseline_stage":"final","baseline_zero_drift":false,"candidate":105675,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":14.397834912043294,"guard":0,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is 14.397834912043294% (value 92375 to 105675). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 046 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":100661,"baseline_stage":"final","baseline_zero_drift":false,"candidate":124301,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":23.48476569873139,"guard":0,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is 23.48476569873139% (value 100661 to 124301). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 047 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":103070,"baseline_stage":"final","baseline_zero_drift":false,"candidate":133990,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":29.99902978558262,"guard":0,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 29.99902978558262% (value 103070 to 133990). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 048 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 open_ns mean

Source row (copied exactly): `{"baseline":93292.7,"baseline_stage":"final","baseline_zero_drift":false,"candidate":108532.53333333334,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":16.335504635768228,"guard":0,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=mean: within-stage repeat drift is 16.335504635768228% (value 93292.7 to 108532.53333333334). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 049 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":13094409,"baseline_stage":"final","baseline_zero_drift":false,"candidate":14113033,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":7.779075787231027,"guard":0,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is 7.779075787231027% (value 13094409 to 14113033). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 050 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":13100969,"baseline_stage":"final","baseline_zero_drift":false,"candidate":14605184,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":11.481707956106145,"guard":0,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is 11.481707956106145% (value 13100969 to 14605184). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 051 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns p50

Source row (copied exactly): `{"baseline":37265434,"baseline_stage":"final","baseline_zero_drift":false,"candidate":43490551,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":16.704802096226757,"guard":0,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=p50: within-stage repeat drift is 16.704802096226757% (value 37265434 to 43490551). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p50","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 052 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns p95

Source row (copied exactly): `{"baseline":37924062,"baseline_stage":"final","baseline_zero_drift":false,"candidate":44471724,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":17.265191687535996,"guard":0,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=p95: within-stage repeat drift is 17.265191687535996% (value 37924062 to 44471724). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p95","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 053 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns p99

Source row (copied exactly): `{"baseline":38080942,"baseline_stage":"final","baseline_zero_drift":false,"candidate":45227737,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":18.767379756519674,"guard":0,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=p99: within-stage repeat drift is 18.767379756519674% (value 38080942 to 45227737). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p99","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 054 — XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns mean

Source row (copied exactly): `{"baseline":37357378.366666675,"baseline_stage":"final","baseline_zero_drift":false,"candidate":43378208.19999999,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_edit_save","change_percent":16.116842499594664,"guard":0,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard0 xlsx_source_backed_cell_values_one_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=mean: within-stage repeat drift is 16.116842499594664% (value 37357378.366666675 to 43378208.19999999). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.mean","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard0","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 055 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":103195,"baseline_stage":"final","baseline_zero_drift":false,"candidate":95415,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-7.539124957604537,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is -7.539124957604537% (value 103195 to 95415). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 056 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":109710,"baseline_stage":"final","baseline_zero_drift":false,"candidate":101971,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-7.05405159055692,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is -7.05405159055692% (value 109710 to 101971). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 057 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":111770,"baseline_stage":"final","baseline_zero_drift":false,"candidate":103051,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-7.800841012794136,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is -7.800841012794136% (value 111770 to 103051). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 058 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns mean

Source row (copied exactly): `{"baseline":103816.29999999999,"baseline_stage":"final","baseline_zero_drift":false,"candidate":95629.13333333332,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-7.886205409619373,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=mean: within-stage repeat drift is -7.886205409619373% (value 103816.29999999999 to 95629.13333333332). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 059 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns p50

Source row (copied exactly): `{"baseline":41457421,"baseline_stage":"final","baseline_zero_drift":false,"candidate":38508067,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-7.114176253269589,"guard":1,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=p50: within-stage repeat drift is -7.114176253269589% (value 41457421 to 38508067). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p50","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 060 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns p95

Source row (copied exactly): `{"baseline":41965987,"baseline_stage":"final","baseline_zero_drift":false,"candidate":39060285,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-6.923945336970149,"guard":1,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=p95: within-stage repeat drift is -6.923945336970149% (value 41965987 to 39060285). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p95","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 061 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns p99

Source row (copied exactly): `{"baseline":42030778,"baseline_stage":"final","baseline_zero_drift":false,"candidate":39162424,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-6.82441329066048,"guard":1,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=p99: within-stage repeat drift is -6.82441329066048% (value 42030778 to 39162424). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p99","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 062 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 reopen_ns mean

Source row (copied exactly): `{"baseline":41420020.43333334,"baseline_stage":"final","baseline_zero_drift":false,"candidate":38593582.900000006,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-6.823843889412274,"guard":1,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=reopen_ns stat=mean: within-stage repeat drift is -6.823843889412274% (value 41420020.43333334 to 38593582.900000006). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.mean","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 063 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":89081,"baseline_stage":"final","baseline_zero_drift":false,"candidate":94265,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.819422772532867,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is 5.819422772532867% (value 89081 to 94265). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"medium","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 064 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":100110,"baseline_stage":"final","baseline_zero_drift":false,"candidate":120750,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":20.61732094695834,"guard":1,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 20.61732094695834% (value 100110 to 120750). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"medium","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 065 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 publication_ns p50

Source row (copied exactly): `{"baseline":6760946,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7270596,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":7.538146288995651,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=publication_ns stat=p50: within-stage repeat drift is 7.538146288995651% (value 6760946 to 7270596). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p50","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"medium","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 066 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":6853226,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7369237,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":7.529461307711149,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is 7.529461307711149% (value 6853226 to 7369237). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"medium","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 067 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":6894276,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7374338,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":6.96319671565222,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is 6.96319671565222% (value 6894276 to 7374338). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"medium","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 068 — XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 publication_ns mean

Source row (copied exactly): `{"baseline":6783602.399999999,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7290761.566666667,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":7.476251359700381,"guard":1,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard1 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=publication_ns stat=mean: within-stage repeat drift is 7.476251359700381% (value 6783602.399999999 to 7290761.566666667). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.mean","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard1","shape":"medium","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 069 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":101590,"baseline_stage":"final","baseline_zero_drift":false,"candidate":91630,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-9.804114578206512,"guard":3,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is -9.804114578206512% (value 101590 to 91630). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard3","shape":"noncompact","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 070 — XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":114290,"baseline_stage":"final","baseline_zero_drift":false,"candidate":96051,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":-15.958526555254181,"guard":3,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard3 xlsx_source_backed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is -15.958526555254181% (value 114290 to 96051). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard3","shape":"noncompact","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 071 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 commit_ns p95

Source row (copied exactly): `{"baseline":9903288,"baseline_stage":"final","baseline_zero_drift":false,"candidate":10431989,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.338641065472394,"guard":4,"kind":"guard","lane":"native","phase":"commit_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=commit_ns stat=p95: within-stage repeat drift is 5.338641065472394% (value 9903288 to 10431989). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"commit_ns.p95","phase":"commit_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 072 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 commit_ns p99

Source row (copied exactly): `{"baseline":9903847,"baseline_stage":"final","baseline_zero_drift":false,"candidate":10450240,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.516977392724254,"guard":4,"kind":"guard","lane":"native","phase":"commit_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=commit_ns stat=p99: within-stage repeat drift is 5.516977392724254% (value 9903847 to 10450240). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"commit_ns.p99","phase":"commit_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 073 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 elapsed_ns p95

Source row (copied exactly): `{"baseline":23605828,"baseline_stage":"final","baseline_zero_drift":false,"candidate":25641688,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":8.624395636535187,"guard":4,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=elapsed_ns stat=p95: within-stage repeat drift is 8.624395636535187% (value 23605828 to 25641688). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"elapsed_ns.p95","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 074 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 elapsed_ns p99

Source row (copied exactly): `{"baseline":23607679,"baseline_stage":"final","baseline_zero_drift":false,"candidate":26214299,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":11.041407331910946,"guard":4,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=elapsed_ns stat=p99: within-stage repeat drift is 11.041407331910946% (value 23607679 to 26214299). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"elapsed_ns.p99","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 075 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 elapsed_ns mean

Source row (copied exactly): `{"baseline":23507068.333333336,"baseline_stage":"final","baseline_zero_drift":false,"candidate":24724375.499999996,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":5.1784729146360675,"guard":4,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=elapsed_ns stat=mean: within-stage repeat drift is 5.1784729146360675% (value 23507068.333333336 to 24724375.499999996). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"elapsed_ns.mean","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 076 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":93831,"baseline_stage":"final","baseline_zero_drift":false,"candidate":114200,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":21.708177468000976,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is 21.708177468000976% (value 93831 to 114200). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 077 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":99720,"baseline_stage":"final","baseline_zero_drift":false,"candidate":156271,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":56.70978740473325,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is 56.70978740473325% (value 99720 to 156271). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 078 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":103640,"baseline_stage":"final","baseline_zero_drift":false,"candidate":156681,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":51.17811655731377,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 51.17811655731377% (value 103640 to 156681). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 079 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 open_ns mean

Source row (copied exactly): `{"baseline":94237.06666666667,"baseline_stage":"final","baseline_zero_drift":false,"candidate":115858.0666666667,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":22.943201401288693,"guard":4,"kind":"guard","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=open_ns stat=mean: within-stage repeat drift is 22.943201401288693% (value 94237.06666666667 to 115858.0666666667). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 080 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 publication_ns p50

Source row (copied exactly): `{"baseline":6103653,"baseline_stage":"final","baseline_zero_drift":false,"candidate":6884986,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":12.801071751621528,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=publication_ns stat=p50: within-stage repeat drift is 12.801071751621528% (value 6103653 to 6884986). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p50","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 081 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":6159843,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7294228,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":18.415810273086496,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is 18.415810273086496% (value 6159843 to 7294228). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 082 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":6222453,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7883250,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":26.690390429626376,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is 26.690390429626376% (value 6222453 to 7883250). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 083 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 publication_ns mean

Source row (copied exactly): `{"baseline":6110254.666666666,"baseline_stage":"final","baseline_zero_drift":false,"candidate":6960236.800000001,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":13.910748073566403,"guard":4,"kind":"guard","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=publication_ns stat=mean: within-stage repeat drift is 13.910748073566403% (value 6110254.666666666 to 6960236.800000001). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.mean","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 084 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 reopen_ns p50

Source row (copied exactly): `{"baseline":32426701,"baseline_stage":"final","baseline_zero_drift":false,"candidate":29495082,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-9.040756258245331,"guard":4,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=reopen_ns stat=p50: within-stage repeat drift is -9.040756258245331% (value 32426701 to 29495082). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.p50","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 085 — XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 reopen_ns mean

Source row (copied exactly): `{"baseline":32434468.23333333,"baseline_stage":"final","baseline_zero_drift":false,"candidate":29609087.666666664,"candidate_stage":"final","case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","change_percent":-8.711043283771136,"guard":4,"kind":"guard","lane":"native","phase":"reopen_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"noncompact","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX guard4 xlsx_source_backed_managed_cell_values_one_percent_edit_save shape=noncompact final repeat 1 to 2 phase=reopen_ns stat=mean: within-stage repeat drift is -8.711043283771136% (value 32434468.23333333 to 29609087.666666664). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_managed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"reopen_ns.mean","phase":"reopen_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard4","shape":"noncompact","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 086 — DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 elapsed_ns p50

Source row (copied exactly): `{"baseline":2042153,"baseline_stage":"final","baseline_zero_drift":false,"candidate":4361701,"candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":113.58345824235498,"guard":5,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 phase=elapsed_ns stat=p50: within-stage repeat drift is 113.58345824235498% (value 2042153 to 4361701). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":false,"metric":"elapsed_ns.p50","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard5","shape":"medium","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 087 — DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 elapsed_ns p95

Source row (copied exactly): `{"baseline":2070918,"baseline_stage":"final","baseline_zero_drift":false,"candidate":5071849,"candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":144.90824841930007,"guard":5,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 phase=elapsed_ns stat=p95: within-stage repeat drift is 144.90824841930007% (value 2070918 to 5071849). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":false,"metric":"elapsed_ns.p95","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard5","shape":"medium","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 088 — DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 elapsed_ns p99

Source row (copied exactly): `{"baseline":2072078,"baseline_stage":"final","baseline_zero_drift":false,"candidate":5716251,"candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":175.8704546836557,"guard":5,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 phase=elapsed_ns stat=p99: within-stage repeat drift is 175.8704546836557% (value 2072078 to 5716251). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":false,"metric":"elapsed_ns.p99","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard5","shape":"medium","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 089 — DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 elapsed_ns mean

Source row (copied exactly): `{"baseline":2041410.8333333337,"baseline_stage":"final","baseline_zero_drift":false,"candidate":4235577.7,"candidate_stage":"final","case":"docx_source_backed_one_edit_save","change_percent":107.48286581216502,"guard":5,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** DOCX guard5 docx_source_backed_one_edit_save shape=medium final repeat 1 to 2 phase=elapsed_ns stat=mean: within-stage repeat drift is 107.48286581216502% (value 2041410.8333333337 to 4235577.7). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"docx_source_backed_one_edit_save","format":"DOCX","material_guard_regression":false,"metric":"elapsed_ns.mean","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard5","shape":"medium","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 090 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium final repeat 1 to 2 elapsed_ns p95

Source row (copied exactly): `{"baseline":6881397,"baseline_stage":"final","baseline_zero_drift":false,"candidate":9034374,"candidate_stage":"final","case":"pptx_source_backed_one_edit_save","change_percent":31.286917467485154,"guard":6,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium final repeat 1 to 2 phase=elapsed_ns stat=p95: within-stage repeat drift is 31.286917467485154% (value 6881397 to 9034374). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":false,"metric":"elapsed_ns.p95","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard6","shape":"medium","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 091 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium final repeat 1 to 2 elapsed_ns p99

Source row (copied exactly): `{"baseline":8372422,"baseline_stage":"final","baseline_zero_drift":false,"candidate":9205515,"candidate_stage":"final","case":"pptx_source_backed_one_edit_save","change_percent":9.950442058462894,"guard":6,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium final repeat 1 to 2 phase=elapsed_ns stat=p99: within-stage repeat drift is 9.950442058462894% (value 8372422 to 9205515). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":false,"metric":"elapsed_ns.p99","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard6","shape":"medium","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 092 — PPTX guard6 pptx_source_backed_one_edit_save shape=medium final repeat 1 to 2 elapsed_ns mean

Source row (copied exactly): `{"baseline":6751934.966666668,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7125360.000000001,"candidate_stage":"final","case":"pptx_source_backed_one_edit_save","change_percent":5.53063729400356,"guard":6,"kind":"guard","lane":"native","phase":"elapsed_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** PPTX guard6 pptx_source_backed_one_edit_save shape=medium final repeat 1 to 2 phase=elapsed_ns stat=mean: within-stage repeat drift is 5.53063729400356% (value 6751934.966666668 to 7125360.000000001). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"pptx_source_backed_one_edit_save","format":"PPTX","material_guard_regression":false,"metric":"elapsed_ns.mean","phase":"elapsed_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":true,"scope":"guard6","shape":"medium","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 093 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p50

Source row (copied exactly): `{"baseline":98715,"baseline_stage":"final","baseline_zero_drift":false,"candidate":104886,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":6.2513295851694295,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p50","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p50: within-stage repeat drift is 6.2513295851694295% (value 98715 to 104886). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p50","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p50"}`

### Drift 094 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p95

Source row (copied exactly): `{"baseline":107820,"baseline_stage":"final","baseline_zero_drift":false,"candidate":137351,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":27.389167130402512,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p95: within-stage repeat drift is 27.389167130402512% (value 107820 to 137351). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p95","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 095 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":115721,"baseline_stage":"final","baseline_zero_drift":false,"candidate":152410,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":31.7047035542382,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 31.7047035542382% (value 115721 to 152410). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 096 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 open_ns mean

Source row (copied exactly): `{"baseline":99584.38500000001,"baseline_stage":"final","baseline_zero_drift":false,"candidate":110622.685,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":11.084368297298797,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"mean","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=open_ns stat=mean: within-stage repeat drift is 11.084368297298797% (value 99584.38500000001 to 110622.685). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.mean","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"mean"}`

### Drift 097 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":13852252,"baseline_stage":"final","baseline_zero_drift":false,"candidate":14555005,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.073203981561991,"guard":null,"kind":"primary","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is 5.073203981561991% (value 13852252 to 14555005). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 098 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":14039843,"baseline_stage":"final","baseline_zero_drift":false,"candidate":14751676,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":5.070092308012275,"guard":null,"kind":"primary","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"dense-sparse","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=dense-sparse final repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is 5.070092308012275% (value 14039843 to 14751676). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"dense-sparse","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 099 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 open_ns p99

Source row (copied exactly): `{"baseline":97420,"baseline_stage":"final","baseline_zero_drift":false,"candidate":104491,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.258263190309999,"guard":null,"kind":"primary","lane":"native","phase":"open_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=open_ns stat=p99: within-stage repeat drift is 7.258263190309999% (value 97420 to 104491). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"open_ns.p99","phase":"open_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-final same-build","stat":"p99"}`

### Drift 100 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 publication_ns p95

Source row (copied exactly): `{"baseline":6814756,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7296518,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":7.06939470760215,"guard":null,"kind":"primary","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p95","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=publication_ns stat=p95: within-stage repeat drift is 7.06939470760215% (value 6814756 to 7296518). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p95","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-final same-build","stat":"p95"}`

### Drift 101 — XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 publication_ns p99

Source row (copied exactly): `{"baseline":6830607,"baseline_stage":"final","baseline_zero_drift":false,"candidate":7446398,"candidate_stage":"final","case":"xlsx_source_backed_cell_values_one_percent_edit_save","change_percent":9.015172443678864,"guard":null,"kind":"primary","lane":"native","phase":"publication_ns","repeat_first":1,"repeat_second":2,"repeat_stage":"final","shape":"medium","stage":"final","stat":"p99","threshold_percent":5.0}`

**Review:** XLSX primary xlsx_source_backed_cell_values_one_percent_edit_save shape=medium final repeat 1 to 2 phase=publication_ns stat=p99: within-stage repeat drift is 9.015172443678864% (value 6830607 to 7446398). This is not a final-vs-baseline result; retain it as repeat-variation evidence, with no blanket causal attribution, and do not use it to erase adverse evidence or authorize a production speedup claim.

**Classification:** `{"case":"xlsx_source_backed_cell_values_one_percent_edit_save","format":"XLSX","material_guard_regression":false,"metric":"publication_ns.p99","phase":"publication_ns","repeat_scope":"final repeat 1 to 2","review_outcome":"retained same-build drift; no speedup claim","same_build_guard_drift":false,"scope":"primary","shape":"medium","stage_scope":"within-final same-build","stat":"p99"}`

## Preserved originals and disposition

The original campaign reports remain separate and unchanged. Their bindings are recorded here so this rejected final-stage review cannot replace earlier evidence.

| path | SHA-256 | role |
| --- | --- | --- |
| `comparison.json` | `99a8b1bb06d3d1e46e7b29882dad39baf4dc1bc4aa7aa1dab28c07299e26026b` | original main campaign comparison; preserved |
| `native-pilot-comparison.json` | `66ac7faa7a37f013e0074113acef6e8a4668e7b7ed5825f520b9056788e5b60f` | recovered original native pilot; preserved |
| `adverse-review.json` | `c7606d7148a1873e25a1ae8e6fe1bb71681fff2a37ddd09cb3f614df6cf36867` | original all-flag adverse review; preserved |
| `reopen-comparison.json` | `5c24c4d57ccca251b883ef2516a1c57f80db0339f8e5f02490de826fabceac97` | supplemental reopen comparison; preserved |
| `reopen-review.json` | `cd853a82b44783c8f40c820f3d0b3d936186dfeb345271dfd79e339db485e715` | supplemental reopen review; preserved |

The complete final native matrix is valid evidence (`status=pass`), but its admission decision is `reject`. The candidate is not authorized for production retention and this review retains no speedup claim. OLE2 and OOXML remain the active optimization priority; ODF remains deferred.
