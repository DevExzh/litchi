# Matched DOCX source-backed SVG lifecycle profile

This comparison uses two isolated committed worktrees and the same frozen public-API harness: control `211806fee` and candidate `ceeecf972`. Each lane used three fresh processes, two warmups, and twenty measured samples per process (60 samples per row). The frozen verifier passed independently for both profiles.

The percentage columns are matched scenario observations computed from p50 values as `(candidate / control - 1) * 100`; they are not a general speedup claim. Allocation traffic, peak live allocation, and process RSS are separate metrics. The temporary Cargo targets and worktrees were removed after verification; per-commit source and harness snapshots remain under `source/`.

- control production commit: `211806feea6e83279975ff77adeb5c8fa7be01a6`
- candidate production commit: `ceeecf972e716be850547917fe348c1437186906`
- frozen harness anchor: `211806feea6e83279975ff77adeb5c8fa7be01a6`
- comparison driver SHA-256: `fd01bf52371e4605c5a96bbb58ef1cf6357e30f723bf9db9db5972226f47d857`
- control frozen-verifier receipt SHA-256: `fcf2bf6efed59c431019ffffd12ac6979af837ee7043cdc02e07695cea7c31e1`
- candidate frozen-verifier receipt SHA-256: `fcf2bf6efed59c431019ffffd12ac6979af837ee7043cdc02e07695cea7c31e1`
- allocator: `CountingAllocator` process-local `GlobalAlloc` observer
- RSS: `/usr/bin/time -v` maximum resident set size sidecars
- hard scratch-budget/OOM claim: none
- retained per-profile reports: [`control/full-report.md`](control/full-report.md), [`candidate/full-report.md`](candidate/full-report.md)

| lane | elapsed p50 ns control/candidate (delta) | requested alloc p50 B control/candidate (delta) | peak live p50 B control/candidate (delta) | RSS KiB control/candidate ranges | dominant p50 phase control → candidate |
|---|---:|---:|---:|---:|---|
| native_svg_capture | 166031/172161 (+3.7%) | 1052550/1052550 (+0.0%) | 210091/210091 (+0.0%) | 5012–5180 / 5120–5332 | capture_ns (162681) → capture_ns (168711) |
| native_floating_capture | 243531/246821 (+1.4%) | 6422961/6422961 (+0.0%) | 438197/438197 (+0.0%) | 5916–6084 / 5764–6080 | capture_ns (239431) → capture_ns (243451) |
| lazy_inventory_1 | 95241/96720 (+1.6%) | 939605/939605 (+0.0%) | 133592/133592 (+0.0%) | 5116–5252 / 5132–5176 | validation_ns (36020) → validation_ns (36070) |
| lazy_inventory_64 | 1554396/1578887 (+1.6%) | 2838947/2838947 (+0.0%) | 602032/602032 (+0.0%) | 5664–5728 / 5664–5816 | capture_ns (896094) → capture_ns (916954) |
| single_attach_1 | 646712/589752 (-8.8%) | 2935189/2860932 (-2.5%) | 489758/489758 (+0.0%) | 5660–5720 / 5640–5724 | publish_ns (208301) → publish_ns (212221) |
| single_attach_16 | 3283144/2777112 (-15.4%) | 6417493/5945682 (-7.4%) | 559600/559600 (+0.0%) | 5628–5784 / 5664–5824 | commit_ns (1055465) → publish_ns (984875) |
| single_attach_64 | 11428289/9441950 (-17.4%) | 17937861/16136354 (-10.0%) | 819480/819480 (+0.0%) | 6156–6296 / 6196–6276 | commit_ns (3771346) → publish_ns (3313074) |
| single_detach_1 | 634032/579142 (-8.7%) | 2927979/2862184 (-2.2%) | 485645/485645 (+0.0%) | 5276–5644 / 5388–5460 | publish_ns (214991) → publish_ns (216061) |
| single_detach_16 | 5175982/4473250 (-13.6%) | 9349796/8720831 (-6.7%) | 622192/622192 (+0.0%) | 5736–6024 / 5732–6120 | commit_ns (1721487) → publish_ns (1618457) |
| single_detach_64 | 19145581/16377210 (-14.5%) | 30297870/27809049 (-8.2%) | 1370318/1370318 (+0.0%) | 6860–7016 / 6876–7032 | commit_ns (6502667) → publish_ns (5908375) |
| batch_attach_1 | 642163/578503 (-9.9%) | 2936327/2862815 (-2.5%) | 493748/494966 (+0.2%) | 5620–5680 / 5732–5848 | publish_ns (206471) → publish_ns (207021) |
| batch_attach_16 | 11928851/4456870 (-62.6%) | 16615985/9899870 (-40.4%) | 810614/815942 (+0.7%) | 5988–5992 / 5848–5924 | stage_ns (7409551) → publish_ns (1553407) |
| batch_attach_64 | 123263912/10476575 (-91.5%) | 125340031/23029940 (-81.6%) | 1183047/1201527 (+1.6%) | 6480–6748 / 6424–6880 | stage_ns (114497894) → readback_ns (3619965) |
| batch_detach_1 | 627633/576813 (-8.1%) | 2929111/2864062 (-2.2%) | 489633/490853 (+0.2%) | 5520–5776 / 5496–5596 | publish_ns (210611) → publish_ns (215581) |
| batch_detach_16 | 10929036/4042538 (-63.0%) | 15660617/9540296 (-39.1%) | 749901/755261 (+0.7%) | 5956–6040 / 5936–6048 | stage_ns (6724498) → publish_ns (1521037) |
| batch_detach_64 | 114883874/11190188 (-90.3%) | 107823271/13867942 (-87.1%) | 1060831/1079439 (+1.8%) | 6320–6624 / 6512–6708 | stage_ns (103059045) → publish_ns (4420559) |
| shared_svg_cleanup | 115236966/11204048 (-90.3%) | 108209705/14254376 (-86.8%) | 1060831/1079439 (+1.8%) | 6484–6520 / 6432–6632 | stage_ns (103515367) → publish_ns (4410269) |
| exact_inverse_single_1 | 786603/730613 (-7.1%) | 4059451/3986685 (-1.8%) | 493748/494966 (+0.2%) | 5700–5848 / 5636–5716 | publish_ns (206831) → publish_ns (210821) |
| exact_inverse_batch_64 | 123890251/10436485 (-91.6%) | 125340031/23029940 (-81.6%) | 1183047/1201527 (+1.6%) | 6440–6636 / 6428–6816 | stage_ns (115092705) → readback_ns (3611106) |
| large_unchanged_media_managed_cap | 6921898/6706619 (-3.1%) | 499845/499845 (+0.0%) | 133864/133864 (+0.0%) | 41940–42004 / 41812–42024 | validation_ns (1964708) → validation_ns (1951668) |
| noop_detach_64 | 3842856/3875017 (+0.8%) | 10778602/10796546 (+0.2%) | 777581/777581 (+0.0%) | 5936–5952 / 5956–6036 | reopen_ns (1301475) → reopen_ns (1312115) |

## Cost and scaling observations

The control profile's 16- and 64-owner attach/detach and shared-cleanup rows are dominated by `stage_ns`, and the stage p50 grows sharply with owner count. The candidate removes most of that repeated source-layout/stage work: candidate 16-owner batch attach/detach rows are dominated by publication, and the 64-owner successful detach/shared-cleanup rows show the same shift. Single-owner and native capture rows remain dominated by publication or capture/validation, so their deltas are smaller and should not be generalized beyond this fixture matrix.

The raw harness receipts intentionally retain `baseline_commit=892441d95db29da4351390716ef5c65b4c7c97de` and `opc_baseline_label=OPC57680dc86`; the outer matched manifests bind those receipts to the control and candidate production commits above.

Both 64-owner refusal lanes still return the typed `DocxError::Opc(OpcError::SourceBackedOverlayUnavailable { .. })` match in every sample. The receipt class `topology_part_bound` is only a diagnostic label; `typed_match=true` is the gate. Failed operations republish byte-identical physical source bytes and matching metadata, and emit zero output bytes.

The candidate still builds a full projected story buffer for each staged operation. The profile therefore does not establish a bounded scratch-memory guarantee: the residual projected-story copy remains a material allocation component as same-story owner counts grow, even where repeated source scanning/layout projection is avoided. The `large_unchanged_media_managed_cap` lane keeps its 8-MiB unchanged member under the 2-MiB managed execution cap and passes exact opaque-preservation checks; RSS remains an independent process-level observation.

Raw receipts, `/usr/bin/time -v` sidecars, per-commit source/harness/fixture snapshots, binary/build metadata, and the two frozen-verifier outputs are retained beside this report. Recompute with `python3 matched/compare.py`.
