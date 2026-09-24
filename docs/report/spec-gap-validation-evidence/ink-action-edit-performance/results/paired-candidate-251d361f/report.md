# Ink-action edit bounded allocator/runtime profile

This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane across 34 bounded lanes. The timer covers the named detached draft or source-backed edit workflow, including its bounded validation and output allocation; fixture construction and post-timer semantic, source, patch, and inverse checks stay outside the timed interval. Requested allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.

| lane | fresh processes | samples | input actions | result actions | queued operations | input bytes | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| draft_small_8 | 3 | 60 | 8 | 8 | 8 | 0 | 28640 / 29711 / 43250 | 24423 / 24423 | 14010 / 14010 | 3816–3848 |
| draft_scaled_128 | 3 | 60 | 128 | 128 | 128 | 0 | 405352 / 416931 / 424872 | 373487 / 373487 | 212903 / 212903 | 3992–4092 |
| draft_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 0 | 3264024 / 3349104 / 3513205 | 2990143 / 2990143 | 1707647 / 1707647 | 5784–6060 |
| draft_opaque_64 | 3 | 60 | 64 | 64 | 64 | 0 | 216421 / 223621 / 224791 | 196316 / 196316 | 107704 / 107704 | 3736–3872 |
| scalar_edit_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 43400 / 45240 / 55680 | 29257 / 29257 | 11958 / 11958 | 3872–3912 |
| scalar_edit_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 472692 / 479862 / 627102 | 332539 / 332539 | 163576 / 163576 | 4104–4132 |
| scalar_edit_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 3651886 / 3671355 / 3824227 | 2605059 / 2605059 | 1300768 / 1300768 | 5352–5636 |
| scalar_batch_scaled_128 | 3 | 60 | 128 | 128 | 128 | 15380 | 520352 / 526943 / 529112 | 447047 / 447047 | 194130 / 194130 | 4092–4156 |
| scalar_batch_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 120260 | 4466409 / 4491489 / 4502409 | 3559495 / 3559495 | 1554522 / 1554522 | 5528–5692 |
| scalar_coalesce_scaled_128 | 3 | 60 | 128 | 128 | 128 | 15380 | 493902 / 506432 / 516602 | 387001 / 387001 | 181022 / 181022 | 4052–4156 |
| scalar_coalesce_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 120260 | 4283048 / 4347919 / 4464699 | 3051108 / 3051108 | 1444584 / 1444584 | 5576–5636 |
| no_op_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 5060 / 5340 / 5450 | 2407 / 2407 | 2366 / 2366 | 3816–3876 |
| no_op_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 61250 / 67500 / 72780 | 34979 / 34979 | 34938 / 34938 | 3816–3848 |
| no_op_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 488632 / 494042 / 495072 | 279635 / 279635 | 279594 / 279594 | 4356–4428 |
| add_small_8 | 3 | 60 | 8 | 9 | 1 | 1528 | 46850 / 49400 / 55310 | 33843 / 33843 | 14197 / 14197 | 3856–3912 |
| add_scaled_128 | 3 | 60 | 128 | 129 | 1 | 15380 | 473622 / 482792 / 498672 | 365953 / 365953 | 180233 / 180233 | 3992–4128 |
| add_near_1024 | 3 | 60 | 1024 | 1025 | 1 | 120260 | 3647626 / 3668665 / 3678825 | 2853531 / 2853531 | 1424958 / 1424958 | 5628–5692 |
| insert_batch_scaled_128 | 3 | 60 | 128 | 256 | 128 | 15380 | 922044 / 931394 / 941844 | 799827 / 799827 | 384414 / 384414 | 4372–4424 |
| insert_batch_near_1024 | 3 | 60 | 1024 | 2048 | 1024 | 120260 | 7293460 / 7362991 / 7510222 | 6388987 / 6388987 | 3079022 / 3079022 | 7696–7796 |
| remove_small_8 | 3 | 60 | 8 | 7 | 1 | 1528 | 41901 / 44550 / 49690 | 27705 / 27705 | 11163 / 11163 | 3784–3876 |
| remove_scaled_128 | 3 | 60 | 128 | 127 | 1 | 15380 | 479942 / 485782 / 487352 | 330975 / 330975 | 162775 / 162775 | 4028–4116 |
| remove_near_1024 | 3 | 60 | 1024 | 1023 | 1 | 120260 | 3735256 / 3758486 / 3765776 | 2603481 / 2603481 | 1299956 / 1299956 | 5332–5404 |
| remove_batch_scaled_128 | 3 | 60 | 128 | 64 | 64 | 15380 | 335532 / 343792 / 535212 | 256172 / 256172 | 113249 / 113249 | 3840–3904 |
| remove_batch_near_1024 | 3 | 60 | 1024 | 512 | 512 | 120260 | 2982963 / 3005292 / 3007273 | 2024852 / 2024852 | 902285 / 902285 | 4760–4900 |
| clear_batch_scaled_128 | 3 | 60 | 128 | 128 | 64 | 15380 | 455582 / 468902 / 472742 | 327282 / 327282 | 138124 / 138124 | 3868–4120 |
| clear_batch_near_1024 | 3 | 60 | 1024 | 1024 | 512 | 120260 | 4062317 / 4100427 / 4113337 | 2596922 / 2596922 | 1103284 / 1103284 | 5384–5440 |
| move_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 42920 / 47630 / 53810 | 29289 / 29289 | 11938 / 11938 | 3840–3848 |
| move_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 470712 / 479612 / 483512 | 332579 / 332579 | 163564 / 163564 | 4052–4096 |
| move_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 3652246 / 3701716 / 3710445 | 2605099 / 2605099 | 1300756 / 1300756 | 5272–5404 |
| move_batch_scaled_128 | 3 | 60 | 128 | 128 | 64 | 15380 | 481582 / 489352 / 492702 | 389487 / 389487 | 176524 / 176524 | 3992–4092 |
| move_batch_near_1024 | 3 | 60 | 1024 | 1024 | 512 | 120260 | 3880326 / 3977987 / 3985957 | 3086839 / 3086839 | 1410484 / 1410484 | 5528–5628 |
| cap_refusal_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 11070 / 11860 / 30520 | 5094 / 5094 | 2821 / 2821 | 3736–3904 |
| cap_refusal_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 96481 / 101220 / 106770 | 28760 / 28760 | 26487 / 26487 | 4052–4156 |
| cap_refusal_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 749843 / 755853 / 764834 | 206192 / 206192 | 203919 / 203919 | 7176–7232 |

The small, scaled, and near-limit source-backed lanes use 8, 128, and 1,024 direct actions. Draft creation adds a 64-action lane with complete namespace-bearing `inkml:definitions` and `inkml:trace` opaque payloads. Scalar replacement, distinct scalar batches, repeated writes to one scalar, root insertion batches, structural add/remove/clear/move, and exact no-op edits retain source comments, identifiers, namespaces, and opaque bytes; the no-op lanes additionally require source allocation sharing. The repeated-write lane reports only the final scalar value and measured cost; the public API exposes no internal coalescing diagnostic, so the report makes no coalescing claim. Caller-cap lanes attempt a bounded property expansion and must refuse at the configured complete-output limit without changing the source snapshot. Batch operation counts and resulting action counts are retained in each raw receipt so the report distinguishes one edit in a large source from many queued edits.

Every successful edit is checked after timing by applying its source-checked patch and inverse. The report records only absolute observations for this detached API and fixture matrix. It makes no before/after speedup, native-application, asymptotic, or host-placement claim. It makes no package-wide performance claim.

Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.
