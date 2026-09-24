# Ink-action edit bounded allocator/runtime profile

This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane across 34 bounded lanes. The timer covers the named detached draft or source-backed edit workflow, including its bounded validation and output allocation; fixture construction and post-timer semantic, source, patch, and inverse checks stay outside the timed interval. Requested allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.

| lane | fresh processes | samples | input actions | result actions | queued operations | input bytes | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| draft_small_8 | 3 | 60 | 8 | 8 | 8 | 0 | 28040 / 29510 / 43090 | 24423 / 24423 | 14010 / 14010 | 3752–3984 |
| draft_scaled_128 | 3 | 60 | 128 | 128 | 128 | 0 | 403002 / 416362 / 553912 | 373487 / 373487 | 212903 / 212903 | 4048–4240 |
| draft_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 0 | 3223583 / 3292484 / 3306324 | 2990143 / 2990143 | 1707647 / 1707647 | 5896–6032 |
| draft_opaque_64 | 3 | 60 | 64 | 64 | 64 | 0 | 214251 / 221321 / 223981 | 196316 / 196316 | 107704 / 107704 | 3776–3908 |
| scalar_edit_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 42380 / 44020 / 50280 | 32340 / 32340 | 14481 / 14481 | 3792–3804 |
| scalar_edit_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 465952 / 475102 / 481132 | 363322 / 363322 | 178976 / 178976 | 4012–4112 |
| scalar_edit_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 3589626 / 3625756 / 3768856 | 2845602 / 2845602 | 1421048 / 1421048 | 5552–5780 |
| scalar_batch_scaled_128 | 3 | 60 | 128 | 128 | 128 | 15380 | 516042 / 526982 / 533032 | 479653 / 479653 | 210442 / 210442 | 4052–4180 |
| scalar_batch_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 120260 | 4374648 / 4401598 / 4580769 | 3816245 / 3816245 | 1682906 / 1682906 | 5804–6228 |
| scalar_coalesce_scaled_128 | 3 | 60 | 128 | 128 | 128 | 15380 | 496982 / 506702 / 512602 | 417801 / 417801 | 196430 / 196430 | 4004–4100 |
| scalar_coalesce_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 120260 | 4238598 / 4276638 / 4406198 | 3291677 / 3291677 | 1564880 / 1564880 | 5796–5896 |
| no_op_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 5030 / 5280 / 5710 | 2407 / 2407 | 2366 / 2366 | 3752–3980 |
| no_op_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 60070 / 64450 / 73301 | 34979 / 34979 | 34938 / 34938 | 3792–3980 |
| no_op_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 482032 / 489572 / 492272 | 279635 / 279635 | 279594 / 279594 | 4312–4360 |
| add_small_8 | 3 | 60 | 8 | 9 | 1 | 1528 | 46180 / 50840 / 57050 | 37138 / 37138 | 16813 / 16813 | 3796–3988 |
| add_scaled_128 | 3 | 60 | 128 | 129 | 1 | 15380 | 465332 / 473082 / 475812 | 396960 / 396960 | 195745 / 195745 | 4048–4180 |
| add_near_1024 | 3 | 60 | 1024 | 1025 | 1 | 120260 | 3602465 / 3625845 / 3641525 | 3094308 / 3094308 | 1545358 / 1545358 | 5804–5904 |
| insert_batch_scaled_128 | 3 | 60 | 128 | 256 | 128 | 15380 | 913834 / 922673 / 924804 | 859603 / 859603 | 414310 / 414310 | 4268–4308 |
| insert_batch_near_1024 | 3 | 60 | 1024 | 2048 | 1024 | 120260 | 7200960 / 7282421 / 7341301 | 6864699 / 6864699 | 3316886 / 3316886 | 8096–8376 |
| remove_small_8 | 3 | 60 | 8 | 7 | 1 | 1528 | 41140 / 46410 / 53491 | 30563 / 30563 | 13575 / 13575 | 3984–3988 |
| remove_scaled_128 | 3 | 60 | 128 | 127 | 1 | 15380 | 473252 / 482762 / 486552 | 361529 / 361529 | 178063 / 178063 | 4048–4048 |
| remove_near_1024 | 3 | 60 | 1024 | 1023 | 1 | 120260 | 3674776 / 3703656 / 3890126 | 2843785 / 2843785 | 1420116 / 1420116 | 5560–5832 |
| remove_batch_scaled_128 | 3 | 60 | 128 | 64 | 64 | 15380 | 329862 / 337291 / 344051 | 272582 / 272582 | 121465 / 121465 | 3756–3796 |
| remove_batch_near_1024 | 3 | 60 | 1024 | 512 | 512 | 120260 | 3011133 / 3034033 / 3213563 | 2148830 / 2148830 | 964285 / 964285 | 5048–5268 |
| clear_batch_scaled_128 | 3 | 60 | 128 | 128 | 64 | 15380 | 447251 / 455032 / 456882 | 352174 / 352174 | 150580 / 150580 | 4020–4104 |
| clear_batch_near_1024 | 3 | 60 | 1024 | 1024 | 512 | 120260 | 4015997 / 4057147 / 4069948 | 2790358 / 2790358 | 1200012 / 1200012 | 5308–5460 |
| move_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 42891 / 44790 / 56161 | 32361 / 32361 | 14450 / 14450 | 3752–3792 |
| move_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 462922 / 472232 / 477782 | 363359 / 363359 | 178964 / 178964 | 3996–4056 |
| move_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 3613575 / 3733366 / 3745185 | 2845639 / 2845639 | 1421036 / 1421036 | 5584–5824 |
| move_batch_scaled_128 | 3 | 60 | 128 | 128 | 64 | 15380 | 474962 / 483402 / 485622 | 420267 / 420267 | 191924 / 191924 | 4052–4244 |
| move_batch_near_1024 | 3 | 60 | 1024 | 1024 | 512 | 120260 | 3841926 / 3892027 / 3918397 | 3327379 / 3327379 | 1530764 / 1530764 | 5780–5856 |
| cap_refusal_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 10940 / 11490 / 19660 | 5094 / 5094 | 2821 / 2821 | 3776–3796 |
| cap_refusal_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 95271 / 102730 / 111971 | 28760 / 28760 | 26487 / 26487 | 4016–4056 |
| cap_refusal_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 733754 / 739643 / 744523 | 206192 / 206192 | 203919 / 203919 | 7084–7316 |

The small, scaled, and near-limit source-backed lanes use 8, 128, and 1,024 direct actions. Draft creation adds a 64-action lane with complete namespace-bearing `inkml:definitions` and `inkml:trace` opaque payloads. Scalar replacement, distinct scalar batches, repeated writes to one scalar, root insertion batches, structural add/remove/clear/move, and exact no-op edits retain source comments, identifiers, namespaces, and opaque bytes; the no-op lanes additionally require source allocation sharing. The repeated-write lane reports only the final scalar value and measured cost; the public API exposes no internal coalescing diagnostic, so the report makes no coalescing claim. Caller-cap lanes attempt a bounded property expansion and must refuse at the configured complete-output limit without changing the source snapshot. Batch operation counts and resulting action counts are retained in each raw receipt so the report distinguishes one edit in a large source from many queued edits.

Every successful edit is checked after timing by applying its source-checked patch and inverse. The report records only absolute observations for this detached API and fixture matrix. It makes no before/after speedup, native-application, asymptotic, or host-placement claim. It makes no package-wide performance claim.

Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.
