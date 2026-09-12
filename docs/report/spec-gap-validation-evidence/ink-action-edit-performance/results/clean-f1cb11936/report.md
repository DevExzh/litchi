# Ink-action edit bounded allocator/runtime profile

This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane across 34 bounded lanes. The timer covers the named detached draft or source-backed edit workflow, including its bounded validation and output allocation; fixture construction and post-timer semantic, source, patch, and inverse checks stay outside the timed interval. Requested allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.

| lane | fresh processes | samples | input actions | result actions | queued operations | input bytes | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| draft_small_8 | 3 | 60 | 8 | 8 | 8 | 0 | 28190 / 29240 / 38280 | 24423 / 24423 | 14010 / 14010 | 3784–3904 |
| draft_scaled_128 | 3 | 60 | 128 | 128 | 128 | 0 | 399822 / 409752 / 422371 | 373487 / 373487 | 212903 / 212903 | 4020–4236 |
| draft_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 0 | 3234343 / 3254983 / 3260323 | 2990143 / 2990143 | 1707647 / 1707647 | 5956–5960 |
| draft_opaque_64 | 3 | 60 | 64 | 64 | 64 | 0 | 214961 / 224221 / 227191 | 196316 / 196316 | 107704 / 107704 | 3900–3916 |
| scalar_edit_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 43200 / 49880 / 62300 | 32340 / 32340 | 14481 / 14481 | 3880–3912 |
| scalar_edit_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 466352 / 474812 / 479052 | 363322 / 363322 | 178976 / 178976 | 4024–4168 |
| scalar_edit_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 3608855 / 3644824 / 3905716 | 2845602 / 2845602 | 1421048 / 1421048 | 5556–5628 |
| scalar_batch_scaled_128 | 3 | 60 | 128 | 128 | 128 | 15380 | 515022 / 523472 / 528162 | 479653 / 479653 | 210442 / 210442 | 4020–4020 |
| scalar_batch_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 120260 | 4417328 / 4509218 / 4633928 | 3816245 / 3816245 | 1682906 / 1682906 | 6064–6212 |
| scalar_coalesce_scaled_128 | 3 | 60 | 128 | 128 | 128 | 15380 | 498392 / 510822 / 518963 | 417801 / 417801 | 196430 / 196430 | 4168–4236 |
| scalar_coalesce_near_1024 | 3 | 60 | 1024 | 1024 | 1024 | 120260 | 4233107 / 4355047 / 4479068 | 3291677 / 3291677 | 1564880 / 1564880 | 5812–5948 |
| no_op_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 5120 / 5490 / 18990 | 2407 / 2407 | 2366 / 2366 | 3756–3912 |
| no_op_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 60720 / 67490 / 70860 | 34979 / 34979 | 34938 / 34938 | 3760–3904 |
| no_op_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 485502 / 494402 / 497572 | 279635 / 279635 | 279594 / 279594 | 4272–4396 |
| add_small_8 | 3 | 60 | 8 | 9 | 1 | 1528 | 46390 / 50270 / 55011 | 37138 / 37138 | 16813 / 16813 | 3768–3904 |
| add_scaled_128 | 3 | 60 | 128 | 129 | 1 | 15380 | 469852 / 479312 / 488862 | 396960 / 396960 | 195745 / 195745 | 4052–4164 |
| add_near_1024 | 3 | 60 | 1024 | 1025 | 1 | 120260 | 3585815 / 3632694 / 3741625 | 3094308 / 3094308 | 1545358 / 1545358 | 5804–5952 |
| insert_batch_scaled_128 | 3 | 60 | 128 | 256 | 128 | 15380 | 913564 / 921914 / 923844 | 859603 / 859603 | 414310 / 414310 | 4276–4492 |
| insert_batch_near_1024 | 3 | 60 | 1024 | 2048 | 1024 | 120260 | 7218889 / 7256929 / 7393210 | 6864699 / 6864699 | 3316886 / 3316886 | 8072–8292 |
| remove_small_8 | 3 | 60 | 8 | 7 | 1 | 1528 | 41290 / 45280 / 52830 | 30563 / 30563 | 13575 / 13575 | 3768–3912 |
| remove_scaled_128 | 3 | 60 | 128 | 127 | 1 | 15380 | 474822 / 480702 / 481492 | 361529 / 361529 | 178063 / 178063 | 4016–4160 |
| remove_near_1024 | 3 | 60 | 1024 | 1023 | 1 | 120260 | 3684055 / 3730855 / 3891455 | 2843785 / 2843785 | 1420116 / 1420116 | 5692–5952 |
| remove_batch_scaled_128 | 3 | 60 | 128 | 64 | 64 | 15380 | 332141 / 338831 / 340231 | 272582 / 272582 | 121465 / 121465 | 3760–4000 |
| remove_batch_near_1024 | 3 | 60 | 1024 | 512 | 512 | 120260 | 2951892 / 3066783 / 3088543 | 2148830 / 2148830 | 964285 / 964285 | 5032–5156 |
| clear_batch_scaled_128 | 3 | 60 | 128 | 128 | 64 | 15380 | 451842 / 458852 / 463922 | 352174 / 352174 | 150580 / 150580 | 4140–4160 |
| clear_batch_near_1024 | 3 | 60 | 1024 | 1024 | 512 | 120260 | 4023856 / 4046366 / 4232697 | 2790358 / 2790358 | 1200012 / 1200012 | 5352–5448 |
| move_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 42610 / 48110 / 52150 | 32361 / 32361 | 14450 / 14450 | 3764–3912 |
| move_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 464132 / 471122 / 472922 | 363359 / 363359 | 178964 / 178964 | 4084–4168 |
| move_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 3602935 / 3645305 / 3673105 | 2845639 / 2845639 | 1421036 / 1421036 | 5692–6084 |
| move_batch_scaled_128 | 3 | 60 | 128 | 128 | 64 | 15380 | 477042 / 484362 / 491422 | 420267 / 420267 | 191924 / 191924 | 4040–4164 |
| move_batch_near_1024 | 3 | 60 | 1024 | 1024 | 512 | 120260 | 3854196 / 3918796 / 3950396 | 3327379 / 3327379 | 1530764 / 1530764 | 5844–5952 |
| cap_refusal_small_8 | 3 | 60 | 8 | 8 | 1 | 1528 | 11010 / 11750 / 21790 | 5094 / 5094 | 2821 / 2821 | 3784–3884 |
| cap_refusal_scaled_128 | 3 | 60 | 128 | 128 | 1 | 15380 | 96430 / 102341 / 105980 | 28760 / 28760 | 26487 / 26487 | 4052–4168 |
| cap_refusal_near_1024 | 3 | 60 | 1024 | 1024 | 1 | 120260 | 741723 / 751093 / 753063 | 206192 / 206192 | 203919 / 203919 | 7168–7308 |

The small, scaled, and near-limit source-backed lanes use 8, 128, and 1,024 direct actions. Draft creation adds a 64-action lane with complete namespace-bearing `inkml:definitions` and `inkml:trace` opaque payloads. Scalar replacement, distinct scalar batches, repeated writes to one scalar, root insertion batches, structural add/remove/clear/move, and exact no-op edits retain source comments, identifiers, namespaces, and opaque bytes; the no-op lanes additionally require source allocation sharing. The repeated-write lane reports only the final scalar value and measured cost; the public API exposes no internal coalescing diagnostic, so the report makes no coalescing claim. Caller-cap lanes attempt a bounded property expansion and must refuse at the configured complete-output limit without changing the source snapshot. Batch operation counts and resulting action counts are retained in each raw receipt so the report distinguishes one edit in a large source from many queued edits.

Every successful edit is checked after timing by applying its source-checked patch and inverse. The report records only absolute observations for this detached API and fixture matrix. It makes no before/after speedup, native-application, asymptotic, or host-placement claim. It makes no package-wide performance claim.

Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.
