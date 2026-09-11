# Theme-family hardening profile

This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane. The timer covers the named shared DrawingML operation with setup and fixture construction outside the timed region. Allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.

| lane | fresh processes | samples | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|
| native_read | 3 | 60 | 137470 / 144430 / 179570 | 100174 / 100174 | 9514 / 9514 | 2716–3004 |
| native_replace | 3 | 60 | 414911 / 428531 / 489572 | 295246 / 295246 | 17956 / 17956 | 2748–3060 |
| native_remove | 3 | 60 | 400201 / 408841 / 412541 | 279248 / 279248 | 16301 / 16301 | 2772–3004 |
| native_add | 3 | 60 | 411662 / 424922 / 426611 | 287290 / 287290 | 17684 / 17684 | 2748–3080 |
| unknown_32 | 3 | 60 | 296851 / 304881 / 314181 | 203762 / 203762 | 36870 / 36870 | 2716–3064 |
| unknown_1000 | 3 | 60 | 1728796 / 1752396 / 1776296 | 859658 / 859658 | 171637 / 171637 | 3004–3080 |
| duplicate | 3 | 60 | 133490 / 140030 / 140511 | 87750 / 87750 | 3232 / 3232 | 3000–3056 |
| limit_replace | 3 | 60 | 132891 / 139911 / 140831 | 88355 / 88355 | 3232 / 3232 | 2748–2996 |
| limit_add | 3 | 60 | 129211 / 135511 / 136520 | 85048 / 85048 | 2397 / 2397 | 3016–3064 |

`native_read`, `native_replace`, `native_remove`, and `native_add` are valid native Theme workflows. `unknown_32` and `unknown_1000` contain an admitted unknown-URI extension with respectively 32 and 1,000 direct family-shaped opaque descendants plus 200 active root namespace declarations; both must read successfully with no typed owner. `duplicate` contains two supported family owners and must reject. `limit_replace` and `limit_add` use a caller output cap of one byte and must reject before producing output.

These are scoped absolute observations. The run does not provide a before/after comparison, an asymptotic proof, or a whole-library performance claim. The unknown-owner lanes are bounded synthetic stress points, and the 200-declaration root remains below the implementation's active namespace ceiling.

Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.
