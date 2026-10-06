# 0921 final audit

This is a read-only audit of the frozen comparison packet. No comparative child, benchmark, recapture, or source edit was run.

The manifest is 7,580 bytes with SHA-256 `cc782ab79078a5eed79114f01328bd64d19c7be0ec873bffff4141f91f42d1d8`; its protocol binding is `2e0b1d10a7b45b403aeb053fa3da24b2371d955fb92b5cb68a312ae7b19b4a53`. Baseline and candidate both contain 938 source files and bind to base revision `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`; the candidate patch is 22,350 bytes with SHA-256 `7705c8ec7f12a5a5083b2ff29408451b0ee2afbbf8a25bb827c01b885aece67d`. The six modified source paths and candidate manifest are bound by the source receipt and build receipt.

The four retained binary descriptors, 12 build artifacts, 157 compiler-fingerprint rows per version, and 144-package dependency graph pass identity/parity checks. The baseline and candidate source hash inventories match their receipts.

All 192 native and 96 allocator child receipts are present and passed. The audit checked 288 raw directories, 864 packet artifacts, 864 external artifacts, and 288 validator summaries. Recomputing the reduction from retained reports exactly matches `comparison-reduced.json` (560,529 bytes, SHA-256 `afc65949f8f2fdeb646bf1f9c7166fa8b8ff3adde50e7d411f3d3b68c2f03dae`). A copied packet hierarchy also passed offline replay: 288 canonical reports, SHA-256 `4f867b4a6f7aafe0119403065ef7b3f7db8201c247f1f5e358aa936ba55cef0d`, with both fixture hashes verified; the temporary copy was then removed.

Adoption is false because the `docfloat`/2-edit native cell has a 95% lower latency ratio bound of `0.9520756186100885`, below the `1/1.05` threshold `0.9523809523809523`. Every native default median ratio is above 1 (minimum `1.0095497236589868`). Capacity fields are unchanged, allocator allocated bytes do not increase (maximum fraction `-0.04189690962781045`), and peak live bytes do not increase (maximum fraction `0.0`).

The review table retains all adverse rows: native latency flags are p50=5, p95=5, p99=9, native RSS=1, and allocator allocation=0, peak-live=0, RSS=3.

See [final-audit.json](final-audit.json), [comparison-manifest.json](comparison-manifest.json), and [comparison-reduced.json](comparison-reduced.json).
