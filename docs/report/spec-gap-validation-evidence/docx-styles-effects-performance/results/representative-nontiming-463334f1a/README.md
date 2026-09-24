# Representative production operation receipts under contention

This directory retains two direct profile-harness operations against the
committed DOCX public API at frozen checkout
`463334f1a32de687ffb2b3357e55816dcfd3102e`, production source
`d000d977b99e03f8542c7dae74acf767a91b1feb`. It is a bounded allocation/copy
and semantic check under observed shared workload, not the authorized 31-row
latency profile.

The operator session was `44845`. The profile binary was built with the
committed locked offline profile-harness manifest and then ran one sample with
no warmup for each native lane:

```text
--lane replace_main --scale native --warmup 0 --samples 1
--lane replace_glossary --scale native --warmup 0 --samples 1
```

Both actual public operations succeeded. Their receipts report semantic,
opaque, exact-inverse, source-backed, and balanced process-local allocator
checks. The raw JSON receipts retain allocation/copy fields and the phase
breakdown. No `/usr/bin/time` wrapper was used; elapsed fields are retained as
raw receipt data but are excluded from interpretation.

A wrapper validator initially looked for `actual_success` at the receipt
root, while this scaffold correctly nests it under `samples[0]`. It exited 1
after both operations had completed. The receipts were then checked at the
correct nested location without rerunning either operation, and the target was
removed after validation. `status.txt` records this sequence. This is a
wrapper bookkeeping event, not an API refusal or production failure.

The before and after process censuses show an unrelated all-feature
`perf-baseline` Cargo build and rustc children, including DOCX/XLS/XLSB and an
ODS test. The operations therefore ran under shared workload. This evidence
supports no quiet-host, latency, RSS, allocation-volume comparison, scaling,
or speedup claim. The full profile remains separately gated by a fresh quiet
census and root authorization.

The external raw result files are listed and hashed in `retained-files.json`.
The disposable Cargo target was
`/var/tmp/litchi-docx-styles-effects-profile-representative-target-463334f1a-20260912-a`
and was removed after receipt validation.
