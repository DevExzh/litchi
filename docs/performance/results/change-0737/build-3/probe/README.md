# 0737 public-format probe

This probe is copied from the sealed predecessor probe and keeps its command-line
existing arguments, semantic oracle, controls, and allocation boundary for the
default `--lifecycle legacy` mode. Its JSON schema is extended with lifecycle
metadata and witness counts; the independently rebuilt 0735 probe is the exact
legacy schema anchor. The local package and
library crate are suffixed `0737`; the two binary names remain
`ole_format_save_probe` and `ole_format_save_probe_alloc`.

The optional `--lifecycle` selector is `legacy` (the default),
`strict-retained`, or `strict-drained`. Strict modes are bounded to the PPT
`format` operation: they accept at most three warmups and fifty measured
samples, and refuse DOC and common-container invocations with a typed scope
error. Every strict warmup and measured operation runs the complete output
oracle. Its full `Sample` is serialized through the same `serde_json::to_value`
receipt path used by both strict arms. Warmup receipts are retained as compact
visible JSON values and the full witness is dropped before the next operation.
Measured receipts have the same visible sample shape in both arms;
`strict-retained` additionally keeps each full witness in a vector until the
last timer, while `strict-drained` drops it immediately. The output records
`lifecycle`, `warmup_receipts`, `retained_witness_count` at the top level, and
the count in every measured sample.

The timing binary has one collection-only wrapper,
`ole_format_save_probe_0737::measured_public_format`. It is marked
`#[inline(never)]` so callgrind can attribute the timed public-format
lifecycle to a stable symbol. `timed_format` calls that wrapper. The wrapper
only delegates to the existing `public_format_edit`; it adds no timer, hook,
allocator behavior, or unsafe code. Oracle construction, expected-output
generation, controls, and allocation measurements continue to call the
existing function directly.

The wrapper is inside the existing timed lifecycle. Source bytes are borrowed
from the input buffer, and the returned output `Vec<u8>` stays alive after the
timed call for validation and JSON reporting, exactly as in the sealed probe.
The collection-only change therefore identifies the public-format call in
callgrind without changing the measured operation or its source/output
lifetime contract.

The PPT selectors are `ppt45543` and `ppt-secondary`. Both perform the same
public edit, removing slide position 1, and run the same CFB, live-persist,
survivor-payload, and public-text oracle. `ppt45543` remains the sealed
11-slide fixture. Qualification tries `ppt-secondary` in this fixed order:
`test-data/poi/test-data/slideshow/41246-1.ppt`, then
`test-data/office-interop/libreoffice-resaved/45543-transition-litchi.ppt`,
then `test-data/ole/ppt/SampleShow.ppt`; unsupported or oracle-failing
fixtures are retained as refusals rather than weakening the selector.
