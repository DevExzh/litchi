# 0731 public-format probe

This probe is copied from the sealed predecessor probe and keeps its command-line
arguments, output schema, semantic oracle, controls, and allocation probe
unchanged. The local package and library crate are suffixed `0731`; the two
binary names remain `ole_format_save_probe` and `ole_format_save_probe_alloc`.

The timing binary has one collection-only wrapper,
`ole_format_save_probe_0731::measured_public_format`. It is marked
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
