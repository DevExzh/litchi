# 0732 public-format and PPT phase probe

This probe is copied from the sealed predecessor probe and keeps its command-line
arguments, output schema, semantic oracle, controls, and allocation probe
unchanged. The local package and library crate are suffixed `0732`; the two
binary names remain `ole_format_save_probe` and `ole_format_save_probe_alloc`.

The timing binary has one collection-only wrapper,
`ole_format_save_probe_0732::measured_public_format`. It is marked
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

The `ppt_phase_probe` binary is the native PPT attribution lane. It accepts
`--route ordinary-opaque|ordinary-split|profiled-empty|profiled-clock`,
`--input`, `--warmups`, and `--samples`; the phase packet fixes the case to
`ppt45543` and removes slide 1 through the public `slide_order` owner. The
ordinary opaque route times the original `public_format_edit` lifecycle. The
split routes keep source open, edit construction, slide removal, commit, and
output byte extraction clocks inside the same outer lifecycle, with local
owners dropped before the outer clock stops. The profiled routes use ordinary
open plus the feature-gated `Transaction::commit_profiled` API; the clock
route records only content-free diagnostic events in a fixed `[Option<Event>;
32]` stack. Trace conversion and exact semantic/inventory validation happen
after the timed lifecycle. Every sample reports a nonnegative residual and
commit-window offsets, and a missing, reordered, overflowing, or failed event
trace fails closed.
