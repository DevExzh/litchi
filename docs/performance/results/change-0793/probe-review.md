# Probe review: change 0793

This packet is an evidence-only probe derived from the exact change-0792
probe source.  The only probe behavior change is the optional
`capture-profile` feature.  With that feature enabled, `run_capture` calls the
`#[inline(never)]` `capture_region_0793(&Package)` wrapper.  The wrapper owns
the one profiled public `Package::opened_presentation` operation and applies
`black_box` to its `Result<opened::Snapshot>` before returning it, so the
profile lane retains a distinct call boundary rather than permitting a tail
call directly into the library operation.  Without the feature, the control
lane keeps the original direct call.

The feature matrix is therefore:

* control: no features;
* profile: `capture-profile`;
* control allocation: `allocator-metrics`;
* profile allocation: `capture-profile,allocator-metrics`.

The corpus, five shapes (`tiny`, `medium`, `large`, `vendor`, and
`unicode-vendor`), schema, tool name, counters, and verification oracles are
unchanged from change 0792.  `verify_output` reopens output through the public
`Package::presentation` path; it does not call `opened_presentation`.
The commit and lifecycle modes retain their original direct public calls and
are outside the capture-wrapper comparison.

Source custody:

* source base: change-0792 probe source, whose recorded provenance is the
  exact change-0785 probe source;
* dependency lock: `probe-src/Cargo.lock` copied byte-for-byte from
  change-0792 (`3828dd46cddb7b5f3b0e838051b1c8d70140c1deb0ead723cd7ce991e0c5fb07`);
* modified source hashes: `probe-src/Cargo.toml.template`
  `7fdb50bb6f5103ddca97ec2fee66561526ca08fd757ff248fdcd0543fc46a11f` and
  `probe-src/src/main.rs`
  `1bc0c66d13982f942c3ec889d498e30623a16b1faba78d8ae81c8dadc7425d28`;
* production code and workspace manifests were not edited;
* this probe was not built or executed in this source-preparation step.
