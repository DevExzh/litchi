# 0957 final build and quality provenance

This v2 supplement preserves the historical pre-capture v1 packet and adds
only completed source-manifest reproduction, allocator integration build, and
quality provenance. It still contains no allocation measurements, comparison
results, adoption decision, capture output, or final performance report.

The frozen source integration is the external `source-v1` tree at commit
`02449f394529766d76794332204c27c4022931fc`, with six approved overlay paths.
The external `source-manifest-v2.json` is 2,267,268 bytes with SHA-256
`f5fd1ee75b734bb7efaa1882be0397defea98c76c835be0bdf1f565c573ef617`.
The standalone packet helper
`tools/reproduce_source_manifest_0957_v1.py` regenerated a fresh external
manifest containing 10,932 files and 273,429,543 source bytes; `cmp` matched
the frozen manifest exactly. Its command and output descriptors are in
`source-build-provenance-v2.json`.

The allocator integration build completed externally with return code 0,
cleared process group, and no workload launch. Its environment records
`RUSTFLAGS`, `CARGO_ENCODED_RUSTFLAGS`, `RUSTC_WRAPPER`, and
`CARGO_BUILD_RUSTC_WRAPPER` as unset. Binaries and targets remain external;
only bounded receipts are copied here.

The initial workspace-wide `cargo fmt --all -- --check` failed on inherited
baseline files. That observation is preserved separately. The scoped follow-up
passed rustfmt for the four approved overlay Rust files, both allocator clippy
checks, allocation-metrics tests, and the allocator binary tests. These are
quality provenance only and do not imply any allocation or performance result.
Capture/reducer/control tooling remains unbound pending root’s final measurement
freeze.
