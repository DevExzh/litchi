# Profile requirements

- Keep the profile separate from production runtime dependencies.
- Use the native `date.xlsb` Theme family and one deterministic opaque-extension-heavy derivative.
- Run three fresh processes and thirty measured samples per lane after three warm-ups.
- Measure allocator-instrumented operation time, requested allocation bytes, exact live balance, incremental peak live, whole-process RSS, and counted source-backed I/O where applicable.
- Prepare input clones and fixture hashes outside timed/allocation regions; perform no-op source hashing and source-pointer checks outside those regions.
- Capture Rust/toolchain/host data, binary SHA-256, and a Cargo package source manifest before and after build/run; fail if the manifests differ.
- Preserve exact commands and raw process results; remove only owned temporary files/targets.
- State component scope explicitly and make no broad speedup claim from absolute lanes.
