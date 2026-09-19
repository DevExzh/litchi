# Legacy CFB directory lookup control

Build with `RUSTFLAGS="-D warnings" cargo build --release --locked --offline --manifest-path docs/performance/results/change-0690/lookup-probe/Cargo.toml`.
Run `cfb-lookup-probe-0690 --case all --format json --repetitions 250000 --warmups 1000`.
Use `--list` for the twelve case names, or select one with `--case NAME`.

Each process deterministically creates and opens its CFB fixture before timing.
Only repeated public `OleFile::stream_len` calls and checksum/outcome accumulation
are timed. Fixture byte length and SHA-256 are emitted with outcomes. This is
a lookup mechanism control, not an open/read/edit/save workflow or a per-call
tail measurement. The main XLS matrix separately exercises SharedOleFile.

The two builds must use identical probe sources and Cargo.lock. Parent drivers
retain baseline A/A and candidate A/B/B/A, nine processes per leg, CPU 12 affinity,
and all raw process values. Adapt the affinity to the measured machine.
