# 0814 protocol review

The unchanged public probe and deterministic fixtures retain the 0806 schema,
36 tests, and exact capture owner. The production census matches sealed 0813
after, so six production quality gates are reused by full file identity. The
probe receives fresh formatting, all-feature tests, and Clippy checks.

Three serial release binaries use the same target/profile: ordinary control,
capture-profile wrapper, and wrapper with force-frame-pointers. CPU 12 runs
six counterbalanced blocks of tiny/medium/large capture, thirty samples and
three warmups per process. Two fp perf children then use cycles:u at 499 Hz,
100 samples, and no warmup. The total is 56 reports and 1,820 samples.

Decode uses each exact binary while it exists; raw perf and decoded frames
are retained with deterministic gzip and both compressed/uncompressed hashes.
Root records bounded scanner/inspector disassembly in all three binaries after
decode. Offline readers independently reproduce numerical perturbation and
sampled leaf/stack counts after captures terminate. They preserve unresolved
frames, lost events, tail/RSS variability, and overlapping inclusive counts.

The six-block median bootstrap uses 10,000 draws, seed 814814, and sorted
endpoints 250/9749. No historical timing pool, native phase fraction, causal
cycle savings, allocation comparison, cross-format extrapolation, or production
adoption is authorized by this diagnostic. Current attribution ranks further
investigation; a future source change needs its own frozen workflow trial.
