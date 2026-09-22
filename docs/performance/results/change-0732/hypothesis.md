# 0732 native PPT commit phase attribution

0731 showed that Valgrind masks SHA acceleration. Its instruction fractions
cannot rank native costs. Add feature-gated synchronous content-free events
around existing PPT commit expression boundaries. Keep the ordinary commit
body unchanged and preserve local owner lifetimes in the diagnostic copy.
No validation, digest, cache, default save policy or output behavior changes.

Compare ordinary opaque, ordinary split, profiled empty and profiled clock
routes in one executable with the diagnostic feature enabled. Three cycles
with three rounds rotate and reverse route order, with 50 samples and three
warmups per process: 36 processes, 1,800 measured lifecycles, CPU 12. Preserve
all raw samples and tails. In each round, compare the fixed pairs ordinary
opaque to ordinary split, ordinary split to profiled empty, and profiled empty
to profiled clock. Absolute p50/mean changes above 5% are interpretation flags;
never discard samples or selectively rerun. Record empty/clock differences
before interpreting same-owner phase/whole fractions. Ordinary-to-empty
controls include implementation and compiler effects as well as observation.

All exact 0728 PPT identities, direct semantic and preservation oracles, and
corruption controls must pass. No default-build code generation, allocation,
RSS, cold I/O, concurrency or speedup claim follows from this attribution.
