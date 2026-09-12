# PPTX existing InkAction owner performance report

Status: **API and correctness scaffold implemented — no release profile or timing run performed**.

The proposed profile is pinned to the full owner source commit
`cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd`. It has exactly 23 named recipes
and 42 named lanes in [`corpus-manifest.json`](corpus-manifest.json), with no
Cartesian expansion; matrix preflight requires every recipe to be used and
every lane to resolve to a real recipe. The retained fixture authority is the synthetic helper
`crates/litchi-pptx/tests/pptx_ink_actions.rs`, SHA-256
`bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e`. The
retained bounded OPC generator is `harness/adapter.rs`, and the executable
source/verifier tools plus isolated lockfile are present; no native PowerPoint
fixture or measurement receipt exists in this scaffold.

The later report may describe measured Litchi package and presentation host
behavior for these complete synthetic OPC graphs. It must not turn those
observations into PowerPoint acceptance, rendering, playback, recognition,
interoperability, or speedup claims. It reports only absolute latency,
requested allocation bytes, incremental peak-live bytes, and process maximum
RSS initially; p50/p95/p99 are descriptive summaries of the captured process
samples. The opaque recipe retains both an unknown internal outbound
diagnostic and an unknown external outbound diagnostic without traversal or
fetch.

The fixed later-run budget is 126 fresh process launches, 252 warm-up calls,
2,520 measured calls, and 2,772 operation calls. The run can begin only after
this scaffold is committed and reviewed, then from a clean isolated checkout
with its own authoritative `harness/Cargo.lock`. Every receipt must carry the
full source/toolchain/host/binary/fixture hashes, disjoint phase allocation
equations, source-preservation checks, and the actual typed error/resource for
each refusal lane.
