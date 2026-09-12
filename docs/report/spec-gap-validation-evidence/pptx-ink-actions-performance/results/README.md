# Results placeholder

No build or measurement has been run. This directory must remain free of
timing receipts until the scaffold commit/review and production-freeze gate
in [`../PLAN.md`](../PLAN.md) pass.

A later sealed run may add raw per-process receipts and verification output
for the exact 23 recipes and 42 lanes only. Its fixed budget is three fresh
processes per lane, two warm-ups, and twenty measured samples per process.
Each receipt must identify the full source commit, helper/generator hashes,
isolated `harness/Cargo.lock` hash, toolchain/host/binary/fixture hashes,
retained OPC and owner limits, actual typed errors, and the disjoint
setup/operation/validation/drop/post-drop allocation equations. The report may
describe measured Litchi host observations from the synthetic OPC graphs; it
must not state native PowerPoint acceptance or a speedup. Before that run, the
implementation must materialize the isolated lockfile and matrix preflight
must prove that all 23 recipes are used and all 42 lane recipe references
resolve.
