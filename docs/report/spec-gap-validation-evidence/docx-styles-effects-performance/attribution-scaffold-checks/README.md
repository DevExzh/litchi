# Attribution scaffold checks

The executable allocator and setup-lifetime tests passed serially (2 tests),
strict all-target Clippy passed, and the Python verifier/source suite passed
31 tests. `verification.json` records the exact commands, source hashes and
raw Rust gate logs. The final documentation review approved the attribution
boundaries and serial test instructions.

These are timing-free scaffold checks on the shared worktree, not a clean
profiling capture. They do not establish performance results or authorize a
timing run. Historical v1 results remain unchanged; new v2 capture still needs
the separately pinned, clean-source gate in the performance plan.
