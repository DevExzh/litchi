# 0362ugja — PPT lazy cold first-commit route
The frozen PPT lazy probe's normal timed route is deliberately warmed.  Each
timed child opens a fresh package, then `timed()` calls `run_batch()` for three
warmups before starting its measured sample.  `primary_noop_commit` therefore
does exercise a never-hashed package, but its first hash/block-root work is in
the warmup interval and is not a first-commit timing observation.
The supplemental consumer route is named `primary_first_commit_cold`.  It
maps to the reviewed probe route `primary_noop_commit` with exactly one
iteration, one sample, and `--warmups 0`, and is available for both the normal
and allocator binaries.  The adapter is
`tools/ppt_lazy_cold_first_commit.py`; it verifies the fixed accepted binary
and default-fixture identities and proves the exact no-op postcondition when
explicitly executed.  The source-only contract tests are
`tools/test_ppt_lazy_cold_first_commit.py`.
This route is supplemental and is not inserted into the frozen 28-route ABBA
manifest or its acceptance claim.  Its future command is emitted with CPU 29,
the coordinator-reserved quiet CPU; this packet performs no native execution,
calibration, allocation measurement, timing, or performance acceptance.
The existing 0357b first-commit residual (1.105×, no budget) remains the only
accepted first-commit performance record until a separately authorized quiet
run produces this route's evidence.
