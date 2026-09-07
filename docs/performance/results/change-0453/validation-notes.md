# Validation and evidence scope

The first harness build succeeded, but its allocator executable only dispatched
`retention`. Two allocator preflights failed before execution with an unrecognized
`provider-lifecycle` command; the first also contained a mistyped revision
argument, which is not used as an identity claim. Root added the missing explicit
command dispatch and retained the draft binaries/build/source receipts.

The next three-sample preflight executed successfully. The old portable lifecycle
oracle intentionally accepts only one-sample controls or thirty-sample reports,
so its check rejected that report's sample count. The raw exploratory report is
retained; a separate 30-sample/3-warmup run and unchanged old oracle count contract
provide the accepted pre-optimization baseline. No failed result is relabeled.

Normal latency runs and allocator-instrumented reports are separate. Allocation
regions cover source open, destination open, planning and publication separately.
The existing serialized callback observer includes process-global allocations
between its boundaries; it excludes allocator-internal realloc overlap and is
not physical RSS. Corpus generation, gate checks, output hashing, diagnostics and
handle-drop work are outside each region and the API timing sum. Absolute process
live bytes include preexisting owners; region peaks include region-entry bytes.

The decoded staging reservation remains conservative and covers a late fallback
copy. This batch targets actual allocations/copies and does not claim reduced
memory admission. Native application, cold I/O and scaling coverage remain open.

Focused production validation initially passed 604 tests. The first intended
physical fallback discriminator edit only stored the source fixture; review
caught that its compression-method assertions had not actually been inserted.
After adding them, the test correctly failed on a wrong assumed output media
name (`image1.png`). The corrected test discovers the fixture's single copied
media/chart target and asserts control Store versus fallback Deflate explicitly.
All intermediate receipts remain visible.

The first broad standalone harness strict check found 29 preexisting style
errors in six unchanged files. They were corrected mechanically without
suppression; both compared binaries use the same final harness sources.
The earlier allocator preflight is exploratory before optimization; the frozen
matched matrix remains the performance decision evidence.

The physical discriminator also initially counted an unrelated destination media
member; final target discovery subtracts the original destination inventory.
The passing r4 discriminator is retained. The final test additionally uses an
independent finite destination budget and checks both budgets release completely.
Full strict harness validation required 50 preexisting mechanical lint repairs;
each failed frontier and the final all-target/all-feature passing check remain.
The final shared enum's incremental inline storage is explicitly included in
checked destination staging admission (full decoded fallback allowance remains).

Portable verification distinguishes historical checks from the required final
source epoch. Earlier passing focused/strict receipts are retained as historical
evidence; only the required final checks and matrix/fuzz commands certify the
final candidate. The full final PPTX suite passes 850 tests and OPC passes 471.

All required final gates pass: 850 PPTX, 471 OPC and 381 harness tests, strict
checks, warning-denied docs, workspace feature check, format and boundaries. The
unchanged OPC ASan/sancov target completes 1,000 runs. All 36 primary/pilot captures
and eight separate fixed plain-tail confirmation captures pass their oracles.
