# 0820 execution notes

The preceding goal turn was progress: commit `096b810f23` established the
admitted real-file ordinary-save baseline, passed six quality gates and exact
committed sealing, and removed owned build/scratch trees. Root rechecked the
current history and working tree before this batch. The three unrelated files
remain excluded. The entire 0819 seal payload and all 35 normative document
hashes revalidated before any current index-document update.

The experiment measures 24 configurations: three real formats, lifecycle and
atomic publication phases, and default/full/file-only/no-sync policy labels.
It preserves the default durability contract and does not pool old timings.
The six native blocks use explicit deterministic group and policy permutations;
ratios compare policy/default in matched blocks. These are configuration
comparisons, not source optimizations or additive phase-cost decomposition.

Before freezing, root found and corrected the admission driver's use of an
indexed default durability field: the producer omits that field for default
save. Both capture and admission now require omission for default and the exact
level for explicit policies. Exact atomic-step strings were checked against the
producer. The generic phase timing description is retained but is not used to
claim skipped synchronization occurred. Each sample publishes to an absent
destination; post-sample removal is outside the timer.

Root owns all Cargo, exporter, workload and profiler processes and waits for
terminal handles. Static preparation and review are delegated. All frozen
packet/driver identities are checked around children. Unfrozen result readers
may be prepared while builds run, but heavy offline analysis waits until every
capture terminates.

## Quality failure and scope change

The original quality process terminated with exit 1 after its Cargo test child
returned 101. The completed library suite had 555 passing tests and one ignored;
the whole partial invocation reached 568 passing tests and two failures across
six summaries. The primary failure was the process-global live-byte increase
assertion in the shared allocator wrapper; the second failure was its poisoned
TEST_LOCK consequence. No release build, exporter, qualification, native, or
observer process was launched.

The batch scope changed to the necessary test-quality repair. Frozen original
plan, origin, drivers, source witness, and failure logs remain untouched. A
separate repair origin allows exactly the cfg(test) section of the allocator
wrapper to change. Root formatted that one file before freezing the repair
source. The focused five-test binary passed; six fresh quality gates use the
same two-test-thread setting and existing owned build target. The revised
assertion checks signed global allocation/deallocation conservation and the
process high-water bound instead of an invalid per-test net-live increase.

Success-path measurement readers were moved into planned-readers as unexecuted
preparation. The root packet validator now verifies the retained failed
experiment and the repaired quality gates, with no invented timing result.

## Completed repair quality

Root retained the repair runner handle through terminal exit 0. The focused
allocator binary passed five tests. The fresh six-gate sequence passed, with
28 full-test summaries totaling 641 passed, zero failed, and one ignored.
No release build, export, qualification, native, or observer command ran.
The independent reader and final results review follow these terminal gates;
original frozen failure evidence remains unchanged.

## Offline validation

The independent review initially found `original frozen-input keys changed`:
the validator expected an extra top-level `provenance` key, while the frozen
writer includes provenance through the packet hash set. The reader was aligned
to the original nine-key witness without modifying frozen evidence. It also
recomputes full/focused test aggregates and enforces serial receipt ordering.
The analyst and root then independently ran the corrected validator to exit 0
with six gates accepted. No workload rerun occurred.

## Cleanup and final replay

Root cleanup removed 5,204 owned target files totaling 3,685,472,529 logical
bytes. Scratch was never created. Final validation exits zero with cleanup
checked; source, frozen input, and unrelated-file guards still pass.
