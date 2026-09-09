# Independent review notes

The source reviewer confirmed that each file-store lifecycle calls sync_data
exactly once under the frozen data-sync policy. The one file-store preflight
precedes all warmups and measured samples. The timer includes source open,
replay creation/write/sync/seal, preparation/publication and handle destruction;
cleanup reopens, hashes and removes the replay only after timing. The replay
store is a benchmark provider, distinct from production OPC atomic save.

Initial helper review found that syscall-summary files were not semantically
parsed. Before freeze/capture, the helper gained validated headers, values,
call/error totals, required syscall families, exactly the expected successful
fdatasync count, and retained parsed counts. Eight tests cover event alignment,
foreign/mutated paths, failed or unfinished calls, order, empty/malformed/wrong
summary data and bad totals. Protocol fields explicitly declare Linux
fdatasync and include fsync only as an unexpected-call guard. Tool hashes,
nonempty artifacts and timeout state are checked.

The final independent review found no blocker in the eight captured profiles.
Each sync-only child has 34 events and every selected sync duration fits within
its paired sample. All selected summary syscall counts match between versions.
The four median sync fractions are 79.52%, 80.73%, 81.61% and 78.99%; the
largest after repeat-1 sync is 8.04 ms. These are traced observations, not a
retrospective proof of the original 0489 outlier's cause.

Strace's help explicitly documents that `-c` defaults to system time and `-w`
selects wall-clock latency. These summaries use `-c` without `-w`; their
reported seconds cannot measure blocked synchronization wall time. Their
3,205 write calls include setup/report output outside the operation timer.
The sync-only `-T` traces provide the separate per-call elapsed-duration data.

The initial formal variance helper was also reviewed before any capture.
Root requested bounded process-group timeout handling, stronger terminal exit
and build-gate checks, and verification that actual receipt chronology matches
the frozen block ordering. The pre-review helper/protocol are retained as
unused development artifacts. Final helper/test receipts bind the revision
used for all formal results.
