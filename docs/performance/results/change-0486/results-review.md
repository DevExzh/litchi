# 0486: remaining DOCX replay metadata callers

This is a diagnostic and documentation batch using the retained 0485 normal
executable. It makes no production change and no new speedup claim. The full
non-iWork goal remains open.

## Capture and custody

Four `perf record -F 499 --call-graph dwarf,8192` children cover the source-heavy
and authored-heavy workloads with owned and file input. Each child uses three
samples and one warmup, CPU 2, and the canonical 0484 axis arguments and output
oracles. The executable SHA-256 is
`fe2d708e10a9070d1223ac854901c377a5f080565653757767093640584edbec`.
The source manifest remains
`c4594536bc2b80c6693be6bf5617ed7e699dd117945bc13183338fb0fd437abe`
before and after capture. The retained 0485 build receipt identifies the build;
the capture did not rebuild it.

`profile_callers.py` retains commands, process results, report validation,
input identities, artifact hashes, and compressed raw perf data and script
exports. Each archive records the original bytes/hash and compressed bytes/hash.
`summarize_callers.py` rechecks those archives, the executable/build/validator
identities, and the canonical report oracles before calculating its summary.
The initial environment command failed with an argument error; `environment-host1`
is retained, and `environment-host2` is the successful replacement.

## Interpretation

The authored-heavy file profile places about 27.38% of whole-child sampled
period weight in stacks containing `statx`. The other three profiles contain no
sampled `statx` frame. Absence from a sampled stack does not mean zero calls:
0485's separate source-heavy file strace diagnostic counted 25,219 calls.
These captures include setup, warmup, timed operations, output oracles, and
launcher samples. Inclusive symbols overlap, and sample periods are weights,
not syscall counts. There is no paired CPU comparison or confidence interval.
Unknown kernel frames and addr2line warnings remain in the retained exports.

The authored-heavy file stacks include two concrete paths:

```text
statx / metadata / version / ensure_current
  / check_splice_external_state / check_external_state
  / write_consumed_range / flush_pending_consumed / consume / audit_splice

statx / metadata / version / ensure_current
  / check_splice_external_state / check_external_state
  / fill_fragment_buffer / fill_buf / audit_splice
```

Source review explains a remaining opportunity: DOCX `EncodingReader::read`
returns after supplying pending event output, so short replay reads can exhaust
the current filled slice while most of the OPC adapter allocation is unused.
0485 then flushes that consumed slice. Capacity alone therefore does not bound
sink metadata checks across multiple replay callbacks.

## Pending candidate and review constraint

A candidate is to retain already-parser-consumed replay bytes in the existing
OPC buffer and append the next read into its unused tail. Keep every provider
callback separate and preserve its pre/post freshness guards. Keep fixed
payloads and source prefix/suffix behavior unchanged. Flush before capacity
reuse, replay completion, or a phase transition. Do not aggregate DOCX cursor
callbacks into one `EncodingReader::read`, which could change error precedence.

Independent read-only review identified an unresolved partial-output issue:
if mutation or cancellation occurs during a subsequent replay callback, earlier
consumed bytes retained in the window would remain unpublished. Current
per-read flushing may already have accepted those bytes. Before implementation,
review the public partial-output contract and explicitly test that boundary.
This batch does not claim the proposed change preserves that behavior.

Required regression cases include one-byte replay reads; payload lengths below,
at, and above capacity; authenticated EOF; ordinary provider failure after a
retained prefix; Work/cancellation precedence; source mutation between reads;
short sink writes; exact replay/candidate digests; and source/suffix transitions.
Ordinary provider errors must not be stored early enough to suppress the final
pending flush. Source changes must still take precedence over provider I/O.
Any implementation needs fresh matched measurements including the small-workload
and adverse-tail controls from 0485.

## Program documentation

The top-level CRUD coverage, phase report, and goal audit now describe accepted
0485 evidence and its adverse observations. The representative index remains
unchanged: 15 categories, 34 rows, 33 selector-backed mappings. Cold-cache,
concurrency, provider intersections, atomic-save, native producer coverage,
broader semantic CRUD, and scaling remain open. The coverage-index and report
classification validators are retained as source-stable gates in this bundle.
