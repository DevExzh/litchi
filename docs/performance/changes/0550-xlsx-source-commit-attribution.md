# 0550: identify the leading source-backed XLSX commit cost

`performance_claim: none; current baseline attribution`

Fresh profiles place source worksheet layout scanning at 53.24–54.94% of
exact `MultiSourceEdit::commit` instruction references. XML validation is
30.40–35.11% and provenance merge 4.64–6.03%. These nested shares are not
latency shares or removable fractions. Source and harness remain unchanged.

The eight profiles isolate one measured commit each, excluding 26 lifecycle
dumps and all external staging/publication work. The native commit phase
includes staging, and neither selected case exercises `SourceEdit::commit`.
The reduced-reader helper names lack separate emitted edges; this does not
establish zero work.

The campaign retains 480 native and 480 allocator samples across one-cell and
one-percent updates on medium, dense-sparse, noncompact and vendor-extension
shapes. All source/output/semantic identities match. All 22 native repeat
flags above 5% and six immediate-child instruction drift rows are reviewed
individually; allocation vectors are exact across repeats. Native counts are descriptive and support no registered speedup.

Both builds, all successful capture jobs and four metadata checks pass.
The unchanged source inventory matches 0549; prior tests are continuity
evidence rather than a fresh XLSX suite. One pre-measurement Valgrind debugger
setup failure is preserved with the frozen remediation. Canonical reports,
negative metrics probes, strict replay and cleanup are retained in the
[evidence bundle](../results/change-0550/README.md).

The next task is a private source-bound layout proof that could avoid the
second complete source scan while preserving lexical spans, scanner facts,
error order, resource limits and output validation. A candidate must improve
planning plus commit and publication, rather than merely shift work. The
rejected 0527 row arena is not revived. OLE2/OOXML remain active, ODF deferred,
and iWork excluded; the full performance goal remains open.
