# Independent review — 0914 DOCX section-layout baseline

The final packet is a descriptive baseline for
`docx_source_backed_existing_section_layout_edit_save`. The runner measures
the sum of six intervals: archive and sink preparation, source-backed open,
section query, edit, commit, and sequential publication. Preparation includes
the archive clone and sink budget reservation. The timer stops when publication
returns; semantic reopen, ZIP/preservation checks, inverse publication,
refusal controls, digest checks, and source immutability checks are outside
the timer. Refusal controls run once during setup before the measured loop.

The source counters are aggregate logical `InstrumentedSource` observations
from the timed path. They are not physical I/O, decompressed bytes, copied
bytes, allocator metrics, RSS, or syscall counts. The query and publication
paths each perform a complete source-artifact fingerprint, and publication
also hashes the emitted output. The counters do not attribute bytes or time to
those individual operations.

The final frozen source is revision
`a4ca85e28cdea151176534b7f8dc912253a1ca1e`. Its external selected-source
archive is 171,120,640 bytes with SHA-256
`7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`.
The independent tree check was run with:

```text
python3 docs/performance/results/change-0914/source-tree-verify.py
```

It reported 9,357 Git entries, 9,357 extracted entries, and 9,357 archive
file entries, with zero missing, extra, extracted content/mode, archive
content/mode, archive-missing, or archive-extra entries. The supplementary
compile fixtures are included in that check. The old custody-drift cohort was
retained under
`runs-rejected-custody-drift/` and was not used; every final child receipt
points to the final protocol and build descriptors.

`capture.py reduce` and `verify` both pass for the final four-child packet.
The retained reduction is 27,697 bytes with SHA-256
`770bc343ce6ab351cfd86bd433657c32be40f5196accf9995c3f1061308ca058`.
Across 800 samples, pooled elapsed p50/p95/p99 are 4,810,184 ns,
4,850,650 ns, and 4,876,280 ns. Every sample reports 202 logical source
reads, 8,409,488 logical source bytes, 138 ordinary-payload reads,
8,397,632 ordinary-payload bytes, and one ordinary-payload materialization.
The fixed source/output identities are archive 2,101,427 bytes with SHA-256
`183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8`
and target `word/document.xml` 18,855 bytes with SHA-256
`67cb9066f112bfee6905a23f3eabdc76b65c86cd0c6613dd60a5a18895ceac81`.
The changed output is 2,101,446 bytes with SHA-256
`5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec`.

The reducer mutation audit used real child reports and updated only the
temporary report descriptor before each check. Wrong output hash, false gate,
wrong corpus archive hash, phase-sum mismatch, section sample-order mismatch,
source-read/cache/materialization counter mismatches, zero sink write size,
and missing max-in-flight vectors were all rejected. The authoritative child
reports, receipts, reduction, and final verification were restored and pass
after the audit.

Raw ZIP preservation remains qualified: untouched local spans and central
records are checked, with the central-directory relative local-header offset
bytes excluded for relocation; member order, names, and archive comment are
checked separately. This does not establish full central-record byte
equality. The baseline makes no candidate comparison or speedup claim.
