# 0694 independent evidence review

The independent read-only audit passes on the frozen pre-cleanup packet. It
checks the 602 Rust source bindings, all 33 constraint hashes, the five build
input hashes and the separate main, allocator, refusal and oracle probe maps.
The baseline oracle build temporarily restored the exact baseline MCE files;
the successful receipt and the superseded copied-lock failure both prove exact
source restoration. The retained workspace lock copy hashes byte-identically
with the root `Cargo.lock`.

- The native matrix contains 13 workflows, two A/A legs plus four A/B/B/A legs,
  and 100 samples per leg: 7,800 samples. Raw TSV hashes, 100/5 sample shape, source archive
  hashes, workflow metadata, target coordinates, revision/semantic digests,
  no-op invariants and reopened markers all pass.
- The counting allocator retains both phases for all 13 workflows, six timed
  phases and three samples per phase. The derived 156-row summary and 78-row
  comparison are deterministic. For `one-real`, calls fall from 190,338 to
  143,848 and requested bytes from 13,815,905 to 12,430,874; reallocations,
  peak-above-start and net live change remain unchanged.
- The refusal matrix has 10 cases, six legs and 100 samples: 6,000 samples.
  Every expected and observed error record agrees, including late graph/MCE
  paths and the disclosed generated-valid p99 trigger.
- The exact-output oracle covers 192 cases, five profiles and both binaries:
  1,920 invocations with zero mismatches. The synthetic controls contain 36
  rows and 24 comparisons; the real DOCX/XLSX/PPTX controls contain 54 rows
  and 36 comparisons, all with 300 raw samples. The audit recomputes every
  p50, mean, p95, p99 and paired delta from the retained TSVs and requires
  identical identity records across all six legs. The PPTX one-name control
  contains the reported mean/p99 outlier and remains visible in the raw
  evidence.
- Profiles, allocation diagnostics, binary-size receipts and native counter
  formulas are bound. The reported RSS change of 5,696 to 5,700 KiB, page-fault
  increase and 920-byte native text reduction remain scoped diagnostics.
- All seven integration gates and all six repository evidence gates exit zero;
  the quality receipts total 5,385 passed, zero failed and 35 ignored test
  results across the repeated default/all-features runs.
  Deterministic summaries reproduce byte-for-byte, and Python packet scripts
  parse successfully.

No material numerical, source-binding, semantic or scope blocker remains. The
retained results support the narrow private allocation change under the stated
warm shared-host workflow scope. They do not support universal Office,
cold-I/O, remote, concurrent, RSS-bound or cross-platform claims. iWork is
outside this review.
