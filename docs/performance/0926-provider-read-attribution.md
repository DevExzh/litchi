# 0926 — separate source fingerprints from provider request amplification

This batch audits the source at `82b0017e42305e6b909dce486ffd913e2da66613`
and reconciles the retained [0925](0925-section-layout-provider-baseline.md)
reports. It runs no performance children, changes no production code, and
does not repair the failed terminal custody of that cohort. The counters remain
descriptive evidence, not qualification or a speedup claim.

## What the source proves

The harness queries `section_layout_snapshot` and later calls
`publish_section_layout_commit_to_stream`. Both paths call
`main_section_layout_snapshot` in `crates/litchi-docx/src/source_backed.rs`.
That helper reads the main part and calls `SourceArtifact::fingerprint`.
For these generic providers, the OPC helper scans the complete source each
time, using a 65,536-byte buffer. The private memo introduced in 0923 exists
only for explicitly owned immutable ingress; wrapping an OwnedSource in a
generic range provider does not enable that memo.

Thus two complete hashes account for 4,202,854 of the observed 6,308,061
returned bytes. The remaining 2,105,207 bytes include catalog and target reads,
preservation validation, and unchanged ZIP record copying. The total is
`3 × 2,101,427 + 3,780`, **not three exact full-source scans**. The retained
operation-end snapshot cannot assign the residual to individual phases or
call sites.

Changed publication enters `write_single_part_overlay_to_stream` and the OPC
preservation planner. `soapberry_zip::preserve::copy_range` copies unchanged
ranges in bounded chunks using `read_exact_at`; the selected main part is
regenerated. Copy-through does not decompress the untouched media. This is a
different operation from hashing, even when both traverse most source bytes.

## Request accounting

The following fingerprint counts are derived from the source loop and provider
caps, not a recorded per-call trace. A normal scan makes 32 full requests and
one 4,275-byte request. After a short read the fingerprint loop advances by
returned bytes and requests up to a fresh 64 KiB, rather than merely requesting
the unfilled suffix of the previous buffer.

| Per operation | Full-return providers | 4 KiB short provider |
| --- | ---: | ---: |
| Observed nonempty calls | 165 | 1,607 |
| Observed requested buffer bytes | 6,308,061 | 84,096,471 |
| Observed returned bytes | 6,308,061 | 6,308,061 |
| Two hashes: modeled calls | 66 | 1,028 |
| Two hashes: modeled requested bytes | 4,202,854 | 66,262,624 |
| Two hashes: modeled returned bytes | 4,202,854 | 4,202,854 |
| Residual calls | 99 | 579 |
| Residual requested bytes | 2,105,207 | 17,833,847 |
| Residual returned bytes | 2,105,207 | 2,105,207 |

The two short-read hashes account for 1,026 of the 1,506 short reads; the
residual is 480. Requested buffer volume is not transferred bytes, copying,
resident memory, or CPU cost. The delegated short-provider boundary receives
only capped requests, so its requested and returned bytes are equal even
though its call count remains 1,607.

To reproduce the modeled counts for one hash, start at offset zero; while
offset is below 2,101,427, request `min(65,536, 2,101,427 - offset)`, return
the smaller of that request and the provider cap, then advance by the return.
Sum calls, requests and returns; double each total for the two hashes. This
models the successful deterministic provider, not arbitrary erroring sources.

The raw histograms contain fixed size buckets, not exact request sizes or
offsets. The independent offline checker verifies every report against the
retained failure inventory, checks source/output/binary identities, and checks
all 300 samples per provider route for identical counters and bucket counts.
It retains the file route's unavailable underlying counter as unavailable.

## Compatible next work

Do not generalize the owned digest memo to FileSource or arbitrary ReadAt.
FileSource metadata is not a strong immutable byte owner; source-version
checks alone cannot justify suppressing the publication fingerprint. Preserve
ADR 0003 source-checked patches, ADR 0005 explicit bounded I/O, ADR 0006
fail-closed publication, and ADR 0032's ownership boundary.

Existing opt-in OPC read-ahead is not a cache for these two hashes:
`source_artifact` retains the exact source snapshot, while changed publication
explicitly disables archive read-ahead. Repeating the older 0492 read-window
experiment would therefore not test the dominant source-byte mechanism here.

The next useful diagnostic should distinguish fingerprint reads from catalog,
preservation-index and copy-through reads at their actual call boundaries,
without putting those observers into native timing claims. Before changing
request windows, establish whether provider latency, hashing CPU, XML snapshot
work, or preservation overhead is the limiting cost on the intended route.
Any subsequent candidate needs a new independently specified experiment with
working terminal receipts; 0925 must not be rerun or promoted retroactively.

## Reproduction and validation

Run the committed offline checker with a fresh output path:

```sh
python3 -B docs/performance/results/change-0926/provider-reconciliation.py \
  --output /tmp/0926-provider-reconciliation.json
```

The retained result is
`results/change-0926/provider-reconciliation-final-v4.json`. It binds all 42
raw reports by SHA-256 and size. No binary, large evidence file, new capture,
or production change belongs to this batch. Source-derived attribution is
kept separate from the checker's directly observed report accounting.

Root and independent reviewer replays match the retained result apart from
creation time. `results/change-0926/shared-owner-review-v2.json` records the
independent replay and source-model review. The packet manifest adds only
runner custody/script descriptors to the authorized manifest; that expected
enrichment does not resolve the missing terminal receipt. Runtime test suites
were not rerun because no runtime code changed. Reproducible scratch JSON was
removed after review; the original 0925 evidence remains intact.
