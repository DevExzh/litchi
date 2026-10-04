# 0905 retained replay method

`replay.py` is a post-cleanup, offline verifier. It is run only after the
root coordinator has retained the frozen main plan, all raw main captures,
the passing `analysis.json`, the frozen micro plan and its raw captures and
analysis. It launches no workload, rebuilds no binary and never opens
`source.root_worktree` or `source.source_root`; paths recorded there are
compared as receipt identifiers only.

Invoke it with a new destination:

```sh
python3 -B docs/performance/results/change-0905/replay.py \
  --output /tmp/litchi-0905-replay.json
```

The destination is create-only. A failed replay leaves no replacement result;
the retained captures and prior failure artifacts remain available for review.

The source replay reads the baseline and candidate manifest descriptors from
`plan.json`, checks their exact file count and one-file delta, and verifies the
retained prior source archive. Every regular archive member must belong to the
manifest and its bytes/hash must match the archived content. The archive may
contain an older version of an overlay file or omit a file added after the
archive. Those names are calculated from the archive/manifest byte difference
and must be supplied by retained overlays.

By default the overlays are `baseline-overlay/` and `candidate-overlay/` in
this packet. The candidate replay also includes the baseline overlay because
the unchanged 0904 harness files are shared between the two manifests. The
plan may provide `baseline_overlay`, `baseline_overlay_root`,
`candidate_overlay` or `candidate_overlay_root` paths instead. A standalone
`candidate-audit.rs` at the packet root is accepted as the one-file candidate
fallback. Overlay paths are compared against manifest hashes; no source-root
directory is needed. The final replay result explicitly records
`source_root_claim: deleted isolated source roots are not reopened`.

For each main native and allocator role, replay verifies the retained binary
descriptor and build receipt, exit status, stable input census and the selected
audit-file hash for the corresponding baseline or candidate manifest. It does
not try to stat any path named inside `inputs_before` or `inputs_after`; this
is what keeps cleanup of `/tmp/litchi-0905-src` compatible with replay. Main
capture rows are then independently re-reviewed through the retained analyzer
against the raw child folders. Replay compares every qualification/native/
allocator row and native reduction with the retained passing `analysis.json`,
while deliberately excluding the earlier live-root source summary.

The micro section is replayed independently of `micro.py`. It checks the
frozen 36-job order, six reverse-order blocks, all 12 start/TSV/stderr/receipt
sets, helper/plan/binary descriptors, exact input/outcome identity, 30 samples
per job and 1,000 iterations per sample. It recomputes p50/p95/p99 rows and
the six-ratio, 10,000-draw bootstrap with seed 905019, then compares the
result to `micro-analysis.json`. The machine-readable `adoption_allowed`
field and material-adverse case list are retained. Semantic parity is also
checked when the main plan binds its passing comparison descriptor.

Replay proves custody and reproducibility of retained evidence. It does not
turn the micro guard into an Office workflow claim, does not reopen deleted
source trees, and does not authorize production adoption by itself.
