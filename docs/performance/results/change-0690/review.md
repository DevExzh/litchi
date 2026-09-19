# 0690 review and disposition

Root owns the serial Cargo/performance lane. Production, tests and independent
review were delegated under the active goal; no iWork files changed.

- `program_priority`: fresh profile confirms a small, low-risk lookup-only
  opportunity; also identifies repeated PPTX MCE passes as the larger next
  end-to-end investigation. Its design is retained separately, not implemented.
- `next_path_review`: no semantic blocker. Length-first ordering, comparison
  tail length, Unicode fallback, cache presence and bounded traversal remain
  intact. Per-node ASCII conversion requires the deep-tree controls.
- `cfb_name_coder`: production-only helper and both CFB lookup owners. Root
  added documentation clarifying that None requires general validation.
- `lookup_test_design`: five shared-reader tests and a frozen original oracle
  cover tree traversal, Unicode equivalence, length ordering, nested/root
  semantics and missing/mismatched cached keys.
- `lookup_key_tests`: two new tests and extended exhaustive matrices compare
  ASCII errors, fallback and ordering against the owned key/reference. Root
  added altered cached comparison lengths to exercise the final tie-break.
- `lookup_probe`: deterministic public legacy-reader control; root corrected
  development CLI compilation/case-count issues before baseline measurements.
  Frozen probe sources, lockfile, fixtures and raw outputs are bound by hashes.
  Root archived the initial warmup-warning run, fixed the explicit discard,
  then rebuilt both supplemental binaries with warnings denied and recaptured
  the full lookup matrix. Both versions are independently audited.
- `cfb_name_profiler`: final scoped retention supported by consistent
  owned repeated-query/counter gains and 31–41% short-name lookup gains.
  Invalid-name costs, missing-query costs, code growth and all tail flags stay
  explicit. Follow-up confirms the tiny missing-query cost and preserves
  the Simple file warm-mean p99 flag; two original 45365 anomalies do not recur.
- `evidence_review`: source semantics, counts, quoted numbers, profiles,
  allocations/I/O, corpus and supplementary results verified. Root clarified
  the 2 MiB long-loop versus 1 MiB native missing case and named the indexed
  54016 workflow explicitly. The warning-denied supplemental recapture and
  original archive also pass independent review. No source or evidence
  blocker remains.

All Rust gates pass (4,407 passed, zero failed, 27 existing ignored). Primary
and follow-up audits preserve outcomes, including all 96 allocation groups,
12 counted routes and 24,000 supplementary owners. Existing fuzz tooling is
unavailable; no campaign is claimed. No universal, maximum-name-length,
physical-cold, remote, cross-platform or concurrency performance claim.
