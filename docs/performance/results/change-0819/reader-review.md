# 0819 reader review

Static review of `analyze.py` and `validate.py` against the frozen capture
receipts, plan, and `ordinary_save` report producer. No workload or heavy reader
execution was run.

## Blocking schema mismatches

1. `analyze.py:761-772` assigns `published = summary.get("published_sha256")`.
   Here `summary` is `ordinary_save.corpus`, whose `published_sha256` is the
   single reference digest (`SaveEvidence.published_sha256`). The per-sample
   publication vector is `ordinary_save.published_sha256`. Consequently edit
   reports reject because the reference string is not `[]`, and publication
   reports reject because the reference string is not a list. The reader must
   use `ordinary.get("published_sha256")` for the vector while retaining
   `summary.get("published_sha256")` for the reference digest checks.

2. `analyze.py:796` searches for `"never subtracted"`. The frozen producer's
   `PROCESS_PROBE_CONTROL_SCOPE` is
   `fixed_32_empty_adjacent_procfs_snapshot_pairs_acquired_before_warmups_and_never_subtracted`
   (`tools/perf-baseline/src/ordinary_save.rs:99-100`). Every observer and
   qualification report will fail this check. Match the producer's exact
   `never_subtracted` token (or validate the full frozen field).

3. `validate.py:86` requires the raw report mean to equal Python
   `statistics.fmean(samples)` exactly. The Rust harness computes the mean with
   sorted-sample Welford updates (`tools/perf-baseline/src/lib.rs:60612-60620`),
   so equal mathematical means can differ by floating-point ULPs. `analyze.py`
   already uses a `<1e-9` tolerance at line 731, but the final validator is
   stricter. Replay the Rust Welford calculation or apply the same justified
   tolerance in both readers.

## Reader coverage gaps to resolve or explicitly accept

- `validate_report` checks phase names but does not bind `timing_scope`,
  `atomic_publication_steps`, default full durability (`save_durability` absent),
  or the format's sink entry point. This leaves the full-durability boundary and
  PPTX `to_bytes()` counting contract weakly checked even though the frozen
  source documents them.
- Counting reports do not validate `ordinary_save.sample_byte_split`; the
  counting phase should retain a split and non-counting phases should omit it.
- `load_artifacts` verifies selector case/source/output identity but does not
  verify selector format/phase/input fields. `load_lane` and capture admission
  currently enforce those fields, but the offline replay should bind them too.
- `cleanup_value` does not enforce the `removed` row count/shape or its source
  descriptor. `validate.py:123-139` likewise checks directory absence and binary
  identity but not the retained directory rows/source witness. `cleanup.py`
  writes two rows and the build source; the final validator should preserve that
  contract.
- `validate.py:142-157` validates hashes for seal entries present in `files`,
  while exact set equality is enforced later by `seal.py --check-index` or
  `--check-head`; the seal driver therefore remains the authoritative exact-set
  gate.

The three numbered mismatches are blocking. The remaining items are bounded
schema-hardening findings; no captures should be interpreted until the numbered
reader checks are corrected and replayed.

## Follow-up after reader corrections

The three blocking mismatches above are resolved in the settled readers:
`ordinary.get("published_sha256")` is used for the sample vector,
`never_subtracted` matches the producer, and both readers use the Rust Welford
replay.

One remaining durability-hardening issue is present at `analyze.py:798-802`.
The check only requires the substrings `litchi_opc::atomic::replace_with` and
`parent-directory sync`. Both frozen weaker-policy descriptions also contain
those substrings: `replace_with_durability(FileOnly/NoSync)` and `no
parent-directory sync`. Although `save_durability` omission is checked, a
malformed report could pair that omission with a weaker step description and
pass. Bind `atomic_publication_steps` to the exact full-durability string (or
also reject `replace_with_durability`, `no sync_all`, and `no parent-directory
sync`) before treating the full-durability reader gate as complete.
