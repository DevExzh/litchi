# 0362ul39b — PPT standalone benchmark consumers

The guarded standalone check on composed source
`9d8b8e5a8c07d325506ea55673c48dcd624d7c63` reaches two remaining PPT
consumers in `tools/perf-baseline`: the history receipt still calls
`CurrentUser::parse` without its required `SourceLimits`, and the snapshot-open
benchmark calls the now-private default constructor. Both consumers migrate
with the RTF/DrawingML integration. The history parser receives the same
default finite limits that the former convenience method selected. The
snapshot benchmark calls `from_bytes_with_options(owned, OpenOptions::new())`,
exactly matching the private helper's delegation. The timer boundaries,
pre-timer archive clone, post-open checks and receipt semantics are unchanged.
Independent source review confirms these equivalences.

The failed capture remains at
`.coord/final-composed-9d8b8e5a-validation/f5-minimal-perf-validation/result.json`
(48,179 bytes, SHA-256
`93102e96b29e1137574bdfa011172dcd983212d888a9d6df50bacee6d864868b`).
Both minimal/non-encryption PPT check and Clippy passed there; standalone
compilation failed with exit 101. Its failure log has SHA-256
`8d61a32bf319ae70082f1a24b6405566e659dd6763af62f48574e67f90d400df`.
Root verifies each log and target trace, clean before/after source identity,
the reviewed `efff87af6c`/`de1f8ae53b`/`989da70f10` lock identities, target
size below the 19 GB live guard, and empty terminal owned-process samples.
The last target size is 2,645,429,840 bytes.

The preceding `200a3e5af3` capture passed all ten PPT/PPTX/XLSX/ops/facade
check and Clippy rows across its two attempts. The first attempt's sparse
XLSX input failure remains preserved; the fully hydrated rerun passes that
row. Receipt identities are:

* `.coord/final-composed-200a3e5a-validation/f5-fixed-validation/result.json`,
  43,535 bytes, SHA-256
  `9dbc2a426fb1352d611a7ef75977e96c290f2839cbf9544c77f8d8864e600846`;
* `.coord/final-composed-200a3e5a-validation/f5-fixed-rehydrated-validation-v2/result.json`,
  74,647 bytes, SHA-256
  `916f1ec91e71430d66f350a2bcaf9428c60d2c0c00d6ce5349eef54d5f67ced3`.

That capture's standalone row was deliberately interrupted before useful
work and is not a pass. Its owned processes were retired before the fresh
9d capture. Root verifies all recorded logs, traces, locks, source and final
ownership observations. A fresh composed standalone check is required after
this repair; no native or performance acceptance is claimed.
