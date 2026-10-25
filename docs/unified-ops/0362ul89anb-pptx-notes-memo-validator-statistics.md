# 0362ul89anb — PPTX notes-memo validator statistics and identity closure

This follow-up is based on `0dc79b4bd590b72878c74ea9abb50f03c2804442` in
the dedicated `uo2/pptx-notes-perf-89an` sparse worktree. It is harness-only;
there is no Cargo, build, native, benchmark, or public-API change.

`tools/pptx_notes_memo_abba.py` now recomputes every retained row's elapsed
statistics with the harness algorithm: sorted samples, Welford mean and
sample standard deviation, the pinned Student's-t critical values and
large-degree expansion, nearest-rank p95/p99, integer midpoint p50, and the
nonnegative 95% interval. Reported floating values are checked with the same
relative/absolute `1e-12` tolerance used by the shared ABBA validator, and
the confidence-interval method must be exactly
`two-sided Student's t interval for the mean`.

The synthetic packet now contains those computed values rather than marker
numbers. Focused negative tests reject altered mean, deviation, interval,
method, fixture recipe, and pinned report metadata. The validator audit also
pins the producer memo/fidelity scopes and continues to enforce exact object
key sets, allocator identities, counter revision, scopes, source/binary
pairing, and raw/reopen fidelity evidence.

Validation performed:

```text
python3 -m unittest tools/test_pptx_notes_memo_abba.py  # 12 passed
python3 -m py_compile tools/pptx_notes_memo_abba.py tools/test_pptx_notes_memo_abba.py
rustfmt --edition 2024 --config-path rustfmt.toml --check tools/perf-baseline/src/pptx_notes_memo.rs
```

No performance or qualification result is claimed. The coordinator must use
the exact frozen endpoint binaries and the quiet four-leg ABBA protocol for
any later measurement.
