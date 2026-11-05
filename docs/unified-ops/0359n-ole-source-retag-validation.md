# 0359n — W5 source retag validation

- Date: 2026-10-22.
- Source tip: `c7bdb60fa3` (`75bae6d478` retag repair plus formatting).
- Status: focused correctness and paired allocator validation complete; native
  latency qualification remains pending, and the generic XLS allocation
  residual remains unresolved.

The retag repair updates a freshly opened cell snapshot's source-version field
in place when its `Inner` is uniquely owned.  A shared snapshot still takes the
existing copy-on-write path.  The focused regression covers pointer reuse,
shared-sibling isolation, retained object limits, CFB limits, and exact bytes.

Validation used one capped offline Cargo process for each command:

```text
cargo fmt --all --check
cargo test --offline -p litchi-xls retag_source_version_reuses_unique_inner_and_copies_shared_inner --all-features --lib
cargo test --offline -p litchi-ole-common -p litchi-xls --all-targets --all-features
cargo clippy --offline -p litchi-ole-common -p litchi-xls --all-targets --all-features -- -D warnings
```

The full test run passed 1,316 XLS unit tests and all selected integration and
example suites with zero failures.  The affected target was
`(deleted-coord-artifact)`.

## Paired allocator receipt

The receipt is outside the checkout at
`(deleted-coord-artifact)`.
It is 256 KiB and contains only reduced TSV output and command metadata.  The
runner is the preserved 0359i runner, copied to
`.../w5-ole-source-arc-retag-20261022/run-counting-allocator-w5-retag.sh`;
its SHA-256 is
`70d07e3b07bce4237536f1239be66aed23621a19b999d79b2230abd9f3c50b93`.
The exact probe source SHA-256 is
`ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a`, the
matched lock SHA-256 is
`bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a`, and
the current candidate source is `c7bdb60fa3`.

The candidate allocator binary was built with one capped Cargo job from the
external matched manifest and has SHA-256
`fbe21f892719d499f85a6310ccd7c57804ee90e474d1526eb3887da7fcb1f35b`.
The preserved baseline binary is SHA-256
`23a7088c3f9a6563adc3ae1de644a0d2b05fe91604824afd7517aa426c981fc0`.
The runner used `--warmup 0 --samples 1 --batch 1` for `read`, `edit`,
`xls_cells`, `xls_cells_stage`, `doc_history`, and `xls_history`; native timing
was not run.

All paired routes emitted the same `output_bytes` values.  The resulting
allocation counters were:

| route | baseline calls/bytes | candidate calls/bytes | retained/live or peak observation |
| --- | ---: | ---: | --- |
| read | 45 / 1,192 | 45 / 1,192 | live 280,629; peak 280,848 on both |
| edit | 33 / 3,523 | 33 / 3,523 | live 280,629; peak 281,935 on both |
| XLS cells | 1,589 / 2,361,904 | 1,585 / 2,137,013 | candidate peak +110 bytes; live bytes equal |
| XLS cells stage | 1,588 / 2,249,264 | 1,584 / 2,024,373 | candidate peak +110 bytes; stage live +40 bytes |
| DOC history | 1,609,900 / 1,830,103,867 | 1,572,032 / 1,770,486,356 | retained live and peak bytes equal |
| XLS history | 387,769 / 441,765,854 | 386,997 / 398,342,748 | retained live +24 and peak +56 bytes; +129 live calls |

The retag function is not used by these generic edit/history routes; it is
specific to source-backed numeric publication.  Therefore these results
correctly leave the generic residual unchanged and do not attribute it to the
retag path.  The remaining cells/history difference follows the common
failure-atomic COW candidate and rendered-publication path described in
0359l; a source repair is still required before accepting the residual.
