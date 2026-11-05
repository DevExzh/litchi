# 0362ui — corpus fidelity and CRUD golden refresh

This review refreshes the two reports that failed only because their checked-in
goldens were behind the frozen `f6fe430a4d19d94aa360c0e9ed03ebd757265225`
source. The report run was isolated to this leaf worktree
(`uo2/final-corpus-fidelity`); the gate checkout was never blessed.

## Report run

The copied lockfile is SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The command was run with `LITCHI_BLESS=1`, one Cargo job, one test thread,
offline locked dependencies, the capped 20 GiB environment, and the external
target `/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target`.
The target closed at 3,483,753,074 bytes. `LITCHI_BLESS` was not set in any
gate checkout.

```text
cargo test --offline --locked -p litchi-ops --all-features \
  --test corpus_fidelity --test crud_matrix -- --nocapture
```

Both tests passed: corpus fidelity opened 1,020 claimed-format files in
121.7 seconds, and CRUD reported zero unreached claimed cells. The external
stdout/stderr receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-receipt/corpus-crud-blessed.log`
(851 bytes, SHA-256
`0478ff68527b96bc51da8c7b7de5653ab7495518270d7e70f1a19240741d14f5`). The
measured run took 2:52.30 and reached 184,132 KiB RSS. The corpus's
`NORMALIZING` list remains empty.

The same command was then rerun without `LITCHI_BLESS` against the committed
goldens. That diagnostic pass was not a bounded validation: its Cargo output
printed relative `target/debug/deps/...` paths, and the later cleanup found
the old leaf's local `target/` at 14,428 files and over 20.3 billion apparent
bytes. Its tracked worktree was unchanged, but its old receipt must not be
used as bounded-build evidence. The corrected explicit-path rerun is recorded
in `0362uia`.

The old diagnostic receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-receipt/corpus-crud-unblessed.log`
(15,377 bytes, SHA-256
`4ba489316f5c09e032fb90ec19ab99cf2cb4739614dd0895b109d17209fd2c23`).

## Row-by-row review

The fidelity TSV has seven changed rows, and every change was reviewed against
the old and new tuple fields:

* `draw-image-link.fodp` and `draw-object-link.fodp` remain `exact`; their
  `/slide[1] name` revert changed from a typed `unsupported_capability`
  refusal to `byte-exact`. This is a real capability improvement: the flat
  ODP projection admission now uses its bounded projection profile (the
  `384205d218` fix), and the updated `odp_unified_write` consumer proves
  rename to `x`, a changed output, rename back to the original value, and
  exact source-byte restoration.
* `field_hide_section.fodt`, `protectedLinkCopy.fodt` and
  `simple-ruby.odt` keep `exact`/`refused`/`one-way`, the same
  `unsupported_capability` code, and the same revert result. Their underlying
  refusal reasons are unchanged: the old cleaned details already named
  `text:variable-set`, `text:a` and `text:ruby`. The refreshed details only
  add the fixed `ODF edit refused: ` display prefix; that 18-character prefix
  pushes the old element-specific tail past the 140-character `clean()` cap,
  so the report now ends in `replace...`. No element reason was newly
  introduced or removed; this is the 0358 refusal-display change
  (`77ab7cab8a`/`e4a0614f5d`).
* The two `tdf169979-litchi.odp` native-resave rows keep
  `exact`/`one-way`, `unsupported_capability`, and the same value/byte
  outcomes. Only the refusal detail now carries the content-free
  `ODF edit refused: odp.scene refused` wrapper from the 0358 refusal
  boundary.

The Markdown summary changes only the corresponding ODP revert counts (+2
`byte-exact`, -2 `none`) and the two native-resave refusal details. No refusal
code, unexpected-fidelity disposition, or capability was lost.

The CRUD TSV has five changed rows, all in the XLSX `run` element:

* `get`, `find` and `view` report two instances instead of one. The checklist
  deliberately adds a run to an empty created cell and then adds the same node
  a second time to form the sibling pair used by the format-level move/swap
  probes. The run implementation inserts into the cell's rich-text sequence;
  it does not duplicate a pre-existing visible value. The existing rich-text
  test also proves that adding to a plain cell retains its original text as
  run 1 and appends the new text as run 2.
* `move` and `swap` are now reached and return the typed
  `unsupported_type` refusal for XLSX runs, replacing the old `unreached`
  observations. The pair is therefore exercised and refused intentionally;
  no claimed cell is lost.

The base `f6fe` consumer test had one stale refusal assertion for
`draw-image-link.fodp`; the test-only patch in this leaf migrates it to the
capability proof and extends the same proof to `draw-object-link.fodp` in one
two-fixture loop. Both fixtures now check rename to `x`, changed output,
rename back, exact source-byte restoration, and source-byte save. The focused
receipt for that patched test is
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-receipt/odp-consumer-two-fixtures.log`
(725 bytes, SHA-256
`fe470e247b86fbdbc1cf37f2161fe6dabdb6027283923005e6ee9921c2c74088`); all
7/7 `odp_unified_write` tests passed. The earlier combined focused receipt
also passed all 6/6 `xlsx_cells_runs` tests:
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-receipt/focused-consumers.log`
(1,250 bytes, SHA-256
`886947dcdd3917cb49444720d39a4e80697cd55f5d02cde7f74984c5e2cffe54`).

That earlier focused pass is also diagnostic rather than bounded provenance:
its command metadata did not include the shim `PATH`, explicit memory
environment, or an explicit `--target-dir`, even though its inherited target
path happened to print externally. The corrected two-fixture focused run with
all of those settings is in `0362uia`.

No production source was changed in this leaf. The only source patch is the
consumer test migration described above; the report goldens contain no
unreviewed rows, no `NORMALIZING` exception, and no new error code or warning.

The coordinator merge retains the TSV's empty final detail fields. A narrow
`.gitattributes` rule for this generated fidelity report treats its trailing
field separator as data; other whitespace checks remain enabled.
