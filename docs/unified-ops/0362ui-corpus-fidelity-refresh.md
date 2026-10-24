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
  `unsupported_capability` code, and the same revert result. Only the
  content-free ODT element reason changed to the static `text:variable-set`,
  `text:a` and `text:ruby` detail required by the 0358a source-backed ODF
  migration (`77ab7cab8a`).
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

The focused consumer receipt for that review is
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-receipt/focused-consumers.log`
(1,250 bytes, SHA-256
`886947dcdd3917cb49444720d39a4e80697cd55f5d02cde7f74984c5e2cffe54`). It
contains 7/7 `odp_unified_write` tests and 6/6 `xlsx_cells_runs` tests passed;
the updated ODP test explicitly checks the rename/inverse proof above. Its
peak RSS was 91,072 KiB.

No production source was changed in this leaf. The only consumer adjustment
is the stale ODP test assertion, migrated to the now-admitted flat projection;
the report goldens contain no unreviewed rows, no `NORMALIZING` exception, and
no new error code or warning.
