# Frozen current DOCX source retention

The retained source head is `637082e318ae6bee2bd47ef317e4a417c1176996` on
`perf/docx-styles-effects-current-baseline-preflight-20260912`. The bundle is an
incremental Git bundle whose exact prerequisite is
`8702fd4db8723acceb7deb51bcb40ff66604bf10`.

`adc38a814f4256d02a7299ef4c3bd85780caa7b6` is **not** an ancestor of the
frozen head. It shares `8702fd4db` as the common ancestor, so a repository that
already contains `adc38a814` also contains the bundle prerequisite
transitively. Creating `adc38..637` is therefore not a valid incremental range;
the retained bundle was created as `HEAD ^8702fd4db`.

Minimal source replay from a repository containing the prerequisite:

```bash
git bundle verify captured-source-637082e31.bundle
git fetch captured-source-637082e31.bundle \
  637082e318ae6bee2bd47ef317e4a417c1176996:refs/heads/docx-current-637082
git worktree add /var/tmp/docx-current-637082 \
  refs/heads/docx-current-637082
git -C /var/tmp/docx-current-637082 rev-parse HEAD
```

The final command must print `637082e318ae6bee2bd47ef317e4a417c1176996`.
Use `current-overlay-files.json` to verify the five current overlay files
byte-for-byte after checkout.

The bundle contains source history only. It does not contain the external
current smoke receipts, the terminated mismatch receipts, Cargo target output,
preflight metadata, or host/process evidence. The approved raw receipts are retained beside this file; rejected receipts
are retained in `../smoke-rejected-source-133c4ede0/`. The original external
paths may be removed after byte verification against committed files. Absolute paths embedded in
receipts are provenance paths; replaying source from a new worktree does not
make those old receipt paths valid for a new run.

The current 52-lane smoke was validated and retained separately. No timing
capture is included here. Before authorizing timing, the profile runner's
historical smoke prerequisite and replay binding must be updated and reviewed
to require the current `8702fd4db` smoke receipt.
