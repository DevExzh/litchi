# Replay layout

The captured source head is
`0aa99593cb609487cf85230c361be23ca34d4367`; the incremental bundle requires
`506e6f8e5fb94a14c390fc262b152c773e8990b1`.

```sh
git bundle verify captured-source.bundle
git fetch captured-source.bundle \
  0aa99593cb609487cf85230c361be23ca34d4367:refs/heads/docx-after-smoke-506e6f8e5
git worktree add --detach /var/tmp/docx-after-smoke-506e6f8e5 \
  refs/heads/docx-after-smoke-506e6f8e5
git -C /var/tmp/docx-after-smoke-506e6f8e5 rev-parse HEAD
```

The final command must print
`0aa99593cb609487cf85230c361be23ca34d4367`. The bundle contains source
history only; it does not replace the 167 retained smoke receipts.

The raw receipts contain absolute paths from
`/var/tmp/litchi-docx-styles-effects-after-smoke-results-506e6f8e5-20260912-a`.
For a verifier replay, recreate that exact external path and copy only the
paths listed in `retained-files.json`, preserving bytes and embedded paths.
The retained repository directory is not a direct verifier output path.
