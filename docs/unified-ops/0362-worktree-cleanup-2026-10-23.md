# Worktree cleanup — 2026-10-23

Two inactive merged leaves were removed with `git worktree remove` without
`--force`. Their branch refs remain available for reproducible source recovery.

- `uo2-xlsb-provenance-review`: HEAD `7b3fd27e2ec9f6be3103807932828f183981b4fb`, ancestor of integration `a814baa1de`; 37,389 bytes reclaimed. It had no ignored or untracked files.
- `uo2-iwa-detection-api`: HEAD `7f3aef217eb9090199b8389a8d70bca9633e4ee3`, ancestor of integration `a814baa1de`; 341,906,569 bytes reclaimed. Eleven generated ignored files (461,243 bytes) were moved first to quarantine, with relative paths and SHA-256 records.

Both trees were clean, had no targets, protected measurement binaries, or active
process/file-descriptor users. Their branch refs were checked after removal;
the full receipt records exact `git worktree add` commands for recreation.

The DOCX section/redaction leaves and the old W8 DOCX/RTF leaves contain ignored
performance captures, native bundles, and portable `.zst` evidence. The replay
error audit leaf has tracked changes. Those trees were retained. Full inventory,
quarantine manifests, hashes, and before/after worktree lists are in
`.coord/worktree-cleanup-2026-10-23/` outside the checkout.

## Second cleanup wave

Four additional inactive source-only leaves were removed on 2026-10-23 after
their commits were confirmed as ancestors of integration `0ef732727d`:

- `uo2/0362rla-opc-policy-sparse` at `a79627ee3593708d8b630acd35fd185d573f0dbf`:
  58,274,082 bytes.
- `uo2/consumer-v8-repair` at
  `65e8f94f74bc0dcaef8c0f1bda240d67ac6db6a4`: 6,258,992,707 bytes.
- `uo2/0362ug-odf-errors` at
  `3b47fca40e5327c808bf5d04e51ee34837866f1f`: 73,609,004 bytes.
- `uo2/odf-consumer-repair` at
  `702accfa6243ad96ef0585087f3f523ea5e73f74`: 93,648,857 bytes.

All four trees had clean tracked state and zero untracked or ignored files.
Target, binary, capture, native, and portable-evidence directory checks found
nothing local to those leaves, and the process/file-descriptor audit found no
active users after the audit commands completed. No files were quarantined.
The consumer repair checkout contained shared tracked `docs/performance/results`
content; it was not changed or removed from Git, and its branch ref remains
available. The four branch refs remain after removal.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362rla-opc-policy-sparse uo2/0362rla-opc-policy-sparse
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-consumer-v8-repair uo2/consumer-v8-repair
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362ug-odf-errors uo2/0362ug-odf-errors
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-odf-consumer-repair uo2/odf-consumer-repair
```

The second wave reclaimed 6,484,524,650 bytes. It used `git worktree remove`
without `--force`; no branch or tracked object was deleted.
