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

## Third cleanup wave

Five further inactive source-only leaves were removed on 2026-10-23 after
their commits were confirmed as ancestors of integration
`18983596a3a1b0f2ab937f4fb1f7d9ef29f18f4c`:

- `uo2/consumer-v9-repair` at
  `59b9a87de1aa8162b1ca9fe8fdc5cca21ebb314b`: 124,336,611 bytes.
- `uo2/xlsx-current-consumers` at
  `1469c148f71cdd7dd494731ac378d60897c98c05`: 6,258,986,092 bytes.
- `uo2/opc-required-policy` at
  `d6ff788d3fe1d8a139b5e26d55102e274b1cd890`: 6,259,140,205 bytes.
- `uo2/0362ui-pptx-consumers` at
  `c077e9f8d19be821cd682d03d550b4d579f7b10a`: 64,327,811 bytes.
- `uo2/0362rnb-opc-recipient-policy` at
  `752a5fa9d1d0ba00ff5e2adb2144e0d46d192988`: 36,456,058 bytes.

The apparent checkout total was 12,743,246,777 bytes. The required-policy
checkout included four generated ignored Python bytecode files totaling
131,671 bytes; those were moved before removal, so the retained source/tree
total was 12,743,115,106 bytes. The quarantine is outside the checkout at
`.coord/worktree-cleanup-2026-10-23/quarantine/third-wave/`, with relative
paths, sizes, and SHA-256 values in `manifest.json`:

- `tools/__pycache__/api_debt_inventory.cpython-314.pyc`, 17,557 bytes,
  `cabb601951fd23533aee89139df55a11d29dd5cf28cd29960aabedad36437522`.
- `tools/__pycache__/check_api_glossary.cpython-314.pyc`, 44,684 bytes,
  `e9255bc1b6625cb53fffb246f8b8b8b3d8e8e09ca882b758648bedc0f82ec0ff`.
- `tools/__pycache__/rustdoc_return_types.cpython-314.pyc`, 15,603 bytes,
  `293c01eeb71adb89ed39789d0f27c469db1996fbc210dee64da3f9d63a0fcb90`.
- `tools/__pycache__/test_check_api_glossary.cpython-314.pyc`, 53,827 bytes,
  `ddff51a86d068e6ef5df9fc0f89356258bf442c9c173b7f8d24a5725afb1a1aa`.

All five trees had clean tracked state. Target, binary, capture, native,
portable-evidence, and other protected measurement checks found no local
artifacts, and the process cwd/file-descriptor audit found no active users.
Shared tracked `docs/performance/results` and `docs/report` evidence was not
changed or removed from Git; all five branch refs remain at the hashes above.
No active/source leaf was touched. Removal used `git worktree remove` without
`--force`.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-consumer-v9-repair uo2/consumer-v9-repair
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-xlsx-current-consumers uo2/xlsx-current-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362rd-substrate-scope uo2/opc-required-policy
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362ui-pptx-consumers uo2/0362ui-pptx-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362rnb-opc-recipient-policy uo2/0362rnb-opc-recipient-policy
```

The third wave reclaimed 12,743,246,777 checkout bytes while retaining all
source refs, tracked evidence, and quarantined generated files.
