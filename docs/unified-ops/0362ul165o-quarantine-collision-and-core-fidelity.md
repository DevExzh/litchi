# 0362ul165o — Quarantine collision and core-properties fidelity

The remaining-29 diagnostic stops at CLI validation after six passing rows
and the two known DOCX assertion failures. The CLI library's three tests pass,
but artifact retirement collides with an earlier suite's
`artifact-quarantine/0000/` destination. The preservation barrier correctly
stops the run; 20 later checks remain unrun. Source, locks and resource checks
are unchanged. The diagnostic and its artifacts remain preserved.

Root merges `28a6967294665e8f097a4a81464d0aba28b26af0`, including its
`9d364a0226` predecessor. Quarantine paths now include a stable namespace
derived from the resolved log directory and gate attempt, so independent
suites may reuse chunk numbers. An existing artifact is reused only when
its bytes, size and mode match; differing or nonregular destinations are
refused without overwrite or a false manifest entry. Receipts retain the
scope, namespace and reuse decision.

Independent review reproduces the public multi-suite flow, then returns a
symlink check performed after canonicalization. The successor checks the
lexical destination and every ancestor before resolution or reuse. Root
reviews that fix and runs all 22 focused tests, including the regression
that preserves the source and leaves the manifest untouched on refusal.
The helper's identity changes, so future gate plans must use fresh pins.
Existing frozen controls and failed receipts are not rewritten.

## DOCX validation

The queued DOCX watcher correctly refuses to start after the diagnostic's
preservation stop. Root verifies termination and launches focused validation
independently. Its first attempt fails before compilation because the sparse
worktree omitted path dependencies. Root materializes tracked crates,
vendored dependencies and test fixtures, preserving HEAD and clean tracked
status, then reruns under the same bounded driver with fresh state/target.

Seven metadata tests pass and one fails. The corrected present-versus-absent
and core-part-projection assertions now pass. The remaining failure is real:
the combined core/app/custom edit's inverse restores the typed core values
but does not restore the original archive bytes. The core patch retains
typed values and regenerates XML rather than retaining the exact part delta.
The exact-byte assertion stays in place. A production fidelity repair is
required before the DOCX leaf can merge; this is no longer merely a test
expectation correction.

## W6 preparation

W6 v8 passes independent source review. All four v6 findings are closed:
safe token arguments, a positive test through the actual producer receipt
path, verified selector patches, and complete manifest/context bindings in
the public recipe/producer/final-validator flow. Seventeen focused tests,
policy, AST/JSON and shell syntax checks pass. Earlier packets and the
successful TDnt83 source preparation remain unchanged.

All 28 qualification preparations are now source-ready. Actual final-C/P
rebinding, native builds, correctness, fidelity, allocations and ABBA
measurements remain pending. The DOCX production repair and fresh final
all-65 gate are still required. Memo 22 alone is deferred. Nothing is pushed.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `rest29-terminal-quarantine-collision-20261026/receipt.json` | `e76dbf7a1d4800f09afd7e7d710a2b7901e8aefb7ab1829b06082feec31bc560` |
| `gate-quarantine-scope-9d364a-source-review-20261026/review-receipt.json` | `6745c93f9431d67c7c80f8569ade4c8d6124435b86ed39b5b4dcfed0af38d7f3` |
| `docx-metadata-focused-ec243-v2-launch-20261026/launch.json` | `8350c016cf7439c04fb7c58862e57f91f3334c3efc67ab873cb644259d1ec3df` |
| `w6-odt-ods-C2a-prep-20261026-v8/seal.json` | `9fb4fedd3ca93d75ac9666026a22d4f084c2d5a606c32c5915142baa59488686` |
| `w6-odt-ods-C2a-prep-20261026-v8-independent-review-fresh-20261026/review-receipt.json` | `2667316d508ee10971e426a200a81f433d748d5873b7261dae7c3323aade6e27` |

Focused DOCX output remains under
`docx-metadata-focused-ec243-v2-20261026/`. The quarantine tests run with:

```sh
python3 -B -m unittest discover -s docs/unified-ops/results/0362ul29/tooling -p test_safe_chunked_tests.py
```
