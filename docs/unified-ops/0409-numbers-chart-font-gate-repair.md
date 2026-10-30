# 0409 — Numbers chart-font reset gate repair

- Base: coordinator source `3ce2a067742b9606e3c2558adee9e022c87f0420`.
- Scope: Numbers chart-font facade, the shared chart-font registry-admission
  path, and focused parity/contract tests; Numbers series-style paths remain
  separate ownership.
- Performance claim: none; no capture or timing was run.
- Runtime source commit: `c498ead7a1415430d5688babfdf029bca7b198b2`.
- Final handoff tip: `1255f04c38f7066d94dcde85c4b69e2cff74a57f` (the later
  commits are documentation-only and include synced coordinator trunk).

## Root cause and bounded repair

The root cause was an unnecessary registry-mutation admission immediately
after removing a style, before the actual registry mutation admission. That
duplicate charge pushed the chart-font field report above the residual default
ceiling; the observed ~108k field count was not the operation's true required
work. The subsequent exact inspection and mutation admissions remain intact.
The facade also reuses one admitted workspace for before-read, chart-font
mutation, and staged readback, and removes redundant immutable-source selector
re-resolution. Candidate publication still resolves the target and performs
strict semantic readback. The original debugger residual maxima were 93,258
and 93,378; the intermediate workspace-only repair reached 96,448 before the
registry double-charge was removed. No limit ceiling or headroom was raised.

## Verification

- `cargo fmt --all -- --check` — pass.
- Numbers chart-font parity set/size/reset/read vectors — pass (5 focused
  parity tests).
- `cargo test -p litchi-numbers --test sheet_chart_font --locked --offline` —
  pass (4 contract tests).
- Typed reset/readback/inverse/no-op behavior remains covered by the existing
  parity and contract vectors. Series-style failures are outside this slice.

External command receipts are under
`/home/zhuhe/code/litchi-worktrees/.coord/resume-20261028/font-gate-repair-1255/`:

- `parity-numbers-chart-font.log` — SHA-256
  `c05f03627d67316fe581b36cdaafa390680608dab5b417739bc29bf651e6c4ee`, exit
  0; five Numbers chart-font parity tests passed.
- `numbers-chart-font-contract.log` — SHA-256
  `37a502f1d49508e94a1dcad82c9a1b9706030ba4fedda6b0bcd5538f0add2e22`, exit
  0; four public contract tests passed.
- `shared-chart-font-unit.log` — SHA-256
  `8c105a2a20a6fa8f7efde8274e7be042257386197160affd370a9af1589226e0`, exit
  0; five shared chart-font unit tests passed.
- `clippy-numbers-style.log` — SHA-256
  `7eaf9ddfc3385b608acb2dba94c33d8c9c1b1ad369f2b524d4e4ba2c7775d509`, exit
  0; strict Clippy for Numbers and shared style passed (vendor warnings only).
- `fmt.log` — SHA-256
  `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`, exit
  0; formatting check was clean.
