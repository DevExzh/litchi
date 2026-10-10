# 0085 — Inline controls in direct DOCX paragraph text setters

Supported direct paragraph command setters now normalize CRLF/CR and vertical
tab to LF and publish native inline breaks. The underlying primary paragraph
setter also accepts LF and tabs. This extends the operation-specific input
policy introduced for run addition in 0084. The full goal remains active.

## Leaf prerequisite

An initial command-only probe failed: `replace_paragraph_text` still used the
plain-text validator, even though its slot encoder and scanner already support
inline control elements. Merely removing command admission was insufficient.
The leaf validator now admits tab and LF for the scalar paragraph setter,
including its managed implementation and owner preflight. CR, VT and other
XML-illegal controls remain refused by the native setter.

The existing encoder emits tab/wrapping-break elements, retains text-slot and
run-formatting distribution, preserves opaque siblings and paragraph count,
and checks semantic readback. Existing preflight size calculations, managed
admission, immutable publication and exact inverse machinery remain in use.
The separate native bulk paragraph replacement API retains its prior plain-text
admission in this checkpoint.

## Command admission

Only commands containing LF/CR/VT need the additional target check. The target
is parsed and resolved against the current candidate. A supported DOCX paragraph
receives the same inline normalization as run addition; its typed target is
reused for mutation. This closes pending formatting before reading the target.
Repeated/case-insensitive properties keep their order. Ordinary setters retain
the existing path without the extra lookup.

Run setters, selector setters and paragraph addition with controls remain typed
refusals. Unsupported paragraph locations still pass through the existing leaf
selection rules and cannot become implicitly writable. The capability reason
and generated delta distinguish the newly supported direct paragraph context.

This is control-input support, not a claim of complete OfficeCLI paragraph
replacement parity: native text-slot distribution stays intact. Flattened output
policy, structural paragraph splitting, selector-wide context admission, native
bulk controls, other formats, Python and performance acceptance remain open.

## Validation

New eager/source transaction tests reopen native output, verify two preserved
runs and paragraph count, repeated exact no-ops, unchanged input, rejected native
CR/VT/illegal controls and exact patch inversion. A managed-source regression
verifies LF/tab readback, repeated no-op and failure isolation. Command tests
cover formatting barriers, case-insensitive paragraph aliases, CRLF/CR/LF/VT/tab,
atomic and best-effort execution, exact inverse bytes, and continued refusal of
run, selector and body targets with original item indices.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --lib
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --test source_backed_managed_paragraph_batch
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx -p litchi-ops --no-default-features --features litchi-ops/docx --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Final validation passed 1,397 native unit tests, 9 managed paragraph tests,
685 DOCX and 418 minimal unified tests/doctests, plus all-target Clippy with
warnings denied. All 29 batch tests and 14 delta tests passed; the generated
delta check, scoped formatting and diff hygiene passed. The rewritten base
remains an ancestor of HEAD, and both required staged evidence checks printed
nothing.

Staging and committing remain blocked by read-only linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created; changes remain in the checkout. Checkpoint scratch files, generated
lockfiles, Python cache and build output were cleaned. No temporary worktree
was created. Earlier benchmark evidence and its intentional lockfile remain.
