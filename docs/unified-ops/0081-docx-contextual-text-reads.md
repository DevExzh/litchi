# 0081 — Contextual DOCX paragraph and run text reads

The typed-content repair in 0080 exposed the same namespace loss in ordinary
`Paragraph::text()` and `Run::text()`. Those methods now pass retained bindings
to the existing streaming scanner. Unified Node text and text predicates inherit
the correction through their leaf calls. The full goal remains active.

## Behavior

A retained run using another inherited Word prefix now yields its text, tabs,
breaks and hyphens. Foreign elements with matching local names remain excluded.
Local rebinding takes precedence and closing its scope restores the inherited
binding. Both eager and source-backed memory owners are covered, including
retained values read after the package is dropped.

The internal namespace tracker can now start with a bounded set of inherited
bindings at its base scope. It validates reserved declarations through the same
primitive as local declarations. No XML wrapper, self-contained fragment copy,
or new public text model is introduced. Paragraph/run parser admission and the
bounded namespace-context recovery remain in use.

Retained semantic reads enable unresolved-fragment fallback only at the root.
A resolved default Word namespace does not make a child with `xmlns=""` into
Word content. Internal raw-fragment helpers keep their historical heuristic;
the existing differential scanner oracles remain unchanged and pass. This is
an internal distinction between context-aware owners and context-free fragments,
not an alternate public compatibility API.

Control-character projection remains unchanged: this does not implement the
operation-specific OfficeCLI text policies documented in 0080. Whole-document,
cell and specialized story paths have not all been migrated to the new context
entry. Scalar typed-child shape validation from the 0080 audit remains open.

## Performance boundary

The scanner consumes the original bytes and retains its streaming token and
chunk behavior. Namespace-context recovery for a retained range can still scan
its owner prefix. This checkpoint makes no performance-acceptance claim: repeated
run reads need context captured/shared at traversal boundaries to avoid repeated
owner scans, followed by paired measurements. That work is required before the
read-path complexity and performance goals can be considered achieved.

## Validation

- 1,393 native DOCX unit tests passed, including all 77 paragraph tests and the
  unchanged raw-fragment differential oracle corpus.
- Five typed/flattened-content integration tests passed. They check inherited
  aliases after owner drop, Strict/default namespaces, local shadowing,
  namespace resets, literal versus structural content, and traversal bounds.
- Ten namespace-tracker tests passed, including base bindings surviving local
  scope pop and invalid reserved binding refusal.
- All-target DOCX/common Clippy passed with warnings denied.
- A unified regression checks paragraph/run Node text and exact text queries
  against expected strings for both eager and source-backed handles.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --lib
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --test text_content_namespaces
CARGO_NET_OFFLINE=true cargo test -p litchi-ooxml-common --no-default-features --lib binding_tracker
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --test sequential_text --test source_backed_managed --test source_read_policy
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx -p litchi-ooxml-common --no-default-features --all-targets -- -D warnings
```

Final unified DOCX validation passed all 680 tests/doctests. The managed-source,
source-policy and sequential-text targets passed all 39 tests. Scoped formatting
and diff hygiene passed. The rewritten base remains an ancestor of HEAD.
Both required staged evidence checks printed nothing.

Staging and committing remain blocked by the read-only linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit was
created; changes remain in the checkout. Checkpoint scratch files, generated
lockfiles and build output were cleaned. No temporary worktree was created;
earlier small benchmark evidence and its intentional lockfile were preserved.
