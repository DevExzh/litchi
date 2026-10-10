# 0083 — Decode typed DOCX controls in the current namespace scope

Typed break and note-reference decoding now uses the enclosing run traversal's
resolver instead of reconstructing a temporary Run and rescanning the owning
XML part for each child. This closes the scalar-subparser scan gap recorded in
0082 and the malformed-child audit item recorded in 0080. The full goal remains
active.

## Behavior

A single shared break-attribute decoder serves `Run::breaks()` and typed
`Run::contents()`. Word attribute namespaces, duplicate expanded names, default
values and valid type/clear values retain the same checks. Note references use
the same namespace-aware attribute primitive for their IDs. Both consume the
active scope at the child's closing or empty event, where local declarations
still apply. No namespace snapshot or owner-prefix traversal is needed per
scalar child.

Before producing a typed scalar, the content reader checks its complete XML
shape. Breaks, tabs, carriage returns, hyphens, note references and note marks
accept empty or paired tags with whitespace and inert comments/processing
instructions. Nested elements, authored text, CDATA and character references
are refused rather than silently discarded from the typed value. Foreign
children remain exact opaque content. Attribute decoding stays checked and
note-ID failures no longer echo the authored attribute value.

These checks apply to typed run content, not schema validation of every DOCX
reader. `Run::breaks()` retains its existing traversal behavior. Images,
embedded objects, literal text shape validation and specialized stories are
outside this checkpoint. OfficeCLI input/output policies, broader context
capture, paired performance evidence, remaining formats and Python are still
required.

## Evidence

New regressions exercise every scalar family with four invalid payload forms,
paired tags, comments, foreign opaque content, numeric references in note IDs,
and a locally rebound Word alias on both breaks and notes. A dense run with
1,000 breaks, 1,000 footnote references and 1,000 endnote references returns all
3,000 typed children with zero owner namespace scans after the paragraph is
dropped. The test-only counter is absent from production. This is direct work
count evidence, not a timing or allocation claim.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --lib --test text_content_namespaces --test source_backed_managed --test source_read_policy
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx --no-default-features --all-targets -- -D warnings
```

Validation passed 1,397 native unit tests, 19 source/content tests and 680
unified tests/doctests. All-target Clippy passed with warnings denied. Scoped
formatting and diff hygiene passed. The rewritten base remains an ancestor of
HEAD. Both required staged evidence checks printed nothing; the artifact audit
checked 365 changed/untracked paths and found no prohibited artifacts.

Staging and committing remain blocked by the read-only linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created. Changes remain in the checkout. Checkpoint scratch files, generated
lockfiles and build output were cleaned; no temporary worktree was created.
Earlier small benchmark evidence and its intentional lockfile were preserved.
