# 0082 — Reuse namespace context across DOCX run reads

Record 0081 corrected text reads but recovered inherited context by scanning
from the start of the owning XML part for every retained run. Paragraph
traversal now captures that context while its resolver is already at the run.
The full unified operations goal remains active.

## Result

`Paragraph::runs`, `inlines` and `addressable_runs` attach captured namespace
bindings to each returned Run. The last of these supplies unified positional
run paths and queries. The existing bounded `NamespaceCapture` reuses an Arc
when the scope is unchanged and bounds newly captured bindings and aggregate
bytes. Local declarations and nested scopes remain authoritative.

`runs` now also seeds its scanner with inherited paragraph bindings, so another
inherited Word alias can name the run itself. Nonempty outer and inline runs
capture at their closing event, before the resolver pops the root scope;
addressable runs capture at the opening event, including nested runs. Empty
runs capture in their own event scope.

Run clones share the captured bindings. Text, typed content and break readers
consume that context instead of recovering it from the full owner. The original
XML range and owner admission remain attached. Cached-context access still
performs parser admission; replacing a Run's XML clears the captured context.
Fallback constructors without captured context retain bounded owner recovery.
No public operation or namespace API was added to DOCX.

## Evidence and limits

Tests check all three traversal families, shared Arc identity for adjacent
scopes, independent local rebinding, empty runs, cheap clones and expected text.
A thread-local test counter at the owner namespace scan proves that reading
1,000 captured outer or addressable runs performs zero owner namespace scans,
after the paragraph has been dropped. The counter is absent from production.
This is an operation-count assertion, not a timing or allocation claim.

Paragraph context recovery itself, separately constructed ranges, typed-content
scalar subparsers and specialized readers still need broader context reuse.
Paired performance measurements, OfficeCLI text-policy completion, remaining
format migrations and Python are still required. This checkpoint does not claim
that the full read-path complexity or performance gate is satisfied.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --lib --test text_content_namespaces --test source_backed_managed --test source_read_policy
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx --no-default-features --all-targets -- -D warnings
```

Final validation passed 1,395 native unit tests, 19 source/content integration
tests and 680 unified tests/doctests. All-target Clippy passed with warnings
denied. Scoped formatting and diff hygiene passed. The rewritten base remains
an ancestor of HEAD, and both required staged evidence checks printed nothing.

The linked Git metadata remains read-only: staging/committing cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created. Changes remain in the checkout. Checkpoint scratch files, generated
lockfiles and build output were removed; no temporary worktree was created.
Earlier benchmark evidence and its intentional lockfile were preserved.
