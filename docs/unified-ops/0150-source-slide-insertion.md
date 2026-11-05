# Source-backed slide insertion

This increment implements primary source-backed PPTX slide insertion and
connects it to unified `File::add`, transactions and batches. It follows the
existing ADR 0034 vocabulary and the native Edit/Commit/Patch lifecycle.

The native editor selects a registered layout by its exact name, reserves new
slide IDs, relationship IDs and part names against physical and effective
identities, and authors one empty slide. It publishes through the prepared OPC
topology owner. Untouched payloads retain source transfers; insertion does not
promote the package to an eager owner.

The v3 structural patch admits empty and disjoint slide catalogs when their
resource changes fully witness the transition. An insertion inverse can restore
an originally empty presentation; direct removal still refuses the last slide.
Exact metadata images support replay and inverse after reopening. Mixed text,
order, insertion and removal edits share the same effective native revision.

OPC now composes removal and append operations within noncanonical relationship
members by removing parsed child spans and appending authored children using
the existing root namespace prefix. It reparses the result against the complete
typed relationship graph. Comments, processing instructions, BOMs, root
attributes and trailing bytes remain intact. Replacement of an existing
relationship still requires the existing supported metadata path.

Insertion and patch replay preserve unrelated opaque relationship targets while
continuing to reject additional owners of the selected slide. Allocators reserve
case-equivalent dangling part targets and relationship IDs referenced through
custom namespace prefixes. Final-slide policy and unrelated-target policy are
separate checks. Durable patches prove exact before/after states and resource
images; they do not claim authenticated editing history.

The facade advertises source slide addition separately from shape/text-box
addition, which remains eager-only. Layout matching is case-sensitive and
ambiguous layout names are refused. Source topology, strict root policy and
dependency admission still constrain individual operations. Missing slide-list
elements and self-closing presentation roots remain outside the existing
lexical insertion grammar.
Structural recipes retain the existing v3 limits of 32 changed slide resources
and 128 MiB encoded patch bytes; this increment does not expand those limits.

Focused coverage includes source immutability, strict and transitional XML,
empty and disjoint catalogs, exact layout ambiguity, local and reopened replay,
inverse after mixed edits, allocation refusal followed by retry, cancellation,
retained patch resources, transaction poisoning and batch behavior. Three OPC
regressions cover mixed lexical preservation and refusal before output under
XML-size and managed-memory limits.

The full goal remains active. Source shape/text-box creation, broader format
CRUD, Python/CLI work and performance acceptance remain open. No paired
performance measurement was added and no performance claim is made.

## Verification

The final OPC suite passed 1,144 tests and doctests; the final native PPTX suite
passed 1,630; the combined DOCX/XLSX/PPTX/encryption facade suite passed 1,089.
The capability report generator passed its 14 tests and its generated delta
check. OPC/native and PPTX facade all-target Clippy passed with warnings denied.
Rustdoc for OPC, PPTX and the PPTX facade passed with warnings denied. Both PPTX
glossary profiles passed (79 minimal and 70 full banned names checked). Scoped
rustfmt and `git diff --check` passed. Crate boundaries passed for 66 packages,
268 internal dependency declarations and 11 explicit migration-debt items.
Facade checks with no features and with PPTX alone also passed.
The final audit of 665 changed/untracked paths found no prohibited large evidence
files or compiled binaries.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-opc -p litchi-pptx --features litchi-pptx/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_crate_boundaries.py
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features pptx
python3 tools/officecli_capability_delta.py --check --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The rewritten base remains an ancestor of the current branch; no rebase was
needed. Both required staged guards printed nothing. Scoped staging failed
because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created or permission workaround attempted.
Implementation and progress changes remain in the working tree.

All verification handles completed before cleanup. The three temporary corpus
fixture copies were removed after checking their recorded hashes. Generated
build output, the root lockfile, this turn's Python caches and scratch files
were removed. No temporary worktree was created. Inherited changes, tracked
lockfiles, the base checkout and other worktrees were preserved.
