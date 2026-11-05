# PPTX source ordering with modeled presentation properties

## Change

Source-backed Move/Swap now accepts the modeled, non-positional extensions in
presentation properties. This makes ordinary Builder output usable directly;
tests no longer replace its stock `presProps.xml` with stripped metadata.
The existing strict native property codec identifies root and slide-show
extensions. Order admission rejects unknown extensions, opaque nested extension
owners, invalid extension placement, markup compatibility, DTDs and processing
instructions. DrawingML color validation additionally checks namespace and
payload shape, including wrapper attributes.

Publication retains the original presentation-properties bytes. Custom shows
already used stable relationship references and remain unchanged during order
publication. The capability description now reflects that support, and the
unused custom-show refusal variant was removed. Sections, outline slide lists,
numeric presentation ranges and unknown or opaque metadata remain explicit
source-order gaps.

## Verification

Positive regressions cover stock Builder properties, all modeled root and
slide-show extensions, exact property bytes, custom-show preservation, logical
slide order and standalone source publication. Negative cases require typed
presentation-properties refusal and prove that a failed edit leaves the source
and committed no-op bytes unchanged. They cover unknown, nested, mismatched and
duplicate extension payloads, markup compatibility and foreign color content.

The final full native suite passed 1,578 tests/doctests across 100 groups, with
three existing ignores. The PPTX/encryption facade suite passed 562 tests and
doctests across 29 groups. Capability goldens were regenerated; the OfficeCLI
delta was regenerated and checked, and all 14 generator tests passed.

The first broad run caught invalid ordering in a new positive unit fixture and
a missing guard for opaque pen-color content. Both were corrected. Read-only
review also identified unmodeled laser-color wrapper attributes; the final guard
rejects them and checks pen, laser and recent-color payloads. Final review found
no remaining concrete blocker.

The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,078 tests and
doctests across 29 groups. The crate-boundary checker passed for 66 workspace
packages and 268 internal dependency declarations (11 existing explicit debt
items). Native and facade all-target Clippy and rustdoc passed with warnings
denied. Both PPTX glossary profiles passed (79 minimal and 70 full banned
names), as did the no-format capability test. Scoped rustfmt and
`git diff --check` passed.

Checks were serialized and offline:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/check_crate_boundaries.py
```

As in record 0141, three missing SVG corpus fixtures were temporarily copied
from the base checkout for the full native suite. Their paths, sizes and SHA-256
hashes are recorded there; they are verification inputs only.

## Scope and checkpoint

Source Add/Remove, richer positional metadata, remaining format and element
coverage, Python bindings and paired performance acceptance remain open. This
record makes no performance or full-goal completion claim.

The rewritten base remains an ancestor, so no rebase was necessary. No temporary
worktree was created. Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
Both required staged guards printed nothing; the changed/untracked artifact
audit found no prohibited evidence or compiled binaries.

All verification handles finished before cleanup. This turn's three temporary
corpus copies, scratch logs, generated root lockfile, build output and generated
Python cache files were removed. Inherited fixtures, tracked lockfiles and
working changes were preserved. The full goal remains active.
