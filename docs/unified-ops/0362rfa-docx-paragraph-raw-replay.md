# 0362rfa — DOCX raw paragraph replay scope

Status: repair complete on `uo2/w8-docx-paragraph-copy`.

The detached `source_backed::paragraph_copy::raw` tree now replays a decoded
copy or removal patch by the patch's recorded main-document XML bytes. A tree
reopened from the exact output can therefore apply the inverse, and a tree
reopened from the exact input can apply the forward patch. A second edit can be
replayed through the same reopened chain. A changed XML base remains a
`PatchConflict`.

Package-family replay keeps its stronger proof: `record::Patch::apply` and
`record::RemovalPatch::apply` still require the complete artifact fingerprint,
source version when present, and exact main-story bytes. The raw wrapper uses a
private `apply_raw` path so detached XML content identity does not weaken
package or durable publication replay.

Patch application now scans the recipient's output under its retained policy
before cloning the output bytes. The recipient's XML-byte, output-byte,
paragraph, event, and depth limits are checked before allocation; a rejected
application leaves the tree unchanged. Raw ingress accepts an owned `Vec<u8>`
so admitting a bounded tree does not hide an unbounded clone of borrowed input.
The compact `TailChange` representation and existing Arc-backed patch storage
remain unchanged.

## Validation

- `cargo test --offline --locked -p litchi-docx --all-features --test source_backed_paragraph_copy --test source_backed_paragraph_removal` — 13 copy and 6 removal tests passed, including reopened forward/inverse chains, changed-base conflicts, and recipient limit failures.
- `cargo fmt --all -- --check` — passed.
- `cargo clippy --offline --locked -p litchi-docx --all-targets --all-features -- -D warnings` — passed.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --offline --locked --no-deps --all-features -p litchi-docx` — passed.

No performance measurement, benchmark binary, or large evidence file is part
of this repair.
