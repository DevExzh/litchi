# 0080 — Namespace-correct DOCX typed run content

The existing `RunContent` enum already distinguishes literal text, wrapping,
page and column breaks, carriage returns, tabs and note references. This
checkpoint repairs that reader instead of introducing another text event API.
The full unified operations goal remains active.

## Result

`Run::contents()` seeds its resolver with inherited namespace bindings.
Break and note-reference subparsers retain their `XmlRef`, including the
namespace context, instead of rebuilding detached byte fragments. Typed content
therefore works with multiple inherited Word aliases after either an eager or
source-backed memory owner has been dropped.

`Run::breaks()` resolves expanded element and attribute names. Foreign `br`,
`type` and `clear` lookalikes no longer become Word breaks or override their
values. Duplicate expanded Word attributes are refused. Strict and transitional
namespaces, default namespaces and local prefix rebinding are covered.
Unresolved standalone fragments retain the existing fallback, but a resolved
Word root no longer enables that fallback for explicitly unqualified children
or note IDs. Default namespaces do not apply to attributes.

Both traversals enforce the existing 128-depth and one-million-element bounds;
empty children now count toward depth. Parser admission remains in place.
These bounds are not a claim that repeated inherited-context scans are optimal.
Nested payload validation for scalar break/note children remains an audit item.

## Text-policy correction and next work

Record 0076's blanket command refusal for LF/CR/VT is conservative, but its
paragraph-splitting rationale was too broad. Inspection of the base checkout's
`3rdparty/OfficeCLI/src/officecli/Handlers/Word/` shows distinct contexts:

| Operation | Reference behavior |
|-----------|--------------------|
| Add paragraph | Block-level LF handling can create paragraphs |
| Set paragraph text | `AppendTextWithBreaks` normalizes CRLF/CR and maps LF/VT to inline breaks |
| Add run text | Uses the same inline break helper |
| Set run text | Validates without soft-break allowance and assigns literal text to the first text element |
| Read run text | Wrapping breaks and carriage-return elements become VT; page/column breaks are omitted |

The relevant sources are `WordHandler.Add.Text.cs` (`AppendTextWithBreaks`),
`WordHandler.Set.Element.cs` (paragraph/run text setters), and
`WordHandler.Helpers.Text.cs` (`GetRunText`). Native flattened text currently
collapses distinctions that this typed reader preserves. Mapping a flattened
string back into controls would lose information.

This checkpoint does not change flattened reads, unified Node text, selectors,
write policies or the conservative batch refusal. Next work must compose the
existing typed run/paragraph model, define operation-specific write handling,
then admit each command context with eager/source round-trip evidence. Remaining
formats, Python and broader performance acceptance also remain open.

## Validation

The retained-alias regression failed before the fix. The initial paragraph
suite passed 41 tests; eager/source owner-drop and Strict namespace integration
checks passed. Review prompted added namespace-reset, unqualified note-ID and
empty-child-depth regressions. Final validation is recorded below.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --no-fail-fast
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --lib --test text_content_namespaces
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx --no-default-features --all-targets -- -D warnings
```

The full leaf run passed 2,257 tests/doctests, with four failures caused by
missing external producer fixtures and 32 ignored tests. The missing fixtures
were temporarily copied from the base checkout for the affected targets:
Open XML SDK `svg.docx`, LibreOffice `tdf164835_nonDummyLineHeight.docx`, and
OfficeCLI `revisions.docx`. Unified DOCX passed all 679 tests/doctests.

All three affected fixture targets subsequently passed (36 tests). After the
review fixes, all 1,393 native unit tests and four typed-content integration
tests passed, and final all-target Clippy passed with warnings denied.

The rewritten base remains an ancestor of HEAD, so no further rebase is needed.
Both required staged evidence checks printed nothing. No temporary worktree was
created. Temporary fixture copies, build output and checkpoint scratch files
were removed; earlier benchmark evidence and its intentional lockfile remain.

Committing is still blocked: linked Git metadata under
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api` is read-only, preventing
creation of `index.lock`. No commit was created; changes remain in the checkout.

The final unified DOCX rerun also passed all 679 tests/doctests after the
review fixes. The artifact audit checked 356 changed/untracked paths and found
no prohibited evidence files or compiled binaries. Both staging and commit
attempts failed at the read-only Git lock.
