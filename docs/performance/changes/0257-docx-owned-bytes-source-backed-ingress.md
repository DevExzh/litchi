# Change 0257: source-backed DOCX owned-byte ingress

## Status

Landed in `2bf6a7447`. This is a correctness- and work-removal integration,
not measured before/after performance evidence.

## Scope

The ordinary `litchi::Document::from_bytes` and
`Document::from_bytes_with_limits` DOCX path now moves the caller's buffer
behind `OwnedSource`, opens one `SourceBackedPackage`, and retains the existing
source-backed DOCX owner in the private facade variant. Opening still performs
the mandatory ZIP/OPC catalog, content-type, relationship, and DOCX owner
admission work. The main document and ordinary unselected Part payloads remain
deferred until a semantic query requests them.

This removes the previous mandatory eager `OpcPackage` materialization and
full `document.text()` validation pass for valid normal-facade DOCX byte
inputs. Consequently malformed main-document XML can now succeed at catalog
open and return its typed parse failure on the first semantic query. Tests use
an independently constructed eager `OpcPackage` owner to pin both error-timing
contracts.

## Arbitration, limits, and preservation

- Non-ZIP and non-DOCX inputs remain under the established smart detector.
  Source-catalog fallback drops every temporary source owner before recovering
  the original `Vec` allocation; tests pin pointer and capacity identity for
  an ordinary ZIP and a valid non-DOCX OPC package.
- A catalog identified as DOCX returns its WordprocessingML owner-admission
  error directly instead of silently retrying through eager materialization.
  Hard OPC source/open failures also bypass format fallback. OPC read-limit
  failures cross the facade as structured `ResourceLimit` values with the
  exact observed and configured counts.
- With ODT enabled, the owned source is MIME-arbitrated before a hard DOCX
  limit is surfaced. Ordinary ODT packages keep their own policy; valid
  OOXML/ODT polyglots keep OOXML precedence; malformed OOXML catalogs inside an
  ODT-marked package fail closed.
- The source-backed facade variant and its Markdown cache are target-independent
  for owned bytes. Filesystem source detection remains limited to platforms
  with `FileSource`.
- Active `altChunk` and unknown body blocks remain typed Markdown refusals on
  both the eager oracle and the new owned source-backed path. Publication and
  edit-preservation behavior are unchanged.

## Harness compatibility

The existing `docx_file_eager_*` controls must remain eager even though the
normal byte facade is now lazy. The harness therefore recreates the historical
facade preparation through smart detection, `OpcPackage`, the typed DOCX
owner, the eager full-text validation boundary, and facade-equivalent metadata
conversion. Timed paragraph enumeration also retains the unified paragraph
wrapper projection. Source selectors still call `Document::open`.

All five former byte-facade ingress sites were updated: prepared query roots,
eager open, both eager lifecycle controls, and the untimed eager/source parity
oracle. Selector names, timing boundaries, source metric classifications, and
historical evidence remain semantically stable. This change does not add a
new owned-byte benchmark selector.

## Verification

- DOCX facade tests: 24/24 passed.
- DOCX plus Markdown facade tests: 31/31 passed.
- DOCX plus ODT facade tests: 34/34 passed.
- Allocation/admission source probes: 3/3 passed.
- The adjacent `pptx,odt` minimal-feature library check passed.
- The performance harness binary check passed. A debug CLI smoke covering one
  `docx_file_eager_open` and one `docx_file_source_open` sample completed with
  two filesystem evidence records and the full untimed semantic/archive gates.
- Targeted rustfmt checks, the crate-boundary gate, and `git diff --check`
  passed. Independent production, cfg, error-contract, and harness reviews
  found no remaining blocker.

The smoke required a repository-local `TMPDIR` because the host `/tmp` quota
was exhausted, and an enlarged debug stack because the debug harness corpus
builder exceeded the default main-thread stack. These are environment notes,
not performance observations. It ran from the final pre-commit worktree, so
its binary descriptor names base revision `d4bb5447e` rather than subsequent
commit `2bf6a7447`; the disposable report is validation only and is not retained
performance evidence for the commit. The dedicated 8.4 GiB harness build
target was cleaned afterward; unrelated worktree changes were not touched.

Workspace-wide fmt, Clippy, rustdoc, and test gates were not rerun. Unrelated
user-owned ODF files remain outside this batch, and the prior narrow Clippy
attempt already stops on two pre-existing `litchi-opc`
`clippy::double_must_use` findings.

## Claim boundary

No latency, allocation, RSS, physical-I/O, decompression-byte, or end-to-end
DOCX speedup is claimed. A future CPU-pinned release A1/B1/B2/A2 run with 20
warmups and 500 retained samples per leg must compare catalog-only open,
selected-query lifecycles, full traversal, and media-heavy cases before
quantifying the removed eager work.
