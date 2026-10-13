# 0252 — PPTX ADR 0034 W2 residue and P0 writes

Status: in progress on branch `uo2/p7c-pptx` (plan 0248 §6 steps 5-pptx and 7c).

## 1. Ingress options and patch conflict (W2 residue, part 1)

- `litchi_pptx::OpenOptions` (`source_limits`, `cache_limits`,
  `execution_context`, feature `encryption`: `encryption_limits`) is the one
  configured-open envelope. `OpenPolicy` names the lazy-only policies.
- Owning `Package`: `open_with_options`, `from_reader_with_options`,
  `from_bytes_with_options`, `open_with_password_and_options`,
  `from_reader_with_password_and_options`,
  `from_bytes_with_password_and_options`. Cache/execution policies are refused
  before input is read with `Error::UnsupportedOpenPolicy` (ops code
  `invalid_argument`).
- Deleted without aliases: `Package::{open_with_limits,
  from_reader_with_limits, from_bytes_with_limits, from_vec,
  from_vec_with_limits, from_vec_with_password,
  from_vec_with_password_and_limits, open_with_password_and_limits,
  from_reader_with_password_and_limits, from_bytes_with_password_and_limits}`
  and the six `open_with_*` combinations on `SourceBackedPresentation` and
  `SourceBackedPresentationEditor` (now `open_with_options`).
- Consumers formerly calling `from_vec(bytes)` now call
  `from_reader(Cursor::new(bytes))`, which keeps physical-source provenance
  (cross-slide copy authorization) at the cost of one bounded copy.
  `from_bytes` keeps its borrowed, provenance-free ingress. Affected
  perf-baseline legs (`pptx_*` routes in `tools/perf-baseline`) therefore
  include that copy; no performance claim about them is made here.
- `Error::StaleSource` is renamed `Error::PatchConflict` (ADR 0034 §6); ops
  still classifies it as `io_error` (the source changed under the handle).
- Glossary: both PPTX profiles ban the removed names; pptx-full also gains the
  primary-transaction bans pptx-minimal already had (plan 0248 §2 hole).

Not done here: positional `from_path_*`/`from_read_at_*` constructor
combinations on the source owners still exist (pending, as for DOCX in 0038).

## 2. Primary family completion (W2 residue, part 2)

- `litchi_pptx::opened` exports `Edit`, `Patch`, `Commit`, `Diagnostics`,
  `Error` and `LimitKind`; the family is registered in both PPTX glossary
  profiles, so the gate now checks one PPTX family.
- `opened::LimitKind` names the seven dimensions of `opened::Limits`
  (`Parts`, `PatchBytes`, `TextBytes`, `HistoryEntries`, `HistoryBytes`,
  `RetainedCandidateBytes`, `RetainedMceBytes`); `Limits::maximum(kind)`
  reads one bound.
- The 48 refusals that reported an `opened::Limits` bound as
  `Error::Limit { resource: "...", limit }` now report
  `Error::EditLimit { kind, maximum }` (ops: `limit_exceeded`). The
  LPTX0003 durable-patch golden transcript changed only in its 13
  `*.limits.*.refused` rows, which print that error.
- `opened::Error` is the crate `Error` re-exported. The ADR's per-module
  error split of the PPTX god-enum, and `observed` on edit limits, are not
  done (follow-up).

## 3. Rich text writes (plan 7c item 1) and table cells (item 2)

Leaf (`litchi-pptx`):
- `opened::{TextEdit, TextTarget, TextProperties, Rgb}` and one verb,
  `set_text_properties`, on the eager `opened::Edit`, the source
  `source::Edit` (composes through `stage_slide_edit`, so source-local
  replay and inverse patches cover it) and `SourceBackedSlideEdit`.
- Targets: a shape's whole body, a paragraph, a run (`a:r`/`a:br`/`a:fld`,
  counted as the text-body reader counts them), or a table cell.
  Properties: `bold`, `italic`, `size` (hundredths of a point,
  100–400000), `color` (`a:solidFill/a:srgbClr`), `font` (latin+ea+cs),
  `latin`, `align` (`a:pPr@algn`), paragraph/run/cell `text`, and cell
  `fill` (`a:tcPr` solid fill). Every property can be removed.
- Byte splices only: untouched attributes, siblings and extension lists keep
  their bytes; inserted children follow the `a:rPr`/`a:tcPr` schema order
  and the receiving element's prefix. The candidate slide is read back as a
  scene with identical shape count and identities. `table::Cell::fill()`
  reads a cell's local RGB fill.

Unified (`litchi-ops`), registry entries flipped to `set`:
- run: `bold`, `italic`, `size`, `color`, `font.latin`, plus write-only
  `font` and `text`;
- paragraph: `align`, `size`, plus write-only `text`, `bold`, `italic`,
  `color`, `font` (applied to every run);
- shape/placeholder: `bold`, `italic`, `size`, `color`, `font`,
  `font.latin`, `align` (every paragraph and run; shape `text`/geometry and
  formatting stage in one native edit);
- `tc`: readable/settable `fill`, plus write-only `text`, `bold`,
  `italic`, `size`, `color`, `font`, `align`. The `tc fill` gap
  (L-PPTX-R8) is closed.
- Colours: only opaque RGB without transforms is writable
  (`unsupported_capability` otherwise, with a suggestion). Run text with a
  line break is `invalid_value`.
- Path and query targets (`/slide[N]/shape[M]/paragraph[K]/run[L]`,
  `/slide[N]/table[M]/tr[R]/tc[C]`, scoped `slide[1] run` queries) fan out
  through the same native verb, 256 edits per call; batch and `edit`
  closures reuse `set` (the shape coalescer declines formatting payloads).
- Tests: `crates/litchi-pptx/src/opened/text_format_tests.rs` (6, both owners
  and inverse replay) and `crates/litchi-ops/tests/pptx_text_format.rs` (6:
  read-back via `get` on both owners and after a byte round trip, clears,
  query fan-out, atomic `edit` closures, typed refusals). Capability golden
  `pptx.json` re-blessed; OfficeCLI delta regenerated (pptx: 153→164
  supported, 280→271 missing).
