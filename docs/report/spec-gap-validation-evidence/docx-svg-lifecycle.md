# Committed DOCX SVG attachment lifecycle

Commit `892441d95` adds source-backed attachment and detachment of SVG media on
existing direct main-story pictures. It uses the prepared OPC topology introduced
by `57680dc86792e38ee12bd6f717ec57c4986034ad`.

## Ordinary API

The entry point is `litchi_docx::source_backed::Package`.
`PictureSelector::new(0, picture)` selects a zero-based direct-picture ordinal
in the main story. The currently supported drawing inventory is `0`.

For one picture:

1. Call `package.edit_svg_attachment(selector)`.
2. Call `edit.attach_svg(SvgInput::borrowed(svg_bytes))` or `edit.detach_svg()`.
   The returned boolean reports whether staging changed the selected state.
3. Call `edit.commit()` to obtain a checked commit.
4. Call `package.publish_svg_attachment_commit_to_stream(output, &commit)`.

For several selected pictures, use `package.edit_svg_attachments(selectors)`.
Its `attach_svg(selector, input)` and `detach_svg(selector)` methods stage one
batch; publish the result with
`publish_svg_attachment_batch_commit_to_stream`. Ordinary input does not require
caller-generated relationship IDs or media Part names.

`svg_picture_sources(StorySelector::Main)` exposes source-backed resource views.
Media data remains deferred until requested. `svg_pictures` is the eager
compatibility inventory; callers concerned with unnecessary media reads should
choose the source-view API. Commit diagnostics expose the staged-operation and
selected-picture counts and whether the commit changes story or dependency state.

## Preservation and supported scope

The operation retains the existing internal PNG fallback, drawing placement,
and unrelated source markup. It changes the admitted SVG extension and its
relationship/content-type/media dependency closure. Shared media remains while
other incoming relationships require it. Source checks cover publication and
inverse application; stale sources and unsafe dependency changes are refused.
The publication APIs also expose source-checked physical inverse publication.

This batch does not create pictures, add subsidiary-story inventories, render
SVG, or execute its contents. Linked/external resources and ambiguous or
unsupported drawing ownership do not become authored owners by inference.
SVG bytes are opaque media, not a promise of rendering or SVG-language validity.

## Verification and remaining performance work

The committed integration target contains 20 tests covering lifecycle edits,
shared cleanup, inverse/source checks, native inline and floating drawings,
Strict namespace handling, opaque preservation, and deferred media access:

```sh
cargo test --locked --offline -p litchi-docx --test drawing_svg_lifecycle
cargo clippy --locked --offline -p litchi-docx --test drawing_svg_lifecycle -- -D warnings
```

The supporting prepared-topology targets cover borrowed candidate validation,
freshness, resource limits, and empty metadata roots. These are correctness and
bounded-resource gates, not general peak-memory or throughput acceptance.
The separate DOCX lifecycle performance harness is still under review. Its
smoke output must not be treated as completed performance evidence or a speedup
claim. Large-batch staging cost and measured allocation/latency scaling remain
part of the performance program in `docs/GOAL.md`.
