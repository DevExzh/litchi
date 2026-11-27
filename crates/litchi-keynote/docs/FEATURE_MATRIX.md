# Keynote Feature Matrix

Feature checklist for Apple Keynote (`.key`) packages. ✅ Supported · 🟡 Partial · ❌ Unsupported ·
N/A Not applicable. Writes edit existing packages unless a row says otherwise.

## Package and semantic projection

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Existing `.key` ZIP ingress from path or bytes | ✅ | ✅ | N/A | |
| Exact retained-artifact stream to a caller sink | ✅ | N/A | ✅ | |
| Archive-free ZIP or package-directory semantic presentation | ✅ | ✅ | N/A | |
| Semantic show, slide order, text, notes, media/build summaries | 🟡 | ✅ | N/A | Selected content only |
| Unknown ZIP members, IWA components, and protobuf fields | 🟡 | ✅ | 🟡 | Preserved, not understood |
| Focused source-bound transactions and reversible patches | 🟡 | ✅ | 🟡 | One operation per commit |
| Fresh `.key` package creation | 🟡 | N/A | 🟡 | |
| Durable patch serialization, composition, merge, and history | ❌ | N/A | ❌ | |
| Durable atomic filesystem save | 🟡 | N/A | 🟡 | |
| Metadata and properties sidecar | 🟡 | ✅ | ❌ | |

## Slides, text, settings, and playback state

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Slide selection by checked position or unique navigator name | ✅ | ✅ | N/A | |
| Existing slide skip/include state | ✅ | ✅ | ✅ | |
| Existing slide move/reorder | 🟡 | ✅ | 🟡 | Flat slide topology only |
| Existing slide deletion | 🟡 | ✅ | 🟡 | The last slide cannot be deleted |
| Slide insertion, duplication, or fresh slide creation | 🟡 | N/A | 🟡 | |
| Existing title/body placeholder text | 🟡 | ✅ | 🟡 | Plain text only |
| Existing speaker notes text | 🟡 | ✅ | 🟡 | Absent notes are not created |
| Per-slide placeholder visibility | 🟡 | ✅ | 🟡 | |
| Rich-text storage inspection and plain-text projection | 🟡 | ✅ | N/A | |
| Rich-text styles, lists, links, revisions, and inline-object CRUD | 🟡 | 🟡 | 🟡 | No revisions or inline-object editing |
| Show dimensions and playback settings | 🟡 | ✅ | 🟡 | No playback |
| Slide transitions | 🟡 | ✅ | 🟡 | |
| Slide backgrounds | 🟡 | ✅ | 🟡 | None, solid and gradient fills only |
| Builds and animation summaries | 🟡 | ✅ | ❌ | |
| Existing build playback order | 🟡 | ✅ | 🟡 | |
| Shapes, groups, lines, transforms, and z-order editing | 🟡 | ✅ | 🟡 | Groups cannot be authored |
| Masters, templates, themes, layouts, guides, and inheritance authoring | ❌ | 🟡 | ❌ | |
| Comments | 🟡 | ✅ | 🟡 | |
| Collaboration and coauthoring | ❌ | ❌ | ❌ | |

## Tables

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Table selection/discovery in an existing slide | 🟡 | 🟡 | N/A | |
| Table name | 🟡 | ✅ | 🟡 | |
| Table title visibility/settings | 🟡 | ✅ | 🟡 | |
| Table appearance and style metadata | 🟡 | ✅ | 🟡 | |
| Table dimensions | 🟡 | ✅ | 🟡 | Sizes only; no row or column insertion |
| Header/footer/freeze/repeat settings | 🟡 | ✅ | 🟡 | |
| Table lock state | 🟡 | ✅ | 🟡 | |
| Persisted table sort configuration | 🟡 | ✅ | 🟡 | Does not move rows |
| Physical table row sorting (“Sort Now”) | 🟡 | 🟡 | 🟡 | Existing tables only |
| Existing table-cell Number display format | 🟡 | ✅ | 🟡 | |
| Slide-table merged-cell geometry | 🟡 | ✅ | 🟡 | |
| Cell mutations, formulas, formats, and row/column topology | 🟡 | ✅ | 🟡 | Formulas not evaluated |

## Charts

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Existing slide chart catalog and visible titles | 🟡 | ✅ | 🟡 | |
| Existing chart captions | 🟡 | ✅ | 🟡 | |
| Existing chart axis titles | 🟡 | ✅ | 🟡 | |
| Existing primary value-axis settings | 🟡 | ✅ | 🟡 | |
| Existing chart Arrange interaction flags | 🟡 | ✅ | 🟡 | |
| Existing chart legend visibility | 🟡 | ✅ | 🟡 | |
| Chart data, series, types, legend layout/styles, and chart CRUD | 🟡 | 🟡 | 🟡 | Formula-backed series not evaluated |

## Movies, images, audio, and soundtrack

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Existing movie/audio drawable inventory and metadata | 🟡 | ✅ | N/A | |
| Existing movie title and caption | 🟡 | ✅ | 🟡 | |
| Existing file-backed movie geometry and transform | 🟡 | ✅ | 🟡 | |
| Existing file-backed movie/audio playback settings | 🟡 | ✅ | 🟡 | |
| Existing slide movie/audio content and movie poster bytes | 🟡 | ✅ | 🟡 | |
| Image/movie/audio bytes and generic asset insertion/replacement/deletion | 🟡 | ✅ | 🟡 | |
| Existing soundtrack playback settings | 🟡 | ✅ | 🟡 | A soundtrack is never created |
| Existing soundtrack media-reference order | 🟡 | ✅ | 🟡 | |
| Soundtrack audio item lifecycle | 🟡 | ✅ | 🟡 | No soundtrack creation |

## Security, limits, and explicit non-goals

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ZIP/IWA/wire/semantic resource ceilings | 🟡 | ✅ | 🟡 | |
| Safe regular-file path ingress | ✅ | ✅ | N/A | |
| Encrypted/password-protected packages | ❌ | ❌ | ❌ | |
| Digital signatures and trust verification | ❌ | ❌ | ❌ | |
| External links/resources, actions, and playback targets | 🟡 | 🟡 | ❌ | Inert; never fetched or played |
| Macros, scripts, controls, and embedded code execution | ❌ | 🟡 | ❌ | |
| Rendering, layout calculation, playback, PDF/HTML export, and automation | ❌ | ❌ | ❌ | |
