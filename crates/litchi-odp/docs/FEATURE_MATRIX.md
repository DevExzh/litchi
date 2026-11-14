# ODP Feature Matrix

Feature checklist for packaged OpenDocument presentations. ✅ Supported · 🟡 Partial ·
❌ Unsupported · N/A Not applicable. Scripts, actions, media and embedded payloads are never
executed; external targets are never fetched.

## Package and shared ODF features

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Open, create, and save ODP packages | ✅ | ✅ | ✅ | OTP templates not supported |
| Password-protected ODP opening | ✅ | ✅ | N/A | No encryption authoring |
| ZIP package, manifest, and MIME validation | ✅ | ✅ | ✅ | |
| `content.xml`, `styles.xml`, `meta.xml`, and `settings.xml` | ✅ | ✅ | ✅ | |
| ODF metadata and statistics | ✅ | ✅ | ✅ | |
| Common styles and data styles | 🟡 | 🟡 | 🟡 | No complete style resolver |
| Images and package media | ✅ | ✅ | ✅ | |
| Embedded objects and OLE-like payloads | 🟡 | ✅ | 🟡 | Inert inventory only |
| Embedded charts | 🟡 | ✅ | 🟡 | No chart rendering |
| Annotations/comments | ✅ | ✅ | ✅ | |
| Hyperlinks and external references | ✅ | ✅ | ✅ | Targets never followed |
| Forms and controls | ✅ | ✅ | ✅ | Inert; never executed |
| Scripts, events, and macros | 🟡 | 🟡 | 🟡 | Inert metadata only |
| RDF metadata graphs | ✅ | ✅ | ✅ | |
| ODF encryption authoring | 🟡 | ✅ | ❌ | Encrypted packages are read-only |
| ODF digital signatures | 🟡 | 🟡 | ❌ | Signed packages are read-only; no verification or signing |
| Unknown package-part preservation | 🟡 | ✅ | 🟡 | Preserved, not understood |

## Presentation content and layout

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Slides, ordering, and slide count | ✅ | ✅ | ✅ | |
| Slide titles and body text | ✅ | ✅ | ✅ | |
| Rich text runs, paragraphs, and lists | ✅ | ✅ | ✅ | |
| Text boxes and basic shapes | ✅ | ✅ | ✅ | |
| Advanced ODF drawing shapes | 🟡 | 🟡 | 🟡 | Geometry retained, not modeled |
| Shape stacking, transforms, and anchors | ✅ | ✅ | ✅ | |
| Images in slide frames | ✅ | ✅ | ✅ | |
| Audio and video plugins | ✅ | ✅ | ✅ | No playback |
| Tables | ✅ | ✅ | ✅ | |
| Charts and chart data ranges | 🟡 | ✅ | 🟡 | No chart rendering |
| Presentation page layouts | ✅ | ✅ | ✅ | |
| Slide page metadata and layout references | ✅ | ✅ | ✅ | |
| Drawing layer declarations | ✅ | ✅ | ✅ | |
| Master pages | ✅ | ✅ | ✅ | |
| Handout master | ✅ | ✅ | ✅ | |
| Headers, footers, and date-time declarations | ✅ | ✅ | ✅ | Fields not expanded |
| Backgrounds and named drawing resources | 🟡 | 🟡 | 🟡 | No renderer |
| Speaker notes | ✅ | ✅ | ✅ | Text only, no notes-layout editor |
| Animations and SMIL timing trees | ✅ | ✅ | ✅ | No playback |
| Slide transitions and automatic timing | ✅ | ✅ | ✅ | No playback |
| Presentation settings and custom slide shows | ✅ | ✅ | ✅ | |
| Presentation actions and event bindings | ✅ | ✅ | ✅ | Never activated |
| Protection and edit restrictions | ❌ | ❌ | ❌ | |

## Explicit gaps from the audited feature families

| Feature family | Status | Read | Write | Notes |
|---|---|---|---|---|
| Full rich presentation authoring | ✅ | ✅ | ✅ | |
| Rendering, layout, and playback | ❌ | ❌ | ❌ | |
| PPTX sections and zoom objects | ❌ | ❌ | ❌ | |
| PPTX vendor transitions and design extensions | ❌ | ❌ | ❌ | |
| PPTX threaded comments, presence, and collaboration commands | ❌ | ❌ | ❌ | |
| PPTX media extensions | ❌ | ❌ | ❌ | |
| PPTX guides and application-specific UI state | ❌ | ❌ | ❌ | |
| Full chart/table/form object models | 🟡 | ✅ | 🟡 | No chart rendering or formula recalculation |
| Macro/script execution and external fetching | ❌ | ❌ | ❌ | |
| Signature writing and security policy | 🟡 | ✅ | ❌ | No verification or signing |
