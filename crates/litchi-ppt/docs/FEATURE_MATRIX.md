# PowerPoint Binary Feature Matrix (PPT)

Feature checklist. ✅ Supported · 🟡 Partial · ❌ Unsupported · N/A Not applicable.

## Additional native metadata coverage

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Photo album settings (`PhotoAlbumInfo10Atom`) | 🟡 | ✅ | ❌ |
| Presentation Advisor (`PP9DocBinaryTagExtension` preferences) | 🟡 | ✅ | 🟡 |
| Document sound collection (`SoundCollection`, `Sound`, `SoundData`) | 🟡 | ✅ | ✅ |
| WMF recoloring (`RecolorInfo`) | 🟡 | ✅ | 🟡 |
| Custom table-style record (`RoundTripCustomTableStyles12Atom`) | 🟡 | ✅ | 🟡 |

## Package, records, and presentation structure

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Open, create, and save | ✅ | ✅ | ✅ |
| CFB streams, storages, and property sets | ✅ | ✅ | ✅ |
| PowerPoint document stream and persist mapping | ✅ | ✅ | ✅ |
| Slide list, ordering, IDs, and visibility | ✅ | ✅ | ✅ |
| Main masters and master text styles | ✅ | ✅ | ✅ |
| Title, notes, and handout master records | 🟡 | ✅ | 🟡 |
| Presentation metadata and summary information | ✅ | ✅ | ✅ |
| View state, guides, and print settings | ✅ | ✅ | ✅ |

## Slides, text, drawing, and themes

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Text boxes, paragraphs, runs, and formatting | ✅ | ✅ | ✅ |
| Text extensions and outline metadata | ✅ | ✅ | ✅ |
| Picture bullets | ✅ | ✅ | ✅ |
| OfficeArt shapes and groups | ✅ | ✅ | ✅ |
| Shape-scoped OfficeArtClientData | 🟡 | ✅ | ✅ |
| Placeholders and inherited layout hints | ✅ | ✅ | ✅ |
| Pictures and BLIP resources | ✅ | ✅ | ✅ |
| Tables | ✅ | ✅ | ✅ |
| Legacy color schemes and theme-like round-trip metadata | 🟡 | ✅ | ✅ |
| Native Graph and Excel charts | 🟡 | ✅ | 🟡 |
| Diagram and SmartArt objects | 🟡 | ✅ | 🟡 |
| Diagram build records | 🟡 | ✅ | ✅ |

## Interaction, notes, review, and slide show

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Hyperlinks and slide navigation | ✅ | ✅ | ✅ |
| Action and interaction settings | 🟡 | ✅ | 🟡 |
| Speaker notes | ✅ | ✅ | ✅ |
| Classic comments | ✅ | ✅ | ✅ |
| Modern comments | ❌ | ❌ | ❌ |
| Animations and timing trees | ✅ | ✅ | ✅ |
| Transitions and slide advance timing | ✅ | ✅ | ✅ |
| Named/custom slide shows | ✅ | ✅ | ✅ |
| Headers and footers | ✅ | ✅ | ✅ |

## Media and embedded content

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Audio and video external-object metadata | 🟡 | ✅ | 🟡 |
| Media playback, rendering, and external activation | ❌ | N/A | N/A |
| Embedded OLE objects and ActiveX controls | ✅ | ✅ | ✅ |
| Custom XML data storage | ✅ | ✅ | ✅ |
| Embedded font records and payload metadata | 🟡 | ✅ | 🟡 |

## Metadata extensions, security, and macros

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Smart tags | ✅ | ✅ | ✅ |
| Programmable tags | ✅ | ✅ | ✅ |
| Slide-library synchronization and round-trip records | ✅ | ✅ | ✅ |
| Broadcast, HTML-publish, routing-slip, envelope, and privacy metadata | 🟡 | ✅ | ✅ |
| Document comparison and review diff metadata | 🟡 | ✅ | 🟡 |
| Modify-password and protection metadata | ✅ | ✅ | ✅ |
| Password encryption | ✅ | ✅ | ✅ |
| VBA projects | ✅ | ✅ | ✅ |
| Digital signatures | ✅ | ✅ | ✅ |

## Important boundaries and explicit gaps

| Feature family exposed by the specifications | Status | Read | Write |
| --- | --- | --- | --- |
| General title-, notes-, and handout-master authoring | 🟡 | ✅ | 🟡 |
| Complete native chart authoring | 🟡 | ✅ | 🟡 |
| Native diagram or SmartArt authoring | ❌ | ❌ | ❌ |
| Full OfficeArt record and vendor-extension semantics | 🟡 | 🟡 | 🟡 |
| External resource resolution and OLE/macro/media execution | ❌ | N/A | N/A |
| Rendering and slide-show playback | ❌ | N/A | N/A |
