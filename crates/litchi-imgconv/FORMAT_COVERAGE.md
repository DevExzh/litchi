# Metafile coverage and fidelity policy

This crate classifies every record identifier defined by the Microsoft
`[MS-EMF]`, `[MS-WMF]`, and `[MS-EMFPLUS]` Open Specifications. A recognized record is never discarded merely because a
match arm is missing.

The output policy is deliberately stricter than best-effort metafile viewers:

- drawing that has a faithful SVG representation stays vector;
- DIB or encoded bitmap sources are embedded as bounded PNG/JPEG data;
- PNG/JPEG output rasterizes the resulting self-contained SVG;
- metadata and query records are validated and produce no drawing;
- operations that depend on a live destination surface, execute driver data,
  or otherwise have no faithful safe SVG representation return
  `Error::Unsupported`;
- bounded EMF+ approximations are reported through
  `ConversionReport::diagnostics`; payload-only convenience helpers and the
  lower-level strict `EmfSvgConverter::convert` reject such diagnostics;
- input that departs from the specification in a way Windows GDI playback
  tolerates without changing the picture is accepted and reported with a
  `noncanonical-*` diagnostic, which the strict helpers also accept.

## Record catalogs

| Family | Specification identifiers | Classification |
|---|---:|---|
| Classic EMF | 119 | All explicitly matched: 89 semantic/state/drawing routes, 25 validated non-rendering routes, 5 unconditional `Unsupported` routes |
| WMF | 70 | All explicitly matched: 65 state/drawing/error routes and 5 validated control/no-render routes |
| EMF+ | 58 | All explicitly matched; the three reserved MultiFormat records are rejected/diagnosed |

The five classic EMF records that are unconditionally rejected are
`EMR_WIDENPATH`, `EMR_EXTFLOODFILL`, `EMR_DRAWESCAPE`, `EMR_EXTESCAPE`, and
`EMR_NAMEDESCAPE`. Additional records can be rejected conditionally when their
specific flags or payload require semantics SVG cannot reproduce—for example,
destination-dependent raster operations, unsupported bitmap masks, glyph-index
text, device-dependent bitmaps, or non-representable clip combinations.

EMF+ `GetDC` requires ordered interleaving with classic EMF records. Both
EMF+-only and dual streams containing it are rejected until the two playback
engines share one ordered drawing sink; accepting them would silently reorder
paint operations. Legal but inexact EMF+ features such as complex brushes,
some clip combinations, effects, or curve approximations produce stable public
diagnostics rather than disappearing silently.

## Bitmap coverage

The shared DIB decoder handles CORE, INFO, V4, and V5 headers; indexed palettes;
1/4/8/16/24/32-bit pixels; RGB, RLE4, RLE8, and bitfields; embedded PNG/JPEG;
top-down images where allowed; alpha policy; crop/mirror/stretch operations;
`AlphaBlend`; and `TransparentBlt`. Every EMF and WMF bitmap record family is
routed through the shared typed normalization layer. Unsupported masks,
device-dependent `Bitmap16` sources, and CMYK decode fail explicitly.

## Raster operations and brushes

- Solid, hatch, monochrome pattern and DIB pattern brushes, tiled in device
  pixels.
- ROP2 mix modes whose result for the pen or brush colour does not read the
  destination; other mix modes fail when a record paints with them.
- XOR brackets (a pass of XOR fills, one fill, and the same XOR pass again),
  painted as the fills clipped to the middle fill.
- ROP3 blits: constants, brush fills and source copies directly; consecutive
  blits on one pixel grid, such as mask and image pairs, composed into one
  image. Pixels that still depend on the destination are painted as on a white
  destination and reported as an approximation.

## Resource and safety model

Strict framing checks run before playback. Caller limits cover encoded and
decoded bytes, dimensions, pixels, records, retained objects, saved graphics
states, path points, SVG elements, and encoded output. Embedded image decoders
receive the same dimension and allocation ceilings. SVG emission accepts only
typed markup, replaces XML 1.0-forbidden controls, ignores external resources,
and never executes metafile escape or driver payloads.
