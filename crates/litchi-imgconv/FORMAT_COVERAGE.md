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
  `noncanonical-*` diagnostic, which the strict helpers also accept;
- input whose picture cannot be reproduced exactly, such as a record stream
  that stops at a corrupt record, text given as font glyph indices, or text
  whose character set cannot be decoded exactly, is still converted and
  reported with a diagnostic that strict helpers reject.

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
destination-dependent raster operations, unsupported bitmap masks,
device-dependent bitmaps, or non-representable clip combinations.

EMF+ records play only when an `EmfPlusHeader` is the first record after the
EMF header, as in GDI+; otherwise the metafile plays as classic EMF. A dual
metafile plays its EMF+ records and falls back to its classic records when the
EMF+ records fail or need an approximation; a dual metafile containing
`GetDC` plays its classic records, which hold the complete picture. An
EMF+-only metafile containing `GetDC` is rejected, because it needs its
classic and EMF+ records interleaved in one drawing order. Legal but inexact
EMF+ features such as complex brushes, some clip combinations, effects, or
curve approximations produce stable public diagnostics rather than
disappearing silently.

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

## Text and pens

- Characters of `ExtTextOut` records are placed at the origins their advance
  arrays give, including vertical advances and double-byte characters, and
  aligned as GDI aligns the whole string.
- Font heights and pen widths in logical units follow the world transform
  and page mapping; zero-width and cosmetic pens stay one device pixel wide.
- Symbol faces (Symbol, MT Extra, Wingdings, Wingdings 3, Webdings, Marlett,
  Monotype Sorts, Zapf Dingbats) draw the Unicode characters their glyphs
  represent where one exists; other glyphs keep their private-use code point.
- `font-family` names the recorded face, then the generic family of its
  LOGFONT family class.
- PNG and JPEG output draws missing Windows faces with installed
  metric-compatible families and symbols with installed symbol faces.
- A raw WMF without a placeable header rasterizes at the size of its SVG
  canvas, which keeps the aspect ratio of its frame.

## Macintosh PICT

PICT version 1, version 2 and extended version 2 pictures are read with or
without the 512-byte file header, following *Inside Macintosh: Imaging With
QuickDraw*, Appendix A. Every opcode is framed by its defined length, so
reserved opcodes are skipped as QuickDraw skips them. Playback emits SVG:

- lines, rectangles, rounded rectangles, ovals, arcs, polygons and regions in
  their frame, paint, erase and fill forms, drawn with the pen size, the
  pen, fill and background patterns, and the port colours;
- one-bit, dither and colour pixel patterns;
- text with font, face, size, ratio and spacing, decoded from Mac OS Roman
  (and from the Symbol font encoding);
- clipping regions and origin changes;
- `BitsRect`, `BitsRgn`, `PackBitsRect`, `PackBitsRgn`, `DirectBitsRect` and
  `DirectBitsRgn` with 1/2/4/8-bit indexed and 16/32-bit direct pixels, every
  packing type, source cropping, scaling, mask regions and the copy, or,
  bit-clear and transparent transfer modes;
- QuickTime compressed images whose codec is JPEG, PNG, GIF, TIFF, BMP,
  uncompressed (`raw `), a Windows DIB or a Windows metafile, placed through
  their matrix and mask region; the drawing QuickTime pictures carry for
  readers without QuickTime is skipped as QuickTime skips it.

The native resolution of an extended version 2 picture sets its size, and
pens keep the one-pixel minimum `DrawPicture` gives them. Video codecs,
mattes, exclusive-or and inverting drawing, and the remaining arithmetic
transfer modes are drawn approximately and reported through
`ConversionReport::diagnostics`; the strict helpers (`pict::convert_pict*`,
and `to_png`, `to_svg` and `convert` for PICT BLIPs) reject such pictures
instead of returning a placeholder image. Input QuickDraw tolerates, such as
a missing end-of-picture opcode or trailing bytes, is reported with
`noncanonical-*` diagnostics.

## Resource and safety model

Record streams are walked the way GDI playback walks them, so declared sizes
and counts that disagree with the records are reported, not trusted. Caller limits cover encoded and
decoded bytes, dimensions, pixels, records, retained objects, saved graphics
states, path points, SVG elements, and encoded output. Embedded image decoders
receive the same dimension and allocation ceilings. SVG emission accepts only
typed markup, replaces XML 1.0-forbidden controls, ignores external resources,
and never executes metafile escape or driver payloads.
