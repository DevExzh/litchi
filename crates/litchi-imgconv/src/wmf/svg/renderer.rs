//! Stateful, bounded WMF record playback to SVG.

use std::{collections::HashMap, mem::size_of, sync::Arc};

use super::super::constants::{bk_mode, brush, ext_text_out, pen, record, stock, text_align};
use super::super::parser::WmfRecord;
use super::BitmapHook;
use super::state::{
    Brush, ClipRegion, DeviceRect, Font, GdiObject, GraphicsState, MappingState, ObjectTable,
    Palette, Pen, Region,
};
use super::style::{
    color_paint, escape_xml_attr_into, escape_xml_text_into, fill_attrs, hatch_definition,
    stroke_attrs,
};
use super::transform::CoordinateTransform;
use crate::svg_utils::{write_color_hex, write_num};
use crate::{
    dib::{ColorUsage, Dib, DibLimits},
    metafile_bitmap::{
        BitmapOp, BitmapPaint, BitmapSource, BlitGrid, BrushPattern, RasterComposite, Rop2,
        Rop2Paint, Rop3, StretchPolicy, convex_outlines_overlap, is_convex, xor_bracket_paint,
    },
};
use bytes::Bytes;
use litchi_core::binary::{read_i16_le, read_u16_le};

#[derive(Debug, Clone, Copy)]
enum ArcType {
    Open,
    Pie,
    Chord,
}

#[derive(Debug, Clone)]
pub(super) struct RenderIssue {
    pub fatal: bool,
    /// Stable diagnostic code reported for non-fatal issues.
    pub code: &'static str,
    pub message: String,
}

/// Diagnostic code of non-fatal playback issues that approximate the source.
pub(super) const PLAYBACK_WARNING: &str = "wmf-playback-warning";

/// Cache key for an emitted hatch pattern: hatch colour, hatch style and the
/// background colour, or `None` for a transparent background.
type HatchKey = (u32, u16, Option<u32>);

/// A bitmap record whose painting waits for the records after it, so that
/// raster operations which combine with each other's pixels are composed
/// into one image.
#[derive(Debug, Clone)]
struct DeferredBlit {
    /// The bitmap record, or `None` for PatBlt.
    record: Option<DeferredRecord>,
    rop3: Rop3,
    /// Mirroring of the destination by the mapping mode.
    mirrored: (bool, bool),
    /// SVG destination rectangle `[x, y, width, height]`.
    destination: [f64; 4],
    /// Solid brush `COLORREF`, or `None` for other brushes.
    brush: Option<u32>,
}

/// What is needed to parse a deferred bitmap record again.
#[derive(Debug, Clone)]
struct DeferredRecord {
    size: u32,
    function: u16,
    params: Bytes,
    stretch: StretchPolicy,
}

/// Stands in for the fill of a shape whose paint is known only once its XOR
/// bracket is complete.
const XOR_FILL: &str = "@xor-fill@";

/// Stands in for the stroke of a middle fill's outline.
const XOR_STROKE: &str = "@xor-stroke@";

/// Most XOR fills whose pixels are composed for a middle fill with an
/// outline; each one repeats the middle fills.
const MAX_STROKED_BRACKET_FILLS: usize = 64;

/// Most fills between the two XOR passes.
const MAX_BRACKET_MIDDLES: usize = 256;

/// A fill painted in an XOR mix mode, waiting for the pass that undoes it.
#[derive(Debug, Clone)]
struct XorShape {
    /// The element, with [`XOR_FILL`] as its fill.
    element: String,
    /// The colour the fill XORs into the destination.
    color: u32,
    /// Convex outline in logical coordinates, when the record has one.
    outline: Option<Vec<(i64, i64)>>,
}

/// How the fills between the two XOR passes paint.
#[derive(Debug, Clone)]
enum XorMiddle {
    Solid {
        rop2: Rop2,
        color: u32,
    },
    Pattern {
        rop2: Rop2,
        pattern: Arc<BrushPattern>,
        text_color: u32,
        background_color: u32,
    },
}

/// A fill between the two XOR passes.
#[derive(Debug, Clone)]
struct MiddleShape {
    /// The element, with [`XOR_FILL`] as its fill and [`XOR_STROKE`] as its
    /// stroke when it has an outline.
    element: String,
    /// Mix mode and `COLORREF` of its outline, when it has one.
    stroke: Option<(Rop2, u32)>,
    /// Logical bounds `[left, top, right, bottom]` of its geometry, when the
    /// record has them.
    bounds: Option<[i64; 4]>,
}

impl XorMiddle {
    fn rop2(&self) -> Rop2 {
        match self {
            Self::Solid { rop2, .. } | Self::Pattern { rop2, .. } => *rop2,
        }
    }

    /// Whether two fills paint the same pixels the same way.
    fn same_paint(&self, other: &Self) -> bool {
        match (self, other) {
            (
                Self::Solid { rop2, color },
                Self::Solid {
                    rop2: other_rop2,
                    color: other_color,
                },
            ) => rop2 == other_rop2 && color == other_color,
            (
                Self::Pattern {
                    rop2,
                    pattern,
                    text_color,
                    background_color,
                },
                Self::Pattern {
                    rop2: other_rop2,
                    pattern: other_pattern,
                    text_color: other_text,
                    background_color: other_background,
                },
            ) => {
                rop2 == other_rop2
                    && Arc::ptr_eq(pattern, other_pattern)
                    && text_color == other_text
                    && background_color == other_background
            },
            _ => false,
        }
    }
}

/// Fills XORed into the destination, the fills drawn after them, and the
/// pass that XORs the same fills again.
///
/// The second pass cancels the first wherever the middle fill did not
/// paint. Inside the middle fills every pixel becomes `m(D ^ a) ^ a`, where
/// `a` is the colour of the XOR fill covering it and `m` the paint of the
/// middle fills. Producers use this to clip gradients and patterns to a
/// shape.
#[derive(Debug, Clone)]
struct XorBracket {
    mapping: MappingState,
    first: Vec<XorShape>,
    /// The paint the middle fills share.
    paint: Option<XorMiddle>,
    middles: Vec<MiddleShape>,
    /// Shapes of the second pass matched so far.
    matched: usize,
}

/// What the current record's shape contributes to an XOR bracket.
#[derive(Debug, Clone)]
enum BracketShape {
    Xor(u32),
    Middle(XorMiddle, Option<(Rop2, u32)>),
}

/// Consecutive deferred bitmap records that share one clip and pixel grid.
#[derive(Debug, Clone)]
struct DeferredBlits {
    clip: String,
    grid: BlitGrid,
    blits: Vec<DeferredBlit>,
    /// At least one record combines with destination pixels.
    composite: bool,
}

pub(super) struct SvgRenderer<'hook> {
    transform: CoordinateTransform,
    state: GraphicsState,
    saved_states: Vec<GraphicsState>,
    objects: ObjectTable,
    definitions: String,
    hatch_ids: HashMap<HatchKey, String>,
    /// Emitted brush patterns by their PNG tile.
    pattern_ids: HashMap<Vec<u8>, String>,
    deferred_blits: Option<DeferredBlits>,
    /// The current record's mix mode leaves every pixel it would paint
    /// unchanged.
    paints_nothing: bool,
    xor_bracket: Option<XorBracket>,
    /// The current record's shape belongs to an XOR bracket.
    bracket_shape: Option<BracketShape>,
    /// Clip reference per clip revision; `Some(None)` records a clip that covers
    /// the whole canvas and therefore needs no reference at all.
    clip_ids: HashMap<u64, Option<String>>,
    next_clip_revision: u64,
    next_definition_id: usize,
    issues: Vec<RenderIssue>,
    bitmap_hook: Option<&'hook BitmapHook>,
    dib_limits: DibLimits,
    max_state_depth: usize,
    max_objects: usize,
    max_path_points: usize,
    max_output_bytes: usize,
    max_object_bytes: usize,
    halted: bool,
    remaining_work: usize,
}

impl<'hook> SvgRenderer<'hook> {
    /// Starts playback under `mapping` instead of the default device context
    /// mapping.
    pub(super) fn set_initial_mapping(&mut self, mapping: MappingState) {
        self.state.mapping = mapping;
    }

    pub(super) fn new(
        transform: CoordinateTransform,
        bitmap_hook: Option<&'hook BitmapHook>,
        dib_limits: DibLimits,
        max_state_depth: usize,
        max_objects: usize,
        max_path_points: usize,
        max_output_bytes: usize,
        max_object_bytes: usize,
    ) -> Self {
        Self {
            transform,
            state: GraphicsState::default(),
            saved_states: Vec::new(),
            objects: ObjectTable::default(),
            definitions: String::new(),
            hatch_ids: HashMap::new(),
            pattern_ids: HashMap::new(),
            deferred_blits: None,
            paints_nothing: false,
            xor_bracket: None,
            bracket_shape: None,
            clip_ids: HashMap::new(),
            next_clip_revision: 1,
            next_definition_id: 0,
            issues: Vec::new(),
            bitmap_hook,
            dib_limits,
            max_state_depth,
            max_objects,
            max_path_points,
            max_output_bytes,
            max_object_bytes,
            halted: false,
            remaining_work: max_path_points,
        }
    }

    pub(super) fn into_parts(self) -> (String, Vec<RenderIssue>) {
        (self.definitions, self.issues)
    }

    /// Paints the bitmap records still waiting at the end of playback.
    pub(super) fn finish(&mut self) -> Option<String> {
        if self.halted {
            return None;
        }
        if self.xor_bracket.is_some() {
            self.fail_xor_bracket();
            return None;
        }
        self.flush_blits()
    }

    pub(super) fn render_record(&mut self, rec: &WmfRecord) -> Option<String> {
        if self.halted {
            return None;
        }
        let function = record::canonical(rec.function);
        self.paints_nothing = false;
        self.bracket_shape = None;
        if self.xor_bracket.is_some() && !bracket_compatible(rec) {
            self.fail_xor_bracket();
            return None;
        }
        let included = self.state.clip.rects.as_ref().map_or(1usize, Vec::len);
        let clip_work = included.checked_mul(self.state.clip.excluded.len().saturating_add(1));
        if clip_work.is_none_or(|work| work > self.max_path_points) {
            self.fatal(format!(
                "WMF clip complexity exceeds limit {}",
                self.max_path_points
            ));
            return None;
        }
        if let Some(points) = record_point_count(function, &rec.params)
            && points > self.max_path_points
        {
            self.fatal(format!(
                "WMF record has {points} points; limit is {}",
                self.max_path_points
            ));
            return None;
        }
        let result = match function {
            // Drawing records whose ordering relative to state updates matters.
            record::LINE_TO => self.render_line_to(rec),
            record::TEXT_OUT => self.render_text_out(rec),
            record::EXT_TEXT_OUT => self.render_ext_text_out(rec),

            record::RECTANGLE => self.render_rectangle(rec),
            record::ROUND_RECT => self.render_round_rect(rec),
            record::ELLIPSE => self.render_ellipse(rec),
            record::ARC => self.render_arc_common(rec, ArcType::Open),
            record::PIE => self.render_arc_common(rec, ArcType::Pie),
            record::CHORD => self.render_arc_common(rec, ArcType::Chord),
            record::POLYGON => self.render_polygon(rec, true),
            record::POLYLINE => self.render_polygon(rec, false),
            record::POLYPOLYGON => self.render_polypolygon(rec),
            record::SET_PIXEL => self.render_set_pixel(rec),
            record::PAT_BLT => return self.render_pat_blt(rec),
            record::FLOOD_FILL | record::EXT_FLOOD_FILL => self.render_flood_fill(rec),
            record::FILL_REGION
            | record::FRAME_REGION
            | record::INVERT_REGION
            | record::PAINT_REGION => self.render_region_record(function, rec),
            // Bitmap records paint the deferred records themselves.
            record::BIT_BLT
            | record::STRETCH_BLT
            | record::DIB_BIT_BLT
            | record::DIB_STRETCH_BLT
            | record::SET_DIB_TO_DEV
            | record::STRETCH_DIB => return self.render_bitmap(function, rec),

            // Object table records.
            record::CREATE_PEN_INDIRECT => {
                self.create_pen(rec);
                None
            },
            record::CREATE_BRUSH_INDIRECT => {
                self.create_brush(rec);
                None
            },
            record::CREATE_FONT_INDIRECT => {
                self.create_font(rec);
                None
            },
            record::CREATE_PALETTE => {
                self.create_palette(rec);
                None
            },
            record::CREATE_REGION => {
                self.create_region(rec);
                None
            },
            record::CREATE_PATTERN_BRUSH | record::CREATE_DIB_PATTERN_BRUSH => {
                self.create_pattern_brush(function, rec);
                None
            },
            record::SELECT_OBJECT => {
                self.select_object(rec);
                None
            },
            record::DELETE_OBJECT => {
                self.delete_object(rec);
                None
            },
            record::SELECT_PALETTE => {
                self.select_palette(rec);
                None
            },
            record::ANIMATE_PALETTE | record::SET_PALETTE_ENTRIES => {
                self.update_palette(rec);
                None
            },
            record::RESIZE_PALETTE => {
                self.resize_palette(rec);
                None
            },

            // Device context, mapping and clipping records.
            record::SAVE_DC => {
                if self.saved_states.len() >= self.max_state_depth {
                    self.fatal(format!(
                        "WMF saved DC depth exceeds limit {}",
                        self.max_state_depth
                    ));
                } else if !self.can_retain_additional_state(self.state.retained_heap_bytes()) {
                    self.fatal(format!(
                        "WMF saved graphics-state data exceeds limit {} bytes",
                        self.max_object_bytes
                    ));
                } else {
                    self.saved_states.push(self.state.clone());
                }
                None
            },
            record::RESTORE_DC => {
                self.restore_dc(rec);
                None
            },
            record::MOVE_TO => {
                if let Some((x, y)) = read_yx(&rec.params) {
                    self.state.position = (x, y);
                } else {
                    self.malformed(function, "point");
                }
                None
            },
            record::SET_BK_COLOR => {
                self.set_color(function, rec, true);
                None
            },
            record::SET_TEXT_COLOR => {
                self.set_color(function, rec, false);
                None
            },
            record::SET_BK_MODE => {
                self.set_u16(function, rec, |state, value| state.bk_mode = value);
                None
            },
            record::SET_POLY_FILL_MODE => {
                self.set_u16(function, rec, |state, value| state.poly_fill_mode = value);
                None
            },
            record::SET_TEXT_ALIGN => {
                self.set_u16(function, rec, |state, value| state.text_align = value);
                None
            },
            record::SET_ROP2 => {
                self.set_rop2(rec);
                None
            },
            record::SET_STRETCH_BLT_MODE => {
                self.set_u16(function, rec, |state, value| state.stretch_mode = value);
                None
            },
            record::SET_TEXT_CHAR_EXTRA => {
                if let Ok(value) = read_i16_le(&rec.params, 0) {
                    self.state.text_char_extra = value;
                } else {
                    self.malformed(function, "character spacing");
                }
                None
            },
            record::SET_TEXT_JUSTIFICATION => {
                if rec.params.len() >= 4 {
                    self.state.break_count = read_i16_le(&rec.params, 0).unwrap_or(0);
                    self.state.break_extra = read_i16_le(&rec.params, 2).unwrap_or(0);
                } else {
                    self.malformed(function, "justification");
                }
                None
            },
            record::SET_MAP_MODE => {
                if let Ok(value) = read_u16_le(&rec.params, 0) {
                    self.state.mapping.set_mode(value);
                } else {
                    self.malformed(function, "mapping mode");
                }
                None
            },
            record::SET_LAYOUT => {
                if let Ok(value) = read_u16_le(&rec.params, 0) {
                    self.state.mapping.layout = value;
                } else {
                    self.malformed(function, "layout");
                }
                None
            },
            record::SET_WINDOW_ORG
            | record::SET_VIEWPORT_ORG
            | record::OFFSET_WINDOW_ORG
            | record::OFFSET_VIEWPORT_ORG => {
                self.update_origin(function, rec);
                None
            },
            record::SET_WINDOW_EXT | record::SET_VIEWPORT_EXT => {
                self.update_extent(function, rec);
                None
            },
            record::SCALE_WINDOW_EXT | record::SCALE_VIEWPORT_EXT => {
                self.scale_extent(function, rec);
                None
            },
            record::INTERSECT_CLIP_RECT | record::EXCLUDE_CLIP_RECT => {
                self.update_clip_rect(function, rec);
                None
            },
            record::OFFSET_CLIP_RGN => {
                self.offset_clip(rec);
                None
            },
            record::SELECT_CLIP_REGION => {
                self.select_clip_region(rec);
                None
            },

            // Explicit no-op/query/control records. SET_REL_ABS is undefined
            // by [MS-WMF] and MUST be ignored. Escape data (including
            // PostScript) remains opaque and is never executed or emitted.
            record::EOF
            | record::REALIZE_PALETTE
            | record::SET_MAPPER_FLAGS
            | record::SET_REL_ABS
            | record::ESCAPE => None,

            // PlayMetaFile skips a function number it does not know.
            _ => {
                self.note(
                    "noncanonical-wmf-unknown-record",
                    format!("ignored unknown WMF record 0x{:04X}", rec.function),
                );
                None
            },
        };
        if self.objects.retained_len() > self.max_objects {
            self.fatal(format!(
                "WMF retained object count exceeds limit {}",
                self.max_objects
            ));
        }
        let result = match (self.bracket_shape.take(), result) {
            (Some(shape), Some(element)) => {
                let outline = logical_outline(function, &rec.params);
                let bounds = logical_bounds(function, &rec.params);
                self.capture_bracket_shape(shape, element, outline, bounds)
            },
            (None, Some(_)) if self.xor_bracket.is_some() && !self.paints_nothing => {
                self.fail_xor_bracket();
                None
            },
            (_, result) => result,
        };
        if self.paints_nothing || self.halted {
            return None;
        }
        // Deferred bitmap records paint before anything drawn after them.
        match result {
            Some(element) if self.deferred_blits.is_some() => {
                let mut output = self.flush_blits().unwrap_or_default();
                output.push_str(&element);
                Some(output)
            },
            other => other,
        }
    }

    fn set_u16(&mut self, function: u16, rec: &WmfRecord, setter: fn(&mut GraphicsState, u16)) {
        if let Ok(value) = read_u16_le(&rec.params, 0) {
            setter(&mut self.state, value);
        } else {
            self.malformed(function, "state value");
        }
    }

    fn set_rop2(&mut self, rec: &WmfRecord) {
        let Ok(mode) = read_u16_le(&rec.params, 0) else {
            self.malformed(record::SET_ROP2, "mix mode");
            return;
        };
        if Rop2::from_mode(u32::from(mode)).is_some() {
            self.state.rop2 = mode;
        } else {
            // SetROP2 fails for an undefined mode and keeps the current one.
            self.note(
                "noncanonical-wmf-invalid-rop2",
                format!("SetROP2 ignores undefined mix mode {mode}"),
            );
        }
    }

    fn set_color(&mut self, function: u16, rec: &WmfRecord, background: bool) {
        if let Some(color) = read_u32(&rec.params, 0) {
            if background {
                self.state.bk_color = color;
            } else {
                self.state.text_color = color;
            }
        } else {
            self.malformed(function, "COLORREF");
        }
    }

    fn update_origin(&mut self, function: u16, rec: &WmfRecord) {
        let Some((x, y)) = read_yx(&rec.params) else {
            self.malformed(function, "origin");
            return;
        };
        let pair = (f64::from(x), f64::from(y));
        match function {
            record::SET_WINDOW_ORG => self.state.mapping.window_origin = pair,
            record::SET_VIEWPORT_ORG => self.state.mapping.viewport_origin = pair,
            record::OFFSET_WINDOW_ORG => {
                self.state.mapping.window_origin.0 += pair.0;
                self.state.mapping.window_origin.1 += pair.1;
            },
            record::OFFSET_VIEWPORT_ORG => {
                self.state.mapping.viewport_origin.0 += pair.0;
                self.state.mapping.viewport_origin.1 += pair.1;
            },
            _ => {},
        }
    }

    fn update_extent(&mut self, function: u16, rec: &WmfRecord) {
        let Some((x, y)) = read_yx(&rec.params) else {
            self.malformed(function, "extent");
            return;
        };
        if !self.state.mapping.scalable_extents() {
            return;
        }
        let pair = (f64::from(x), f64::from(y));
        if function == record::SET_WINDOW_EXT {
            self.state.mapping.window_extent = pair;
        } else {
            self.state.mapping.viewport_extent = pair;
        }
    }

    fn scale_extent(&mut self, function: u16, rec: &WmfRecord) {
        if rec.params.len() < 8 {
            self.malformed(function, "extent scale");
            return;
        }
        if !self.state.mapping.scalable_extents() {
            return;
        }
        let y_den = f64::from(read_i16_le(&rec.params, 0).unwrap_or(0));
        let y_num = f64::from(read_i16_le(&rec.params, 2).unwrap_or(0));
        let x_den = f64::from(read_i16_le(&rec.params, 4).unwrap_or(0));
        let x_num = f64::from(read_i16_le(&rec.params, 6).unwrap_or(0));
        if x_den == 0.0 || y_den == 0.0 {
            self.warn(format!(
                "ignored zero divisor in WMF record 0x{function:04X}"
            ));
            return;
        }
        let extent = if function == record::SCALE_WINDOW_EXT {
            &mut self.state.mapping.window_extent
        } else {
            &mut self.state.mapping.viewport_extent
        };
        extent.0 *= x_num / x_den;
        extent.1 *= y_num / y_den;
    }

    fn restore_dc(&mut self, rec: &WmfRecord) {
        let Ok(level) = read_i16_le(&rec.params, 0) else {
            self.malformed(record::RESTORE_DC, "restore level");
            return;
        };
        let target = if level < 0 {
            self.saved_states
                .len()
                .checked_sub(usize::from(level.unsigned_abs()))
        } else if level > 0 {
            usize::from(level as u16).checked_sub(1)
        } else {
            None
        };
        let Some(target) = target.filter(|&index| index < self.saved_states.len()) else {
            // GDI fails a RestoreDC level that names no saved state and leaves
            // the device context unchanged.
            self.note(
                "noncanonical-wmf-restore-dc-level",
                format!("ignored RestoreDC level {level}, which names no saved state"),
            );
            return;
        };
        // The target snapshot becomes the live state. Moving it avoids an
        // otherwise unaccounted duplicate of its clip/font allocations.
        self.saved_states.truncate(target + 1);
        if let Some(restored) = self.saved_states.pop() {
            self.state = restored;
        }
    }

    fn create_pen(&mut self, rec: &WmfRecord) {
        if rec.params.len() < 10 {
            self.malformed(record::CREATE_PEN_INDIRECT, "pen");
            return;
        }
        self.objects.insert(GdiObject::Pen(Pen {
            style: read_u16_le(&rec.params, 0).unwrap_or(0),
            width: (
                read_i16_le(&rec.params, 2).unwrap_or(0),
                read_i16_le(&rec.params, 4).unwrap_or(0),
            ),
            color: read_u32(&rec.params, 6).unwrap_or(0),
        }));
    }

    fn create_brush(&mut self, rec: &WmfRecord) {
        if rec.params.len() < 8 {
            self.malformed(record::CREATE_BRUSH_INDIRECT, "brush");
            return;
        }
        self.objects.insert(GdiObject::Brush(Brush {
            style: read_u16_le(&rec.params, 0).unwrap_or(brush::BS_NULL),
            color: read_u32(&rec.params, 2).unwrap_or(0),
            hatch: read_u16_le(&rec.params, 6).unwrap_or(0),
            pattern: None,
        }));
    }

    fn create_pattern_brush(&mut self, function: u16, rec: &WmfRecord) {
        let (style, pattern) = if function == record::CREATE_PATTERN_BRUSH {
            // [MS-WMF] 2.3.4.4 pads the Bitmap16 header to a legacy 32-byte
            // BITMAP structure before the bits.
            (
                brush::BS_PATTERN,
                bitmap16_pattern(&rec.params, 32, self.dib_limits),
            )
        } else {
            let (Ok(style), Ok(usage)) = (read_u16_le(&rec.params, 0), read_u16_le(&rec.params, 2))
            else {
                self.malformed(function, "pattern brush header");
                return;
            };
            let target = &rec.params[4..];
            // [MS-WMF] 2.3.4.8 pairs BS_PATTERN with a Bitmap16, but
            // generators also write a DIB there; a DIB starts with the size of
            // its header.
            let pattern = if style == brush::BS_PATTERN && !starts_with_dib_header(target) {
                bitmap16_pattern(target, 10, self.dib_limits)
            } else {
                match ColorUsage::from_raw(u32::from(usage))
                    .and_then(|usage| Dib::parse_with_usage(target, usage, self.dib_limits))
                {
                    Ok(dib) => {
                        self.note_dib_tolerances(&dib);
                        BrushPattern::from_dib(&dib).map_err(|error| error.to_string())
                    },
                    Err(error) => Err(error.to_string()),
                }
            };
            (style, pattern)
        };
        let pattern = match pattern {
            Ok(pattern) => {
                if !self.reserve_object_bytes(pattern.retained_bytes()) {
                    return;
                }
                Ok(Arc::new(pattern))
            },
            // The object keeps its table slot; painting with it fails.
            Err(error) => Err(Arc::from(error.as_str())),
        };
        self.objects.insert(GdiObject::Brush(Brush {
            style,
            color: 0,
            hatch: 0,
            pattern: Some(pattern),
        }));
    }

    fn create_font(&mut self, rec: &WmfRecord) {
        if rec.params.len() < 18 {
            self.malformed(record::CREATE_FONT_INDIRECT, "font");
            return;
        }
        let name_end = rec.params.len().min(18 + 32);
        let name_bytes = &rec.params[18..name_end];
        let name_len = name_bytes
            .iter()
            .position(|&byte| byte == 0)
            .unwrap_or(name_bytes.len());
        // LOGFONT facenames are byte strings; Latin-1 preserves every byte and
        // avoids accepting ill-formed UTF-8 into XML attributes.
        let name: String = name_bytes[..name_len]
            .iter()
            .map(|&byte| char::from(byte))
            .collect();
        // The recorded face stays first in the SVG font family so a system
        // that has it uses it; the generic class follows as a stand-in.
        let name = if name.is_empty() {
            "serif".to_owned()
        } else {
            name
        };
        if !self.reserve_object_bytes(name.len()) {
            return;
        }
        self.objects.insert(GdiObject::Font(Font {
            height: read_i16_le(&rec.params, 0).unwrap_or(-12),
            width: read_i16_le(&rec.params, 2).unwrap_or(0),
            escapement: read_i16_le(&rec.params, 4).unwrap_or(0),
            orientation: read_i16_le(&rec.params, 6).unwrap_or(0),
            weight: read_u16_le(&rec.params, 8).unwrap_or(400),
            italic: rec.params[10] != 0,
            underline: rec.params[11] != 0,
            strike_out: rec.params[12] != 0,
            charset: rec.params[13],
            pitch_and_family: rec.params[17],
            name,
        }));
    }

    fn create_palette(&mut self, rec: &WmfRecord) {
        if rec.params.len() < 4 {
            self.malformed(record::CREATE_PALETTE, "palette");
            return;
        }
        let count = usize::from(read_u16_le(&rec.params, 2).unwrap_or(0));
        let available = (rec.params.len() - 4) / 4;
        if count > available {
            self.malformed(record::CREATE_PALETTE, "palette entries");
            return;
        }
        if !self.reserve_object_bytes(count.saturating_mul(4)) {
            return;
        }
        let mut entries = Vec::new();
        if entries.try_reserve_exact(count).is_err() {
            self.fatal("failed to allocate bounded WMF palette".to_owned());
            return;
        }
        for index in 0..count {
            let offset = 4 + index * 4;
            entries.push(
                u32::from(rec.params[offset])
                    | (u32::from(rec.params[offset + 1]) << 8)
                    | (u32::from(rec.params[offset + 2]) << 16),
            );
        }
        self.objects.insert(GdiObject::Palette(Palette { entries }));
    }

    fn create_region(&mut self, rec: &WmfRecord) {
        let Some(region) = parse_region(&rec.params, self.max_path_points) else {
            self.malformed(record::CREATE_REGION, "region scans");
            return;
        };
        if !self.reserve_object_bytes(region.rects.len().saturating_mul(size_of::<DeviceRect>())) {
            return;
        }
        self.objects.insert(GdiObject::Region(region));
    }

    fn select_object(&mut self, rec: &WmfRecord) {
        let Ok(handle) = read_u16_le(&rec.params, 0) else {
            self.malformed(record::SELECT_OBJECT, "object index");
            return;
        };
        if handle & stock::FLAG != 0 {
            let stock_font_bytes = match handle & !stock::FLAG {
                stock::OEM_FIXED_FONT | stock::ANSI_FIXED_FONT | stock::SYSTEM_FIXED_FONT => 9,
                stock::ANSI_VAR_FONT
                | stock::SYSTEM_FONT
                | stock::DEVICE_DEFAULT_FONT
                | stock::DEFAULT_GUI_FONT => 5,
                _ => 0,
            };
            if stock_font_bytes != 0 && !self.can_retain_additional_state(stock_font_bytes) {
                self.fatal(format!(
                    "WMF selected stock-font data exceeds limit {} bytes",
                    self.max_object_bytes
                ));
                return;
            }
            if !self.state.select_stock(handle) {
                self.fatal(format!("unknown stock WMF object 0x{handle:04X}"));
            }
            return;
        }
        match self.objects.get(usize::from(handle)) {
            Some(GdiObject::Pen(value)) => self.state.pen = value.clone(),
            Some(GdiObject::Brush(value)) => self.state.brush = value.clone(),
            Some(GdiObject::Font(value)) => {
                if !self.can_retain_additional_state(value.name.len()) {
                    self.fatal(format!(
                        "WMF selected font data exceeds limit {} bytes",
                        self.max_object_bytes
                    ));
                    return;
                }
                self.state.font = value.clone();
            },
            // SelectObject with a region selects a copy of it as the clip
            // region.
            Some(GdiObject::Region(_)) => self.select_clip_region_object(handle, false),
            // A palette is selected with SelectPalette; SelectObject refuses it.
            Some(GdiObject::Palette(_)) => self.note(
                "noncanonical-wmf-select-palette-object",
                format!("SelectObject ignored palette object {handle}"),
            ),
            // A handle without an object makes SelectObject fail, which
            // leaves the selection unchanged.
            None => self.note(
                "noncanonical-wmf-select-missing-object",
                format!("SelectObject ignored missing object {handle}"),
            ),
        }
    }

    fn delete_object(&mut self, rec: &WmfRecord) {
        let Ok(index) = read_u16_le(&rec.params, 0) else {
            self.malformed(record::DELETE_OBJECT, "object index");
            return;
        };
        if index & stock::FLAG == 0 && !self.objects.delete(usize::from(index)) {
            // GDI fails deleting a handle that holds no object.
            self.note(
                "noncanonical-wmf-delete-missing-object",
                format!("DeleteObject ignored missing object {index}"),
            );
        }
    }

    fn select_palette(&mut self, rec: &WmfRecord) {
        let Ok(index) = read_u16_le(&rec.params, 0) else {
            self.malformed(record::SELECT_PALETTE, "palette index");
            return;
        };
        if index == stock::FLAG | stock::DEFAULT_PALETTE {
            self.state.palette_index = None;
        } else if matches!(
            self.objects.get(usize::from(index)),
            Some(GdiObject::Palette(_))
        ) {
            self.state.palette_index = Some(usize::from(index));
        } else {
            // A handle without a palette makes SelectPalette fail, which
            // leaves the selected palette unchanged.
            self.note(
                "noncanonical-wmf-select-missing-palette",
                format!("SelectPalette ignored missing palette {index}"),
            );
        }
    }

    fn update_palette(&mut self, rec: &WmfRecord) {
        let Some(index) = self.state.palette_index else {
            return;
        };
        if rec.params.len() < 4 {
            self.malformed(record::canonical(rec.function), "palette range");
            return;
        }
        let start = usize::from(read_u16_le(&rec.params, 0).unwrap_or(0));
        let count = usize::from(read_u16_le(&rec.params, 2).unwrap_or(0));
        if rec.params.len() < 4 + count.saturating_mul(4) {
            self.malformed(record::canonical(rec.function), "palette entries");
            return;
        }
        let Some(GdiObject::Palette(palette)) = self.objects.get(index) else {
            return;
        };
        let target = palette.entries.len().max(start.saturating_add(count));
        let additional = target
            .saturating_sub(palette.entries.len())
            .saturating_mul(4);
        if !self.reserve_object_bytes(additional) {
            return;
        }
        let Some(GdiObject::Palette(palette)) = self.objects.get_mut(index) else {
            return;
        };
        if palette
            .entries
            .try_reserve_exact(target.saturating_sub(palette.entries.len()))
            .is_err()
        {
            self.fatal("failed to allocate bounded WMF palette update".to_owned());
            return;
        }
        palette.entries.resize(target, 0);
        for entry in 0..count {
            let offset = 4 + entry * 4;
            palette.entries[start + entry] = u32::from(rec.params[offset])
                | (u32::from(rec.params[offset + 1]) << 8)
                | (u32::from(rec.params[offset + 2]) << 16);
        }
    }

    fn resize_palette(&mut self, rec: &WmfRecord) {
        let Ok(size) = read_u16_le(&rec.params, 0) else {
            self.malformed(record::RESIZE_PALETTE, "palette size");
            return;
        };
        if let Some(index) = self.state.palette_index
            && let Some(GdiObject::Palette(palette)) = self.objects.get(index)
        {
            let target = usize::from(size);
            let additional = target
                .saturating_sub(palette.entries.len())
                .saturating_mul(4);
            if !self.reserve_object_bytes(additional) {
                return;
            }
            if let Some(GdiObject::Palette(palette)) = self.objects.get_mut(index) {
                if palette
                    .entries
                    .try_reserve_exact(target.saturating_sub(palette.entries.len()))
                    .is_err()
                {
                    self.fatal("failed to allocate bounded WMF palette resize".to_owned());
                    return;
                }
                palette.entries.resize(target, 0);
            }
        }
    }

    fn update_clip_rect(&mut self, function: u16, rec: &WmfRecord) {
        let Some((left, top, right, bottom)) = read_box(&rec.params, 0) else {
            self.malformed(function, "clip rectangle");
            return;
        };
        let rect = self.state.mapping.rect(left, top, right, bottom);
        if function == record::INTERSECT_CLIP_RECT {
            let existing = self.state.clip.rects.as_ref().map_or(0, Vec::len);
            if !self.charge_work(existing) {
                return;
            }
            if let Some(rects) = &mut self.state.clip.rects {
                // Intersect in place so a selected region is never duplicated
                // transiently outside the aggregate heap budget.
                rects.retain_mut(|current| {
                    if let Some(intersection) = current.intersect(rect) {
                        *current = intersection;
                        true
                    } else {
                        false
                    }
                });
            } else {
                let mut rects = Vec::new();
                if !rect.is_empty() {
                    if !self.can_retain_additional_state(size_of::<DeviceRect>()) {
                        self.fatal(format!(
                            "WMF intersected clip data exceeds limit {} bytes",
                            self.max_object_bytes
                        ));
                        return;
                    }
                    if rects.try_reserve_exact(1).is_err() {
                        self.fatal("failed to allocate bounded WMF clip rectangle".to_owned());
                        return;
                    }
                    rects.push(rect);
                }
                self.state.clip.rects = Some(rects);
            }
        } else if !rect.is_empty() {
            if !self.can_retain_additional_state(size_of::<DeviceRect>()) {
                self.fatal(format!(
                    "WMF excluded clip data exceeds limit {} bytes",
                    self.max_object_bytes
                ));
                return;
            }
            if self.state.clip.excluded.try_reserve_exact(1).is_err() {
                self.fatal("failed to allocate bounded WMF excluded clip rectangle".to_owned());
                return;
            }
            self.state.clip.excluded.push(rect);
        }
        self.bump_clip_revision();
    }

    fn offset_clip(&mut self, rec: &WmfRecord) {
        let Some((x, y)) = read_yx(&rec.params) else {
            self.malformed(record::OFFSET_CLIP_RGN, "clip offset");
            return;
        };
        let (dx, dy) = self.state.mapping.vector(f64::from(x), f64::from(y));
        self.state.clip.offset(dx, dy);
        self.bump_clip_revision();
    }

    fn select_clip_region(&mut self, rec: &WmfRecord) {
        let Ok(index) = read_u16_le(&rec.params, 0) else {
            self.malformed(record::SELECT_CLIP_REGION, "region index");
            return;
        };
        self.select_clip_region_object(index, true);
    }

    /// Selects the region at `index` as the clip region. `zero_is_null`
    /// makes index 0 pass no region, as META_SELECTCLIPREGION does.
    fn select_clip_region_object(&mut self, index: u16, zero_is_null: bool) {
        // Playback hands SelectClipRgn the handle stored at the index. Index
        // 0, which metafile recorders write for SelectClipRgn(NULL), and an
        // empty slot both pass no region, which removes the clip region.
        let object = if index == 0 && zero_is_null {
            None
        } else {
            self.objects.get(usize::from(index))
        };
        let rect_count = match object {
            Some(GdiObject::Region(region)) => region.rects.len(),
            None => {
                if index != 0 || !zero_is_null {
                    self.note(
                        "noncanonical-wmf-select-missing-region",
                        format!(
                            "SelectClipRegion found no region {index}; the clip region is removed"
                        ),
                    );
                }
                self.state.clip = ClipRegion::default();
                self.bump_clip_revision();
                return;
            },
            Some(_) => {
                // SelectClipRgn fails for a handle that is not a region.
                self.note(
                    "noncanonical-wmf-clip-region-not-region",
                    format!("SelectClipRegion ignores object {index}, which is not a region"),
                );
                return;
            },
        };
        if !self.charge_work(rect_count) {
            return;
        }
        let additional = rect_count.saturating_mul(size_of::<DeviceRect>());
        if !self.can_retain_additional_state(additional) {
            self.fatal(format!(
                "WMF selected clip data exceeds limit {} bytes",
                self.max_object_bytes
            ));
            return;
        }
        let mapping = self.state.mapping;
        let mut rects = Vec::new();
        if rects.try_reserve_exact(rect_count).is_err() {
            self.fatal("failed to allocate bounded WMF clip region".to_owned());
            return;
        }
        let Some(GdiObject::Region(region)) = self.objects.get(usize::from(index)) else {
            return;
        };
        for rect in &region.rects {
            let (x1, y1) = mapping.point_f64(rect.left, rect.top);
            let (x2, y2) = mapping.point_f64(rect.right, rect.bottom);
            rects.push(DeviceRect::new(x1, y1, x2, y2));
        }
        self.state.clip = ClipRegion::from_rects(rects);
        self.bump_clip_revision();
    }

    fn render_rectangle(&mut self, rec: &WmfRecord) -> Option<String> {
        let (left, top, right, bottom) = self.require_box(record::RECTANGLE, rec, 0)?;
        let rect = self.logical_rect(left, top, right, bottom);
        let mut output = String::with_capacity(192);
        write_rect_start(&mut output, rect);
        self.append_shape_attrs(&mut output, true, true);
        output.push_str("/>");
        Some(output)
    }

    fn render_round_rect(&mut self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 12 {
            self.malformed(record::ROUND_RECT, "rounded rectangle");
            return None;
        }
        let height = read_i16_le(&rec.params, 0).unwrap_or(0);
        let width = read_i16_le(&rec.params, 2).unwrap_or(0);
        let (left, top, right, bottom) = read_box(&rec.params, 4)?;
        let rect = self.logical_rect(left, top, right, bottom);
        let (rx, ry) = self
            .state
            .mapping
            .vector(f64::from(width).abs() / 2.0, f64::from(height).abs() / 2.0);
        let mut output = String::with_capacity(224);
        write_rect_start(&mut output, rect);
        output.push_str(r#" rx=""#);
        write_num(
            &mut output,
            self.transform
                .device_width(rx)
                .min((rect.right - rect.left) / 2.0),
        );
        output.push_str(r#"" ry=""#);
        write_num(
            &mut output,
            self.transform
                .device_height(ry)
                .min((rect.bottom - rect.top) / 2.0),
        );
        output.push('"');
        self.append_shape_attrs(&mut output, true, true);
        output.push_str("/>");
        Some(output)
    }

    fn render_ellipse(&mut self, rec: &WmfRecord) -> Option<String> {
        let (left, top, right, bottom) = self.require_box(record::ELLIPSE, rec, 0)?;
        Some(self.ellipse_element(left, top, right, bottom, true, true))
    }

    fn ellipse_element(
        &mut self,
        left: i16,
        top: i16,
        right: i16,
        bottom: i16,
        fill: bool,
        stroke: bool,
    ) -> String {
        let rect = self.logical_rect(left, top, right, bottom);
        let mut output = String::with_capacity(192);
        output.push_str(r#"<ellipse cx=""#);
        write_num(&mut output, (rect.left + rect.right) / 2.0);
        output.push_str(r#"" cy=""#);
        write_num(&mut output, (rect.top + rect.bottom) / 2.0);
        output.push_str(r#"" rx=""#);
        write_num(&mut output, (rect.right - rect.left) / 2.0);
        output.push_str(r#"" ry=""#);
        write_num(&mut output, (rect.bottom - rect.top) / 2.0);
        output.push('"');
        self.append_shape_attrs(&mut output, fill, stroke);
        output.push_str("/>");
        output
    }

    fn render_polygon(&mut self, rec: &WmfRecord, closed: bool) -> Option<String> {
        let function = if closed {
            record::POLYGON
        } else {
            record::POLYLINE
        };
        let Ok(count) = read_u16_le(&rec.params, 0).map(usize::from) else {
            self.malformed(function, "point count");
            return None;
        };
        if rec.params.len() < 2 + count.saturating_mul(4) {
            self.malformed(function, "points");
            return None;
        }
        if count < 2 {
            // GDI refuses a Polygon or Polyline with fewer than two points
            // and draws nothing.
            self.note(
                "noncanonical-wmf-degenerate-polyline",
                format!("ignored WMF record 0x{function:04X} with {count} points"),
            );
            return None;
        }
        let mut xs = Vec::with_capacity(count);
        let mut ys = Vec::with_capacity(count);
        for index in 0..count {
            xs.push(read_i16_le(&rec.params, 2 + index * 4).unwrap_or(0));
            ys.push(read_i16_le(&rec.params, 4 + index * 4).unwrap_or(0));
        }
        let mut output = String::with_capacity(160 + count * 16);
        output.push_str(if closed {
            r#"<polygon points=""#
        } else {
            r#"<polyline points=""#
        });
        self.transform
            .transform_and_format_points(&self.state.mapping, &xs, &ys, &mut output, ' ');
        output.push('"');
        self.append_shape_attrs(&mut output, closed, true);
        output.push_str("/>");
        Some(output)
    }

    fn render_polypolygon(&mut self, rec: &WmfRecord) -> Option<String> {
        let Ok(polygon_count) = read_u16_le(&rec.params, 0).map(usize::from) else {
            self.malformed(record::POLYPOLYGON, "polygon count");
            return None;
        };
        if polygon_count == 0 || rec.params.len() < 2 + polygon_count.saturating_mul(2) {
            self.malformed(record::POLYPOLYGON, "polygon counts");
            return None;
        }
        let mut counts = Vec::with_capacity(polygon_count);
        let mut point_total = 0usize;
        for index in 0..polygon_count {
            let count = usize::from(read_u16_le(&rec.params, 2 + index * 2).unwrap_or(0));
            point_total = point_total.checked_add(count)?;
            counts.push(count);
        }
        let mut offset: usize = 2 + polygon_count * 2;
        if rec.params.len() < offset.saturating_add(point_total.saturating_mul(4)) {
            self.malformed(record::POLYPOLYGON, "polygon points");
            return None;
        }
        let mut path = String::with_capacity(point_total * 16);
        for count in counts {
            if count == 0 {
                continue;
            }
            for index in 0..count {
                let x = read_i16_le(&rec.params, offset + index * 4).unwrap_or(0);
                let y = read_i16_le(&rec.params, offset + index * 4 + 2).unwrap_or(0);
                let (x, y) = self.transform.point(&self.state.mapping, x, y);
                path.push(if index == 0 { 'M' } else { 'L' });
                write_num(&mut path, x);
                path.push(',');
                write_num(&mut path, y);
            }
            path.push('Z');
            offset += count * 4;
        }
        let mut output = String::with_capacity(path.len() + 160);
        output.push_str(r#"<path d=""#);
        output.push_str(&path);
        output.push('"');
        self.append_shape_attrs(&mut output, true, true);
        output.push_str("/>");
        Some(output)
    }

    fn render_line_to(&mut self, rec: &WmfRecord) -> Option<String> {
        let Some((x2, y2)) = read_yx(&rec.params) else {
            self.malformed(record::LINE_TO, "point");
            return None;
        };
        let (x1_svg, y1_svg) = self.transform.point(
            &self.state.mapping,
            self.state.position.0,
            self.state.position.1,
        );
        let (x2_svg, y2_svg) = self.transform.point(&self.state.mapping, x2, y2);
        self.state.position = (x2, y2);
        let mut output = String::with_capacity(192);
        output.push_str(r#"<line x1=""#);
        write_num(&mut output, x1_svg);
        output.push_str(r#"" y1=""#);
        write_num(&mut output, y1_svg);
        output.push_str(r#"" x2=""#);
        write_num(&mut output, x2_svg);
        output.push_str(r#"" y2=""#);
        write_num(&mut output, y2_svg);
        output.push('"');
        self.append_shape_attrs(&mut output, false, true);
        output.push_str("/>");
        Some(output)
    }

    fn render_set_pixel(&mut self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 8 {
            self.malformed(record::SET_PIXEL, "pixel");
            return None;
        }
        let color = read_u32(&rec.params, 0).unwrap_or(0);
        let y = read_i16_le(&rec.params, 4).unwrap_or(0);
        let x = read_i16_le(&rec.params, 6).unwrap_or(0);
        let (x, y) = self.transform.point(&self.state.mapping, x, y);
        // One device pixel; the SVG canvas stands for the playback device.
        let size = 1.0;
        let mut output = String::with_capacity(144);
        output.push_str(r#"<rect x=""#);
        write_num(&mut output, x - size / 2.0);
        output.push_str(r#"" y=""#);
        write_num(&mut output, y - size / 2.0);
        output.push_str(r#"" width=""#);
        write_num(&mut output, size);
        output.push_str(r#"" height=""#);
        write_num(&mut output, size);
        output.push_str(r#"" fill=""#);
        write_color_hex(&mut output, color);
        output.push('"');
        output.push_str(&self.clip_attr());
        output.push_str("/>");
        Some(output)
    }

    fn render_pat_blt(&mut self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 12 {
            self.malformed(record::PAT_BLT, "PatBlt rectangle");
            return None;
        }
        let rop3 = Rop3::from_code(read_u32(&rec.params, 0).unwrap_or(0));
        let height = read_i16_le(&rec.params, 4).unwrap_or(0);
        let width = read_i16_le(&rec.params, 6).unwrap_or(0);
        let y = read_i16_le(&rec.params, 8).unwrap_or(0);
        let x = read_i16_le(&rec.params, 10).unwrap_or(0);
        if width == 0 || height == 0 {
            self.note(
                "noncanonical-wmf-empty-bitmap-record",
                "PatBlt with an empty rectangle paints no pixels".to_owned(),
            );
            return None;
        }
        if rop3.depends_on_source() {
            // PatBlt fails for a raster operation that reads a source.
            self.note(
                "noncanonical-wmf-patblt-source-rop",
                format!(
                    "PatBlt ignores raster operation 0x{:02X}, which needs a source",
                    rop3.table()
                ),
            );
            return None;
        }
        let rect = self.logical_origin_size(x, y, width, height);
        let outline = rect_outline(
            i32::from(x),
            i32::from(y),
            i32::from(width),
            i32::from(height),
        );
        if let Some(output) = self.bracket_blit(rop3, rect, outline) {
            return output;
        }
        if self.xor_bracket.is_some() {
            self.fail_xor_bracket();
            return None;
        }
        // PatBlt paints nothing with a null brush.
        if rop3.table() == 0xaa || self.state.brush.style == brush::BS_NULL {
            return None;
        }
        let fill = match rop3.table() {
            0x00 => r##" fill="#000""##.to_owned(),
            0xff => r##" fill="#fff""##.to_owned(),
            0xf0 => {
                let brush = self.state.brush.clone();
                self.brush_fill(&brush, self.hatch_background())
            },
            _ => {
                let blit = DeferredBlit {
                    record: None,
                    rop3,
                    mirrored: (false, false),
                    destination: [
                        rect.left,
                        rect.top,
                        rect.right - rect.left,
                        rect.bottom - rect.top,
                    ],
                    brush: self.solid_brush_color(),
                };
                return self.defer_blit(blit, None);
            },
        };
        let element = self.clipped_rect_with(rect, &fill);
        self.paint_after_blits(element)
    }

    fn render_flood_fill(&mut self, rec: &WmfRecord) -> Option<String> {
        let required = if record::canonical(rec.function) == record::EXT_FLOOD_FILL {
            10
        } else {
            8
        };
        if rec.params.len() < required {
            self.malformed(record::canonical(rec.function), "flood fill");
            return None;
        }
        let y = read_i16_le(&rec.params, 4).unwrap_or(0);
        let x = read_i16_le(&rec.params, 6).unwrap_or(0);
        let (x, y) = self.transform.point(&self.state.mapping, x, y);
        let brush = self.state.brush.clone();
        let mut output = String::with_capacity(224);
        output.push_str(r#"<circle cx=""#);
        write_num(&mut output, x);
        output.push_str(r#"" cy=""#);
        write_num(&mut output, y);
        output.push_str(r#"" r="2" data-wmf-approximation="flood-fill""#);
        let attrs = self.paint_attrs(true, false, &brush);
        output.push_str(&attrs);
        output.push_str(&self.clip_attr());
        output.push_str("/>");
        self.fatal("WMF flood fill cannot be faithfully represented in SVG".to_owned());
        Some(output)
    }

    fn render_region_record(&mut self, function: u16, rec: &WmfRecord) -> Option<String> {
        let (region_index, brush_value, frame) = match function {
            record::FILL_REGION => {
                if rec.params.len() < 4 {
                    self.malformed(function, "region and brush indexes");
                    return None;
                }
                let brush_index = usize::from(read_u16_le(&rec.params, 0).unwrap_or(0));
                let region_index = usize::from(read_u16_le(&rec.params, 2).unwrap_or(0));
                match self.object_brush(brush_index)? {
                    Some(brush) => (region_index, brush, None),
                    None => return None,
                }
            },
            record::FRAME_REGION => {
                if rec.params.len() < 8 {
                    self.malformed(function, "frame region");
                    return None;
                }
                let height = read_i16_le(&rec.params, 0).unwrap_or(0);
                let width = read_i16_le(&rec.params, 2).unwrap_or(0);
                let brush_index = usize::from(read_u16_le(&rec.params, 4).unwrap_or(0));
                let region_index = usize::from(read_u16_le(&rec.params, 6).unwrap_or(0));
                match self.object_brush(brush_index)? {
                    Some(brush) => (region_index, brush, Some((width, height))),
                    None => return None,
                }
            },
            record::PAINT_REGION | record::INVERT_REGION => {
                let Ok(index) = read_u16_le(&rec.params, 0) else {
                    self.malformed(function, "region index");
                    return None;
                };
                (usize::from(index), self.state.brush.clone(), None)
            },
            _ => return None,
        };
        let Some(GdiObject::Region(region)) = self.objects.get(region_index).cloned() else {
            // GDI fails a region record whose region handle is invalid and
            // paints nothing for it.
            self.note(
                "noncanonical-wmf-region-missing",
                format!("ignored region record with missing region {region_index}"),
            );
            return None;
        };
        let path = self.region_path(&region);
        let mut output = String::with_capacity(path.len() + 224);
        output.push_str(r#"<path d=""#);
        output.push_str(&path);
        output.push('"');
        if function == record::INVERT_REGION {
            self.fatal("InvertRegion requires destination-dependent raster composition".to_owned());
            output.push_str(r##" fill="#fff" style="mix-blend-mode:difference""##);
        } else if let Some((width, height)) = frame {
            let Some((frame_brush, background)) = self.mixed_brush(self.rop2(), &brush_value)
            else {
                self.rop2_attr();
                return None;
            };
            output.push_str(r#" fill="none" stroke=""#);
            let paint = self.brush_paint(&frame_brush, background);
            output.push_str(&paint);
            let (width, height) = self
                .state
                .mapping
                .vector(f64::from(width).abs(), f64::from(height).abs());
            output.push_str(r#"" stroke-width=""#);
            write_num(
                &mut output,
                (self.transform.device_width(width) + self.transform.device_height(height)) / 2.0,
            );
            output.push('"');
        } else {
            let attrs = self.paint_attrs(true, false, &brush_value);
            output.push_str(&attrs);
        }
        output.push_str(&self.clip_attr());
        output.push_str("/>");
        Some(output)
    }

    /// Looks up the brush of a region record.
    ///
    /// GDI fails the record and paints nothing when the brush handle is
    /// invalid, so a missing brush is not fatal for the whole picture.
    fn object_brush(&mut self, index: usize) -> Option<Option<Brush>> {
        match self.objects.get(index).cloned() {
            Some(GdiObject::Brush(value)) => Some(Some(value)),
            _ => {
                self.note(
                    "noncanonical-wmf-region-brush-missing",
                    format!("ignored region record with missing brush {index}"),
                );
                Some(None)
            },
        }
    }

    fn render_arc_common(&mut self, rec: &WmfRecord, arc_type: ArcType) -> Option<String> {
        if rec.params.len() < 16 {
            self.malformed(record::canonical(rec.function), "arc");
            return None;
        }
        let y_end = read_i16_le(&rec.params, 0).unwrap_or(0);
        let x_end = read_i16_le(&rec.params, 2).unwrap_or(0);
        let y_start = read_i16_le(&rec.params, 4).unwrap_or(0);
        let x_start = read_i16_le(&rec.params, 6).unwrap_or(0);
        let bottom = read_i16_le(&rec.params, 8).unwrap_or(0);
        let right = read_i16_le(&rec.params, 10).unwrap_or(0);
        let top = read_i16_le(&rec.params, 12).unwrap_or(0);
        let left = read_i16_le(&rec.params, 14).unwrap_or(0);

        if x_start == x_end && y_start == y_end {
            return Some(self.ellipse_element(
                left,
                top,
                right,
                bottom,
                !matches!(arc_type, ArcType::Open),
                true,
            ));
        }
        let rect = self.logical_rect(left, top, right, bottom);
        let rx = (rect.right - rect.left) / 2.0;
        let ry = (rect.bottom - rect.top) / 2.0;
        let cx = (rect.left + rect.right) / 2.0;
        let cy = (rect.top + rect.bottom) / 2.0;
        let (start_guide_x, start_guide_y) =
            self.transform.point(&self.state.mapping, x_start, y_start);
        let (end_guide_x, end_guide_y) = self.transform.point(&self.state.mapping, x_end, y_end);
        if rx <= 0.0 || ry <= 0.0 {
            let mut output = String::with_capacity(192);
            output.push_str(r#"<line x1=""#);
            write_num(&mut output, start_guide_x);
            output.push_str(r#"" y1=""#);
            write_num(&mut output, start_guide_y);
            output.push_str(r#"" x2=""#);
            write_num(&mut output, end_guide_x);
            output.push_str(r#"" y2=""#);
            write_num(&mut output, end_guide_y);
            output.push('"');
            self.append_shape_attrs(&mut output, false, true);
            output.push_str("/>");
            return Some(output);
        }
        let start_angle = ((start_guide_y - cy) / ry).atan2((start_guide_x - cx) / rx);
        let end_angle = ((end_guide_y - cy) / ry).atan2((end_guide_x - cx) / rx);
        let start = (cx + rx * start_angle.cos(), cy + ry * start_angle.sin());
        let end = (cx + rx * end_angle.cos(), cy + ry * end_angle.sin());
        let mirrored = self.transform.determinant_sign(&self.state.mapping) < 0.0;
        let sweep_flag = usize::from(mirrored);
        let directed_angle = if mirrored {
            positive_angle(end_angle - start_angle)
        } else {
            positive_angle(start_angle - end_angle)
        };
        let large_arc = usize::from(directed_angle > std::f64::consts::PI);

        let mut output = String::with_capacity(256);
        output.push_str(r#"<path d="M"#);
        write_num(&mut output, start.0);
        output.push(',');
        write_num(&mut output, start.1);
        output.push('A');
        write_num(&mut output, rx);
        output.push(',');
        write_num(&mut output, ry);
        let _ =
            std::fmt::Write::write_fmt(&mut output, format_args!(" 0 {large_arc},{sweep_flag} "));
        write_num(&mut output, end.0);
        output.push(',');
        write_num(&mut output, end.1);
        match arc_type {
            ArcType::Pie => {
                output.push('L');
                write_num(&mut output, cx);
                output.push(',');
                write_num(&mut output, cy);
                output.push('Z');
            },
            ArcType::Chord => output.push('Z'),
            ArcType::Open => {},
        }
        output.push('"');
        self.append_shape_attrs(&mut output, !matches!(arc_type, ArcType::Open), true);
        output.push_str("/>");
        Some(output)
    }

    fn render_text_out(&mut self, rec: &WmfRecord) -> Option<String> {
        let Ok(length) = read_u16_le(&rec.params, 0).map(usize::from) else {
            self.malformed(record::TEXT_OUT, "text length");
            return None;
        };
        let padded = length.checked_add(1).map(|value| value & !1)?;
        let coordinate_offset = 2usize.checked_add(padded)?;
        if rec.params.len() < coordinate_offset.saturating_add(4)
            || rec.params.len() < 2usize.saturating_add(length)
        {
            self.malformed(record::TEXT_OUT, "text payload");
            return None;
        }
        let y = read_i16_le(&rec.params, coordinate_offset).unwrap_or(0);
        let x = read_i16_le(&rec.params, coordinate_offset + 2).unwrap_or(0);
        self.render_text_bytes(&rec.params[2..2 + length], x, y, 0, None, &[])
    }

    fn render_ext_text_out(&mut self, rec: &WmfRecord) -> Option<String> {
        if rec.params.len() < 8 {
            self.malformed(record::EXT_TEXT_OUT, "extended text header");
            return None;
        }
        let y = read_i16_le(&rec.params, 0).unwrap_or(0);
        let x = read_i16_le(&rec.params, 2).unwrap_or(0);
        let length = usize::from(read_u16_le(&rec.params, 4).unwrap_or(0));
        let options = read_u16_le(&rec.params, 6).unwrap_or(0);
        if options & ext_text_out::ETO_GLYPH_INDEX != 0 {
            // The string holds glyph numbers of the recording system's font,
            // which no character mapping can recover; the text is omitted.
            self.note(
                "wmf-glyph-index-text-omitted",
                "ExtTextOut text given as font glyph indices is not drawn".to_owned(),
            );
            return None;
        }
        let has_rect = options & (ext_text_out::ETO_CLIPPED | ext_text_out::ETO_OPAQUE) != 0;
        let (rect, string_offset) = if has_rect {
            if rec.params.len() < 16 {
                self.malformed(record::EXT_TEXT_OUT, "option rectangle");
                return None;
            }
            let left = read_i16_le(&rec.params, 8).unwrap_or(0);
            let top = read_i16_le(&rec.params, 10).unwrap_or(0);
            let right = read_i16_le(&rec.params, 12).unwrap_or(0);
            let bottom = read_i16_le(&rec.params, 14).unwrap_or(0);
            (Some((left, top, right, bottom)), 16usize)
        } else {
            (None, 8usize)
        };
        if rec.params.len() < string_offset.saturating_add(length) {
            self.malformed(record::EXT_TEXT_OUT, "text payload");
            return None;
        }
        // The dx array holds the distance between the origins of adjacent
        // character cells, one value per byte of the string, or an x and a y
        // value with ETO_PDY ([MS-WMF] 2.3.3.5).
        let dx_offset = string_offset + ((length + 1) & !1);
        let stride = if options & ext_text_out::ETO_PDY != 0 {
            2
        } else {
            1
        };
        let dx_count = (rec.params.len().saturating_sub(dx_offset) / 2 / stride).min(length);
        let advance =
            |index: usize| f64::from(read_i16_le(&rec.params, dx_offset + index * 2).unwrap_or(0));
        let advances: Vec<(f64, f64)> = (0..dx_count)
            .map(|index| {
                let first = index * stride;
                let dy = if stride == 2 { advance(first + 1) } else { 0.0 };
                (advance(first), dy)
            })
            .collect();
        self.render_text_bytes(
            &rec.params[string_offset..string_offset + length],
            x,
            y,
            options,
            rect,
            &advances,
        )
    }

    fn render_text_bytes(
        &mut self,
        bytes: &[u8],
        supplied_x: i16,
        supplied_y: i16,
        options: u16,
        option_rect: Option<(i16, i16, i16, i16)>,
        advances: &[(f64, f64)],
    ) -> Option<String> {
        if bytes.is_empty() {
            return None;
        }
        let mut decoded = crate::text_charset::decode_logfont_text(
            bytes,
            self.state.font.charset,
            &self.state.font.name,
        );
        if !decoded.exact {
            self.note(
                "wmf-text-charset-inexact",
                format!(
                    "text in font charset {} could not be decoded exactly; Windows-1252 or replacement characters were used",
                    self.state.font.charset
                ),
            );
        }
        let use_face = !decoded.unicode_face;
        let update_position = self.state.text_align & text_align::TA_UPDATECP != 0;
        let (logical_x, logical_y) = if update_position {
            self.state.position
        } else {
            (supplied_x, supplied_y)
        };
        let (logical_x, logical_y) = (f64::from(logical_x), f64::from(logical_y));
        let device_position = |x: f64, y: f64| {
            let (device_x, device_y) = self.state.mapping.point_f64(x, y);
            self.transform.device_point(device_x, device_y)
        };
        let (x, y) = device_position(logical_x, logical_y);
        let (_, font_height) = self
            .state
            .mapping
            .vector(0.0, f64::from(self.state.font.height).abs().max(1.0));
        let font_size = self.transform.device_height(font_height).max(1.0);
        let logical_advance = if advances.is_empty() {
            let nominal = if self.state.font.width == 0 {
                f64::from(self.state.font.height).abs() * 0.6
            } else {
                f64::from(self.state.font.width).abs()
            };
            (nominal + f64::from(self.state.text_char_extra)) * bytes.len() as f64
        } else {
            advances.iter().map(|advance| advance.0).sum()
        };

        // Each character is placed at the origin the advances give it; SVG
        // would add a `dx` value to the glyph's own advance instead.
        let mut positions: Vec<(f64, f64)> = decoded
            .origin_offsets(advances)
            .into_iter()
            .map(|(dx, dy)| device_position(logical_x + dx, logical_y + dy))
            .collect();
        // GDI aligns the whole string on the reference point, while SVG
        // applies text-anchor to every absolutely positioned character; with
        // recorded advances the string extent is known, so the positions move
        // instead.
        let anchor = self.state.text_align & text_align::HORIZONTAL_MASK;
        let positioned = positions.len() > 1;
        if positioned {
            let extent = advances.iter().fold((0.0, 0.0), |sum, advance| {
                (sum.0 + advance.0, sum.1 + advance.1)
            });
            let (width, height) =
                self.transform
                    .logical_vector(&self.state.mapping, extent.0, extent.1);
            let shift = match anchor {
                text_align::TA_RIGHT => 1.0,
                text_align::TA_CENTER => 0.5,
                _ => 0.0,
            };
            for position in &mut positions {
                position.0 -= width * shift;
                position.1 -= height * shift;
            }
        }
        let canvas = self.transform.canvas_rect();
        let positions = decoded.retain_visible(&positions, |position| {
            crate::text_charset::origin_reaches_canvas(
                position,
                (x, y),
                (canvas.left, canvas.top, canvas.right, canvas.bottom),
                4.0 * font_size,
            )
        });
        let positioned = positions.len() > 1;
        let text = decoded.text;

        let mut text_element = String::with_capacity(text.len() + positions.len() * 16 + 320);
        text_element.push_str(r#"<text x=""#);
        if positioned {
            write_num_list(
                &mut text_element,
                positions.iter().map(|position| position.0),
            );
        } else {
            write_num(&mut text_element, x);
        }
        text_element.push_str(r#"" y=""#);
        if positioned {
            write_num_list(
                &mut text_element,
                positions.iter().map(|position| position.1),
            );
        } else {
            write_num(&mut text_element, y);
        }
        text_element.push_str(r#"" font-size=""#);
        write_num(&mut text_element, font_size);
        text_element.push_str(r#"" fill=""#);
        write_color_hex(&mut text_element, self.state.text_color);
        text_element.push('"');
        if !positioned {
            match anchor {
                text_align::TA_RIGHT => text_element.push_str(r#" text-anchor="end""#),
                text_align::TA_CENTER => text_element.push_str(r#" text-anchor="middle""#),
                _ => {},
            }
        }
        match self.state.text_align & text_align::VERTICAL_MASK {
            text_align::TA_BOTTOM => {
                text_element.push_str(r#" dominant-baseline="text-after-edge""#)
            },
            text_align::TA_BASELINE => {},
            _ => text_element.push_str(r#" dominant-baseline="text-before-edge""#),
        }
        if use_face && self.state.font.name != "serif" {
            text_element.push_str(r#" font-family=""#);
            escape_xml_attr_into(
                &mut text_element,
                &crate::text_charset::svg_font_family(
                    &self.state.font.name,
                    self.state.font.pitch_and_family,
                ),
            );
            text_element.push('"');
        }
        if self.state.font.italic {
            text_element.push_str(r#" font-style="italic""#);
        }
        if self.state.font.weight != 400 {
            text_element.push_str(r#" font-weight=""#);
            let _ = std::fmt::Write::write_fmt(
                &mut text_element,
                format_args!("{}", self.state.font.weight),
            );
            text_element.push('"');
        }
        match (self.state.font.underline, self.state.font.strike_out) {
            (true, true) => text_element.push_str(r#" text-decoration="underline line-through""#),
            (true, false) => text_element.push_str(r#" text-decoration="underline""#),
            (false, true) => text_element.push_str(r#" text-decoration="line-through""#),
            _ => {},
        }
        let text_angle = if self.state.font.escapement != 0 {
            self.state.font.escapement
        } else {
            self.state.font.orientation
        };
        if text_angle != 0 {
            let angle = -f64::from(text_angle) / 10.0;
            text_element.push_str(r#" transform="rotate("#);
            write_num(&mut text_element, angle);
            text_element.push(' ');
            write_num(&mut text_element, x);
            text_element.push(' ');
            write_num(&mut text_element, y);
            text_element.push_str(r#")""#);
        }
        if self.state.font.orientation != 0 {
            text_element.push_str(r#" data-wmf-font-orientation=""#);
            let _ = std::fmt::Write::write_fmt(
                &mut text_element,
                format_args!("{}", self.state.font.orientation),
            );
            text_element.push('"');
            if self.state.font.escapement != 0
                && self.state.font.orientation != self.state.font.escapement
            {
                self.warn(
                    "independent WMF glyph orientation retained as SVG diagnostic metadata"
                        .to_owned(),
                );
            }
        }
        if self.state.mapping.layout & 1 != 0
            || self.state.text_align & text_align::TA_RTLREADING != 0
            || options & ext_text_out::ETO_RTLREADING != 0
        {
            text_element.push_str(r#" direction="rtl" unicode-bidi="bidi-override""#);
        }
        if advances.is_empty() && self.state.text_char_extra != 0 {
            let (spacing, _) = self.transform.logical_vector(
                &self.state.mapping,
                f64::from(self.state.text_char_extra),
                0.0,
            );
            text_element.push_str(r#" letter-spacing=""#);
            write_num(&mut text_element, spacing);
            text_element.push('"');
        }
        if self.state.break_count > 0 && self.state.break_extra != 0 {
            let (spacing, _) = self.transform.logical_vector(
                &self.state.mapping,
                f64::from(self.state.break_extra) / f64::from(self.state.break_count),
                0.0,
            );
            text_element.push_str(r#" word-spacing=""#);
            write_num(&mut text_element, spacing);
            text_element.push('"');
        }
        // GDI draws and advances over every space; XML would collapse runs of
        // them and shift the characters after them onto earlier positions.
        text_element.push_str(r#" xml:space="preserve">"#);
        escape_xml_text_into(&mut text_element, &text);
        text_element.push_str("</text>");
        if text.is_empty() {
            text_element.clear();
        }

        let explicit_opaque = options & ext_text_out::ETO_OPAQUE != 0;
        let background = if explicit_opaque || self.state.bk_mode == bk_mode::OPAQUE {
            let rect = if let Some((left, top, right, bottom)) = option_rect {
                self.logical_rect(left, top, right, bottom)
            } else {
                let (advance, _) =
                    self.transform
                        .logical_vector(&self.state.mapping, logical_advance, 0.0);
                let (left, right) = match self.state.text_align & text_align::HORIZONTAL_MASK {
                    text_align::TA_RIGHT => (x - advance, x),
                    text_align::TA_CENTER => (x - advance / 2.0, x + advance / 2.0),
                    _ => (x, x + advance),
                };
                DeviceRect::new(left, y - font_size, right, y + font_size * 0.2)
            };
            let mut background = String::with_capacity(160);
            write_rect_start(&mut background, rect);
            background.push_str(r#" fill=""#);
            write_color_hex(&mut background, self.state.bk_color);
            background.push_str(r#"" stroke="none"/>"#);
            background
        } else {
            String::new()
        };

        if update_position {
            let advance = logical_advance.round();
            self.state.position.0 = if advance > f64::from(i16::MAX) {
                i16::MAX
            } else if advance < f64::from(i16::MIN) {
                i16::MIN
            } else {
                self.state.position.0.saturating_add(advance as i16)
            };
        }

        if text_element.is_empty() && background.is_empty() {
            return None;
        }
        let needs_group = !background.is_empty()
            || options & ext_text_out::ETO_CLIPPED != 0
            || !self.state.clip.is_unbounded();
        if !needs_group {
            return Some(text_element);
        }
        let mut output = String::with_capacity(background.len() + text_element.len() + 128);
        output.push_str("<g");
        output.push_str(&self.clip_attr());
        output.push('>');
        let mut local_clip = false;
        if options & ext_text_out::ETO_CLIPPED != 0
            && let Some((left, top, right, bottom)) = option_rect
        {
            let clip_id = self.add_rect_clip(self.logical_rect(left, top, right, bottom));
            output.push_str(r#"<g clip-path="url(#"#);
            output.push_str(&clip_id);
            output.push_str(r#")">"#);
            local_clip = true;
        }
        output.push_str(&background);
        output.push_str(&text_element);
        if local_clip {
            output.push_str("</g>");
        }
        output.push_str("</g>");
        Some(output)
    }

    fn render_bitmap(&mut self, function: u16, rec: &WmfRecord) -> Option<String> {
        let Some((rect, mirrored)) = bitmap_destination(rec, &self.state.mapping, &self.transform)
        else {
            self.malformed(function, "bitmap destination");
            return None;
        };
        let request = [
            rect.left,
            rect.top,
            rect.right - rect.left,
            rect.bottom - rect.top,
        ];
        if let Some(hook) = self.bitmap_hook {
            let element = match hook(function, &rec.params, request) {
                Ok(Some(svg)) => self.clipped_group(&svg),
                Ok(None) => {
                    self.fatal(format!(
                        "DIB rendering hook declined WMF record 0x{function:04X}"
                    ));
                    return None;
                },
                Err(message) => {
                    self.fatal(format!("DIB rendering hook failed: {message}"));
                    return None;
                },
            };
            return self.paint_after_blits(element);
        }

        let stretch = stretch_policy(self.state.stretch_mode);
        let paint = BitmapOp::parse_wmf(
            rec.size,
            rec.function,
            &rec.params,
            stretch,
            self.dib_limits,
        )
        .and_then(|operation| {
            let destination = operation.destination;
            let outline = rect_outline(
                destination.x,
                destination.y,
                destination.width,
                destination.height,
            );
            if operation.rop3.depends_on_source() && operation.source.is_none() {
                return Ok((operation.rop3, BitmapPaint::Composite, None, outline));
            }
            let window = if operation.rop3.depends_on_source()
                && operation.paint()? != BitmapPaint::Nothing
            {
                operation.blit_window()?
            } else {
                None
            };
            if let Some(BitmapSource::Dib(dib)) = operation.source {
                self.note_dib_tolerances(&dib);
            }
            Ok((operation.rop3, operation.paint()?, window, outline))
        });
        let (rop3, paint, window, outline) = match paint {
            Ok(result) => result,
            Err(error) => {
                self.fatal(format!("WMF bitmap record 0x{function:04X}: {error}"));
                return None;
            },
        };
        if !rop3.depends_on_source()
            && let Some(output) = self.bracket_blit(rop3, rect, outline)
        {
            return output;
        }
        if self.xor_bracket.is_some() {
            self.fail_xor_bracket();
            return None;
        }
        if paint == BitmapPaint::Composite && rop3.depends_on_source() && window.is_none() {
            // A blit whose raster operation reads a source fails without one.
            self.note(
                "noncanonical-wmf-blit-without-source",
                format!(
                    "WMF bitmap record 0x{function:04X} needs a source bitmap but carries none"
                ),
            );
            return None;
        }
        match paint {
            BitmapPaint::Nothing => {
                self.note(
                    "noncanonical-wmf-empty-bitmap-record",
                    format!("WMF bitmap record 0x{function:04X} paints no pixels"),
                );
                None
            },
            BitmapPaint::Solid(color) => {
                let element = self.clipped_rect(rect, color);
                self.paint_after_blits(element)
            },
            BitmapPaint::Brush => {
                if self.state.brush.style == brush::BS_NULL {
                    return None;
                }
                let brush = self.state.brush.clone();
                let fill = self.brush_fill(&brush, self.hatch_background());
                let element = self.clipped_rect_with(rect, &fill);
                self.paint_after_blits(element)
            },
            BitmapPaint::Image | BitmapPaint::Composite => {
                let pixels = if rop3.depends_on_source() {
                    let Some(window) = window else {
                        self.note(
                            "noncanonical-wmf-blit-outside-bitmap",
                            format!(
                                "WMF bitmap record 0x{function:04X} reads no pixel inside its bitmap"
                            ),
                        );
                        return None;
                    };
                    Some(window.pixels())
                } else {
                    None
                };
                let blit = DeferredBlit {
                    record: Some(DeferredRecord {
                        size: rec.size,
                        function: rec.function,
                        params: rec.params.clone(),
                        stretch,
                    }),
                    rop3,
                    mirrored,
                    destination: request,
                    brush: self.solid_brush_color(),
                };
                self.defer_blit(blit, pixels)
            },
        }
    }

    /// Reports DIB header fields that GDI ignores.
    fn note_dib_tolerances(&mut self, dib: &Dib<'_>) {
        for (code, message) in dib.info().tolerances.diagnostics() {
            self.note(code, message.to_owned());
        }
    }

    /// Wraps an element in a group clipped by the current clip region.
    fn clipped_group(&mut self, element: &str) -> String {
        let clip = self.clip_attr();
        wrap_clip(&clip, element)
    }

    /// Paints the deferred bitmap records, then `element`.
    fn paint_after_blits(&mut self, element: String) -> Option<String> {
        match self.flush_blits() {
            Some(mut output) => {
                output.push_str(&element);
                Some(output)
            },
            None => Some(element),
        }
    }

    /// Defers a bitmap record that may combine with the records after it.
    ///
    /// `pixels` is the number of source pixels across and down, or `None`
    /// for a record without a source. Returns the output of the deferred
    /// records that this one cannot join.
    fn defer_blit(&mut self, blit: DeferredBlit, pixels: Option<(u32, u32)>) -> Option<String> {
        let painted = self.blit_painted_area(&blit)?;
        let composite =
            blit.rop3.depends_on_destination() || !matches!(blit.rop3.table(), 0xcc | 0x33);
        let clip = self.clip_attr();
        let max_cells = usize::try_from(self.dib_limits.max_pixels).unwrap_or(usize::MAX);
        if let Some(group) = &mut self.deferred_blits
            && group.clip == clip
            && group.grid.extend(painted, pixels, max_cells)
        {
            group.composite |= composite;
            group.blits.push(blit);
            return None;
        }
        let flushed = self.flush_blits();
        let grid = match pixels {
            Some((columns, rows)) => BlitGrid::for_pixels(painted, columns, rows),
            None => BlitGrid::for_area(painted),
        };
        let Some(grid) = grid else {
            // A destination that collapses to nothing in SVG units paints
            // nothing visible.
            return flushed;
        };
        self.deferred_blits = Some(DeferredBlits {
            clip,
            grid,
            blits: vec![blit],
            composite,
        });
        flushed
    }

    /// The SVG area a deferred record paints.
    fn blit_painted_area(&mut self, blit: &DeferredBlit) -> Option<[f64; 4]> {
        let Some(record) = &blit.record else {
            return Some(blit.destination);
        };
        if !blit.rop3.depends_on_source() {
            return Some(blit.destination);
        }
        match self
            .parse_deferred(record, blit.mirrored)
            .and_then(|operation| operation.blit_window())
        {
            Ok(Some(window)) => Some(window.painted(blit.destination)),
            Ok(None) => None,
            Err(error) => {
                self.fatal(format!(
                    "WMF bitmap record 0x{:04X}: {error}",
                    record::canonical(record.function)
                ));
                None
            },
        }
    }

    /// Parses a deferred bitmap record again, oriented by the mapping mode.
    fn parse_deferred<'record>(
        &self,
        record: &'record DeferredRecord,
        mirrored: (bool, bool),
    ) -> litchi_core::error::Result<BitmapOp<'record>> {
        let mut operation = BitmapOp::parse_wmf(
            record.size,
            record.function,
            &record.params,
            record.stretch,
            self.dib_limits,
        )?;
        // GDI mirrors the bitmap when the mapping mode mirrors the
        // destination rectangle in device space.
        operation.destination.width =
            operation.destination.width.abs() * if mirrored.0 { -1 } else { 1 };
        operation.destination.height =
            operation.destination.height.abs() * if mirrored.1 { -1 } else { 1 };
        Ok(operation)
    }

    /// Paints the deferred bitmap records.
    fn flush_blits(&mut self) -> Option<String> {
        let group = self.deferred_blits.take()?;
        let output = if group.composite {
            self.compose_blits(&group)
        } else {
            // Plain source copies paint exactly as they would one by one.
            let mut output = String::new();
            for blit in &group.blits {
                let Some(record) = &blit.record else {
                    continue;
                };
                let image = self
                    .parse_deferred(record, blit.mirrored)
                    .and_then(|operation| operation.to_svg_image_at(Some(blit.destination)));
                match image {
                    Ok(Some(image)) => output.push_str(&wrap_clip(&group.clip, &image.element)),
                    Ok(None) => {},
                    Err(error) => {
                        self.fatal(format!(
                            "WMF bitmap record 0x{:04X}: {error}",
                            record::canonical(record.function)
                        ));
                        return None;
                    },
                }
            }
            Ok(output)
        };
        match output {
            Ok(output) if !output.is_empty() => Some(output),
            Ok(_) => None,
            Err(error) => {
                self.fatal(error);
                None
            },
        }
    }

    /// Composes deferred records that combine with destination pixels into
    /// one image.
    fn compose_blits(&mut self, group: &DeferredBlits) -> Result<String, String> {
        let mut composite = RasterComposite::new(group.grid).map_err(|error| error.to_string())?;
        let mut brush_approximated = false;
        for blit in &group.blits {
            let pattern = match blit.brush {
                Some(color) => color,
                None => {
                    brush_approximated |= blit.rop3.depends_on_pattern();
                    0
                },
            };
            let source = match &blit.record {
                Some(record) if blit.rop3.depends_on_source() => {
                    let operation = self
                        .parse_deferred(record, blit.mirrored)
                        .map_err(|error| error.to_string())?;
                    match operation.blit_pixels().map_err(|error| error.to_string())? {
                        Some(pixels) => Some(pixels),
                        None => continue,
                    }
                },
                _ => None,
            };
            let result = match &source {
                Some(pixels) => composite.apply(
                    blit.rop3,
                    pixels.window.painted(blit.destination),
                    Some(&pixels.image),
                    pattern,
                ),
                None => composite.apply(blit.rop3, blit.destination, None, pattern),
            };
            result.map_err(|error| error.to_string())?;
        }
        let image = composite
            .to_svg_image(self.max_output_bytes)
            .map_err(|error| error.to_string())?;
        if image.approximated || brush_approximated {
            self.note(
                "wmf-raster-operation-approximated",
                "raster operations that combine with destination pixels were painted as on a white destination"
                    .to_owned(),
            );
        }
        Ok(wrap_clip(&group.clip, &image.image.element))
    }

    /// Rect element covering a bitmap destination, clipped by the current clip.
    fn clipped_rect(&mut self, rect: DeviceRect, color: u32) -> String {
        let mut fill = String::with_capacity(16);
        fill.push_str(r#" fill=""#);
        write_color_hex(&mut fill, color);
        fill.push('"');
        self.clipped_rect_with(rect, &fill)
    }

    /// Rect element covering a bitmap destination with explicit fill attributes.
    fn clipped_rect_with(&mut self, rect: DeviceRect, fill: &str) -> String {
        let mut output = String::with_capacity(160);
        write_rect_start(&mut output, rect);
        output.push_str(fill);
        output.push_str(&self.clip_attr());
        output.push_str("/>");
        output
    }

    fn append_shape_attrs(&mut self, output: &mut String, fill: bool, stroke: bool) {
        let brush = self.state.brush.clone();
        output.push_str(&self.paint_attrs(fill, stroke, &brush));
        output.push_str(&self.clip_attr());
    }

    /// The active mix mode.
    fn rop2(&self) -> Rop2 {
        Rop2::from_mode(u32::from(self.state.rop2)).unwrap_or(Rop2::COPY_PEN)
    }

    /// Fill and stroke attributes of a shape under the current mix mode.
    fn paint_attrs(&mut self, fill: bool, stroke: bool, brush_value: &Brush) -> String {
        let rop2 = self.rop2();
        if let Some(shape) = self.bracket_paint(rop2, fill, stroke, brush_value) {
            let mut attrs = fill_attrs(XOR_FILL, self.state.poly_fill_mode);
            if let BracketShape::Middle(_, Some(_)) = shape {
                let outline = stroke_attrs(&self.state.pen, &self.state.mapping, &self.transform);
                let color = format!(r#" stroke="{}""#, color_paint(self.state.pen.color));
                attrs.push_str(&outline.replacen(&color, &format!(r#" stroke="{XOR_STROKE}""#), 1));
            } else {
                attrs.push_str(r#" stroke="none""#);
            }
            self.bracket_shape = Some(shape);
            return attrs;
        }
        let mut attrs = String::with_capacity(192);
        if fill {
            let Some((brush_value, background)) = self.mixed_brush(rop2, brush_value) else {
                return self.rop2_attr();
            };
            attrs.push_str(&self.brush_fill(&brush_value, background));
        } else {
            attrs.push_str(r#" fill="none""#);
        }
        let mut paints = fill && !attrs.starts_with(r#" fill="none""#);
        if stroke {
            let Some(pen_value) = self.mixed_pen(rop2) else {
                return self.rop2_attr();
            };
            paints |= pen_value.style & pen::PS_STYLE_MASK != pen::PS_NULL;
            attrs.push_str(&stroke_attrs(
                &pen_value,
                &self.state.mapping,
                &self.transform,
            ));
        }
        if rop2 != Rop2::COPY_PEN && !paints {
            self.paints_nothing = true;
        }
        attrs
    }

    /// Whether the selected pen paints nothing under `rop2`.
    fn pen_paints_nothing(&self, rop2: Rop2) -> bool {
        self.mixed_pen(rop2)
            .is_some_and(|pen_value| pen_value.style & pen::PS_STYLE_MASK == pen::PS_NULL)
    }

    /// Identifies a fill that belongs to an XOR bracket: an XOR fill, or a
    /// fill after a first XOR pass.
    fn bracket_paint(
        &self,
        rop2: Rop2,
        fill: bool,
        stroke: bool,
        brush_value: &Brush,
    ) -> Option<BracketShape> {
        if !fill {
            return None;
        }
        let outline = stroke && !self.pen_paints_nothing(rop2);
        if matches!(rop2.mode(), 7 | 10) {
            if outline || brush_value.style != brush::BS_SOLID || brush_value.pattern.is_some() {
                return None;
            }
            // R2_NOTXORPEN XORs the inverse colour.
            let color = if rop2.mode() == 7 {
                brush_value.color
            } else {
                !brush_value.color
            } & 0x00ff_ffff;
            return (color != 0).then_some(BracketShape::Xor(color));
        }
        if rop2.ignores_color() && rop2.paint(0) == Rop2Paint::Unchanged {
            return None;
        }
        let middle = match (&brush_value.pattern, brush_value.style) {
            (None, brush::BS_SOLID) => XorMiddle::Solid {
                rop2,
                color: brush_value.color & 0x00ff_ffff,
            },
            (Some(Ok(pattern)), _) => XorMiddle::Pattern {
                rop2,
                pattern: Arc::clone(pattern),
                text_color: self.state.text_color & 0x00ff_ffff,
                background_color: self.state.bk_color & 0x00ff_ffff,
            },
            _ => return None,
        };
        let outline = if outline {
            let pen_value = &self.state.pen;
            if pen_value.style & pen::PS_STYLE_MASK != pen::PS_SOLID {
                return None;
            }
            Some((rop2, pen_value.color & 0x00ff_ffff))
        } else {
            None
        };
        self.accepts_middle(&middle)
            .then_some(BracketShape::Middle(middle, outline))
    }

    /// Whether the open bracket takes another middle fill painted with
    /// `middle`.
    fn accepts_middle(&self, middle: &XorMiddle) -> bool {
        let Some(bracket) = &self.xor_bracket else {
            return false;
        };
        if bracket.matched != 0 || bracket.middles.len() >= MAX_BRACKET_MIDDLES {
            return false;
        }
        // Overlapping middle fills paint their common pixels twice, which
        // only an idempotent mix mode makes harmless.
        bracket
            .paint
            .as_ref()
            .is_none_or(|paint| paint.same_paint(middle) && middle.rop2().is_idempotent())
    }

    /// Adds a shape to the XOR bracket, returning the bracket's output once
    /// the second pass is complete.
    fn capture_bracket_shape(
        &mut self,
        shape: BracketShape,
        element: String,
        outline: Option<Vec<(i64, i64)>>,
        bounds: Option<[i64; 4]>,
    ) -> Option<String> {
        let mapping = self.state.mapping;
        match shape {
            BracketShape::Xor(color) => {
                let Some(bracket) = &mut self.xor_bracket else {
                    self.xor_bracket = Some(XorBracket {
                        mapping,
                        first: vec![XorShape {
                            element,
                            color,
                            outline,
                        }],
                        paint: None,
                        middles: Vec::new(),
                        matched: 0,
                    });
                    return None;
                };
                if bracket.mapping != mapping {
                    self.fail_xor_bracket();
                    return None;
                }
                if bracket.middles.is_empty() {
                    bracket.first.push(XorShape {
                        element,
                        color,
                        outline,
                    });
                    return None;
                }
                let expected = &bracket.first[bracket.matched];
                if expected.element != element || expected.color != color {
                    self.fail_xor_bracket();
                    return None;
                }
                bracket.matched += 1;
                if bracket.matched < bracket.first.len() {
                    return None;
                }
                let bracket = self.xor_bracket.take()?;
                let output = self.resolve_xor_bracket(bracket);
                if output.is_none() {
                    self.rop2_attr();
                }
                output
            },
            BracketShape::Middle(middle, stroke) => {
                if self
                    .xor_bracket
                    .as_ref()
                    .is_some_and(|bracket| bracket.mapping != mapping)
                {
                    self.fail_xor_bracket();
                    return None;
                }
                if let Some(bracket) = &mut self.xor_bracket {
                    bracket.paint.get_or_insert(middle);
                    bracket.middles.push(MiddleShape {
                        element,
                        stroke,
                        bounds,
                    });
                }
                None
            },
        }
    }

    /// Adds a blit without a source to an XOR bracket: a PATINVERT with a
    /// solid brush is an XOR fill, and another brush raster operation right
    /// after a first XOR pass is the middle fill. Returns `None` when the
    /// blit is not part of a bracket.
    fn bracket_blit(
        &mut self,
        rop3: Rop3,
        rect: DeviceRect,
        outline: Vec<(i64, i64)>,
    ) -> Option<Option<String>> {
        let shape = if rop3.table() == 0x5a {
            let color = self.solid_brush_color()? & 0x00ff_ffff;
            if color == 0 {
                // XOR with black keeps every pixel.
                return Some(None);
            }
            BracketShape::Xor(color)
        } else {
            self.xor_bracket.as_ref()?;
            let rop2 = rop3_as_rop2(rop3);
            if rop2.ignores_color() && rop2.paint(0) == Rop2Paint::Unchanged {
                return Some(None);
            }
            let brush_value = &self.state.brush;
            let middle = match (&brush_value.pattern, brush_value.style) {
                (None, brush::BS_SOLID) => XorMiddle::Solid {
                    rop2,
                    color: brush_value.color & 0x00ff_ffff,
                },
                (Some(Ok(pattern)), _) => XorMiddle::Pattern {
                    rop2,
                    pattern: Arc::clone(pattern),
                    text_color: self.state.text_color & 0x00ff_ffff,
                    background_color: self.state.bk_color & 0x00ff_ffff,
                },
                _ => return None,
            };
            if !self.accepts_middle(&middle) {
                return None;
            }
            BracketShape::Middle(middle, None)
        };
        let mut element = String::with_capacity(160);
        write_rect_start(&mut element, rect);
        element.push_str(&fill_attrs(XOR_FILL, self.state.poly_fill_mode));
        element.push_str(r#" stroke="none""#);
        element.push_str(&self.clip_attr());
        element.push_str("/>");
        let bounds = outline_bounds(&outline);
        let output = self.capture_bracket_shape(shape, element, Some(outline), Some(bounds));
        if self.halted {
            return Some(None);
        }
        Some(output.and_then(|output| self.paint_after_blits(output)))
    }

    /// Fails playback for an XOR fill that no bracket undoes.
    fn fail_xor_bracket(&mut self) {
        self.xor_bracket = None;
        self.fatal(
            "ROP2 mode 7 depends on destination pixels and cannot be faithfully represented in SVG"
                .to_owned(),
        );
    }

    /// Paints a complete XOR bracket: the middle fills as if no XOR fill
    /// covered them, then the part under each XOR fill. Returns `None` when
    /// a pixel would still depend on the destination.
    fn resolve_xor_bracket(&mut self, bracket: XorBracket) -> Option<String> {
        let paint = bracket.paint?;
        // Each pixel takes the colour of the one XOR fill covering it only
        // when no two fills overlap.
        if bracket.first.len() > 1 {
            let outlines = bracket
                .first
                .iter()
                .map(|shape| shape.outline.as_deref())
                .collect::<Option<Vec<_>>>()?;
            if outlines.len() > 4096
                || !self.charge_work(outlines.len().saturating_mul(outlines.len()) / 64)
                || convex_outlines_overlap(&outlines)
            {
                return None;
            }
        }
        let stroked = bracket.middles.iter().any(|middle| middle.stroke.is_some());
        // The middle fills outside every XOR fill. When the XOR fills tile a
        // rectangle that holds every middle fill nothing is left of them;
        // otherwise a mask keeps them from painting under the XOR fills,
        // which paint those pixels themselves.
        let mut output = String::new();
        if stroked || !xor_fills_cover(&bracket.first, &bracket.middles) {
            output = self.paint_middles(&bracket.middles, &paint, 0)?;
            if !output.is_empty() {
                let id = self.bracket_mask(&bracket.first);
                output = format!(r#"<g mask="url(#{id})">{output}</g>"#);
            }
        }
        if stroked {
            // Outlines reach past the middle fills, so the middle fills are
            // painted again under each XOR fill.
            if bracket.first.len() > MAX_STROKED_BRACKET_FILLS {
                return None;
            }
            for shape in &bracket.first {
                let middles = self.paint_middles(&bracket.middles, &paint, shape.color)?;
                if middles.is_empty() {
                    continue;
                }
                let id = self.bracket_clip(&shape.element);
                output.push_str(&format!(r#"<g clip-path="url(#{id})">{middles}</g>"#));
            }
            return Some(output);
        }
        let mut strips = String::new();
        for shape in &bracket.first {
            if let Some(fill) = self.bracket_region_paint(&paint, shape.color)? {
                strips.push_str(&shape.element.replace(XOR_FILL, &fill));
            }
        }
        if !strips.is_empty() {
            let geometry = bracket
                .middles
                .iter()
                .map(|middle| middle.element.as_str())
                .collect::<String>();
            let id = self.bracket_clip(&geometry);
            // The XOR fills tile without overlap; edges without antialiasing
            // keep the seams between them closed, as GDI draws them.
            output.push_str(&format!(
                r#"<g clip-path="url(#{id})"><g shape-rendering="crispEdges">{strips}</g></g>"#
            ));
        }
        Some(output)
    }

    /// The middle fills as they paint where the XOR fills total `xor`.
    fn paint_middles(
        &mut self,
        middles: &[MiddleShape],
        paint: &XorMiddle,
        xor: u32,
    ) -> Option<String> {
        let fill = self.bracket_region_paint(paint, xor)?;
        let mut output = String::new();
        for middle in middles {
            let stroke = match middle.stroke {
                Some((rop2, color)) => xor_bracket_paint(rop2, color, xor)?.map(color_paint),
                None => None,
            };
            if fill.is_none() && stroke.is_none() {
                continue;
            }
            output.push_str(
                &middle
                    .element
                    .replace(XOR_FILL, fill.as_deref().unwrap_or("none"))
                    .replace(XOR_STROKE, stroke.as_deref().unwrap_or("none")),
            );
        }
        Some(output)
    }

    /// Defines a mask that hides the XOR fills, returning its id.
    fn bracket_mask(&mut self, shapes: &[XorShape]) -> String {
        let id = self.next_id("wmf-xor-mask");
        let canvas = self.transform.canvas_rect();
        let mut definition = format!(r#"<mask id="{id}" maskUnits="userSpaceOnUse" x=""#);
        write_num(&mut definition, canvas.left);
        definition.push_str(r#"" y=""#);
        write_num(&mut definition, canvas.top);
        definition.push_str(r#"" width=""#);
        write_num(&mut definition, canvas.right - canvas.left);
        definition.push_str(r#"" height=""#);
        write_num(&mut definition, canvas.bottom - canvas.top);
        definition.push_str(r#"">"#);
        write_rect_start(&mut definition, canvas);
        definition.push_str(r##" fill="#ffffff"/>"##);
        for shape in shapes {
            definition.push_str(&shape.element.replace(XOR_FILL, "#000000"));
        }
        definition.push_str("</mask>");
        self.push_definition(&definition);
        id
    }

    /// Defines a clip path from bracket elements, returning its id.
    fn bracket_clip(&mut self, elements: &str) -> String {
        let id = self.next_id("wmf-xor-clip");
        let geometry = elements
            .replace(XOR_FILL, "#000000")
            .replace(XOR_STROKE, "none")
            .replace(" fill-rule=", " clip-rule=");
        self.push_definition(&format!(
            r#"<clipPath id="{id}" clipPathUnits="userSpaceOnUse">{geometry}</clipPath>"#
        ));
        id
    }

    /// The SVG paint of the middle region whose XOR fills total `xor`:
    /// `Some(None)` when the region keeps the destination, `None` when it
    /// still depends on it.
    fn bracket_region_paint(&mut self, middle: &XorMiddle, xor: u32) -> Option<Option<String>> {
        let resolve = |rop2: Rop2, color: u32| xor_bracket_paint(rop2, color, xor);
        match middle {
            XorMiddle::Solid { rop2, color } => Some(resolve(*rop2, *color)?.map(color_paint)),
            XorMiddle::Pattern {
                rop2,
                pattern,
                text_color,
                background_color,
            } => {
                let png = pattern
                    .mapped(*text_color, *background_color, |color| {
                        resolve(*rop2, color)
                    })?
                    .tile_png(0, 0, self.max_output_bytes)
                    .ok()?;
                if let Some(id) = self.pattern_ids.get(&png) {
                    return Some(Some(format!("url(#{id})")));
                }
                let id = self.next_id("wmf-pattern");
                let size = (f64::from(pattern.width()), f64::from(pattern.height()));
                let definition = BrushPattern::svg_definition(&id, &png, size, (0.0, 0.0)).ok()?;
                self.push_definition(&definition);
                self.pattern_ids.insert(png, id.clone());
                Some(Some(format!("url(#{id})")))
            },
        }
    }

    /// Background colour of hatched brushes, or `None` when transparent.
    fn hatch_background(&self) -> Option<u32> {
        (self.state.bk_mode == bk_mode::OPAQUE).then_some(self.state.bk_color)
    }

    /// `COLORREF` of the selected brush when it is solid.
    fn solid_brush_color(&self) -> Option<u32> {
        (self.state.brush.style == brush::BS_SOLID && self.state.brush.pattern.is_none())
            .then_some(self.state.brush.color)
    }

    /// The brush and hatch background that paint like `brush_value` under
    /// `rop2`, or `None` when the result depends on destination pixels.
    fn mixed_brush(&self, rop2: Rop2, brush_value: &Brush) -> Option<(Brush, Option<u32>)> {
        let background = self.hatch_background();
        if rop2 == Rop2::COPY_PEN || brush_value.style == brush::BS_NULL {
            return Some((brush_value.clone(), background));
        }
        let unchanged = Brush {
            style: brush::BS_NULL,
            ..Brush::default()
        };
        let solid = |color| Brush {
            style: brush::BS_SOLID,
            color,
            hatch: 0,
            pattern: None,
        };
        if let Some(Ok(pattern)) = &brush_value.pattern {
            // Every pixel of the pattern goes through the mix mode on its
            // own.
            let pattern = pattern.under_rop2(
                rop2,
                self.state.text_color & 0x00ff_ffff,
                self.state.bk_color & 0x00ff_ffff,
            )?;
            return Some((
                Brush {
                    pattern: Some(Ok(Arc::new(pattern))),
                    ..brush_value.clone()
                },
                None,
            ));
        }
        if brush_value.pattern.is_some()
            || !matches!(brush_value.style, brush::BS_SOLID | brush::BS_HATCHED)
        {
            // A brush without its pattern bitmap fails when it paints.
            return Some((brush_value.clone(), background));
        }
        let mix = |color| match rop2.paint(color) {
            Rop2Paint::Unchanged => Some(None),
            Rop2Paint::Color(color) => Some(Some(color)),
            Rop2Paint::DestinationDependent => None,
        };
        let foreground = mix(brush_value.color)?;
        if brush_value.style == brush::BS_SOLID {
            return Some(match foreground {
                Some(color) => (solid(color), None),
                None => (unchanged, None),
            });
        }
        let background = match background {
            Some(color) => mix(color)?,
            None => None,
        };
        match (foreground, background) {
            (Some(color), background) => Some((
                Brush {
                    color,
                    ..brush_value.clone()
                },
                background,
            )),
            (None, None) => Some((unchanged, None)),
            // A hatch whose lines keep the destination needs a pattern with
            // transparent lines, which the hatch definitions do not draw.
            (None, Some(_)) => None,
        }
    }

    /// The pen that paints like the selected pen under `rop2`, or `None`
    /// when the result depends on destination pixels.
    fn mixed_pen(&self, rop2: Rop2) -> Option<Pen> {
        let pen_value = &self.state.pen;
        if rop2 == Rop2::COPY_PEN || pen_value.style & pen::PS_STYLE_MASK == pen::PS_NULL {
            return Some(pen_value.clone());
        }
        match rop2.paint(pen_value.color) {
            Rop2Paint::Unchanged => Some(Pen {
                style: pen::PS_NULL,
                ..pen_value.clone()
            }),
            Rop2Paint::Color(color) => Some(Pen {
                color,
                ..pen_value.clone()
            }),
            Rop2Paint::DestinationDependent => None,
        }
    }

    /// Fill attributes of a brush whose hatch background is `background`.
    fn brush_fill(&mut self, brush_value: &Brush, background: Option<u32>) -> String {
        let paint = self.brush_paint(brush_value, background);
        fill_attrs(&paint, self.state.poly_fill_mode)
    }

    /// The SVG paint of a brush: `none`, a colour, or a pattern reference.
    fn brush_paint(&mut self, brush_value: &Brush, background: Option<u32>) -> String {
        match (&brush_value.pattern, brush_value.style) {
            (_, brush::BS_NULL) => "none".to_owned(),
            (Some(Ok(pattern)), _) => {
                let pattern = Arc::clone(pattern);
                match self.ensure_pattern(pattern.as_ref()) {
                    Some(id) => format!("url(#{id})"),
                    None => "none".to_owned(),
                }
            },
            (Some(Err(reason)), _) => {
                self.fatal(format!("WMF pattern brush has no usable bitmap: {reason}"));
                "none".to_owned()
            },
            (None, brush::BS_SOLID) => color_paint(brush_value.color),
            (None, brush::BS_HATCHED) => {
                format!("url(#{})", self.ensure_hatch(brush_value, background))
            },
            (None, style) => {
                self.fatal(format!(
                    "brush style {style} needs a pattern bitmap that its record does not carry"
                ));
                "none".to_owned()
            },
        }
    }

    /// Emits the SVG pattern of a bitmap brush.
    ///
    /// GDI tiles pattern brushes in device pixels from the device origin;
    /// one pattern pixel becomes one SVG pixel, as for hatch brushes.
    /// Monochrome patterns paint their 0 bits with the text colour and their
    /// 1 bits with the background colour.
    fn ensure_pattern(&mut self, pattern: &BrushPattern) -> Option<String> {
        let tile = pattern.tile_png(
            self.state.text_color & 0x00ff_ffff,
            self.state.bk_color & 0x00ff_ffff,
            self.max_output_bytes,
        );
        let png = match tile {
            Ok(png) => png,
            Err(error) => {
                self.fatal(format!("WMF brush pattern cannot be tiled: {error}"));
                return None;
            },
        };
        if let Some(id) = self.pattern_ids.get(&png) {
            return Some(id.clone());
        }
        let id = self.next_id("wmf-pattern");
        let size = (f64::from(pattern.width()), f64::from(pattern.height()));
        match BrushPattern::svg_definition(&id, &png, size, (0.0, 0.0)) {
            Ok(definition) => {
                self.push_definition(&definition);
                self.pattern_ids.insert(png, id.clone());
                Some(id)
            },
            Err(error) => {
                self.fatal(format!("WMF brush pattern cannot be tiled: {error}"));
                None
            },
        }
    }

    fn ensure_hatch(&mut self, brush_value: &Brush, background: Option<u32>) -> String {
        let key = (brush_value.color, brush_value.hatch, background);
        if let Some(id) = self.hatch_ids.get(&key) {
            return id.clone();
        }
        let id = self.next_id("wmf-hatch");
        let definition = hatch_definition(&id, brush_value, background);
        self.push_definition(&definition);
        self.hatch_ids.insert(key, id.clone());
        id
    }

    /// Fails playback for a mix mode whose result needs destination pixels.
    fn rop2_attr(&mut self) -> String {
        let mode = self.state.rop2;
        self.fatal(format!(
            "ROP2 mode {mode} depends on destination pixels and cannot be faithfully represented in SVG"
        ));
        r#" fill="none" stroke="none""#.to_owned()
    }

    fn clip_attr(&mut self) -> String {
        if self.state.clip.is_unbounded() {
            return String::new();
        }
        match self.clip_ids.get(&self.state.clip_revision) {
            Some(None) => return String::new(),
            Some(Some(id)) => return format!(r#" clip-path="url(#{id})""#),
            None => {},
        }
        let included_count = self.state.clip.rects.as_ref().map_or(1usize, Vec::len);
        let clip_work = included_count
            .checked_mul(self.state.clip.excluded.len().saturating_add(1))
            .unwrap_or(usize::MAX);
        if !self.charge_work(clip_work) {
            return String::new();
        }
        let id = self.next_id("wmf-clip");
        let mut path = String::with_capacity(256);
        let included: Vec<DeviceRect> = match &self.state.clip.rects {
            Some(rects) => rects
                .iter()
                .map(|rect| self.transform.rect(*rect))
                .collect(),
            None => vec![self.transform.canvas_rect()],
        };
        // A region that already covers the canvas clips nothing away. Dropping
        // the reference keeps one mask per drawn element out of the rasterizer,
        // which dominates the cost of a busy picture.
        let canvas = self.transform.canvas_rect();
        let covers_canvas = included
            .iter()
            .any(|rect| rect.intersect(canvas) == Some(canvas))
            && !self
                .state
                .clip
                .excluded
                .iter()
                .any(|rect| self.transform.rect(*rect).intersect(canvas).is_some());
        if covers_canvas {
            self.clip_ids.insert(self.state.clip_revision, None);
            return String::new();
        }
        for rect in &included {
            append_rect_path(&mut path, *rect);
        }
        // Only the part of an exclusion inside an included rectangle is a
        // hole. Emitting an outside exclusion as another even-odd subpath
        // would incorrectly enlarge the clip.
        for excluded in &self.state.clip.excluded {
            let excluded = self.transform.rect(*excluded);
            for included in &included {
                if let Some(hole) = included.intersect(excluded) {
                    append_rect_path(&mut path, hole);
                }
            }
        }
        let definition = format!(
            r#"<clipPath id="{id}" clipPathUnits="userSpaceOnUse"><path d="{path}" clip-rule="evenodd" fill-rule="evenodd"/></clipPath>"#
        );
        self.push_definition(&definition);
        self.clip_ids
            .insert(self.state.clip_revision, Some(id.clone()));
        format!(r#" clip-path="url(#{id})""#)
    }

    fn add_rect_clip(&mut self, rect: DeviceRect) -> String {
        let id = self.next_id("wmf-text-clip");
        let mut definition = format!(r#"<clipPath id="{id}"><rect x=""#);
        write_num(&mut definition, rect.left);
        definition.push_str(r#"" y=""#);
        write_num(&mut definition, rect.top);
        definition.push_str(r#"" width=""#);
        write_num(&mut definition, rect.right - rect.left);
        definition.push_str(r#"" height=""#);
        write_num(&mut definition, rect.bottom - rect.top);
        definition.push_str(r#""/></clipPath>"#);
        self.push_definition(&definition);
        id
    }

    fn next_id(&mut self, prefix: &str) -> String {
        let id = format!("{prefix}-{}", self.next_definition_id);
        self.next_definition_id += 1;
        id
    }

    fn region_path(&self, region: &Region) -> String {
        let mut path = String::with_capacity(region.rects.len() * 48);
        for rect in &region.rects {
            append_rect_path(&mut path, self.map_logical_rect_to_svg(*rect));
        }
        path
    }

    fn logical_rect(&self, left: i16, top: i16, right: i16, bottom: i16) -> DeviceRect {
        self.transform
            .rect(self.state.mapping.rect(left, top, right, bottom))
    }

    fn logical_origin_size(&self, x: i16, y: i16, width: i16, height: i16) -> DeviceRect {
        let (x, y) = self.state.mapping.point(x, y);
        let (width, height) = self
            .state
            .mapping
            .vector(f64::from(width), f64::from(height));
        self.transform
            .rect(DeviceRect::new(x, y, x + width, y + height))
    }

    fn map_logical_rect_to_device(&self, rect: DeviceRect) -> DeviceRect {
        let (x1, y1) = self.state.mapping.point_f64(rect.left, rect.top);
        let (x2, y2) = self.state.mapping.point_f64(rect.right, rect.bottom);
        DeviceRect::new(x1, y1, x2, y2)
    }

    fn map_logical_rect_to_svg(&self, rect: DeviceRect) -> DeviceRect {
        self.transform.rect(self.map_logical_rect_to_device(rect))
    }

    fn require_box(
        &mut self,
        function: u16,
        rec: &WmfRecord,
        offset: usize,
    ) -> Option<(i16, i16, i16, i16)> {
        let value = read_box(&rec.params, offset);
        if value.is_none() {
            self.malformed(function, "bounding rectangle");
        }
        value
    }

    fn malformed(&mut self, function: u16, field: &str) {
        self.fatal(format!(
            "malformed WMF record 0x{function:04X}: missing or invalid {field}"
        ));
    }

    fn fatal(&mut self, message: String) {
        self.halted = true;
        self.issues.push(RenderIssue {
            fatal: true,
            code: PLAYBACK_WARNING,
            message,
        });
    }

    fn push_definition(&mut self, definition: &str) {
        if self
            .definitions
            .len()
            .checked_add(definition.len())
            .is_some_and(|size| size <= self.max_output_bytes)
        {
            self.definitions.push_str(definition);
        } else {
            self.fatal(format!(
                "WMF SVG definitions exceed limit {} bytes",
                self.max_output_bytes
            ));
        }
    }

    fn reserve_object_bytes(&mut self, additional: usize) -> bool {
        let fits = self.can_retain_additional_state(additional);
        if !fits {
            self.fatal(format!(
                "WMF retained object data exceeds limit {} bytes",
                self.max_object_bytes
            ));
        }
        fits
    }

    fn can_retain_additional_state(&self, additional: usize) -> bool {
        self.saved_state_bytes_with_current()
            .checked_add(additional)
            .is_some_and(|bytes| bytes <= self.max_object_bytes)
    }

    fn saved_state_bytes_with_current(&self) -> usize {
        self.saved_states
            .iter()
            .fold(self.objects.retained_heap_bytes(), |bytes, state| {
                bytes.saturating_add(state.retained_heap_bytes())
            })
            .saturating_add(self.state.retained_heap_bytes())
    }

    fn bump_clip_revision(&mut self) {
        self.state.clip_revision = self.next_clip_revision;
        self.next_clip_revision = self.next_clip_revision.saturating_add(1);
    }

    fn charge_work(&mut self, amount: usize) -> bool {
        if amount > self.remaining_work {
            self.fatal("WMF cumulative geometry work exceeds configured path limit".to_owned());
            false
        } else {
            self.remaining_work -= amount;
            true
        }
    }

    /// Records an approximation of the source picture.
    fn warn(&mut self, message: String) {
        self.note(PLAYBACK_WARNING, message);
    }

    /// Records a non-fatal issue under a stable diagnostic code.
    ///
    /// Use a `noncanonical-` code when Windows GDI playback tolerates the
    /// input without changing the picture; any other code reports an
    /// approximation that strict conversion rejects.
    fn note(&mut self, code: &'static str, message: String) {
        if !self
            .issues
            .iter()
            .any(|issue| issue.code == code && issue.message == message)
        {
            self.issues.push(RenderIssue {
                fatal: false,
                code,
                message,
            });
        }
    }
}

fn record_point_count(function: u16, params: &[u8]) -> Option<usize> {
    match function {
        record::POLYGON | record::POLYLINE => Some(usize::from(read_u16_le(params, 0).ok()?)),
        record::POLYPOLYGON => {
            let polygons = usize::from(read_u16_le(params, 0).ok()?);
            (0..polygons).try_fold(0usize, |total, index| {
                let offset = 2usize.checked_add(index.checked_mul(2)?)?;
                total.checked_add(usize::from(read_u16_le(params, offset).ok()?))
            })
        },
        _ => None,
    }
}

fn read_u32(data: &[u8], offset: usize) -> Option<u32> {
    let bytes = data.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes(bytes.try_into().ok()?))
}

/// Records model Win16's pushed parameter order: Y precedes X.
fn read_yx(data: &[u8]) -> Option<(i16, i16)> {
    Some((read_i16_le(data, 2).ok()?, read_i16_le(data, 0).ok()?))
}

/// Rectangle parameters used by shape/state records are bottom, right, top,
/// left in the record's parameter array.
fn read_box(data: &[u8], offset: usize) -> Option<(i16, i16, i16, i16)> {
    Some((
        read_i16_le(data, offset + 6).ok()?,
        read_i16_le(data, offset + 4).ok()?,
        read_i16_le(data, offset + 2).ok()?,
        read_i16_le(data, offset).ok()?,
    ))
}

/// Writes space-separated numbers, as SVG coordinate lists take them.
fn write_num_list(output: &mut String, values: impl Iterator<Item = f64>) {
    for (index, value) in values.enumerate() {
        if index != 0 {
            output.push(' ');
        }
        write_num(output, value);
    }
}

fn write_rect_start(output: &mut String, rect: DeviceRect) {
    output.push_str(r#"<rect x=""#);
    write_num(output, rect.left);
    output.push_str(r#"" y=""#);
    write_num(output, rect.top);
    output.push_str(r#"" width=""#);
    write_num(output, rect.right - rect.left);
    output.push_str(r#"" height=""#);
    write_num(output, rect.bottom - rect.top);
    output.push('"');
}

fn append_rect_path(output: &mut String, rect: DeviceRect) {
    output.push('M');
    write_num(output, rect.left);
    output.push(',');
    write_num(output, rect.top);
    output.push('H');
    write_num(output, rect.right);
    output.push('V');
    write_num(output, rect.bottom);
    output.push('H');
    write_num(output, rect.left);
    output.push('Z');
}

fn positive_angle(angle: f64) -> f64 {
    angle.rem_euclid(std::f64::consts::TAU)
}

fn parse_region(data: &[u8], max_rects: usize) -> Option<Region> {
    if data.len() < 22 || read_u16_le(data, 2).ok()? != 6 {
        return None;
    }
    let scan_count = usize::from(read_u16_le(data, 10).ok()?);
    let mut offset = 22usize;
    let mut rects = Vec::new();
    for _ in 0..scan_count {
        let count = usize::from(read_u16_le(data, offset).ok()?);
        if count % 2 != 0 || data.len() < offset.checked_add(8 + count * 2)? {
            return None;
        }
        let top = f64::from(read_i16_le(data, offset + 2).ok()?);
        let bottom = f64::from(read_i16_le(data, offset + 4).ok()?);
        for pair in 0..count / 2 {
            if rects.len() >= max_rects {
                return None;
            }
            let left = f64::from(read_i16_le(data, offset + 6 + pair * 4).ok()?);
            let right = f64::from(read_i16_le(data, offset + 8 + pair * 4).ok()?);
            rects.push(DeviceRect::new(left, top, right, bottom));
        }
        let count2_offset = offset + 6 + count * 2;
        if read_u16_le(data, count2_offset).ok()? != count as u16 {
            return None;
        }
        offset = count2_offset + 2;
    }
    Some(Region { rects })
}

/// Whether `data` starts with the header size of a DIB header.
fn starts_with_dib_header(data: &[u8]) -> bool {
    read_u32(data, 0).is_some_and(|size| matches!(size, 12 | 40 | 52 | 56 | 64 | 108 | 124))
}

/// Decodes a monochrome Bitmap16 ([MS-WMF] 2.2.2.1) pattern whose bits start
/// `header_len` bytes into `data`.
fn bitmap16_pattern(
    data: &[u8],
    header_len: usize,
    limits: DibLimits,
) -> Result<BrushPattern, String> {
    let (Ok(width), Ok(height), Ok(stride)) = (
        read_i16_le(data, 2),
        read_i16_le(data, 4),
        read_i16_le(data, 6),
    ) else {
        return Err("pattern bitmap header is truncated".to_owned());
    };
    if data.get(8..10) != Some(&[1, 1]) {
        return Err("only monochrome Bitmap16 patterns are supported".to_owned());
    }
    let (Ok(width), Ok(height), Ok(stride)) = (
        u32::try_from(width),
        u32::try_from(height),
        usize::try_from(stride),
    ) else {
        return Err("pattern bitmap geometry is invalid".to_owned());
    };
    if u64::from(width) * u64::from(height) > limits.max_pixels {
        return Err("pattern bitmap exceeds the configured pixel limit".to_owned());
    }
    let bits = data
        .get(header_len..)
        .ok_or_else(|| "pattern bitmap bits are truncated".to_owned())?;
    // Device-dependent bitmaps store their rows from the top.
    BrushPattern::from_monochrome_bits(width, height, stride, bits, true)
        .map_err(|error| error.to_string())
}

/// Whether a record may occur inside an XOR bracket: it paints nothing,
/// paints a shape whose fill the bracket inspects, or paints the brush over
/// a rectangle.
fn bracket_compatible(rec: &WmfRecord) -> bool {
    match record::canonical(rec.function) {
        record::PAT_BLT => true,
        record::BIT_BLT | record::STRETCH_BLT | record::DIB_BIT_BLT | record::DIB_STRETCH_BLT => {
            rec.size == u32::from(rec.function >> 8) + 3
        },
        record::TEXT_OUT
        | record::EXT_TEXT_OUT
        | record::SET_PIXEL
        | record::FLOOD_FILL
        | record::EXT_FLOOD_FILL
        | record::INVERT_REGION
        | record::SET_DIB_TO_DEV
        | record::STRETCH_DIB => false,
        _ => true,
    }
}

/// The mix mode that a ROP3 without a source applies to its brush.
fn rop3_as_rop2(rop3: Rop3) -> Rop2 {
    // The ROP2 truth table is indexed by (P, D); the ROP3 one by (P, S, D).
    let table = rop3.table();
    let mode = (table & 0b11) | ((table >> 2) & 0b1100);
    Rop2::from_mode(u32::from(mode) + 1).unwrap_or(Rop2::COPY_PEN)
}

/// The corners of a logical `(x, y, width, height)` rectangle.
fn rect_outline(x: i32, y: i32, width: i32, height: i32) -> Vec<(i64, i64)> {
    let (left, top) = (i64::from(x), i64::from(y));
    let (right, bottom) = (left + i64::from(width), top + i64::from(height));
    vec![(left, top), (right, top), (right, bottom), (left, bottom)]
}

/// Whether the XOR fills tile the rectangle that bounds them and that
/// rectangle holds every middle fill.
fn xor_fills_cover(fills: &[XorShape], middles: &[MiddleShape]) -> bool {
    let Some(outlines) = fills
        .iter()
        .map(|fill| fill.outline.as_deref())
        .collect::<Option<Vec<_>>>()
    else {
        return false;
    };
    let bounds = outlines
        .iter()
        .map(|outline| outline_bounds(outline))
        .reduce(|a, b| {
            [
                a[0].min(b[0]),
                a[1].min(b[1]),
                a[2].max(b[2]),
                a[3].max(b[3]),
            ]
        });
    let Some(bounds) = bounds else {
        return false;
    };
    // Disjoint fills tile their bounds exactly when their areas add up to
    // its area; both sides are doubled to stay in integers.
    let area = outlines.iter().try_fold(0i64, |total, outline| {
        total.checked_add(doubled_area(outline))
    });
    let bounds_area = (bounds[2] - bounds[0])
        .checked_mul(bounds[3] - bounds[1])
        .and_then(|value| value.checked_mul(2));
    if area.is_none() || area != bounds_area {
        return false;
    }
    middles.iter().all(|middle| {
        middle.bounds.is_some_and(|inner| {
            inner[0] >= bounds[0]
                && inner[1] >= bounds[1]
                && inner[2] <= bounds[2]
                && inner[3] <= bounds[3]
        })
    })
}

/// Twice the area of a simple polygon outline.
fn doubled_area(outline: &[(i64, i64)]) -> i64 {
    let mut sum = 0i64;
    for index in 0..outline.len() {
        let (x0, y0) = outline[index];
        let (x1, y1) = outline[(index + 1) % outline.len()];
        sum = sum.saturating_add(x0 * y1 - x1 * y0);
    }
    sum.abs()
}

/// The bounds `[left, top, right, bottom]` of an outline.
fn outline_bounds(outline: &[(i64, i64)]) -> [i64; 4] {
    outline.iter().fold(
        [i64::MAX, i64::MAX, i64::MIN, i64::MIN],
        |[left, top, right, bottom], &(x, y)| {
            [left.min(x), top.min(y), right.max(x), bottom.max(y)]
        },
    )
}

/// The logical bounds of a shape record's geometry.
fn logical_bounds(function: u16, params: &[u8]) -> Option<[i64; 4]> {
    let point = |offset: usize| {
        Some((
            i64::from(read_i16_le(params, offset).ok()?),
            i64::from(read_i16_le(params, offset + 2).ok()?),
        ))
    };
    let points = match function {
        record::RECTANGLE | record::ELLIPSE => {
            let (left, top, right, bottom) = read_box(params, 0)?;
            vec![
                (i64::from(left), i64::from(top)),
                (i64::from(right), i64::from(bottom)),
            ]
        },
        record::ROUND_RECT => {
            let (left, top, right, bottom) = read_box(params, 4)?;
            vec![
                (i64::from(left), i64::from(top)),
                (i64::from(right), i64::from(bottom)),
            ]
        },
        record::POLYGON => {
            let count = usize::from(read_u16_le(params, 0).ok()?);
            (0..count)
                .map(|index| point(2 + index * 4))
                .collect::<Option<Vec<_>>>()?
        },
        record::POLYPOLYGON => {
            let polygons = usize::from(read_u16_le(params, 0).ok()?);
            let total = (0..polygons).try_fold(0usize, |total, index| {
                total.checked_add(usize::from(read_u16_le(params, 2 + index * 2).ok()?))
            })?;
            let start = 2 + polygons * 2;
            (0..total)
                .map(|index| point(start + index * 4))
                .collect::<Option<Vec<_>>>()?
        },
        _ => return None,
    };
    (!points.is_empty()).then(|| outline_bounds(&points))
}

/// The convex outline of a polygon or rectangle record in logical
/// coordinates, or `None` for other records and non-convex polygons.
fn logical_outline(function: u16, params: &[u8]) -> Option<Vec<(i64, i64)>> {
    let outline = match function {
        record::RECTANGLE => {
            let (left, top, right, bottom) = read_box(params, 0)?;
            let (left, top, right, bottom) = (
                i64::from(left),
                i64::from(top),
                i64::from(right),
                i64::from(bottom),
            );
            vec![(left, top), (right, top), (right, bottom), (left, bottom)]
        },
        record::POLYGON => {
            let count = usize::from(read_u16_le(params, 0).ok()?);
            (0..count)
                .map(|index| {
                    Some((
                        i64::from(read_i16_le(params, 2 + index * 4).ok()?),
                        i64::from(read_i16_le(params, 4 + index * 4).ok()?),
                    ))
                })
                .collect::<Option<Vec<_>>>()?
        },
        _ => return None,
    };
    is_convex(&outline).then_some(outline)
}

/// Wraps an element in a group with a clip-path attribute, if any.
fn wrap_clip(clip: &str, element: &str) -> String {
    if clip.is_empty() {
        return element.to_owned();
    }
    let mut output = String::with_capacity(element.len() + clip.len() + 8);
    output.push_str("<g");
    output.push_str(clip);
    output.push('>');
    output.push_str(element);
    output.push_str("</g>");
    output
}

/// Stretch policy of a SetStretchBltMode value.
fn stretch_policy(mode: u16) -> StretchPolicy {
    match mode {
        1 => StretchPolicy::BlackOnWhite,
        2 => StretchPolicy::WhiteOnBlack,
        3 => StretchPolicy::ColorOnColor,
        _ => StretchPolicy::Halftone,
    }
}

/// The SVG destination rectangle of a bitmap record, and whether the
/// mapping mode mirrors it horizontally and vertically in device space.
fn bitmap_destination(
    rec: &WmfRecord,
    mapping: &MappingState,
    transform: &CoordinateTransform,
) -> Option<(DeviceRect, (bool, bool))> {
    let function = record::canonical(rec.function);
    // [MS-WMF] 2.3.1: blits without a source bitmap carry a reserved field
    // before the destination extent.
    let no_source = rec.size == u32::from(rec.function >> 8) + 3;
    let reserved = usize::from(no_source) * 2;
    let (height_offset, width_offset, y_offset, x_offset) = match function {
        record::BIT_BLT | record::DIB_BIT_BLT => {
            (8 + reserved, 10 + reserved, 12 + reserved, 14 + reserved)
        },
        record::STRETCH_BLT | record::DIB_STRETCH_BLT => {
            (12 + reserved, 14 + reserved, 16 + reserved, 18 + reserved)
        },
        record::SET_DIB_TO_DEV => (10, 12, 14, 16),
        record::STRETCH_DIB => (14, 16, 18, 20),
        _ => return None,
    };
    let data = &rec.params;
    let height = read_i16_le(data, height_offset).ok()?;
    let width = read_i16_le(data, width_offset).ok()?;
    let y = read_i16_le(data, y_offset).ok()?;
    let x = read_i16_le(data, x_offset).ok()?;
    let (x1, y1) = mapping.point(x, y);
    let (dx, dy) = mapping.vector(f64::from(width), f64::from(height));
    Some((
        transform.rect(DeviceRect::new(x1, y1, x1 + dx, y1 + dy)),
        (dx < 0.0, dy < 0.0),
    ))
}

#[cfg(test)]
mod tests {
    use bytes::Bytes;

    use super::*;

    fn rec(function: u16, params: &[u8]) -> WmfRecord {
        WmfRecord {
            size: 3 + params.len().div_ceil(2) as u32,
            function,
            params: Bytes::copy_from_slice(params),
        }
    }

    fn font_params(charset: u8, face: &str) -> Vec<u8> {
        let mut params = vec![0u8; 18];
        params[0..2].copy_from_slice(&(-12_i16).to_le_bytes());
        params[13] = charset;
        let mut name = [0u8; 32];
        name[..face.len()].copy_from_slice(face.as_bytes());
        params.extend_from_slice(&name);
        params
    }

    fn renderer() -> SvgRenderer<'static> {
        renderer_with(CoordinateTransform::new(
            (0.0, 0.0, 100.0, 100.0),
            100.0,
            100.0,
        ))
    }

    fn renderer_with(transform: CoordinateTransform) -> SvgRenderer<'static> {
        SvgRenderer::new(
            transform,
            None,
            DibLimits::default(),
            crate::Limits::default().max_state_depth,
            crate::Limits::default().max_objects,
            crate::Limits::default().max_path_points,
            crate::Limits::default().max_output_bytes,
            crate::Limits::default().max_uncompressed_bytes,
        )
    }

    #[test]
    fn line_to_draws_from_old_position_then_updates_current_position() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::MOVE_TO, &[20, 0, 10, 0]));
        let first = renderer
            .render_record(&rec(record::LINE_TO, &[40, 0, 30, 0]))
            .unwrap();
        let second = renderer
            .render_record(&rec(record::LINE_TO, &[60, 0, 50, 0]))
            .unwrap();
        assert!(first.contains(r#"x1="10" y1="20" x2="30" y2="40""#));
        assert!(second.contains(r#"x1="30" y1="40" x2="50" y2="60""#));
    }

    #[test]
    fn save_restore_restores_selected_drawing_state() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::SET_TEXT_COLOR, &[1, 2, 3, 0]));
        renderer.render_record(&rec(record::SAVE_DC, &[]));
        renderer.render_record(&rec(record::SET_TEXT_COLOR, &[4, 5, 6, 0]));
        renderer.render_record(&rec(record::RESTORE_DC, &[0xff, 0xff]));
        assert_eq!(renderer.state.text_color, 0x0003_0201);
    }

    #[test]
    fn clip_that_covers_the_canvas_needs_no_reference() {
        let mut renderer = renderer();
        renderer.state.clip =
            ClipRegion::from_rects(vec![DeviceRect::new(-5.0, -5.0, 110.0, 110.0)]);

        assert_eq!(renderer.clip_attr(), "");
        assert!(renderer.definitions.is_empty());
    }

    #[test]
    fn clip_that_cuts_into_the_canvas_keeps_its_reference() {
        let mut renderer = renderer();
        renderer.state.clip = ClipRegion::from_rects(vec![DeviceRect::new(0.0, 0.0, 50.0, 50.0)]);

        assert_eq!(renderer.clip_attr(), r#" clip-path="url(#wmf-clip-0)""#);

        renderer
            .state
            .clip
            .excluded
            .push(DeviceRect::new(10.0, 10.0, 20.0, 20.0));
        renderer.bump_clip_revision();
        assert_eq!(renderer.clip_attr(), r#" clip-path="url(#wmf-clip-1)""#);
    }

    #[test]
    fn saved_and_selected_clip_copies_obey_aggregate_heap_limit() {
        let rect = DeviceRect::new(0.0, 0.0, 10.0, 10.0);

        let mut save = renderer();
        save.state.clip.excluded.push(rect);
        save.max_object_bytes = save.saved_state_bytes_with_current();
        save.render_record(&rec(record::SAVE_DC, &[]));
        assert!(save.halted);
        assert!(save.saved_states.is_empty());

        let mut select = renderer();
        select.state.clip.excluded.push(rect);
        // Index 0 selects no region, so the region takes the second slot.
        select.objects.insert(GdiObject::Pen(Pen::default()));
        let handle = select
            .objects
            .insert(GdiObject::Region(Region { rects: vec![rect] }));
        select.max_object_bytes = select
            .saved_state_bytes_with_current()
            .checked_add(size_of::<DeviceRect>())
            .unwrap()
            - 1;
        select.render_record(&rec(
            record::SELECT_CLIP_REGION,
            &u16::try_from(handle).unwrap().to_le_bytes(),
        ));
        assert!(select.halted);
        assert_eq!(select.state.clip.excluded.len(), 1);

        let clip_params = [10, 0, 10, 0, 0, 0, 0, 0];
        let mut exclude = renderer();
        exclude.max_object_bytes = exclude.saved_state_bytes_with_current();
        exclude.render_record(&rec(record::EXCLUDE_CLIP_RECT, &clip_params));
        assert!(exclude.halted);
        assert!(exclude.state.clip.excluded.is_empty());

        let mut intersect = renderer();
        intersect.max_object_bytes = intersect.saved_state_bytes_with_current();
        intersect.render_record(&rec(record::INTERSECT_CLIP_RECT, &clip_params));
        assert!(intersect.halted);
        assert!(intersect.state.clip.rects.is_none());
    }

    #[test]
    fn set_pixel_uses_its_color_and_does_not_move_current_position() {
        let mut renderer = renderer();
        renderer.state.position = (9, 8);
        let output = renderer
            .render_record(&rec(
                record::SET_PIXEL,
                &[0x11, 0x22, 0x33, 0, 20, 0, 10, 0],
            ))
            .unwrap();
        assert!(output.contains(r##"fill="#112233""##));
        assert_eq!(renderer.state.position, (9, 8));
    }

    /// A bottom-up DIB with one 24-bpp row per entry of `rows`, listed from
    /// the top.
    fn rgb_dib(rows: &[&[u32]]) -> Vec<u8> {
        let width = rows[0].len();
        let stride = (width * 3).div_ceil(4) * 4;
        let mut d = vec![0; 40];
        d[0..4].copy_from_slice(&40u32.to_le_bytes());
        d[4..8].copy_from_slice(&i32::try_from(width).unwrap().to_le_bytes());
        d[8..12].copy_from_slice(&i32::try_from(rows.len()).unwrap().to_le_bytes());
        d[12..14].copy_from_slice(&1u16.to_le_bytes());
        d[14..16].copy_from_slice(&24u16.to_le_bytes());
        for row in rows.iter().rev() {
            let mut line = vec![0; stride];
            for (index, color) in row.iter().enumerate() {
                let [red, green, blue, _] = color.to_le_bytes();
                line[index * 3..index * 3 + 3].copy_from_slice(&[blue, green, red]);
            }
            d.extend_from_slice(&line);
        }
        d
    }

    /// A DIBBitBlt record copying all of `dib` to the destination rectangle.
    fn dib_bitblt_of(
        dib: &[u8],
        rop3: u32,
        left: i16,
        top: i16,
        width: i16,
        height: i16,
    ) -> WmfRecord {
        let mut params = Vec::new();
        params.extend_from_slice(&rop3.to_le_bytes());
        // SourceY, SourceX, Height, Width, DestinationY, DestinationX.
        for value in [0i16, 0, height, width, top, left] {
            params.extend_from_slice(&value.to_le_bytes());
        }
        params.extend_from_slice(dib);
        rec(record::DIB_BIT_BLT, &params)
    }

    fn dib_bitblt(rop3: u32, left: i16, top: i16, width: i16, height: i16) -> WmfRecord {
        dib_bitblt_of(&rgb_dib(&[&[0x0003_0201]]), rop3, left, top, width, height)
    }

    const RECTANGLE: [u8; 8] = [10, 0, 20, 0, 0, 0, 0, 0];
    const PATCOPY: u32 = 0x00f0_0021;
    const SRCAND: u32 = 0x0088_00c6;
    const SRCPAINT: u32 = 0x00ee_0086;

    fn fatal(renderer: &SvgRenderer<'_>) -> bool {
        renderer.issues.iter().any(|issue| issue.fatal)
    }

    fn create_solid_brush(renderer: &mut SvgRenderer<'_>, color: u32) {
        let mut params = vec![0, 0];
        params.extend_from_slice(&color.to_le_bytes());
        params.extend_from_slice(&[0, 0]);
        renderer.render_record(&rec(record::CREATE_BRUSH_INDIRECT, &params));
    }

    #[test]
    fn rop2_nop_paints_nothing_without_failing_playback() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::SET_ROP2, &[11, 0]));
        assert!(
            renderer
                .render_record(&rec(record::RECTANGLE, &RECTANGLE))
                .is_none()
        );
        // LineTo still moves the current position.
        assert!(
            renderer
                .render_record(&rec(record::LINE_TO, &[40, 0, 30, 0]))
                .is_none()
        );
        assert_eq!(renderer.state.position, (30, 40));
        assert!(renderer.issues.is_empty());
    }

    #[test]
    fn rop2_modes_that_ignore_the_destination_paint_exact_colours() {
        let paint = |mode: u8| {
            let mut renderer = renderer();
            renderer.render_record(&rec(record::SET_ROP2, &[mode, 0]));
            let output = renderer.render_record(&rec(record::RECTANGLE, &RECTANGLE));
            assert!(renderer.issues.is_empty(), "{:?}", renderer.issues);
            output.unwrap()
        };
        // Default white brush and black pen.
        let black = paint(1);
        assert!(black.contains(r##"fill="#000000""##) && black.contains(r##"stroke="#000000""##));
        let white = paint(16);
        assert!(white.contains(r##"fill="#ffffff""##) && white.contains(r##"stroke="#ffffff""##));
        let inverted = paint(4);
        assert!(inverted.contains(r##"fill="#000000""##));
        assert!(inverted.contains(r##"stroke="#ffffff""##));
        // R2_MASKPEN keeps the destination under the white brush and paints
        // the black pen; R2_XORPEN does the opposite.
        let masked = paint(9);
        assert!(masked.contains(r#"fill="none""#) && masked.contains(r##"stroke="#000000""##));
    }

    #[test]
    fn rop2_mode_that_combines_with_the_destination_fails_only_when_painting() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::SET_ROP2, &[6, 0]));
        renderer.render_record(&rec(record::SET_ROP2, &[13, 0]));
        assert!(
            renderer
                .render_record(&rec(record::RECTANGLE, &RECTANGLE))
                .is_some()
        );
        // R2_NOT inverts the destination under any pen and brush.
        renderer.render_record(&rec(record::SET_ROP2, &[6, 0]));
        renderer.render_record(&rec(record::RECTANGLE, &RECTANGLE));
        assert!(fatal(&renderer));
    }

    #[test]
    fn undefined_rop2_mode_keeps_the_current_mode() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::SET_ROP2, &[0, 0]));
        assert_eq!(renderer.state.rop2, 13);
        assert!(
            renderer
                .issues
                .iter()
                .any(|issue| issue.code == "noncanonical-wmf-invalid-rop2")
        );
    }

    #[test]
    fn select_clip_region_zero_removes_the_clip_region() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::INTERSECT_CLIP_RECT, &RECTANGLE));
        assert!(!renderer.state.clip.is_unbounded());
        renderer.render_record(&rec(record::SELECT_CLIP_REGION, &[0, 0]));
        assert!(renderer.state.clip.is_unbounded());
        assert!(renderer.issues.is_empty());

        // An index naming another kind of object leaves the clip unchanged.
        create_solid_brush(&mut renderer, 0);
        create_solid_brush(&mut renderer, 0);
        renderer.render_record(&rec(record::INTERSECT_CLIP_RECT, &RECTANGLE));
        renderer.render_record(&rec(record::SELECT_CLIP_REGION, &[1, 0]));
        assert!(!renderer.state.clip.is_unbounded());
        assert!(!fatal(&renderer));
        // An empty slot passes no region either.
        renderer.render_record(&rec(record::SELECT_CLIP_REGION, &[7, 0]));
        assert!(renderer.state.clip.is_unbounded());
    }

    #[test]
    fn patcopy_blit_fills_the_destination_with_the_brush() {
        let mut renderer = renderer();
        create_solid_brush(&mut renderer, 0x0000_00ff);
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = renderer
            .render_record(&dib_bitblt(PATCOPY, 5, 6, 20, 10))
            .unwrap_or_else(|| panic!("{:?}", renderer.issues));
        assert!(output.starts_with(r#"<rect x="5" y="6" width="20" height="10""#));
        assert!(output.contains(r##"fill="#ff0000""##));
        assert!(renderer.issues.is_empty());
    }

    #[test]
    fn zero_size_bitmap_destination_paints_nothing() {
        let mut renderer = renderer();
        assert!(
            renderer
                .render_record(&dib_bitblt(PATCOPY, 5, 6, 0, 10))
                .is_none()
        );
        assert!(!fatal(&renderer));
        assert!(
            renderer
                .issues
                .iter()
                .any(|issue| issue.code == "noncanonical-wmf-empty-bitmap-record")
        );
    }

    #[test]
    fn source_copies_wait_for_the_next_drawing_and_keep_their_order() {
        let mut renderer = renderer();
        let image = rgb_dib(&[&[0x0000_00ff]]);
        assert!(
            renderer
                .render_record(&dib_bitblt_of(&image, 0x00cc_0020, 0, 0, 10, 10))
                .is_none()
        );
        let output = renderer
            .render_record(&rec(record::RECTANGLE, &RECTANGLE))
            .unwrap();
        let image_at = output.find("<image").unwrap();
        assert!(image_at < output.find("<rect").unwrap());
        assert!(renderer.finish().is_none());
    }

    #[test]
    fn mask_blits_compose_into_one_exact_image() {
        let mut composed = renderer();
        let mask = rgb_dib(&[&[0, 0x00ff_ffff]]);
        let image = rgb_dib(&[&[0x0000_ff00, 0]]);
        assert!(
            composed
                .render_record(&dib_bitblt_of(&mask, SRCAND, 10, 10, 2, 1))
                .is_none()
        );
        assert!(
            composed
                .render_record(&dib_bitblt_of(&image, SRCPAINT, 10, 10, 2, 1))
                .is_none()
        );
        let output = composed.finish().unwrap();
        assert_eq!(output.matches("<image").count(), 1);
        assert!(output.contains(r#"x="10" y="10" width="2" height="1""#));
        assert!(composed.issues.is_empty(), "{:?}", composed.issues);

        // A lone SRCPAINT of a coloured image depends on the destination.
        let mut lone = renderer();
        lone.render_record(&dib_bitblt_of(&image, SRCPAINT, 10, 10, 2, 1));
        assert!(lone.finish().is_some());
        assert!(
            lone.issues
                .iter()
                .any(|issue| issue.code == "wmf-raster-operation-approximated")
        );
    }

    #[test]
    fn dib_pattern_brush_with_bs_pattern_style_tiles_its_dib() {
        let mut renderer = renderer();
        let mut params = vec![3, 0, 0, 0];
        params.extend_from_slice(&rgb_dib(&[&[0x0000_00ff, 0], &[0, 0x0000_00ff]]));
        renderer.render_record(&rec(record::CREATE_DIB_PATTERN_BRUSH, &params));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = renderer
            .render_record(&rec(record::RECTANGLE, &RECTANGLE))
            .unwrap();
        assert!(output.contains(r#"fill="url(#wmf-pattern-"#));
        assert!(
            renderer
                .definitions
                .contains(r#"patternUnits="userSpaceOnUse" x="0" y="0" width="2" height="2""#)
        );
        assert!(renderer.issues.is_empty(), "{:?}", renderer.issues);
    }

    #[test]
    fn mapping_mode_mirroring_flips_the_drawn_bitmap() {
        let mut renderer = renderer();
        // Window extent (-10, 10) mirrors x under MM_ANISOTROPIC.
        renderer.render_record(&rec(record::SET_MAP_MODE, &[8, 0]));
        renderer.render_record(&rec(record::SET_WINDOW_EXT, &[10, 0, 0xf6, 0xff]));
        renderer.render_record(&rec(record::SET_VIEWPORT_EXT, &[10, 0, 10, 0]));
        let image = rgb_dib(&[&[0x0000_00ff, 0x0000_ff00]]);
        renderer.render_record(&dib_bitblt_of(&image, 0x00cc_0020, 0, 0, 2, 1));
        let blit = renderer.deferred_blits.as_ref().unwrap().blits[0].clone();
        assert_eq!(blit.mirrored, (true, false));
        let record = blit.record.unwrap();
        let operation = renderer.parse_deferred(&record, blit.mirrored).unwrap();
        assert!(operation.blit_window().unwrap().unwrap().flip_x);
    }

    /// A polygon record through the given logical points.
    fn polygon(points: &[(i16, i16)]) -> WmfRecord {
        let mut params = u16::try_from(points.len()).unwrap().to_le_bytes().to_vec();
        for (x, y) in points {
            params.extend_from_slice(&x.to_le_bytes());
            params.extend_from_slice(&y.to_le_bytes());
        }
        rec(record::POLYGON, &params)
    }

    /// Selects a null pen and a new solid brush, returning the brush index.
    fn select_null_pen_and_brush(renderer: &mut SvgRenderer<'_>, color: u32) {
        renderer.render_record(&rec(record::SELECT_OBJECT, &[8, 0x80]));
        create_solid_brush(renderer, color);
        let index = u16::try_from(renderer.objects.retained_len() - 1).unwrap();
        renderer.render_record(&rec(record::SELECT_OBJECT, &index.to_le_bytes()));
    }

    fn xor_bracket(renderer: &mut SvgRenderer<'_>, strips: &[((i16, i16), u32)]) -> String {
        let mut output = String::new();
        let mut pass = |renderer: &mut SvgRenderer<'_>| {
            renderer.render_record(&rec(record::SET_ROP2, &[7, 0]));
            for ((left, right), color) in strips {
                select_null_pen_and_brush(renderer, *color);
                let strip = polygon(&[(*left, 0), (*right, 0), (*right, 50), (*left, 50)]);
                if let Some(element) = renderer.render_record(&strip) {
                    output.push_str(&element);
                }
            }
        };
        pass(renderer);
        // The middle fill: a black triangle drawn with R2_COPYPEN.
        renderer.render_record(&rec(record::SET_ROP2, &[13, 0]));
        select_null_pen_and_brush(renderer, 0);
        assert!(
            renderer
                .render_record(&polygon(&[(0, 0), (40, 0), (0, 40)]))
                .is_none()
        );
        pass(renderer);
        output
    }

    #[test]
    fn xor_bracket_paints_its_fills_clipped_to_the_middle_fill() {
        let mut renderer = renderer();
        let output = xor_bracket(
            &mut renderer,
            &[((0, 20), 0x0000_00ff), ((20, 40), 0x0000_ff00)],
        );
        assert!(renderer.issues.is_empty(), "{:?}", renderer.issues);
        // The strips tile the rectangle holding the middle triangle, so only
        // the strips paint, in their own colours, inside a clip of the
        // triangle.
        assert!(output.starts_with(r#"<g clip-path="url(#wmf-xor-clip-"#));
        assert!(!output.contains(r##"fill="#000000""##));
        assert!(output.contains(r##"fill="#ff0000""##) && output.contains(r##"fill="#00ff00""##));
        assert!(renderer.definitions.contains(r#"clip-rule="evenodd""#));
        assert!(!output.contains(XOR_FILL) && !renderer.definitions.contains(XOR_FILL));
    }

    #[test]
    fn middle_fill_outside_the_xor_fills_keeps_its_own_paint() {
        let mut renderer = renderer();
        // The strips cover x 0..40 only; the triangle reaches x 60.
        renderer.render_record(&rec(record::SET_ROP2, &[7, 0]));
        select_null_pen_and_brush(&mut renderer, 0x0000_00ff);
        let strip = polygon(&[(0, 0), (40, 0), (40, 50), (0, 50)]);
        renderer.render_record(&strip);
        renderer.render_record(&rec(record::SET_ROP2, &[13, 0]));
        select_null_pen_and_brush(&mut renderer, 0);
        renderer.render_record(&polygon(&[(0, 0), (60, 0), (0, 40)]));
        renderer.render_record(&rec(record::SET_ROP2, &[7, 0]));
        select_null_pen_and_brush(&mut renderer, 0x0000_00ff);
        let output = renderer.render_record(&strip).unwrap();
        assert!(renderer.issues.is_empty(), "{:?}", renderer.issues);
        // The black triangle under a mask of the strip, then the strip.
        assert!(output.starts_with(r#"<g mask="url(#wmf-xor-mask-"#));
        assert!(output.contains(r##"fill="#000000""##) && output.contains(r##"fill="#ff0000""##));
        assert!(renderer.definitions.contains(r#"<mask id="wmf-xor-mask-"#));
    }

    #[test]
    fn xor_bracket_with_overlapping_fills_fails() {
        let mut renderer = renderer();
        xor_bracket(
            &mut renderer,
            &[((0, 30), 0x0000_00ff), ((20, 40), 0x0000_ff00)],
        );
        assert!(fatal(&renderer));
    }

    #[test]
    fn xor_fill_that_no_second_pass_undoes_fails() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::SET_ROP2, &[7, 0]));
        select_null_pen_and_brush(&mut renderer, 0x0000_00ff);
        renderer.render_record(&polygon(&[(0, 0), (10, 0), (10, 10)]));
        assert!(!fatal(&renderer));
        assert!(renderer.finish().is_none());
        assert!(fatal(&renderer));
    }

    /// A DIBBitBlt record without a source bitmap.
    fn brush_bitblt(rop3: u32, left: i16, top: i16, width: i16, height: i16) -> WmfRecord {
        let mut params = rop3.to_le_bytes().to_vec();
        // SourceY, SourceX, Reserved, Height, Width, DestinationY, DestinationX.
        for value in [0i16, 0, 0, height, width, top, left] {
            params.extend_from_slice(&value.to_le_bytes());
        }
        rec(record::DIB_BIT_BLT, &params)
    }

    #[test]
    fn blit_without_source_reads_its_destination_after_the_reserved_field() {
        let mut renderer = renderer();
        create_solid_brush(&mut renderer, 0x0000_00ff);
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = renderer
            .render_record(&brush_bitblt(PATCOPY, 5, 6, 20, 10))
            .unwrap_or_else(|| panic!("{:?}", renderer.issues));
        assert!(output.starts_with(r#"<rect x="5" y="6" width="20" height="10""#));
        assert!(renderer.issues.is_empty(), "{:?}", renderer.issues);
    }

    #[test]
    fn patinvert_bracket_paints_the_middle_pattern_in_the_brush_colour() {
        let mut renderer = renderer();
        create_solid_brush(&mut renderer, 0x0000_00ff);
        // A black and white DIB pattern: black pixels take the XOR colour.
        let mut params = vec![5, 0, 0, 0];
        params.extend_from_slice(&rgb_dib(&[&[0, 0x00ff_ffff], &[0x00ff_ffff, 0]]));
        renderer.render_record(&rec(record::CREATE_DIB_PATTERN_BRUSH, &params));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[8, 0x80]));
        let blit = brush_bitblt(0x005a_0049, 0, 0, 40, 40);
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        assert!(renderer.render_record(&blit).is_none());
        renderer.render_record(&rec(record::SET_ROP2, &[9, 0]));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[1, 0]));
        assert!(
            renderer
                .render_record(&polygon(&[(0, 0), (30, 0), (0, 30)]))
                .is_none()
        );
        renderer.render_record(&rec(record::SET_ROP2, &[13, 0]));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = renderer.render_record(&blit).unwrap();
        assert!(renderer.issues.is_empty(), "{:?}", renderer.issues);
        // Inside the XOR rectangle, which holds the middle triangle, the
        // black pattern pixels become red and white ones stay transparent.
        assert_eq!(output.matches("url(#wmf-pattern-").count(), 1);
        assert!(output.contains(r#"<g clip-path="url(#wmf-xor-clip-"#));
        assert!(!output.contains("mask="));
    }

    #[test]
    fn select_object_with_a_region_selects_the_clip_region() {
        let mut renderer = renderer();
        // A region with one scan covering x 10..20, y 0..5.
        let mut region = vec![0u8; 22];
        region[2..4].copy_from_slice(&6u16.to_le_bytes());
        region[10..12].copy_from_slice(&1u16.to_le_bytes());
        for value in [2u16, 0, 5, 10, 20, 2] {
            region.extend_from_slice(&value.to_le_bytes());
        }
        renderer.render_record(&rec(record::CREATE_REGION, &region));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let rects = renderer.state.clip.rects.clone().unwrap();
        assert_eq!(rects, vec![DeviceRect::new(10.0, 0.0, 20.0, 5.0)]);
        assert!(renderer.issues.is_empty(), "{:?}", renderer.issues);
    }

    #[test]
    fn brush_raster_operations_map_to_mix_modes() {
        assert_eq!(rop3_as_rop2(Rop3::from_code(0x00a0_00c9)).mode(), 9);
        assert_eq!(rop3_as_rop2(Rop3::from_code(0x00f0_0021)).mode(), 13);
        assert_eq!(rop3_as_rop2(Rop3::from_code(0x005a_0049)).mode(), 7);
        assert_eq!(rop3_as_rop2(Rop3::from_code(0x0055_0009)).mode(), 6);
    }

    #[test]
    fn dib_header_fields_that_gdi_ignores_are_noncanonical() {
        let mut renderer = renderer();
        let mut dib = rgb_dib(&[&[0x0000_00ff]]);
        // biSizeImage smaller than the scanlines.
        dib[20..24].copy_from_slice(&1u32.to_le_bytes());
        let output = renderer.render_record(&dib_bitblt_of(&dib, 0x00cc_0020, 0, 0, 1, 1));
        assert!(output.is_none());
        assert!(renderer.finish().unwrap().contains("<image"));
        assert!(
            renderer
                .issues
                .iter()
                .all(|issue| !issue.fatal && issue.code == "noncanonical-dib-image-size")
        );
    }

    #[test]
    fn polyline_with_one_point_paints_nothing() {
        let mut renderer = renderer();
        assert!(
            renderer
                .render_record(&rec(record::POLYLINE, &[1, 0, 5, 0, 5, 0]))
                .is_none()
        );
        assert!(!fatal(&renderer));
        assert!(
            renderer
                .issues
                .iter()
                .any(|issue| issue.code == "noncanonical-wmf-degenerate-polyline")
        );
        // Truncated points stay malformed.
        renderer.render_record(&rec(record::POLYLINE, &[3, 0, 5, 0, 5, 0]));
        assert!(fatal(&renderer));
    }

    #[test]
    fn set_pixel_paints_one_canvas_unit_under_a_magnifying_transform() {
        // The whole 1-by-1 device extent fills a 100-unit canvas.
        let mut renderer =
            renderer_with(CoordinateTransform::new((0.0, 0.0, 1.0, 1.0), 100.0, 100.0));
        let output = renderer
            .render_record(&rec(record::SET_PIXEL, &[0, 0, 0, 0, 0, 0, 0, 0]))
            .unwrap();

        assert!(output.contains(r#"width="1" height="1""#));
    }

    #[test]
    fn malformed_bitmap_is_a_fatal_explicit_diagnostic() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::DIB_BIT_BLT, &[0; 16]));
        assert!(renderer.issues.iter().any(|issue| issue.fatal));
    }

    #[test]
    fn escape_payload_is_never_emitted_or_diagnosed_as_drawing() {
        let mut renderer = renderer();
        assert!(
            renderer
                .render_record(&rec(record::ESCAPE, b"%!PS dangerous"))
                .is_none()
        );
        assert!(renderer.issues.is_empty());
    }

    #[test]
    fn every_ms_wmf_record_type_has_an_explicit_classification() {
        for &function in record::ALL {
            let mut renderer = renderer();
            renderer.render_record(&rec(function, &[]));
            assert!(
                renderer
                    .issues
                    .iter()
                    .all(|issue| issue.code != "noncanonical-wmf-unknown-record"),
                "record 0x{function:04X} fell through classification"
            );
        }
    }

    #[test]
    fn unknown_records_and_degenerate_polylines_are_ignored_like_gdi() {
        let mut renderer = renderer();
        assert!(renderer.render_record(&rec(0x0d32, &[0; 8])).is_none());
        // A Polyline of one point, a Polygon of none.
        assert!(
            renderer
                .render_record(&rec(record::POLYLINE, &[1, 0, 5, 0, 5, 0]))
                .is_none()
        );
        assert!(
            renderer
                .render_record(&rec(record::POLYGON, &[0, 0]))
                .is_none()
        );
        assert_eq!(
            issue_codes(&renderer),
            [
                "noncanonical-wmf-unknown-record",
                "noncanonical-wmf-degenerate-polyline",
                "noncanonical-wmf-degenerate-polyline",
            ]
        );
        assert!(renderer.issues.iter().all(|issue| !issue.fatal));
    }

    #[test]
    fn same_arc_points_emit_complete_ellipse_without_nan() {
        let mut renderer = renderer();
        let params = [50, 0, 100, 0, 50, 0, 100, 0, 100, 0, 100, 0, 0, 0, 0, 0];
        let output = renderer.render_record(&rec(record::ARC, &params)).unwrap();
        assert!(output.starts_with("<ellipse"));
        assert!(!output.contains("NaN"));
    }

    #[test]
    fn windows_1252_text_and_xml_metacharacters_are_safe() {
        let decoded = crate::text_charset::decode_logfont_text(&[0x80, b'<', b'&'], 0, "");
        assert!(decoded.exact);
        let mut escaped = String::new();
        escape_xml_text_into(&mut escaped, &decoded.text);
        assert_eq!(escaped, "€&lt;&amp;");
    }

    #[test]
    fn symbol_font_text_uses_a_unicode_face_and_private_use_keeps_the_face() {
        let mut greek = renderer();
        greek.render_record(&rec(
            record::CREATE_FONT_INDIRECT,
            &font_params(2, "Symbol"),
        ));
        greek.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = greek
            .render_record(&rec(record::TEXT_OUT, &[1, 0, b'a', 0, 0, 0, 0, 0]))
            .unwrap();
        assert!(output.contains("α"));
        assert!(!output.contains("Symbol"));

        let mut dingbat = renderer();
        dingbat.render_record(&rec(
            record::CREATE_FONT_INDIRECT,
            &font_params(2, "Wingdings"),
        ));
        dingbat.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = dingbat
            .render_record(&rec(record::TEXT_OUT, &[1, 0, b'a', 0, 0, 0, 0, 0]))
            .unwrap();
        assert!(output.contains('\u{264b}'), "{output}");
        assert!(output.contains("Wingdings"), "{output}");

        let mut unmapped = renderer();
        unmapped.render_record(&rec(
            record::CREATE_FONT_INDIRECT,
            &font_params(2, "Webdings"),
        ));
        unmapped.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = unmapped
            .render_record(&rec(record::TEXT_OUT, &[1, 0, b'a', 0, 0, 0, 0, 0]))
            .unwrap();
        assert!(output.contains("\u{f061}"), "{output}");
        assert!(output.contains("Webdings"), "{output}");
    }

    fn region_params(left: i16, top: i16, right: i16, bottom: i16) -> Vec<u8> {
        let mut params = vec![0u8; 22];
        params[2..4].copy_from_slice(&6_u16.to_le_bytes());
        params[10..12].copy_from_slice(&1_u16.to_le_bytes());
        for value in [2_i16, top, bottom, left, right, 2] {
            params.extend_from_slice(&value.to_le_bytes());
        }
        params
    }

    fn issue_codes(renderer: &SvgRenderer<'_>) -> Vec<&'static str> {
        renderer.issues.iter().map(|issue| issue.code).collect()
    }

    #[test]
    fn selecting_a_region_object_sets_the_clip_region() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::CREATE_REGION, &region_params(0, 0, 10, 10)));
        assert!(renderer.state.clip.is_unbounded());
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        assert!(!renderer.state.clip.is_unbounded());
        assert!(renderer.issues.is_empty());
    }

    #[test]
    fn object_table_misuse_that_gdi_ignores_reports_noncanonical_codes() {
        let mut renderer = renderer();
        renderer.render_record(&rec(record::CREATE_PALETTE, &[0, 3, 1, 0, 1, 2, 3, 0]));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[9, 0]));
        renderer.render_record(&rec(record::DELETE_OBJECT, &[9, 0]));
        renderer.render_record(&rec(record::SELECT_PALETTE, &[9, 0]));
        renderer.render_record(&rec(record::RESTORE_DC, &[0xff, 0xff]));
        renderer.render_record(&rec(record::SELECT_CLIP_REGION, &[9, 0]));
        assert_eq!(
            issue_codes(&renderer),
            [
                "noncanonical-wmf-select-palette-object",
                "noncanonical-wmf-select-missing-object",
                "noncanonical-wmf-delete-missing-object",
                "noncanonical-wmf-select-missing-palette",
                "noncanonical-wmf-restore-dc-level",
                "noncanonical-wmf-select-missing-region",
            ]
        );
        assert!(renderer.issues.iter().all(|issue| !issue.fatal));
    }

    #[test]
    fn glyph_index_text_is_omitted_with_an_approximation_diagnostic() {
        let mut renderer = renderer();
        let params = [5, 0, 5, 0, 2, 0, 0x10, 0, 3, 0, 4, 0];
        assert!(
            renderer
                .render_record(&rec(record::EXT_TEXT_OUT, &params))
                .is_none()
        );
        assert_eq!(issue_codes(&renderer), ["wmf-glyph-index-text-omitted"]);
        assert!(renderer.issues.iter().all(|issue| !issue.fatal));
    }

    /// ExtTextOut with a dx array: y, x, count, options, the string padded to
    /// a word boundary, then one advance per character.
    fn ext_text_out_params(x: i16, options: u16, text: &[u8], dx: &[i16]) -> Vec<u8> {
        let mut params = Vec::new();
        for value in [0_i16, x] {
            params.extend_from_slice(&value.to_le_bytes());
        }
        params.extend_from_slice(&(text.len() as u16).to_le_bytes());
        params.extend_from_slice(&options.to_le_bytes());
        params.extend_from_slice(text);
        if text.len() % 2 != 0 {
            params.push(0);
        }
        for value in dx {
            params.extend_from_slice(&value.to_le_bytes());
        }
        params
    }

    #[test]
    fn the_recorded_face_stays_first_in_the_font_family() {
        let mut renderer = renderer();
        let mut params = font_params(0, "Arial");
        params[17] = 0x20; // FF_SWISS
        renderer.render_record(&rec(record::CREATE_FONT_INDIRECT, &params));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let output = renderer
            .render_record(&rec(record::TEXT_OUT, &[1, 0, b'a', 0, 0, 0, 0, 0]))
            .unwrap();
        assert!(
            output.contains(r#"font-family="Arial, sans-serif""#),
            "{output}"
        );
    }

    #[test]
    fn ext_text_out_dx_array_places_each_glyph_at_its_own_origin() {
        let mut renderer = renderer();
        // GDI advances 10 then 20 logical units; SVG must not add these to the
        // natural advances as well.
        let params = ext_text_out_params(0, 0, b"abc", &[10, 20]);
        let element = renderer
            .render_record(&rec(record::EXT_TEXT_OUT, &params))
            .expect("text element");
        assert!(element.contains(r#"<text x="0 10 30"#), "{element}");
        assert!(element.contains(r#"y="0 0 0"#), "{element}");
        assert!(!element.contains(" dx="), "{element}");
    }

    #[test]
    fn ext_text_out_dx_array_keeps_gdi_string_anchoring() {
        let mut renderer = renderer();
        renderer.state.text_align = text_align::TA_CENTER | text_align::TA_TOP;
        let params = ext_text_out_params(40, 0, b"abc", &[10, 20]);
        let element = renderer
            .render_record(&rec(record::EXT_TEXT_OUT, &params))
            .expect("text element");
        // The 30 unit string is centered on 40: glyphs at 25, 35 and 55.
        assert!(element.contains(r#"<text x="25 35 55"#), "{element}");
        assert!(!element.contains("text-anchor"), "{element}");
    }

    #[test]
    fn ext_text_out_advances_cover_every_byte_of_a_double_byte_character() {
        let mut renderer = renderer();
        renderer.render_record(&rec(
            record::CREATE_FONT_INDIRECT,
            &font_params(128, "MS Mincho"),
        ));
        renderer.render_record(&rec(record::SELECT_OBJECT, &[0, 0]));
        let params = ext_text_out_params(0, 0, b"a\x82\xa0b", &[10, 6, 6, 8]);
        let element = renderer
            .render_record(&rec(record::EXT_TEXT_OUT, &params))
            .expect("text element");
        assert!(
            element.contains(r#"<text x="0 10 22" y="0 0 0""#),
            "{element}"
        );
        assert!(element.contains("aあb"), "{element}");
    }

    #[test]
    fn ext_text_out_vertical_advances_are_applied() {
        let mut renderer = renderer();
        let params = ext_text_out_params(0, ext_text_out::ETO_PDY, b"ab", &[10, 4, 10, 4]);
        let element = renderer
            .render_record(&rec(record::EXT_TEXT_OUT, &params))
            .expect("text element");
        assert!(element.contains(r#"<text x="0 10" y="0 4""#), "{element}");
        assert!(renderer.issues.is_empty());
    }
}
