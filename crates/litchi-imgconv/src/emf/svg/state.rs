//! SVG playback state for classic EMF records.

use super::path::PathBuilder;
use crate::emf::records::{
    ColorRef, XForm, brush_style, font_weight, hatch_style, pen_style, stock_objects,
};
use crate::metafile_bitmap::BrushPattern;
use crate::svg_utils::{write_num, write_xml_escaped};
use litchi_core::error::{Error, Result};
use std::collections::HashMap;
use std::sync::Arc;

#[derive(Debug, Clone)]
pub enum GdiObject {
    Pen(Pen),
    Brush(Brush),
    Font(Font),
    Palette,
}

/// Complete rendering state. The object table and definition table are not DC
/// state and therefore intentionally do not participate in SaveDC/RestoreDC.
pub struct RenderState {
    pub dc_stack: Vec<DeviceContext>,
    pub dc: DeviceContext,
    pub path_builder: Option<PathBuilder>,
    pub in_path: bool,
    pub objects: HashMap<u32, GdiObject>,
    pub defs: Vec<String>,
    next_definition: u64,
    /// Non-fatal playback facts, as `(code, message)` pairs.
    pub diagnostics: Vec<(&'static str, String)>,
    /// Bitmap records whose painting waits for the records after them.
    pub(super) deferred_blits: Option<super::converter::DeferredBlits>,
    /// Emitted brush pattern definitions by PNG tile and brush origin.
    pattern_definitions: HashMap<(Vec<u8>, (i32, i32)), String>,
    /// Fills XORed into the destination, waiting for the pass that undoes
    /// them.
    pub(super) xor_bracket: Option<super::converter::XorBracket>,
    /// The current record's shape and elements, when they belong to an XOR
    /// bracket.
    pub(super) bracket_shape: Option<(super::converter::BracketShape, Vec<String>)>,
}

/// Device-space rectangle covered by the output canvas.
///
/// A clip region never needs to cover ground outside the canvas, and bounding
/// "every point" to it keeps clip definitions inside the coordinate range SVG
/// rasterizers handle reliably: geometry near a billion units renders as
/// nothing, and larger coordinates overflow the rasterizer.
#[derive(Debug, Clone, Copy)]
pub struct Canvas {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

/// Fill rule that decides which points of a clip path are inside the clip.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ClipRule {
    NonZero,
    EvenOdd,
}

/// Most distinct diagnostics retained for one classic EMF playback.
const MAX_DIAGNOSTICS: usize = 64;

impl RenderState {
    pub fn new() -> Self {
        Self::with_device_metrics(96.0, 96.0)
    }

    pub fn with_device_metrics(dpi_x: f64, dpi_y: f64) -> Self {
        Self {
            dc_stack: Vec::new(),
            dc: DeviceContext::with_device_metrics(dpi_x, dpi_y),
            path_builder: None,
            in_path: false,
            objects: HashMap::new(),
            defs: Vec::new(),
            next_definition: 0,
            diagnostics: Vec::new(),
            deferred_blits: None,
            pattern_definitions: HashMap::new(),
            xor_bracket: None,
            bracket_shape: None,
        }
    }

    /// Records a non-fatal playback fact under a stable diagnostic code.
    ///
    /// Use a `noncanonical-` code when Windows GDI playback tolerates the
    /// input without changing the picture; any other code reports an
    /// approximation that strict conversion rejects. Repeated facts are
    /// reported once, and the list is bounded.
    pub fn note(&mut self, code: &'static str, message: String) {
        if self.diagnostics.len() < MAX_DIAGNOSTICS
            && !self
                .diagnostics
                .iter()
                .any(|(seen, text)| *seen == code && *text == message)
        {
            self.diagnostics.push((code, message));
        }
    }

    pub fn push_dc(&mut self) {
        self.dc_stack.push(self.dc.clone());
    }

    /// Restore the requested saved state and discard that state and every
    /// state saved after it, as specified by RestoreDC.
    pub fn pop_dc(&mut self, saved_dc: i32) -> bool {
        let target = if saved_dc < 0 {
            let distance = match saved_dc.checked_abs().and_then(|v| usize::try_from(v).ok()) {
                Some(value) => value,
                None => return false,
            };
            self.dc_stack.len().checked_sub(distance)
        } else if saved_dc > 0 {
            usize::try_from(saved_dc - 1).ok()
        } else {
            None
        };
        let Some(target) = target.filter(|&index| index < self.dc_stack.len()) else {
            return false;
        };
        self.dc = self.dc_stack[target].clone();
        self.dc_stack.truncate(target);
        true
    }

    pub fn begin_path(&mut self) {
        self.path_builder = Some(PathBuilder::new());
        self.in_path = true;
    }

    pub fn end_path(&mut self) {
        self.in_path = false;
    }

    pub fn take_path(&mut self) -> Option<PathBuilder> {
        self.in_path = false;
        self.path_builder.take()
    }

    pub fn abort_path(&mut self) {
        self.in_path = false;
        self.path_builder = None;
    }

    /// Stores an object under a handle and reports whether it replaced a live
    /// object.
    pub fn insert_object(&mut self, handle: u32, object: GdiObject) -> bool {
        self.objects.insert(handle, object).is_some()
    }

    /// Frees a handle-table slot.
    ///
    /// GDI frees the slot even when the object is still selected into a
    /// device context: the context keeps drawing with it, and the handle can
    /// be created again.
    pub fn delete_object(&mut self, handle: u32) -> bool {
        !stock_objects::is_stock_object(handle) && self.objects.remove(&handle).is_some()
    }

    pub fn select_object(&mut self, handle: u32) -> bool {
        let object = if stock_objects::is_stock_object(handle) {
            stock_object(handle)
        } else {
            self.objects.get(&handle).cloned()
        };
        match object {
            Some(GdiObject::Pen(pen)) => {
                self.dc.pen = pen;
                self.dc.pen_handle = Some(handle);
            },
            Some(GdiObject::Brush(brush)) => {
                self.dc.brush = brush;
                self.dc.brush_handle = Some(handle);
            },
            Some(GdiObject::Font(font)) => {
                self.dc.font = font;
                self.dc.font_handle = Some(handle);
            },
            Some(GdiObject::Palette) => return false,
            None => return false,
        }
        true
    }

    pub fn select_brush(&mut self, handle: u32) -> bool {
        let object = if stock_objects::is_stock_object(handle) {
            stock_object(handle)
        } else {
            self.objects.get(&handle).cloned()
        };
        if let Some(GdiObject::Brush(brush)) = object {
            self.dc.brush = brush;
            self.dc.brush_handle = Some(handle);
            true
        } else {
            false
        }
    }

    pub fn fresh_id(&mut self, prefix: &str) -> String {
        let id = format!("emf-{prefix}-{}", self.next_definition);
        self.next_definition += 1;
        id
    }

    pub fn add_definition(&mut self, definition: String) {
        self.defs.push(definition);
    }

    /// Emits the fill pattern definition of the selected brush, once.
    ///
    /// GDI tiles brush patterns in device pixels from the brush origin, and
    /// monochrome patterns paint their 0 bits with the text colour and their
    /// 1 bits with the background colour. `maximum` bounds the encoded
    /// pattern image.
    pub fn prepare_brush_pattern(&mut self, maximum: usize) -> Result<()> {
        if !self.dc.brush.needs_pattern() || self.dc.brush.pattern_id.is_some() {
            return Ok(());
        }
        if let Some(pattern) = self.dc.brush.pattern.clone() {
            let pattern = pattern.map_err(|reason| {
                Error::Unsupported(format!("pattern brush has no usable bitmap: {reason}"))
            })?;
            let png = pattern.tile_png(
                self.dc.text_color.value & 0x00ff_ffff,
                self.dc.bg_color.value & 0x00ff_ffff,
                maximum,
            )?;
            let id = self.pattern_definition(
                png,
                (f64::from(pattern.width()), f64::from(pattern.height())),
                self.dc.brush_org,
            )?;
            self.dc.brush.pattern_id = Some(id);
            return Ok(());
        }
        let id = self.fresh_id("brush");
        let definition = self
            .dc
            .brush
            .svg_pattern(
                &id,
                self.dc.bg_color,
                self.dc.bg_mode == 2,
                self.dc.brush_org,
            )
            .ok_or_else(|| {
                Error::Unsupported(format!(
                    "hatch style {:?} has no pattern definition",
                    self.dc.brush.hatch
                ))
            })?;
        self.dc.brush.pattern_id = Some(id);
        self.add_definition(definition);
        Ok(())
    }

    /// The id of the pattern definition tiling `png` from the device-space
    /// `origin`, emitted once.
    pub fn pattern_definition(
        &mut self,
        png: Vec<u8>,
        tile: (f64, f64),
        origin: (i32, i32),
    ) -> Result<String> {
        let key = (png, origin);
        if let Some(id) = self.pattern_definitions.get(&key) {
            return Ok(id.clone());
        }
        let id = self.fresh_id("brush");
        let definition = BrushPattern::svg_definition(
            &id,
            &key.0,
            tile,
            (f64::from(origin.0), f64::from(origin.1)),
        )?;
        self.add_definition(definition);
        self.pattern_definitions.insert(key, id.clone());
        Ok(id)
    }

    /// Combines `shape` with the current clip under the RGN_* `mode` and
    /// makes the result the current clip. Returns `false` for a combine mode
    /// that SVG clipping cannot express exactly.
    ///
    /// Rectangles intersect geometrically, so only other clip paths chain
    /// clip references; renderers mask every clipped element once per link.
    /// Past [`MAX_CLIP_CHAIN_DEPTH`] links a further clip path is not applied
    /// and a lossy diagnostic reports it.
    pub fn install_clip(&mut self, shape: ClipShape<'_>, mode: u32, canvas: Canvas) -> bool {
        let previous = self.dc.clip_region.clone();
        let mut region = match (mode, previous.is_unclipped()) {
            (5, _) | (1 | 2, true) => ClipRegion::default(),
            (1, false) => previous,
            // OR, XOR and DIFF need a stored copy of the old geometry or mask
            // semantics; do not pretend a lossy clip is equivalent.
            _ => return false,
        };
        match shape {
            ClipShape::Rect(rect) => {
                region.rect = Some(match region.rect {
                    Some(old) => intersect_rects(old, rect),
                    None => normalized_rect(rect),
                });
            },
            ClipShape::Path(path, rule) => {
                let depth = region.chain.as_ref().map_or(0, |(_, depth)| *depth);
                if depth >= MAX_CLIP_CHAIN_DEPTH {
                    self.note(
                        "emf-clip-depth-approximated",
                        format!(
                            "a clip region intersecting more than {MAX_CLIP_CHAIN_DEPTH} clip paths was drawn without the later ones"
                        ),
                    );
                    return true;
                }
                let id = self.fresh_id("clip");
                let mut definition = String::from("<clipPath id=\"");
                definition.push_str(&id);
                definition.push_str("\" clipPathUnits=\"userSpaceOnUse\"><path d=\"");
                definition.push_str(path);
                definition.push('"');
                if rule == ClipRule::EvenOdd {
                    definition.push_str(" clip-rule=\"evenodd\"");
                }
                // The clipPath content model admits shape elements but not
                // `<g>`, so the earlier clip paths are referenced from the
                // child shape; renderers drop a clip path that holds a group
                // and then clip everything away.
                push_clip_reference(&mut definition, region.chain.as_ref());
                definition.push_str("/></clipPath>");
                self.add_definition(definition);
                region.chain = Some((id, depth + 1));
            },
        }
        self.set_clip_region(region, canvas);
        true
    }

    /// Moves the current clip by the device-space vector `(x, y)`.
    pub fn offset_clip(&mut self, x: f64, y: f64, canvas: Canvas) {
        let mut region = self.dc.clip_region.clone();
        if region.is_unclipped() {
            return;
        }
        if let Some(rect) = &mut region.rect {
            *rect = [rect[0] + x, rect[1] + y, rect[2] + x, rect[3] + y];
        }
        if let Some((old, depth)) = region.chain.clone() {
            if depth >= MAX_CLIP_CHAIN_DEPTH {
                self.note(
                    "emf-clip-depth-approximated",
                    format!(
                        "a clip region intersecting more than {MAX_CLIP_CHAIN_DEPTH} clip paths was not moved by OffsetClipRgn"
                    ),
                );
                return;
            }
            let id = self.fresh_id("clip");
            // The translate moves the previous clip; the rectangle that carries
            // the reference stands in for "every point" and is placed so that it
            // covers the canvas after the translate. The reference sits on a shape
            // element because the clipPath content model does not admit `<g>`.
            let mut definition = String::from("<clipPath id=\"");
            definition.push_str(&id);
            definition.push_str("\" clipPathUnits=\"userSpaceOnUse\" transform=\"translate(");
            write_num(&mut definition, x);
            definition.push(' ');
            write_num(&mut definition, y);
            definition.push_str(")\"><rect x=\"");
            write_num(&mut definition, canvas.left - x);
            definition.push_str("\" y=\"");
            write_num(&mut definition, canvas.top - y);
            definition.push_str("\" width=\"");
            write_num(&mut definition, canvas.right - canvas.left);
            definition.push_str("\" height=\"");
            write_num(&mut definition, canvas.bottom - canvas.top);
            definition.push_str("\" clip-path=\"url(#");
            definition.push_str(&old);
            definition.push_str(")\"/></clipPath>");
            self.add_definition(definition);
            region.chain = Some((id, depth + 1));
        }
        self.set_clip_region(region, canvas);
    }

    /// Removes the clip region, so that drawing reaches the whole canvas.
    pub fn clear_clip(&mut self) {
        self.dc.clip_region = ClipRegion::default();
        self.dc.clip_id = None;
    }

    /// Makes `region` the current clip, defining the clip that elements
    /// reference.
    fn set_clip_region(&mut self, region: ClipRegion, canvas: Canvas) {
        if region == self.dc.clip_region {
            // Elements that share a clip reference share one mask.
            return;
        }
        self.dc.clip_id = match (&region.rect, &region.chain) {
            (None, None) => None,
            (None, Some((chain, _))) => Some(chain.clone()),
            // Nothing outside the canvas is visible, so a rectangle that
            // holds the canvas clips nothing away. Leaving it out keeps one
            // mask per drawn element out of the rasterizer.
            (Some(rect), None) if rect_holds_canvas(*rect, canvas) => None,
            (Some(rect), Some((chain, _))) if rect_holds_canvas(*rect, canvas) => {
                Some(chain.clone())
            },
            (Some(rect), chain) => {
                let id = self.fresh_id("clip");
                let mut definition = String::from("<clipPath id=\"");
                definition.push_str(&id);
                definition.push_str("\" clipPathUnits=\"userSpaceOnUse\"><path d=\"");
                let mut outline = PathBuilder::new();
                outline.move_to(rect[0], rect[1]);
                outline.line_to(rect[2], rect[1]);
                outline.line_to(rect[2], rect[3]);
                outline.line_to(rect[0], rect[3]);
                outline.close();
                definition.push_str(&outline.build());
                definition.push('"');
                push_clip_reference(&mut definition, chain.as_ref());
                definition.push_str("/></clipPath>");
                self.add_definition(definition);
                Some(id)
            },
        };
        self.dc.clip_region = region;
    }
}

/// Most clip paths, other than rectangles, that one clip region intersects
/// through chained clip references.
pub const MAX_CLIP_CHAIN_DEPTH: usize = 16;

/// The geometry a clip record combines with the current clip.
#[derive(Debug, Clone, Copy)]
pub enum ClipShape<'a> {
    /// A device-space rectangle `[left, top, right, bottom]`.
    Rect([f64; 4]),
    /// SVG path data in device space and the rule that fills it.
    Path(&'a str, ClipRule),
}

/// The clip region of a device context: a device-space rectangle
/// intersected with a chain of other clip paths. `None` for both clips
/// nothing away.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ClipRegion {
    /// Device-space rectangle `[left, top, right, bottom]`, normalized.
    pub rect: Option<[f64; 4]>,
    /// The clip definition that intersects the other clip paths, and how
    /// many clip paths it chains.
    pub chain: Option<(String, usize)>,
}

impl ClipRegion {
    fn is_unclipped(&self) -> bool {
        self.rect.is_none() && self.chain.is_none()
    }
}

fn push_clip_reference(definition: &mut String, chain: Option<&(String, usize)>) {
    if let Some((chain, _)) = chain {
        definition.push_str(" clip-path=\"url(#");
        definition.push_str(chain);
        definition.push_str(")\"");
    }
}

fn normalized_rect(rect: [f64; 4]) -> [f64; 4] {
    [
        rect[0].min(rect[2]),
        rect[1].min(rect[3]),
        rect[0].max(rect[2]),
        rect[1].max(rect[3]),
    ]
}

/// The intersection of two rectangles; an empty one keeps zero size.
fn intersect_rects(old: [f64; 4], new: [f64; 4]) -> [f64; 4] {
    let new = normalized_rect(new);
    let (left, top) = (old[0].max(new[0]), old[1].max(new[1]));
    [
        left,
        top,
        old[2].min(new[2]).max(left),
        old[3].min(new[3]).max(top),
    ]
}

fn rect_holds_canvas(rect: [f64; 4], canvas: Canvas) -> bool {
    rect[0] <= canvas.left
        && rect[1] <= canvas.top
        && rect[2] >= canvas.right
        && rect[3] >= canvas.bottom
}

impl Default for RenderState {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone)]
pub struct DeviceContext {
    /// Current position is maintained in logical coordinates. Each operation
    /// transforms it using the state active for that operation.
    pub current_pos: (f64, f64),
    pub world_transform: XForm,
    pub window_org: (i32, i32),
    pub window_ext: (i32, i32),
    pub viewport_org: (i32, i32),
    pub viewport_ext: (i32, i32),
    pub brush_org: (i32, i32),
    pub map_mode: u32,
    pub dpi: (f64, f64),
    pub pen: Pen,
    pub pen_handle: Option<u32>,
    pub brush: Brush,
    pub brush_handle: Option<u32>,
    pub font: Font,
    pub font_handle: Option<u32>,
    pub text_color: ColorRef,
    pub bg_color: ColorRef,
    pub bg_mode: u32,
    pub poly_fill_mode: u32,
    pub text_align: u32,
    /// Extra width that `EMR_SETTEXTJUSTIFICATION` adds to break characters.
    pub text_break_extra: i32,
    pub rop2: u32,
    pub stretch_mode: u32,
    pub arc_direction: bool,
    pub miter_limit: f64,
    pub layout: u32,
    /// The clip definition that drawn elements reference.
    pub clip_id: Option<String>,
    pub clip_region: ClipRegion,
}

impl Default for DeviceContext {
    fn default() -> Self {
        Self::with_device_metrics(96.0, 96.0)
    }
}

impl DeviceContext {
    pub fn with_device_metrics(dpi_x: f64, dpi_y: f64) -> Self {
        Self {
            current_pos: (0.0, 0.0),
            world_transform: XForm::default(),
            window_org: (0, 0),
            window_ext: (1, 1),
            viewport_org: (0, 0),
            viewport_ext: (1, 1),
            brush_org: (0, 0),
            map_mode: 1,
            dpi: (dpi_x.max(1.0), dpi_y.max(1.0)),
            pen: Pen::default(),
            pen_handle: Some(stock_objects::BLACK_PEN),
            brush: Brush::default(),
            brush_handle: Some(stock_objects::WHITE_BRUSH),
            font: Font::default(),
            font_handle: Some(stock_objects::SYSTEM_FONT),
            text_color: ColorRef::from_rgb(0, 0, 0),
            bg_color: ColorRef::from_rgb(255, 255, 255),
            bg_mode: 2,
            poly_fill_mode: 1,
            text_align: 0,
            text_break_extra: 0,
            rop2: 13,
            stretch_mode: 3,
            arc_direction: false,
            miter_limit: 10.0,
            layout: 0,
            clip_id: None,
            clip_region: ClipRegion::default(),
        }
    }

    /// Transform logical to device coordinates in the required order:
    /// world transform (logical -> page), then mapping transform (page -> device).
    pub fn transform_point(&self, x: f64, y: f64) -> (f64, f64) {
        let (page_x, page_y) = self.world_transform.transform_point(x, y);
        let (scale_x, scale_y) = self.page_scale();
        (
            (page_x - f64::from(self.window_org.0)) * scale_x + f64::from(self.viewport_org.0),
            (page_y - f64::from(self.window_org.1)) * scale_y + f64::from(self.viewport_org.1),
        )
    }

    pub fn transform_vector(&self, x: f64, y: f64) -> (f64, f64) {
        let origin = self.transform_point(0.0, 0.0);
        let end = self.transform_point(x, y);
        (end.0 - origin.0, end.1 - origin.1)
    }

    /// Length in device units of the transformed logical x and y axes.
    ///
    /// GDI measures font heights and the widths of most pens in logical
    /// units, so the world transform and page mapping scale them.
    pub fn device_scale(&self) -> (f64, f64) {
        let (xx, xy) = self.transform_vector(1.0, 0.0);
        let (yx, yy) = self.transform_vector(0.0, 1.0);
        (xx.hypot(xy), yx.hypot(yy))
    }

    /// SVG `font-size` of the selected font in device units.
    ///
    /// A LOGFONT height is in logical units, so the world transform and page
    /// mapping scale the glyphs ([MS-EMF] 2.2.13 LogFont). A zero height asks
    /// GDI for its default size, which the crate models with the 12 device
    /// units of [`Font::default`].
    pub fn font_device_size(&self) -> f64 {
        let height = self.font.height.abs();
        if height == 0.0 {
            return Font::default().height;
        }
        (height * self.device_scale().1).max(0.01)
    }

    fn page_scale(&self) -> (f64, f64) {
        let ratio = || {
            let x = if self.window_ext.0 == 0 {
                1.0
            } else {
                f64::from(self.viewport_ext.0) / f64::from(self.window_ext.0)
            };
            let y = if self.window_ext.1 == 0 {
                1.0
            } else {
                f64::from(self.viewport_ext.1) / f64::from(self.window_ext.1)
            };
            (x, y)
        };
        match self.map_mode {
            1 => (1.0, 1.0),
            2 => (self.dpi.0 / 254.0, -self.dpi.1 / 254.0),
            3 => (self.dpi.0 / 2540.0, -self.dpi.1 / 2540.0),
            4 => (self.dpi.0 / 100.0, -self.dpi.1 / 100.0),
            5 => (self.dpi.0 / 1000.0, -self.dpi.1 / 1000.0),
            6 => (self.dpi.0 / 1440.0, -self.dpi.1 / 1440.0),
            7 => {
                let (x, y) = ratio();
                let magnitude = x.abs().min(y.abs());
                (magnitude.copysign(x), magnitude.copysign(y))
            },
            8 => ratio(),
            _ => (1.0, 1.0),
        }
    }

    pub fn get_stroke_attrs(&self) -> String {
        if self.pen.style == pen_style::NULL {
            return "stroke=\"none\"".to_string();
        }
        let mut attrs = format!("stroke=\"{}\"", self.pen.color.to_svg_color());
        self.push_pen_shape_attrs(&mut attrs);
        attrs
    }

    /// Appends width, dash, cap, join and miter attributes of the pen.
    fn push_pen_shape_attrs(&self, attrs: &mut String) {
        let width = self.pen_device_width();
        if width > 1.0 {
            attrs.push_str(" stroke-width=\"");
            write_num(attrs, width);
            attrs.push('"');
        }
        if let Some(dash) = &self.pen.dash_pattern {
            let (lengths, scale) = match dash {
                DashPattern::Width(multiples) => (multiples, width),
                DashPattern::Units(lengths) => (lengths, self.pen_unit_scale()),
            };
            attrs.push_str(" stroke-dasharray=\"");
            for (index, length) in lengths.iter().enumerate() {
                if index != 0 {
                    attrs.push(' ');
                }
                write_num(attrs, length * scale);
            }
            attrs.push('"');
        }
        let cap = match self.pen.end_cap {
            pen_style::ENDCAP_ROUND => "round",
            pen_style::ENDCAP_SQUARE => "square",
            _ => "butt",
        };
        let join = match self.pen.line_join {
            pen_style::JOIN_ROUND => "round",
            pen_style::JOIN_BEVEL => "bevel",
            _ => "miter",
        };
        attrs.push_str(&format!(
            " stroke-linecap=\"{cap}\" stroke-linejoin=\"{join}\" stroke-miterlimit=\"{}\"",
            self.miter_limit
        ));
    }

    /// Device units per unit of the selected pen's width and style lengths.
    ///
    /// A non-uniform mapping turns the round nib of a logical width into an
    /// ellipse; the scale returned here gives the circle of the same area.
    fn pen_unit_scale(&self) -> f64 {
        if self.pen.logical_width {
            let (x, y) = self.device_scale();
            (x * y).sqrt()
        } else {
            1.0
        }
    }

    /// Width of the selected pen in device units; GDI draws every line at
    /// least one device pixel wide, and a zero-width pen exactly that wide.
    fn pen_device_width(&self) -> f64 {
        (self.pen.width * self.pen_unit_scale()).max(1.0)
    }

    pub fn get_fill_attr(&self) -> String {
        let fill = if self.brush.style == brush_style::NULL {
            "none".to_string()
        } else if let Some(id) = &self.brush.pattern_id {
            format!("url(#{id})")
        } else {
            self.brush.color.to_svg_color()
        };
        format!("fill=\"{}\"", fill)
    }

    pub fn get_fill_rule(&self) -> Option<String> {
        (self.poly_fill_mode == 1).then(|| "fill-rule=\"evenodd\"".to_string())
    }

    pub fn clip_attr(&self) -> String {
        self.clip_id
            .as_ref()
            .map(|id| format!("clip-path=\"url(#{id})\""))
            .unwrap_or_default()
    }
}

#[derive(Debug, Clone)]
pub struct Pen {
    pub style: u32,
    /// Width as recorded; zero draws lines one device pixel wide.
    pub width: f64,
    pub color: ColorRef,
    pub end_cap: u32,
    pub line_join: u32,
    /// Dash lengths, in the units the pen style measures them in.
    pub dash_pattern: Option<DashPattern>,
    /// Whether the width and style lengths are logical units, which the world
    /// transform and page mapping scale, rather than device units.
    pub logical_width: bool,
}

/// Dash lengths of a pen.
#[derive(Debug, Clone)]
pub enum DashPattern {
    /// Multiples of the pen width, as the fixed pen styles define them.
    Width(Arc<[f64]>),
    /// Lengths in the units of the pen width, as an extended pen's style
    /// array records them ([MS-EMF] 2.2.20 LogPenEx).
    Units(Arc<[f64]>),
}

impl Default for Pen {
    /// The stock black pen: solid and zero width.
    fn default() -> Self {
        Self::from_create_pen(pen_style::SOLID, 0, ColorRef::from_rgb(0, 0, 0))
    }
}

impl Pen {
    /// A pen created from an [MS-EMF] 2.2.19 LogPen.
    ///
    /// GDI creates such a pen with its width in logical units whatever pen
    /// type bits the style carries, so the world transform and page mapping
    /// scale it; a zero width draws one device pixel wide.
    pub fn from_create_pen(style: u32, width: u32, color: ColorRef) -> Self {
        let base = style & 0xff;
        let dash_pattern = match base {
            pen_style::DASH => Some(DashPattern::Width(Arc::from([3.0, 1.0]))),
            pen_style::DOT => Some(DashPattern::Width(Arc::from([1.0, 1.0]))),
            pen_style::DASHDOT => Some(DashPattern::Width(Arc::from([3.0, 1.0, 1.0, 1.0]))),
            pen_style::DASHDOTDOT => Some(DashPattern::Width(Arc::from([
                3.0, 1.0, 1.0, 1.0, 1.0, 1.0,
            ]))),
            _ => None,
        };
        Self {
            style: base,
            width: f64::from(width),
            color,
            end_cap: style & 0x0f00,
            line_join: style & 0xf000,
            dash_pattern,
            logical_width: true,
        }
    }

    /// A pen created from an [MS-EMF] 2.2.20 LogPenEx, whose width is in
    /// logical units for a geometric pen and in device units for a cosmetic
    /// one.
    pub fn from_ext_create_pen(style: u32, width: u32, color: ColorRef) -> Self {
        Self {
            logical_width: style & pen_style::GEOMETRIC != 0,
            ..Self::from_create_pen(style, width, color)
        }
    }
}

#[derive(Debug, Clone)]
pub struct Brush {
    pub style: u32,
    pub color: ColorRef,
    pub hatch: Option<u32>,
    /// Pattern bitmap of a pattern brush, or why its record carries no
    /// usable bitmap.
    pub pattern: Option<std::result::Result<Arc<BrushPattern>, Arc<str>>>,
    pub pattern_id: Option<String>,
}

impl Default for Brush {
    fn default() -> Self {
        Self::from_create_brush(brush_style::SOLID, ColorRef::from_rgb(255, 255, 255), 0)
    }
}

impl Brush {
    pub fn from_create_brush(style: u32, color: ColorRef, hatch: u32) -> Self {
        Self {
            style,
            color,
            hatch: (style == brush_style::HATCHED).then_some(hatch),
            pattern: None,
            pattern_id: None,
        }
    }

    /// Creates a brush painted with a bitmap pattern.
    pub fn from_pattern(
        style: u32,
        pattern: std::result::Result<Arc<BrushPattern>, Arc<str>>,
    ) -> Self {
        Self {
            style,
            color: ColorRef::from_rgb(0, 0, 0),
            hatch: None,
            pattern: Some(pattern),
            pattern_id: None,
        }
    }

    pub fn needs_pattern(&self) -> bool {
        self.pattern.is_some() || (self.style == brush_style::HATCHED && self.hatch.is_some())
    }

    fn svg_pattern(
        &self,
        id: &str,
        background: ColorRef,
        opaque: bool,
        origin: (i32, i32),
    ) -> Option<String> {
        let hatch = self.hatch?;
        let mut content = String::new();
        if opaque {
            content.push_str(&format!(
                "<rect width=\"8\" height=\"8\" fill=\"{}\"/>",
                background.to_svg_color()
            ));
        }
        let color = self.color.to_svg_color();
        let line = |x1, y1, x2, y2| {
            format!("<path d=\"M{x1} {y1}L{x2} {y2}\" stroke=\"{color}\" stroke-width=\"1\"/>")
        };
        match hatch {
            hatch_style::HORIZONTAL => content.push_str(&line(0, 4, 8, 4)),
            hatch_style::VERTICAL => content.push_str(&line(4, 0, 4, 8)),
            hatch_style::FDIAGONAL => content.push_str(&line(0, 0, 8, 8)),
            hatch_style::BDIAGONAL => content.push_str(&line(0, 8, 8, 0)),
            hatch_style::CROSS => {
                content.push_str(&line(0, 4, 8, 4));
                content.push_str(&line(4, 0, 4, 8));
            },
            hatch_style::DIAGCROSS => {
                content.push_str(&line(0, 0, 8, 8));
                content.push_str(&line(0, 8, 8, 0));
            },
            _ => return None,
        }
        Some(format!(
            "<pattern id=\"{}\" patternUnits=\"userSpaceOnUse\" width=\"8\" height=\"8\" x=\"{}\" y=\"{}\">{}</pattern>",
            id, origin.0, origin.1, content
        ))
    }
}

#[derive(Debug, Clone)]
pub struct Font {
    pub height: f64,
    pub width: f64,
    pub escapement: f64,
    pub orientation: f64,
    pub weight: i32,
    pub italic: bool,
    pub underline: bool,
    pub strike_out: bool,
    pub charset: u8,
    /// `lfPitchAndFamily` byte, whose family class names the generic SVG
    /// family to fall back to.
    pub pitch_and_family: u8,
    pub face_name: String,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            height: 12.0,
            width: 0.0,
            escapement: 0.0,
            orientation: 0.0,
            weight: font_weight::NORMAL,
            italic: false,
            underline: false,
            strike_out: false,
            charset: 1,
            pitch_and_family: 0,
            face_name: "Arial".to_string(),
        }
    }
}

impl Font {
    /// Returns the SVG font attributes for the font drawn at `font_size`
    /// device units, naming the recorded face only when the text is drawn
    /// with it.
    ///
    /// Text that a symbol face draws through a glyph order — the Adobe Symbol
    /// encoding in particular — is emitted as the Unicode characters those
    /// glyphs represent, which the recorded symbol face does not contain.
    pub fn svg_attrs(&self, font_size: f64, use_face: bool) -> String {
        let mut attrs = String::from("font-size=\"");
        write_num(&mut attrs, font_size);
        attrs.push('"');
        if use_face {
            let family =
                crate::text_charset::svg_font_family(&self.face_name, self.pitch_and_family);
            attrs.push_str(" font-family=\"");
            write_xml_escaped(&mut attrs, &family);
            attrs.push('"');
        }
        if self.weight != font_weight::NORMAL {
            attrs.push_str(&format!(" font-weight=\"{}\"", self.weight.clamp(1, 1000)));
        }
        if self.italic {
            attrs.push_str(" font-style=\"italic\"");
        }
        let mut decorations = Vec::new();
        if self.underline {
            decorations.push("underline");
        }
        if self.strike_out {
            decorations.push("line-through");
        }
        if !decorations.is_empty() {
            attrs.push_str(&format!(" text-decoration=\"{}\"", decorations.join(" ")));
        }
        attrs
    }
}

fn stock_object(handle: u32) -> Option<GdiObject> {
    let gray = |value| {
        Brush::from_create_brush(
            brush_style::SOLID,
            ColorRef::from_rgb(value, value, value),
            0,
        )
    };
    match handle {
        stock_objects::WHITE_BRUSH => Some(GdiObject::Brush(gray(255))),
        stock_objects::LTGRAY_BRUSH => Some(GdiObject::Brush(gray(192))),
        stock_objects::GRAY_BRUSH => Some(GdiObject::Brush(gray(128))),
        stock_objects::DKGRAY_BRUSH => Some(GdiObject::Brush(gray(64))),
        stock_objects::BLACK_BRUSH | stock_objects::DC_BRUSH => Some(GdiObject::Brush(gray(0))),
        stock_objects::NULL_BRUSH => Some(GdiObject::Brush(Brush::from_create_brush(
            brush_style::NULL,
            ColorRef::from_rgb(0, 0, 0),
            0,
        ))),
        stock_objects::WHITE_PEN => Some(GdiObject::Pen(Pen::from_create_pen(
            pen_style::SOLID,
            0,
            ColorRef::from_rgb(255, 255, 255),
        ))),
        stock_objects::BLACK_PEN | stock_objects::DC_PEN => Some(GdiObject::Pen(Pen::default())),
        stock_objects::NULL_PEN => Some(GdiObject::Pen(Pen::from_create_pen(
            pen_style::NULL,
            0,
            ColorRef::from_rgb(0, 0, 0),
        ))),
        stock_objects::OEM_FIXED_FONT
        | stock_objects::ANSI_FIXED_FONT
        | stock_objects::ANSI_VAR_FONT
        | stock_objects::SYSTEM_FONT
        | stock_objects::DEVICE_DEFAULT_FONT
        | stock_objects::SYSTEM_FIXED_FONT
        | stock_objects::DEFAULT_GUI_FONT => Some(GdiObject::Font(Font::default())),
        stock_objects::DEFAULT_PALETTE => Some(GdiObject::Palette),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn world_transform_precedes_page_mapping() {
        let mut dc = DeviceContext {
            map_mode: 8,
            window_ext: (10, 10),
            viewport_ext: (100, 100),
            ..DeviceContext::default()
        };
        dc.world_transform.dx = 2.0;
        assert_eq!(dc.transform_point(1.0, 1.0), (30.0, 10.0));
    }

    #[test]
    fn restore_discards_restored_and_newer_states() {
        let mut state = RenderState::new();
        state.push_dc();
        state.dc.current_pos = (1.0, 0.0);
        state.push_dc();
        state.dc.current_pos = (2.0, 0.0);
        assert!(state.pop_dc(-2));
        assert_eq!(state.dc.current_pos, (0.0, 0.0));
        assert!(state.dc_stack.is_empty());
    }

    fn path(data: &str) -> ClipShape<'_> {
        ClipShape::Path(data, ClipRule::NonZero)
    }

    fn canvas_rect() -> Canvas {
        Canvas {
            left: 0.0,
            top: 0.0,
            right: 10.0,
            bottom: 10.0,
        }
    }

    /// Rasterizes a black square of the given geometry under the state's
    /// definitions and returns how many pixels it covers.
    fn covered_pixels(state: &RenderState, clip: &str) -> usize {
        let svg = format!(
            concat!(
                r#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" "#,
                r##"viewBox="0 0 10 10"><defs>{}</defs>"##,
                r##"<rect width="10" height="10" fill="#000000" clip-path="url(#{})"/></svg>"##,
            ),
            state.defs.join(""),
            clip
        );
        let png = crate::raster::rasterize_svg(
            &svg,
            10,
            10,
            image::Rgba([255, 255, 255, 255]),
            image::ImageFormat::Png,
            &crate::raster::RasterLimits::default(),
        )
        .expect("clip markup must rasterize");
        let image = image::load_from_memory_with_format(&png, image::ImageFormat::Png)
            .expect("rasterized PNG must decode")
            .to_rgba8();
        image.pixels().filter(|pixel| pixel.0 != [255; 4]).count()
    }

    #[test]
    fn intersected_clip_keeps_both_regions() {
        let mut state = RenderState::new();
        assert!(state.install_clip(path("M0 0L10 0 10 10 0 10Z"), 1, canvas_rect()));
        assert!(state.install_clip(path("M0 0L5 0 5 10 0 10Z"), 1, canvas_rect()));
        let clip = state.dc.clip_id.clone().unwrap();

        // The intersection covers the left half, not nothing and not all.
        assert_eq!(covered_pixels(&state, &clip), 50);
    }

    #[test]
    fn offset_clip_moves_the_region() {
        let mut state = RenderState::new();
        assert!(state.install_clip(path("M0 0L2 0 2 2 0 2Z"), 1, canvas_rect()));
        state.offset_clip(4.0, 3.0, canvas_rect());
        let clip = state.dc.clip_id.clone().unwrap();

        assert_eq!(covered_pixels(&state, &clip), 4);
    }

    #[test]
    fn offset_clip_brings_an_off_canvas_region_onto_the_canvas() {
        let mut state = RenderState::new();
        assert!(state.install_clip(path("M-6 0L-2 0 -2 2 -6 2Z"), 1, canvas_rect()));
        state.offset_clip(6.0, 0.0, canvas_rect());
        let clip = state.dc.clip_id.clone().unwrap();

        assert_eq!(covered_pixels(&state, &clip), 8);
    }

    #[test]
    fn even_odd_clip_cuts_out_an_equally_oriented_inner_rectangle() {
        let mut state = RenderState::new();
        assert!(state.install_clip(
            ClipShape::Path("M0 0L10 0 10 10 0 10Z M2 2L8 2 8 8 2 8Z", ClipRule::EvenOdd),
            1,
            canvas_rect()
        ));
        let clip = state.dc.clip_id.clone().unwrap();

        assert_eq!(covered_pixels(&state, &clip), 100 - 36);
    }

    #[test]
    fn rectangle_clips_intersect_without_chaining_references() {
        let mut state = RenderState::new();
        for step in 0..10_000 {
            let inset = f64::from(step % 4);
            assert!(state.install_clip(
                ClipShape::Rect([inset, 0.0, 10.0 - inset, 8.0]),
                1,
                canvas_rect()
            ));
        }
        let clip = state.dc.clip_id.clone().unwrap();
        assert!(
            state
                .defs
                .iter()
                .all(|definition| !definition.contains("url("))
        );
        // Every rectangle intersected: x 3..7, y 0..8.
        assert_eq!(covered_pixels(&state, &clip), 32);

        // A rectangle that holds the canvas needs no clip reference, but
        // still bounds the clip paths intersected with it later.
        let mut state = RenderState::new();
        assert!(state.install_clip(ClipShape::Rect([-5.0, -5.0, 20.0, 20.0]), 1, canvas_rect()));
        assert_eq!(state.dc.clip_id, None);
        assert!(state.install_clip(ClipShape::Rect([0.0, 0.0, 4.0, 4.0]), 1, canvas_rect()));
        assert!(state.install_clip(path("M2 0L10 0 10 10 2 10Z"), 1, canvas_rect()));
        let clip = state.dc.clip_id.clone().unwrap();
        assert_eq!(covered_pixels(&state, &clip), 8);
        state.offset_clip(1.0, 0.0, canvas_rect());
        let clip = state.dc.clip_id.clone().unwrap();
        assert_eq!(covered_pixels(&state, &clip), 8);
    }

    #[test]
    fn clip_path_chain_stops_growing_at_its_depth_limit() {
        let mut state = RenderState::new();
        for _ in 0..MAX_CLIP_CHAIN_DEPTH + 100 {
            assert!(state.install_clip(path("M0 0L5 0 5 10 0 10Z"), 1, canvas_rect()));
            state.offset_clip(0.0, 0.0, canvas_rect());
        }
        assert_eq!(
            state.dc.clip_region.chain.as_ref().unwrap().1,
            MAX_CLIP_CHAIN_DEPTH
        );
        assert!(state.defs.len() <= MAX_CLIP_CHAIN_DEPTH);
        assert_eq!(
            state
                .diagnostics
                .iter()
                .map(|(code, _)| *code)
                .collect::<Vec<_>>(),
            ["emf-clip-depth-approximated", "emf-clip-depth-approximated"]
        );
        let clip = state.dc.clip_id.clone().unwrap();
        assert_eq!(covered_pixels(&state, &clip), 50);
    }
}
