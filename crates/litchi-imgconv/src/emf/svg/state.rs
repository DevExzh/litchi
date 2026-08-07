//! SVG playback state for classic EMF records.

use super::path::PathBuilder;
use crate::emf::records::{
    ColorRef, XForm, brush_style, font_weight, hatch_style, pen_style, stock_objects,
};
use crate::svg_utils::{write_num, write_xml_escaped};
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
}

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

    pub fn insert_object(&mut self, handle: u32, object: GdiObject) -> bool {
        self.objects.insert(handle, object).is_none()
    }

    pub fn delete_object(&mut self, handle: u32) -> bool {
        !stock_objects::is_stock_object(handle)
            && !self.dc.references(handle)
            && !self.dc_stack.iter().any(|dc| dc.references(handle))
            && self.objects.remove(&handle).is_some()
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

    pub fn prepare_brush_pattern(&mut self) {
        if !self.dc.brush.needs_pattern() || self.dc.brush.pattern_id.is_some() {
            return;
        }
        let id = self.fresh_id("hatch");
        if let Some(definition) = self.dc.brush.svg_pattern(
            &id,
            self.dc.bg_color,
            self.dc.bg_mode == 2,
            self.dc.brush_org,
        ) {
            self.dc.brush.pattern_id = Some(id);
            self.add_definition(definition);
        }
    }

    pub fn install_clip(&mut self, path: &str, mode: u32) -> bool {
        let previous = self.dc.clip_id.clone();
        let id = self.fresh_id("clip");
        let definition = match (mode, previous) {
            (5, _) | (1 | 2, None) => format!(
                "<clipPath id=\"{}\" clipPathUnits=\"userSpaceOnUse\"><path d=\"{}\"/></clipPath>",
                id, path
            ),
            (1, Some(old)) => format!(
                "<clipPath id=\"{}\" clipPathUnits=\"userSpaceOnUse\"><g clip-path=\"url(#{})\"><path d=\"{}\"/></g></clipPath>",
                id, old, path
            ),
            // OR, XOR and DIFF need a stored copy of the old geometry or mask
            // semantics; do not pretend a lossy clip is equivalent.
            _ => return false,
        };
        self.add_definition(definition);
        self.dc.clip_id = Some(id);
        true
    }

    pub fn offset_clip(&mut self, x: f64, y: f64) {
        let Some(old) = self.dc.clip_id.clone() else {
            return;
        };
        let id = self.fresh_id("clip");
        self.add_definition(format!(
            "<clipPath id=\"{}\" clipPathUnits=\"userSpaceOnUse\"><g transform=\"translate({} {})\" clip-path=\"url(#{})\"><rect x=\"-1000000000\" y=\"-1000000000\" width=\"2000000000\" height=\"2000000000\"/></g></clipPath>",
            id, x, y, old
        ));
        self.dc.clip_id = Some(id);
    }
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
    pub rop2: u32,
    pub stretch_mode: u32,
    pub arc_direction: bool,
    pub miter_limit: f64,
    pub layout: u32,
    pub clip_id: Option<String>,
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
            rop2: 13,
            stretch_mode: 3,
            arc_direction: false,
            miter_limit: 10.0,
            layout: 0,
            clip_id: None,
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
        if self.pen.width > 1.0 {
            attrs.push_str(" stroke-width=\"");
            write_num(&mut attrs, self.pen.width);
            attrs.push('"');
        }
        if let Some(dash) = &self.pen.dash_pattern {
            attrs.push_str(" stroke-dasharray=\"");
            attrs.push_str(dash);
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
        attrs
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

    fn references(&self, handle: u32) -> bool {
        self.pen_handle == Some(handle)
            || self.brush_handle == Some(handle)
            || self.font_handle == Some(handle)
    }
}

#[derive(Debug, Clone)]
pub struct Pen {
    pub style: u32,
    pub width: f64,
    pub color: ColorRef,
    pub end_cap: u32,
    pub line_join: u32,
    pub dash_pattern: Option<Arc<str>>,
}

impl Default for Pen {
    fn default() -> Self {
        Self::from_create_pen(pen_style::SOLID, 1, ColorRef::from_rgb(0, 0, 0))
    }
}

impl Pen {
    pub fn from_create_pen(style: u32, width: u32, color: ColorRef) -> Self {
        let base = style & 0xff;
        let unit = f64::from(width.max(1));
        let dash_pattern = match base {
            pen_style::DASH => Some(Arc::from(format!("{} {}", unit * 3.0, unit))),
            pen_style::DOT => Some(Arc::from(format!("{} {}", unit, unit))),
            pen_style::DASHDOT => Some(Arc::from(format!(
                "{} {} {} {}",
                unit * 3.0,
                unit,
                unit,
                unit
            ))),
            pen_style::DASHDOTDOT => Some(
                format!(
                    "{} {} {} {} {} {}",
                    unit * 3.0,
                    unit,
                    unit,
                    unit,
                    unit,
                    unit
                )
                .into(),
            ),
            _ => None,
        };
        Self {
            style: base,
            width: unit,
            color,
            end_cap: style & 0x0f00,
            line_join: style & 0xf000,
            dash_pattern,
        }
    }
}

#[derive(Debug, Clone)]
pub struct Brush {
    pub style: u32,
    pub color: ColorRef,
    pub hatch: Option<u32>,
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
            pattern_id: None,
        }
    }

    pub fn needs_pattern(&self) -> bool {
        self.style == brush_style::HATCHED && self.hatch.is_some()
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
            face_name: "Arial".to_string(),
        }
    }
}

impl Font {
    pub fn to_svg_attrs(&self) -> String {
        let mut family = String::with_capacity(self.face_name.len());
        write_xml_escaped(&mut family, &self.face_name);
        let mut attrs = format!(
            "font-size=\"{}\" font-family=\"{}\"",
            self.height.abs().max(1.0),
            family
        );
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
            1,
            ColorRef::from_rgb(255, 255, 255),
        ))),
        stock_objects::BLACK_PEN | stock_objects::DC_PEN => Some(GdiObject::Pen(Pen::default())),
        stock_objects::NULL_PEN => Some(GdiObject::Pen(Pen::from_create_pen(
            pen_style::NULL,
            1,
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
        let mut dc = DeviceContext::default();
        dc.map_mode = 8;
        dc.window_ext = (10, 10);
        dc.viewport_ext = (100, 100);
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
}
