//! Stateful, bounded WMF record playback to SVG.

use std::{collections::HashMap, mem::size_of, sync::Arc};

use super::super::constants::{bk_mode, brush, ext_text_out, record, rop3, stock, text_align};
use super::super::parser::WmfRecord;
use super::BitmapHook;
use super::state::{
    Brush, ClipRegion, DeviceRect, Font, GdiObject, GraphicsState, ObjectTable, Palette, Pen,
    Region,
};
use super::style::{
    escape_xml_attr_into, escape_xml_text_into, fill_attrs, hatch_definition, map_font_family,
    stroke_attrs,
};
use super::transform::CoordinateTransform;
use crate::svg_utils::{write_color_hex, write_num};
use crate::{
    dib::DibLimits,
    metafile_bitmap::{BitmapOp, StretchPolicy},
};
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
    pub message: String,
}

pub(super) struct SvgRenderer<'hook> {
    transform: CoordinateTransform,
    state: GraphicsState,
    saved_states: Vec<GraphicsState>,
    objects: ObjectTable,
    definitions: String,
    hatch_ids: HashMap<(u32, u16, u16, u32), String>,
    clip_ids: HashMap<u64, String>,
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

    pub(super) fn render_record(&mut self, rec: &WmfRecord) -> Option<String> {
        if self.halted {
            return None;
        }
        let function = record::canonical(rec.function);
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
            record::PAT_BLT => self.render_pat_blt(rec),
            record::FLOOD_FILL | record::EXT_FLOOD_FILL => self.render_flood_fill(rec),
            record::FILL_REGION
            | record::FRAME_REGION
            | record::INVERT_REGION
            | record::PAINT_REGION => self.render_region_record(function, rec),
            record::BIT_BLT
            | record::STRETCH_BLT
            | record::DIB_BIT_BLT
            | record::DIB_STRETCH_BLT
            | record::SET_DIB_TO_DEV
            | record::STRETCH_DIB => self.render_bitmap(function, rec),

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
                self.set_u16(function, rec, |state, value| state.rop2 = value);
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

            _ => {
                self.fatal(format!(
                    "unsupported output-affecting WMF record 0x{:04X}",
                    rec.function
                ));
                None
            },
        };
        if self.objects.retained_len() > self.max_objects {
            self.fatal(format!(
                "WMF retained object count exceeds limit {}",
                self.max_objects
            ));
        }
        result
    }

    fn set_u16(&mut self, function: u16, rec: &WmfRecord, setter: fn(&mut GraphicsState, u16)) {
        if let Ok(value) = read_u16_le(&rec.params, 0) {
            setter(&mut self.state, value);
        } else {
            self.malformed(function, "state value");
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
            self.warn(format!("ignored invalid RestoreDC level {level}"));
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
        let payload_offset = usize::from(function == record::CREATE_DIB_PATTERN_BRUSH) * 4;
        if rec.params.len() <= payload_offset {
            self.malformed(function, "pattern bitmap");
            return;
        }
        let payload = &rec.params[payload_offset..];
        if !self.reserve_object_bytes(payload.len()) {
            return;
        }
        self.objects.insert(GdiObject::Brush(Brush {
            style: if function == record::CREATE_DIB_PATTERN_BRUSH {
                brush::BS_DIBPATTERN
            } else {
                brush::BS_PATTERN
            },
            color: 0,
            hatch: 0,
            pattern: Some(Arc::from(payload)),
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
        let name = if name.is_empty() {
            "serif".to_owned()
        } else {
            map_font_family(&name).to_owned()
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
            Some(GdiObject::Palette(_)) | Some(GdiObject::Region(_)) => {
                self.warn(format!(
                    "SelectObject ignored non-selectable object {handle}"
                ));
            },
            None => self.fatal(format!("SelectObject references missing object {handle}")),
        }
    }

    fn delete_object(&mut self, rec: &WmfRecord) {
        let Ok(index) = read_u16_le(&rec.params, 0) else {
            self.malformed(record::DELETE_OBJECT, "object index");
            return;
        };
        if index & stock::FLAG == 0 && !self.objects.delete(usize::from(index)) {
            self.warn(format!("DeleteObject references missing object {index}"));
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
            self.fatal(format!("SelectPalette references missing palette {index}"));
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
        let Some(GdiObject::Region(region)) = self.objects.get(usize::from(index)) else {
            self.fatal(format!(
                "SelectClipRegion references missing region {index}"
            ));
            return;
        };
        let rect_count = region.rects.len();
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
        let minimum = if closed { 2 } else { 2 };
        if count < minimum || rec.params.len() < 2 + count.saturating_mul(4) {
            self.malformed(function, "points");
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
        let size = self.transform.device_width(1.0).max(1.0);
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
        let operation = read_u32(&rec.params, 0).unwrap_or(0);
        let height = read_i16_le(&rec.params, 4).unwrap_or(0);
        let width = read_i16_le(&rec.params, 6).unwrap_or(0);
        let y = read_i16_le(&rec.params, 8).unwrap_or(0);
        let x = read_i16_le(&rec.params, 10).unwrap_or(0);
        let rect = self.logical_origin_size(x, y, width, height);
        let mut output = String::with_capacity(224);
        write_rect_start(&mut output, rect);
        match operation {
            rop3::PATCOPY => {
                let brush = self.state.brush.clone();
                output.push_str(&self.brush_fill(&brush));
            },
            rop3::BLACKNESS => output.push_str(r##" fill="#000""##),
            rop3::WHITENESS => output.push_str(r##" fill="#fff""##),
            rop3::DSTINVERT => {
                self.fatal(
                    "PatBlt DSTINVERT requires destination-dependent raster composition".to_owned(),
                );
                output.push_str(r##" fill="#fff" style="mix-blend-mode:difference""##)
            },
            rop3::PATINVERT => {
                self.fatal(
                    "PatBlt PATINVERT requires destination-dependent raster composition".to_owned(),
                );
                let brush = self.state.brush.clone();
                output.push_str(&self.brush_fill(&brush));
                output.push_str(r#" style="mix-blend-mode:difference""#);
            },
            _ => {
                self.fatal(format!(
                    "unsupported PatBlt raster operation 0x{operation:08X}"
                ));
                output.push_str(r#" fill="none" data-wmf-unsupported-rop=""#);
                let _ = std::fmt::Write::write_fmt(&mut output, format_args!("{operation:08X}"));
                output.push('"');
            },
        }
        output.push_str(&self.clip_attr());
        output.push_str("/>");
        Some(output)
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
        output.push_str(&self.brush_fill(&brush));
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
                (region_index, self.object_brush(brush_index)?, None)
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
                (
                    region_index,
                    self.object_brush(brush_index)?,
                    Some((width, height)),
                )
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
            self.fatal(format!(
                "region record references missing region {region_index}"
            ));
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
            output.push_str(r#" fill="none" stroke=""#);
            if brush_value.style == brush::BS_HATCHED {
                let id = self.ensure_hatch(&brush_value);
                output.push_str("url(#");
                output.push_str(&id);
                output.push(')');
            } else {
                write_color_hex(&mut output, brush_value.color);
            }
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
            output.push_str(&self.brush_fill(&brush_value));
        }
        output.push_str(&self.clip_attr());
        output.push_str("/>");
        Some(output)
    }

    fn object_brush(&mut self, index: usize) -> Option<Brush> {
        match self.objects.get(index).cloned() {
            Some(GdiObject::Brush(value)) => Some(value),
            _ => {
                self.fatal(format!("region record references missing brush {index}"));
                None
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
            self.fatal("ExtTextOut glyph-index strings are unsupported".to_owned());
        }
        if options & ext_text_out::ETO_PDY != 0 {
            self.fatal("ExtTextOut ETO_PDY displacements are unsupported".to_owned());
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
        let dx_offset = string_offset + ((length + 1) & !1);
        let dx_count = (rec.params.len().saturating_sub(dx_offset)) / 2;
        let mut dx = Vec::with_capacity(dx_count.min(length));
        for index in 0..dx_count.min(length) {
            dx.push(read_i16_le(&rec.params, dx_offset + index * 2).unwrap_or(0));
        }
        self.render_text_bytes(
            &rec.params[string_offset..string_offset + length],
            x,
            y,
            options,
            rect,
            &dx,
        )
    }

    fn render_text_bytes(
        &mut self,
        bytes: &[u8],
        supplied_x: i16,
        supplied_y: i16,
        options: u16,
        option_rect: Option<(i16, i16, i16, i16)>,
        dx: &[i16],
    ) -> Option<String> {
        if bytes.is_empty() {
            return None;
        }
        let (text, exact_charset) = decode_text(bytes, self.state.font.charset);
        if !exact_charset {
            self.warn(format!(
                "font charset {} decoded with a safe Windows-1252 fallback",
                self.state.font.charset
            ));
        }
        let update_position = self.state.text_align & text_align::TA_UPDATECP != 0;
        let (logical_x, logical_y) = if update_position {
            self.state.position
        } else {
            (supplied_x, supplied_y)
        };
        let (x, y) = self
            .transform
            .point(&self.state.mapping, logical_x, logical_y);
        let (_, font_height) = self
            .state
            .mapping
            .vector(0.0, f64::from(self.state.font.height).abs().max(1.0));
        let font_size = self.transform.device_height(font_height).max(1.0);
        let logical_advance = if dx.is_empty() {
            let nominal = if self.state.font.width == 0 {
                f64::from(self.state.font.height).abs() * 0.6
            } else {
                f64::from(self.state.font.width).abs()
            };
            (nominal + f64::from(self.state.text_char_extra)) * bytes.len() as f64
        } else {
            dx.iter().map(|&value| f64::from(value)).sum()
        };

        let mut text_element = String::with_capacity(text.len() + dx.len() * 8 + 320);
        text_element.push_str(r#"<text x=""#);
        write_num(&mut text_element, x);
        text_element.push_str(r#"" y=""#);
        write_num(&mut text_element, y);
        text_element.push_str(r#"" font-size=""#);
        write_num(&mut text_element, font_size);
        text_element.push_str(r#"" fill=""#);
        write_color_hex(&mut text_element, self.state.text_color);
        text_element.push('"');
        match self.state.text_align & text_align::HORIZONTAL_MASK {
            text_align::TA_RIGHT => text_element.push_str(r#" text-anchor="end""#),
            text_align::TA_CENTER => text_element.push_str(r#" text-anchor="middle""#),
            _ => {},
        }
        match self.state.text_align & text_align::VERTICAL_MASK {
            text_align::TA_BOTTOM => {
                text_element.push_str(r#" dominant-baseline="text-after-edge""#)
            },
            text_align::TA_BASELINE => {},
            _ => text_element.push_str(r#" dominant-baseline="text-before-edge""#),
        }
        if self.state.font.name != "serif" {
            text_element.push_str(r#" font-family=""#);
            escape_xml_attr_into(&mut text_element, &self.state.font.name);
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
        if !dx.is_empty() {
            text_element.push_str(r#" dx="0"#);
            for value in dx.iter().take(bytes.len().saturating_sub(1)) {
                text_element.push(' ');
                let (advance, _) =
                    self.transform
                        .logical_vector(&self.state.mapping, f64::from(*value), 0.0);
                write_num(&mut text_element, advance);
            }
            text_element.push('"');
        } else if self.state.text_char_extra != 0 {
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
        text_element.push('>');
        escape_xml_text_into(&mut text_element, &text);
        text_element.push_str("</text>");

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
        let Some(rect) =
            bitmap_destination(function, &rec.params, &self.state.mapping, &self.transform)
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
            return match hook(function, &rec.params, request) {
                Ok(Some(svg)) => {
                    if self.state.clip.is_unbounded() {
                        Some(svg)
                    } else {
                        let mut output = String::with_capacity(svg.len() + 64);
                        output.push_str("<g");
                        output.push_str(&self.clip_attr());
                        output.push('>');
                        output.push_str(&svg);
                        output.push_str("</g>");
                        Some(output)
                    }
                },
                Ok(None) => {
                    self.fatal(format!(
                        "DIB rendering hook declined WMF record 0x{function:04X}"
                    ));
                    None
                },
                Err(message) => {
                    self.fatal(format!("DIB rendering hook failed: {message}"));
                    None
                },
            };
        }

        let stretch = match self.state.stretch_mode {
            1 => StretchPolicy::BlackOnWhite,
            2 => StretchPolicy::WhiteOnBlack,
            3 => StretchPolicy::ColorOnColor,
            _ => StretchPolicy::Halftone,
        };
        let rendered = BitmapOp::parse_wmf(
            rec.size,
            rec.function,
            &rec.params,
            stretch,
            self.dib_limits,
        )
        .and_then(|operation| operation.to_svg_image_at(Some(request)));
        match rendered {
            Ok(image) if self.state.clip.is_unbounded() => Some(image.element),
            Ok(image) => {
                let mut output = String::with_capacity(image.element.len() + 64);
                output.push_str("<g");
                output.push_str(&self.clip_attr());
                output.push('>');
                output.push_str(&image.element);
                output.push_str("</g>");
                Some(output)
            },
            Err(error) => {
                self.fatal(format!("WMF bitmap record 0x{function:04X}: {error}"));
                None
            },
        }
    }

    fn append_shape_attrs(&mut self, output: &mut String, fill: bool, stroke: bool) {
        if fill {
            let brush = self.state.brush.clone();
            output.push_str(&self.brush_fill(&brush));
        } else {
            output.push_str(r#" fill="none""#);
        }
        if stroke {
            output.push_str(&stroke_attrs(
                &self.state.pen,
                &self.state.mapping,
                &self.transform,
            ));
            output.push_str(&self.rop2_attr());
        }
        output.push_str(&self.clip_attr());
    }

    fn brush_fill(&mut self, brush_value: &Brush) -> String {
        if brush_value.pattern.is_some()
            || !matches!(
                brush_value.style,
                brush::BS_SOLID | brush::BS_NULL | brush::BS_HATCHED
            )
        {
            self.fatal(
                "selected bitmap/indexed pattern brush needs DIB pattern integration".to_owned(),
            );
            return fill_attrs(brush_value, self.state.poly_fill_mode, None);
        }
        let pattern =
            (brush_value.style == brush::BS_HATCHED).then(|| self.ensure_hatch(brush_value));
        fill_attrs(brush_value, self.state.poly_fill_mode, pattern.as_deref())
    }

    fn ensure_hatch(&mut self, brush_value: &Brush) -> String {
        let key = (
            brush_value.color,
            brush_value.hatch,
            self.state.bk_mode,
            self.state.bk_color,
        );
        if let Some(id) = self.hatch_ids.get(&key) {
            return id.clone();
        }
        let id = self.next_id("wmf-hatch");
        let definition =
            hatch_definition(&id, brush_value, self.state.bk_mode, self.state.bk_color);
        self.push_definition(&definition);
        self.hatch_ids.insert(key, id.clone());
        id
    }

    fn rop2_attr(&mut self) -> String {
        if self.state.rop2 == 13 {
            return String::new();
        }
        let mut attr = format!(r#" data-wmf-rop2="{}""#, self.state.rop2);
        if self.state.rop2 == 7 {
            attr.push_str(r#" style="mix-blend-mode:difference""#);
        }
        self.fatal(format!(
            "ROP2 mode {} cannot be faithfully represented in SVG",
            self.state.rop2
        ));
        attr
    }

    fn clip_attr(&mut self) -> String {
        if self.state.clip.is_unbounded() {
            return String::new();
        }
        if let Some(id) = self.clip_ids.get(&self.state.clip_revision) {
            return format!(r#" clip-path="url(#{id})""#);
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
        self.clip_ids.insert(self.state.clip_revision, id.clone());
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

    fn warn(&mut self, message: String) {
        if !self.issues.iter().any(|issue| issue.message == message) {
            self.issues.push(RenderIssue {
                fatal: false,
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

fn decode_text(bytes: &[u8], charset: u8) -> (String, bool) {
    if charset == 2 {
        return (
            bytes
                .iter()
                .map(|&byte| char::from_u32(0xf000 + u32::from(byte)).unwrap_or('\u{fffd}'))
                .collect(),
            true,
        );
    }
    let exact = matches!(charset, 0 | 1);
    let text = bytes
        .iter()
        .map(|&byte| match byte {
            0x80 => '€',
            0x82 => '‚',
            0x83 => 'ƒ',
            0x84 => '„',
            0x85 => '…',
            0x86 => '†',
            0x87 => '‡',
            0x88 => 'ˆ',
            0x89 => '‰',
            0x8a => 'Š',
            0x8b => '‹',
            0x8c => 'Œ',
            0x8e => 'Ž',
            0x91 => '‘',
            0x92 => '’',
            0x93 => '“',
            0x94 => '”',
            0x95 => '•',
            0x96 => '–',
            0x97 => '—',
            0x98 => '˜',
            0x99 => '™',
            0x9a => 'š',
            0x9b => '›',
            0x9c => 'œ',
            0x9e => 'ž',
            0x9f => 'Ÿ',
            0x81 | 0x8d | 0x8f | 0x90 | 0x9d => '\u{fffd}',
            _ => char::from(byte),
        })
        .collect();
    (text, exact)
}

fn bitmap_destination(
    function: u16,
    data: &[u8],
    mapping: &super::state::MappingState,
    transform: &CoordinateTransform,
) -> Option<DeviceRect> {
    let (height_offset, width_offset, y_offset, x_offset) = match function {
        record::BIT_BLT | record::DIB_BIT_BLT => (8, 10, 12, 14),
        record::STRETCH_BLT | record::DIB_STRETCH_BLT => (12, 14, 16, 18),
        record::SET_DIB_TO_DEV => (10, 12, 14, 16),
        record::STRETCH_DIB => (14, 16, 18, 20),
        _ => return None,
    };
    let height = read_i16_le(data, height_offset).ok()?;
    let width = read_i16_le(data, width_offset).ok()?;
    let y = read_i16_le(data, y_offset).ok()?;
    let x = read_i16_le(data, x_offset).ok()?;
    let (x1, y1) = mapping.point(x, y);
    let (dx, dy) = mapping.vector(f64::from(width), f64::from(height));
    Some(transform.rect(DeviceRect::new(x1, y1, x1 + dx, y1 + dy)))
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

    fn renderer() -> SvgRenderer<'static> {
        SvgRenderer::new(
            CoordinateTransform::new((0.0, 0.0, 100.0, 100.0), 100.0, 100.0),
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
                renderer.issues.iter().all(|issue| !issue
                    .message
                    .starts_with("unsupported output-affecting WMF record")),
                "record 0x{function:04X} fell through classification"
            );
        }
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
        let (text, exact) = decode_text(&[0x80, b'<', b'&'], 0);
        assert!(exact);
        let mut escaped = String::new();
        escape_xml_text_into(&mut escaped, &text);
        assert_eq!(escaped, "€&lt;&amp;");
    }
}
