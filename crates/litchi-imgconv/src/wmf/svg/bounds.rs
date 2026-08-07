//! Device-space bounding-box replay for WMF records.
//!
//! A WMF without a placeable header has no declared drawing rectangle.  The
//! records therefore need to be replayed far enough to apply the active
//! logical-to-device mapping before estimating the output bounds.  This is
//! deliberately a safe, conservative replay: malformed records are ignored,
//! and no input value can create a non-finite output bound.

use super::super::constants::{record, text_align};
use super::super::parser::WmfRecord;
use super::state::MappingState;
use litchi_core::binary::{read_i16_le, read_u16_le};

const FALLBACK_LEFT: f64 = 0.0;
const FALLBACK_TOP: f64 = 0.0;
const FALLBACK_RIGHT: f64 = 1000.0;
const FALLBACK_BOTTOM: f64 = 1000.0;
const MAX_COORDINATE: f64 = 1.0e12;

/// A normalized device-space rectangle.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct Bounds {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

impl Bounds {
    pub(super) const fn fallback() -> Self {
        Self {
            left: FALLBACK_LEFT,
            top: FALLBACK_TOP,
            right: FALLBACK_RIGHT,
            bottom: FALLBACK_BOTTOM,
        }
    }

    pub(super) fn as_tuple(self) -> (f64, f64, f64, f64) {
        (self.left, self.top, self.right, self.bottom)
    }

    fn include_point(&mut self, x: f64, y: f64) {
        let Some((x, y)) = finite_point(x, y) else {
            return;
        };
        self.left = self.left.min(x);
        self.top = self.top.min(y);
        self.right = self.right.max(x);
        self.bottom = self.bottom.max(y);
    }

    fn include_rect(&mut self, left: f64, top: f64, right: f64, bottom: f64) {
        self.include_point(left, top);
        self.include_point(right, bottom);
    }
}

#[derive(Debug, Clone, Copy)]
struct FontMetrics {
    height: f64,
    width: f64,
    escapement: f64,
}

#[derive(Debug, Clone, Copy)]
enum ReplayObject {
    Other,
    Font(FontMetrics),
}

impl Default for FontMetrics {
    fn default() -> Self {
        Self {
            height: 12.0,
            width: 0.0,
            escapement: 0.0,
        }
    }
}

#[derive(Debug, Clone)]
struct ReplayState {
    mapping: MappingState,
    current_position: (f64, f64),
    font: FontMetrics,
    objects: Vec<Option<ReplayObject>>,
    text_align: u16,
    text_char_extra: f64,
}

impl Default for ReplayState {
    fn default() -> Self {
        Self {
            mapping: MappingState::default(),
            current_position: (0.0, 0.0),
            font: FontMetrics::default(),
            objects: Vec::new(),
            text_align: text_align::TA_LEFT | text_align::TA_TOP,
            text_char_extra: 0.0,
        }
    }
}

/// SaveDC snapshots device-context selection and mapping, but not the shared
/// GDI object table. Objects created or deleted inside a saved DC remain so
/// after RestoreDC.
#[derive(Debug, Clone)]
struct ReplaySnapshot {
    mapping: MappingState,
    current_position: (f64, f64),
    font: FontMetrics,
    text_align: u16,
    text_char_extra: f64,
}

impl ReplaySnapshot {
    fn capture(state: &ReplayState) -> Self {
        Self {
            mapping: state.mapping,
            current_position: state.current_position,
            font: state.font,
            text_align: state.text_align,
            text_char_extra: state.text_char_extra,
        }
    }

    fn restore_into(self, state: &mut ReplayState) {
        state.mapping = self.mapping;
        state.current_position = self.current_position;
        state.font = self.font;
        state.text_align = self.text_align;
        state.text_char_extra = self.text_char_extra;
    }
}

/// Calculates the effective device-space bounds of a WMF record stream.
pub(super) struct BoundsCalculator;

impl BoundsCalculator {
    /// Scan records in playback order.  The result is always finite and has
    /// normalized edges, even for reversed rectangles and `i16` extrema.
    pub(super) fn scan_records(records: &[WmfRecord]) -> Bounds {
        let mut state = ReplayState::default();
        let mut saves = Vec::<ReplaySnapshot>::new();
        let mut bounds: Option<Bounds> = None;

        for rec in records {
            let function = record::canonical(rec.function);
            match function {
                record::SAVE_DC => saves.push(ReplaySnapshot::capture(&state)),
                record::RESTORE_DC => Self::restore_dc(&mut state, &mut saves, rec),

                record::SET_MAP_MODE => {
                    if let Some(mode) = u16_at(rec, 0) {
                        state.mapping.set_mode(mode);
                    }
                },
                record::SET_LAYOUT => {
                    if let Some(layout) = u16_at(rec, 0) {
                        state.mapping.layout = layout;
                    }
                },
                record::SET_WINDOW_ORG => {
                    if let Some((x, y)) = point_at(rec, 0) {
                        state.mapping.window_origin = (x, y);
                    }
                },
                record::SET_WINDOW_EXT => {
                    if state.mapping.scalable_extents()
                        && let Some((x, y)) = point_at(rec, 0)
                    {
                        state.mapping.window_extent = (finite_or_one(x), finite_or_one(y));
                    }
                },
                record::SET_VIEWPORT_ORG => {
                    if let Some((x, y)) = point_at(rec, 0) {
                        state.mapping.viewport_origin = (x, y);
                    }
                },
                record::SET_VIEWPORT_EXT => {
                    if state.mapping.scalable_extents()
                        && let Some((x, y)) = point_at(rec, 0)
                    {
                        state.mapping.viewport_extent = (finite_or_one(x), finite_or_one(y));
                    }
                },
                record::OFFSET_WINDOW_ORG => {
                    if let Some((x, y)) = point_at(rec, 0) {
                        state.mapping.window_origin.0 =
                            bounded_add(state.mapping.window_origin.0, x);
                        state.mapping.window_origin.1 =
                            bounded_add(state.mapping.window_origin.1, y);
                    }
                },
                record::OFFSET_VIEWPORT_ORG => {
                    if let Some((x, y)) = point_at(rec, 0) {
                        state.mapping.viewport_origin.0 =
                            bounded_add(state.mapping.viewport_origin.0, x);
                        state.mapping.viewport_origin.1 =
                            bounded_add(state.mapping.viewport_origin.1, y);
                    }
                },
                record::SCALE_WINDOW_EXT => {
                    if state.mapping.scalable_extents() {
                        Self::scale_extents(rec, &mut state.mapping.window_extent);
                    }
                },
                record::SCALE_VIEWPORT_EXT => {
                    if state.mapping.scalable_extents() {
                        Self::scale_extents(rec, &mut state.mapping.viewport_extent);
                    }
                },

                record::CREATE_FONT_INDIRECT => Self::create_font(rec, &mut state),
                record::CREATE_PEN_INDIRECT
                | record::CREATE_BRUSH_INDIRECT
                | record::CREATE_PALETTE
                | record::CREATE_PATTERN_BRUSH
                | record::CREATE_DIB_PATTERN_BRUSH
                | record::CREATE_REGION => Self::insert_object(&mut state, ReplayObject::Other),
                record::SELECT_OBJECT => Self::select_object(rec, &mut state),
                record::DELETE_OBJECT => Self::delete_object(rec, &mut state),
                record::SET_TEXT_ALIGN => {
                    if let Some(align) = u16_at(rec, 0) {
                        state.text_align = align;
                    }
                },
                record::SET_TEXT_CHAR_EXTRA => {
                    if let Some(extra) = i16_at(rec, 0) {
                        state.text_char_extra = f64::from(extra);
                    }
                },

                record::MOVE_TO => {
                    if let Some(point) = point_at(rec, 0) {
                        state.current_position = point;
                    }
                },
                record::LINE_TO => {
                    if let Some(end) = point_at(rec, 0) {
                        Self::include_logical_point(
                            &mut bounds,
                            &state.mapping,
                            state.current_position,
                        );
                        Self::include_logical_point(&mut bounds, &state.mapping, end);
                        state.current_position = end;
                    }
                },
                record::SET_PIXEL => {
                    // ColorRef precedes the y/x pixel coordinates.
                    if let Some(point) = point_at(rec, 4) {
                        Self::include_logical_point(&mut bounds, &state.mapping, point);
                    }
                },
                record::FLOOD_FILL => {
                    // ColorRef precedes the y/x seed coordinates.
                    if let Some(point) = point_at(rec, 4) {
                        Self::include_logical_point(&mut bounds, &state.mapping, point);
                    }
                },
                record::EXT_FLOOD_FILL => {
                    // ColorRef precedes y/x; the flood mode follows the point.
                    if let Some(point) = point_at(rec, 4) {
                        Self::include_logical_point(&mut bounds, &state.mapping, point);
                    }
                },

                record::RECTANGLE | record::ELLIPSE => {
                    if let Some((left, top, right, bottom)) = rect_at(rec, 0) {
                        Self::include_logical_rect(
                            &mut bounds,
                            &state.mapping,
                            left,
                            top,
                            right,
                            bottom,
                        );
                    }
                },
                record::ROUND_RECT => {
                    if let Some((left, top, right, bottom)) = rect_at(rec, 4) {
                        Self::include_logical_rect(
                            &mut bounds,
                            &state.mapping,
                            left,
                            top,
                            right,
                            bottom,
                        );
                    }
                },
                record::ARC | record::PIE | record::CHORD => {
                    if let Some((left, top, right, bottom)) = rect_at(rec, 8) {
                        Self::include_logical_rect(
                            &mut bounds,
                            &state.mapping,
                            left,
                            top,
                            right,
                            bottom,
                        );
                    }
                    // Broken producers occasionally provide endpoints outside
                    // the enclosing rectangle; including them is conservative.
                    for offset in [0, 4] {
                        if let Some(point) = point_at(rec, offset) {
                            Self::include_logical_point(&mut bounds, &state.mapping, point);
                        }
                    }
                },
                record::POLYGON | record::POLYLINE => {
                    Self::include_points(rec, &state.mapping, &mut bounds, 2)
                },
                record::POLYPOLYGON => Self::include_polypolygon(rec, &state.mapping, &mut bounds),
                record::TEXT_OUT => Self::include_text_out(rec, &mut state, &mut bounds),
                record::EXT_TEXT_OUT => Self::include_ext_text_out(rec, &mut state, &mut bounds),

                record::PAT_BLT => {
                    Self::include_bitmap_destination(rec, &state.mapping, &mut bounds, 4, 6, 8, 10)
                },
                record::BIT_BLT | record::DIB_BIT_BLT => Self::include_bitmap_destination(
                    rec,
                    &state.mapping,
                    &mut bounds,
                    12,
                    14,
                    8,
                    10,
                ),
                record::STRETCH_BLT | record::DIB_STRETCH_BLT => Self::include_bitmap_destination(
                    rec,
                    &state.mapping,
                    &mut bounds,
                    16,
                    18,
                    12,
                    14,
                ),
                record::SET_DIB_TO_DEV => {
                    // ColorUsage, ScanCount, StartScan, yDib, xDib, Height,
                    // Width, yDest, xDest.
                    Self::include_bitmap_destination(
                        rec,
                        &state.mapping,
                        &mut bounds,
                        14,
                        16,
                        10,
                        12,
                    )
                },
                record::STRETCH_DIB => {
                    // ROP, usage, source dimensions/origin, destination
                    // height/width/origin (all coordinates are WORD ordered).
                    Self::include_bitmap_destination(
                        rec,
                        &state.mapping,
                        &mut bounds,
                        18,
                        20,
                        14,
                        16,
                    )
                },
                _ => {},
            }
        }

        bounds
            .unwrap_or_else(Bounds::fallback)
            .normalized_or_fallback()
    }

    fn restore_dc(state: &mut ReplayState, saves: &mut Vec<ReplaySnapshot>, rec: &WmfRecord) {
        let Some(level) = i16_at(rec, 0) else {
            return;
        };
        let index = if level < 0 {
            saves.len().checked_sub(level.unsigned_abs() as usize)
        } else if level > 0 {
            (level as usize)
                .checked_sub(1)
                .filter(|&index| index < saves.len())
        } else {
            None
        };
        if let Some(index) = index {
            saves[index].clone().restore_into(state);
            saves.truncate(index);
        }
    }

    fn scale_extents(rec: &WmfRecord, extent: &mut (f64, f64)) {
        // Win16 stack order is yDenom, yNum, xDenom, xNum.
        let (Some(y_den), Some(y_num), Some(x_den), Some(x_num)) = (
            i16_at(rec, 0),
            i16_at(rec, 2),
            i16_at(rec, 4),
            i16_at(rec, 6),
        ) else {
            return;
        };
        if x_den != 0 {
            extent.0 = bounded_product(extent.0, f64::from(x_num) / f64::from(x_den));
        }
        if y_den != 0 {
            extent.1 = bounded_product(extent.1, f64::from(y_num) / f64::from(y_den));
        }
        extent.0 = finite_or_one(extent.0);
        extent.1 = finite_or_one(extent.1);
    }

    fn create_font(rec: &WmfRecord, state: &mut ReplayState) {
        let Some(height) = i16_at(rec, 0) else {
            return;
        };
        let width = i16_at(rec, 2).unwrap_or(0);
        let metrics = FontMetrics {
            height: f64::from(height).abs().max(1.0),
            width: f64::from(width).abs(),
            escapement: f64::from(i16_at(rec, 4).unwrap_or(0)),
        };
        Self::insert_object(state, ReplayObject::Font(metrics));
    }

    fn insert_object(state: &mut ReplayState, object: ReplayObject) {
        if let Some(index) = state.objects.iter().position(Option::is_none) {
            state.objects[index] = Some(object);
        } else {
            state.objects.push(Some(object));
        }
    }

    fn select_object(rec: &WmfRecord, state: &mut ReplayState) {
        let Some(handle) = u16_at(rec, 0) else {
            return;
        };
        if handle & 0x8000 == 0 {
            if let Some(Some(ReplayObject::Font(font))) = state.objects.get(handle as usize) {
                state.font = *font;
            }
        }
    }

    fn delete_object(rec: &WmfRecord, state: &mut ReplayState) {
        if let Some(handle) = u16_at(rec, 0) {
            if let Some(object) = state.objects.get_mut(handle as usize) {
                *object = None;
            }
        }
    }

    fn include_points(
        rec: &WmfRecord,
        mapping: &MappingState,
        bounds: &mut Option<Bounds>,
        offset: usize,
    ) {
        let Some(count) = u16_at(rec, 0) else {
            return;
        };
        let count = usize::from(count).min(rec.params.len().saturating_sub(offset) / 4);
        for index in 0..count {
            if let Some(point) = array_point_at(rec, offset + index * 4) {
                Self::include_logical_point(bounds, mapping, point);
            }
        }
    }

    fn include_polypolygon(rec: &WmfRecord, mapping: &MappingState, bounds: &mut Option<Bounds>) {
        let Some(polygons) = u16_at(rec, 0) else {
            return;
        };
        let polygons = usize::from(polygons);
        let counts_len = polygons.saturating_mul(2);
        if polygons == 0 || rec.params.len() < 2 + counts_len {
            return;
        }
        let mut offset = 2 + counts_len;
        for index in 0..polygons {
            let Some(count) = u16_at(rec, 2 + index * 2) else {
                return;
            };
            let count = usize::from(count).min(rec.params.len().saturating_sub(offset) / 4);
            for point_index in 0..count {
                if let Some(point) = array_point_at(rec, offset + point_index * 4) {
                    Self::include_logical_point(bounds, mapping, point);
                }
            }
            offset = offset.saturating_add(count.saturating_mul(4));
        }
    }

    fn include_text_out(rec: &WmfRecord, state: &mut ReplayState, bounds: &mut Option<Bounds>) {
        let Some(count) = u16_at(rec, 0) else {
            return;
        };
        let count = usize::from(count);
        let text_end = 2usize.saturating_add(count);
        if text_end > rec.params.len() {
            return;
        }
        let offset = text_end.saturating_add(count & 1);
        if let Some(point) = point_at(rec, offset) {
            let anchor = if state.text_align & text_align::TA_UPDATECP != 0 {
                state.current_position
            } else {
                point
            };
            Self::include_text(state, bounds, anchor, count);
        }
    }

    fn include_ext_text_out(rec: &WmfRecord, state: &mut ReplayState, bounds: &mut Option<Bounds>) {
        let Some(point) = point_at(rec, 0) else {
            return;
        };
        let count = usize::from(u16_at(rec, 4).unwrap_or(0));
        let options = u16_at(rec, 6).unwrap_or(0);
        // The optional opaque/clipping rectangle itself paints output.
        if options & 0x0002 != 0 {
            // ExtTextOut embeds a Rect object in left, top, right, bottom
            // order rather than Win16's reversed call-parameter order.
            if let (Some(left), Some(top), Some(right), Some(bottom)) = (
                i16_at(rec, 8),
                i16_at(rec, 10),
                i16_at(rec, 12),
                i16_at(rec, 14),
            ) {
                Self::include_logical_rect(
                    bounds,
                    &state.mapping,
                    f64::from(left),
                    f64::from(top),
                    f64::from(right),
                    f64::from(bottom),
                );
            }
        }
        let text_offset: usize = if options & 0x0006 != 0 { 16 } else { 8 };
        if count != 0 && rec.params.len() >= text_offset.saturating_add(count) {
            let anchor = if state.text_align & text_align::TA_UPDATECP != 0 {
                state.current_position
            } else {
                point
            };
            Self::include_text(state, bounds, anchor, count);
        }
    }

    fn include_text(
        state: &mut ReplayState,
        bounds: &mut Option<Bounds>,
        anchor: (f64, f64),
        count: usize,
    ) {
        let height = state.font.height.max(1.0);
        let glyph_width = if state.font.width > 0.0 {
            state.font.width
        } else {
            height * 0.6
        };
        let width = bounded_add(
            glyph_width * count as f64,
            state.text_char_extra * count.saturating_sub(1) as f64,
        )
        .abs();

        let (mut left, mut right) = (anchor.0, bounded_add(anchor.0, width));
        match state.text_align & text_align::HORIZONTAL_MASK {
            text_align::TA_RIGHT => (left, right) = (bounded_add(anchor.0, -width), anchor.0),
            text_align::TA_CENTER => {
                let half = width / 2.0;
                (left, right) = (bounded_add(anchor.0, -half), bounded_add(anchor.0, half));
            },
            _ => {},
        }
        let (top, bottom) = match state.text_align & text_align::VERTICAL_MASK {
            text_align::TA_BOTTOM => (bounded_add(anchor.1, -height), anchor.1),
            text_align::TA_BASELINE => (
                bounded_add(anchor.1, -height * 0.8),
                bounded_add(anchor.1, height * 0.2),
            ),
            _ => (anchor.1, bounded_add(anchor.1, height)),
        };
        Self::include_logical_rect(bounds, &state.mapping, left, top, right, bottom);

        // The renderer applies escapement as a rotation about the text anchor.
        // A radius around that anchor safely covers every rotated rectangle,
        // including non-uniform logical-to-device mappings.
        if state.font.escapement != 0.0 {
            let (dx, dy) = state.mapping.vector(width, height);
            let radius = dx.hypot(dy).abs();
            let (x, y) = state.mapping.point_f64(anchor.0, anchor.1);
            Self::include_device_rect(bounds, x - radius, y - radius, x + radius, y + radius);
        }

        if state.text_align & text_align::TA_UPDATECP != 0 {
            state.current_position = (right, anchor.1);
        }
    }

    fn include_bitmap_destination(
        rec: &WmfRecord,
        mapping: &MappingState,
        bounds: &mut Option<Bounds>,
        y_offset: usize,
        x_offset: usize,
        height_offset: usize,
        width_offset: usize,
    ) {
        // Bitmap destination fields are ordered y, x, height, width.
        let (Some(y), Some(x), Some(height), Some(width)) = (
            i16_at(rec, y_offset),
            i16_at(rec, x_offset),
            i16_at(rec, height_offset),
            i16_at(rec, width_offset),
        ) else {
            return;
        };
        let left = f64::from(x);
        let top = f64::from(y);
        Self::include_logical_rect(
            bounds,
            mapping,
            left,
            top,
            bounded_add(left, f64::from(width)),
            bounded_add(top, f64::from(height)),
        );
    }

    fn include_logical_point(
        bounds: &mut Option<Bounds>,
        mapping: &MappingState,
        (x, y): (f64, f64),
    ) {
        let (x, y) = mapping.point_f64(x, y);
        Self::include_device_point(bounds, x, y);
    }

    fn include_logical_rect(
        bounds: &mut Option<Bounds>,
        mapping: &MappingState,
        left: f64,
        top: f64,
        right: f64,
        bottom: f64,
    ) {
        let (x1, y1) = mapping.point_f64(left, top);
        let (x2, y2) = mapping.point_f64(right, bottom);
        Self::include_device_rect(bounds, x1, y1, x2, y2);
    }

    fn include_device_point(bounds: &mut Option<Bounds>, x: f64, y: f64) {
        let Some((x, y)) = finite_point(x, y) else {
            return;
        };
        match bounds {
            Some(bounds) => bounds.include_point(x, y),
            None => {
                *bounds = Some(Bounds {
                    left: x,
                    top: y,
                    right: x,
                    bottom: y,
                })
            },
        }
    }

    fn include_device_rect(bounds: &mut Option<Bounds>, x1: f64, y1: f64, x2: f64, y2: f64) {
        let (Some((x1, y1)), Some((x2, y2))) = (finite_point(x1, y1), finite_point(x2, y2)) else {
            return;
        };
        let rect = Bounds {
            left: x1.min(x2),
            top: y1.min(y2),
            right: x1.max(x2),
            bottom: y1.max(y2),
        };
        match bounds {
            Some(bounds) => bounds.include_rect(rect.left, rect.top, rect.right, rect.bottom),
            None => *bounds = Some(rect),
        }
    }
}

impl Bounds {
    fn normalized_or_fallback(self) -> Self {
        if !(self.left.is_finite()
            && self.top.is_finite()
            && self.right.is_finite()
            && self.bottom.is_finite())
        {
            return Self::fallback();
        }
        Self {
            left: self.left.min(self.right),
            top: self.top.min(self.bottom),
            right: self.left.max(self.right),
            bottom: self.top.max(self.bottom),
        }
    }
}

#[inline]
fn i16_at(rec: &WmfRecord, offset: usize) -> Option<i16> {
    read_i16_le(&rec.params, offset).ok()
}

#[inline]
fn u16_at(rec: &WmfRecord, offset: usize) -> Option<u16> {
    read_u16_le(&rec.params, offset).ok()
}

/// WMF stores a point as y followed by x.
#[inline]
fn point_at(rec: &WmfRecord, offset: usize) -> Option<(f64, f64)> {
    Some((
        f64::from(i16_at(rec, offset + 2)?),
        f64::from(i16_at(rec, offset)?),
    ))
}

/// PointS objects embedded in polygon arrays are x followed by y.
#[inline]
fn array_point_at(rec: &WmfRecord, offset: usize) -> Option<(f64, f64)> {
    Some((
        f64::from(i16_at(rec, offset)?),
        f64::from(i16_at(rec, offset + 2)?),
    ))
}

/// WMF stores a rectangle as bottom, right, top, left.
#[inline]
fn rect_at(rec: &WmfRecord, offset: usize) -> Option<(f64, f64, f64, f64)> {
    let bottom = f64::from(i16_at(rec, offset)?);
    let right = f64::from(i16_at(rec, offset + 2)?);
    let top = f64::from(i16_at(rec, offset + 4)?);
    let left = f64::from(i16_at(rec, offset + 6)?);
    Some((left, top, right, bottom))
}

#[inline]
fn finite_point(x: f64, y: f64) -> Option<(f64, f64)> {
    (x.is_finite() && y.is_finite()).then_some((
        x.clamp(-MAX_COORDINATE, MAX_COORDINATE),
        y.clamp(-MAX_COORDINATE, MAX_COORDINATE),
    ))
}

#[inline]
fn finite_or_one(value: f64) -> f64 {
    if value.is_finite() {
        value.clamp(-MAX_COORDINATE, MAX_COORDINATE)
    } else {
        1.0
    }
}

#[inline]
fn bounded_add(left: f64, right: f64) -> f64 {
    let value = left + right;
    if value.is_finite() {
        value.clamp(-MAX_COORDINATE, MAX_COORDINATE)
    } else {
        left.signum().max(right.signum()) * MAX_COORDINATE
    }
}

#[inline]
fn bounded_product(left: f64, right: f64) -> f64 {
    let value = left * right;
    if value.is_finite() {
        value.clamp(-MAX_COORDINATE, MAX_COORDINATE)
    } else {
        left.signum() * right.signum() * MAX_COORDINATE
    }
}

#[cfg(test)]
mod tests {
    use super::super::super::constants::map_mode;
    use super::*;
    use bytes::Bytes;

    fn record(function: u16, words: &[i16]) -> WmfRecord {
        let mut params = Vec::with_capacity(words.len() * 2);
        for word in words {
            params.extend_from_slice(&word.to_le_bytes());
        }
        WmfRecord {
            size: 3 + words.len() as u32,
            function,
            params: Bytes::from(params),
        }
    }

    #[test]
    fn maps_reversed_rectangles_to_normalized_device_bounds() {
        let records = [
            record(record::SET_MAP_MODE, &[map_mode::MM_ANISOTROPIC as i16]),
            record(record::SET_WINDOW_EXT, &[100, 200]),
            record(record::SET_VIEWPORT_EXT, &[-200, 400]),
            // bottom, right, top, left (both axes deliberately reversed)
            record(record::RECTANGLE, &[10, 20, 90, 180]),
        ];
        let bounds = BoundsCalculator::scan_records(&records);
        assert_eq!(bounds.as_tuple(), (40.0, -180.0, 360.0, -20.0));
    }

    #[test]
    fn line_to_includes_the_saved_current_position() {
        let records = [
            record(record::MOVE_TO, &[20, 10]),
            record(record::LINE_TO, &[40, 30]),
        ];
        assert_eq!(
            BoundsCalculator::scan_records(&records).as_tuple(),
            (10.0, 20.0, 30.0, 40.0)
        );
    }

    #[test]
    fn polypolygon_and_arc_use_all_vector_extents() {
        let records = [
            record(
                record::POLYPOLYGON,
                &[2, 3, 3, 0, 0, 10, 10, 20, -5, 100, 100, 110, 110, 120, 90],
            ),
            record(record::ARC, &[7, 11, 3, 2, 20, 10, 30, -10]),
        ];
        assert_eq!(
            BoundsCalculator::scan_records(&records).as_tuple(),
            (-10.0, -5.0, 120.0, 110.0)
        );
    }

    #[test]
    fn text_and_opaque_ext_text_have_estimated_bounds() {
        let records = [
            record(record::CREATE_FONT_INDIRECT, &[-20, 10]),
            record(record::SELECT_OBJECT, &[0]),
            record(record::TEXT_OUT, &[3, 0x6261, 0x0063, 40, 50]),
            // y/x, len/options, then opaque bottom/right/top/left
            record(record::EXT_TEXT_OUT, &[5, 6, 1, 2, 50, 60, 10, 20, 0x0041]),
        ];
        assert_eq!(
            BoundsCalculator::scan_records(&records).as_tuple(),
            (6.0, 5.0, 80.0, 60.0)
        );
    }

    #[test]
    fn bitmap_destination_setpixel_and_floodfill_contribute() {
        let records = [
            // ROP (two words), y/x/height/width
            record(record::PAT_BLT, &[0, 0, 20, 10, 30, 40]),
            record(record::SET_PIXEL, &[0, 0, -2, -3]),
            record(record::FLOOD_FILL, &[0, 0, 70, 80]),
        ];
        assert_eq!(
            BoundsCalculator::scan_records(&records).as_tuple(),
            (-3.0, -2.0, 80.0, 70.0)
        );
    }

    #[test]
    fn bitmap_destination_layouts_use_their_own_field_offsets() {
        let records = [
            // ROP, ySrc/xSrc, height/width, yDst/xDst
            record(record::BIT_BLT, &[0, 0, 0, 0, 7, 11, 30, 20]),
            // ROP, source height/width/origin, destination height/width/origin
            record(record::STRETCH_BLT, &[0, 0, 1, 1, 0, 0, 8, 12, 40, 50]),
            // ColorUsage, scan range, DIB origin, source height/width, destination origin
            record(record::SET_DIB_TO_DEV, &[0, 0, 0, 0, 0, 9, 13, 60, 70]),
            // ROP, ColorUsage, source rectangle, destination height/width/origin
            record(record::STRETCH_DIB, &[0, 0, 0, 1, 1, 0, 0, 10, 14, 80, 90]),
        ];
        assert_eq!(
            BoundsCalculator::scan_records(&records).as_tuple(),
            (20.0, 30.0, 104.0, 90.0)
        );
    }

    #[test]
    fn extreme_coordinates_and_degenerate_mapping_stay_finite() {
        let records = [
            record(record::SET_MAP_MODE, &[map_mode::MM_ANISOTROPIC as i16]),
            record(record::SET_WINDOW_EXT, &[0, 0]),
            record(record::RECTANGLE, &[i16::MIN, i16::MIN, i16::MAX, i16::MAX]),
        ];
        let bounds = BoundsCalculator::scan_records(&records);
        assert_eq!(bounds.as_tuple(), (-32768.0, -32768.0, 32767.0, 32767.0));
        assert!(
            bounds.left.is_finite()
                && bounds.top.is_finite()
                && bounds.right.is_finite()
                && bounds.bottom.is_finite()
        );
    }

    #[test]
    fn save_restore_reinstates_mapping_before_later_draws() {
        let records = [
            record(record::SET_MAP_MODE, &[map_mode::MM_ANISOTROPIC as i16]),
            record(record::SET_VIEWPORT_EXT, &[2, 2]),
            record(record::SAVE_DC, &[]),
            record(record::SET_VIEWPORT_EXT, &[10, 10]),
            record(record::RESTORE_DC, &[-1]),
            record(record::SET_PIXEL, &[0, 0, 3, 4]),
        ];
        assert_eq!(
            BoundsCalculator::scan_records(&records).as_tuple(),
            (8.0, 6.0, 8.0, 6.0)
        );
    }
}
