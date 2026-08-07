//! WMF playback device-context and object-table state.

use std::{mem::size_of, sync::Arc};

use super::super::constants::{bk_mode, brush, fill_mode, layout, map_mode, pen, stock};

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct DeviceRect {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

impl DeviceRect {
    pub(super) fn new(left: f64, top: f64, right: f64, bottom: f64) -> Self {
        Self {
            left: left.min(right),
            top: top.min(bottom),
            right: left.max(right),
            bottom: top.max(bottom),
        }
    }

    pub(super) fn is_empty(self) -> bool {
        self.right <= self.left || self.bottom <= self.top
    }

    pub(super) fn intersect(self, other: Self) -> Option<Self> {
        let left = self.left.max(other.left);
        let top = self.top.max(other.top);
        let right = self.right.min(other.right);
        let bottom = self.bottom.min(other.bottom);
        (right > left && bottom > top).then_some(Self {
            left,
            top,
            right,
            bottom,
        })
    }

    pub(super) fn offset(&mut self, dx: f64, dy: f64) {
        self.left += dx;
        self.right += dx;
        self.top += dy;
        self.bottom += dy;
    }
}

/// A rectilinear clipping region. `None` means the complete output surface.
#[derive(Debug, Clone, Default)]
pub(super) struct ClipRegion {
    pub rects: Option<Vec<DeviceRect>>,
    pub excluded: Vec<DeviceRect>,
}

impl ClipRegion {
    pub(super) fn from_rects(mut rects: Vec<DeviceRect>) -> Self {
        rects.retain(|rect| !rect.is_empty());
        Self {
            rects: Some(rects),
            excluded: Vec::new(),
        }
    }

    pub(super) fn offset(&mut self, dx: f64, dy: f64) {
        if let Some(rects) = &mut self.rects {
            for rect in rects {
                rect.offset(dx, dy);
            }
        }
        for rect in &mut self.excluded {
            rect.offset(dx, dy);
        }
    }

    pub(super) fn is_unbounded(&self) -> bool {
        self.rects.is_none() && self.excluded.is_empty()
    }
}

/// Logical-to-device mapping in the playback DC.
#[derive(Debug, Clone, Copy)]
pub(super) struct MappingState {
    pub mode: u16,
    pub window_origin: (f64, f64),
    pub window_extent: (f64, f64),
    pub viewport_origin: (f64, f64),
    pub viewport_extent: (f64, f64),
    pub layout: u16,
}

impl Default for MappingState {
    fn default() -> Self {
        Self {
            mode: map_mode::MM_TEXT,
            window_origin: (0.0, 0.0),
            window_extent: (1.0, 1.0),
            viewport_origin: (0.0, 0.0),
            viewport_extent: (1.0, 1.0),
            layout: layout::LTR,
        }
    }
}

impl MappingState {
    pub(super) fn set_mode(&mut self, mode: u16) {
        if (map_mode::MM_TEXT..=map_mode::MM_ANISOTROPIC).contains(&mode) {
            self.mode = mode;
        }
    }

    /// Mapping scale in device units per logical unit.  The physical mapping
    /// modes use a conventional 96-DPI playback surface.
    pub(super) fn scale(&self) -> (f64, f64) {
        let (mut sx, sy) = match self.mode {
            map_mode::MM_LOMETRIC => (96.0 / 254.0, -96.0 / 254.0),
            map_mode::MM_HIMETRIC => (96.0 / 2540.0, -96.0 / 2540.0),
            map_mode::MM_LOENGLISH => (0.96, -0.96),
            map_mode::MM_HIENGLISH => (0.096, -0.096),
            map_mode::MM_TWIPS => (96.0 / 1440.0, -96.0 / 1440.0),
            map_mode::MM_ISOTROPIC | map_mode::MM_ANISOTROPIC => {
                let sx = safe_ratio(self.viewport_extent.0, self.window_extent.0);
                let sy = safe_ratio(self.viewport_extent.1, self.window_extent.1);
                if self.mode == map_mode::MM_ISOTROPIC {
                    let magnitude = sx.abs().min(sy.abs());
                    (magnitude.copysign(sx), magnitude.copysign(sy))
                } else {
                    (sx, sy)
                }
            },
            _ => (1.0, 1.0),
        };
        if self.layout & layout::RTL != 0 {
            sx = -sx;
        }
        (sx, sy)
    }

    pub(super) fn point_f64(&self, x: f64, y: f64) -> (f64, f64) {
        let (sx, sy) = self.scale();
        (
            self.viewport_origin.0 + (x - self.window_origin.0) * sx,
            self.viewport_origin.1 + (y - self.window_origin.1) * sy,
        )
    }

    pub(super) fn point(&self, x: i16, y: i16) -> (f64, f64) {
        self.point_f64(f64::from(x), f64::from(y))
    }

    pub(super) fn rect(&self, left: i16, top: i16, right: i16, bottom: i16) -> DeviceRect {
        let (x1, y1) = self.point(left, top);
        let (x2, y2) = self.point(right, bottom);
        DeviceRect::new(x1, y1, x2, y2)
    }

    pub(super) fn vector(&self, dx: f64, dy: f64) -> (f64, f64) {
        let (sx, sy) = self.scale();
        (dx * sx, dy * sy)
    }

    pub(super) fn scalable_extents(&self) -> bool {
        matches!(self.mode, map_mode::MM_ISOTROPIC | map_mode::MM_ANISOTROPIC)
    }
}

#[inline]
fn safe_ratio(numerator: f64, denominator: f64) -> f64 {
    if denominator == 0.0 || !numerator.is_finite() || !denominator.is_finite() {
        1.0
    } else {
        numerator / denominator
    }
}

#[derive(Debug, Clone)]
pub(super) struct Pen {
    pub style: u16,
    pub width: (i16, i16),
    pub color: u32,
}

impl Default for Pen {
    fn default() -> Self {
        Self {
            style: pen::PS_SOLID,
            width: (1, 0),
            color: 0,
        }
    }
}

#[derive(Debug, Clone)]
pub(super) struct Brush {
    pub style: u16,
    pub color: u32,
    pub hatch: u16,
    /// Opaque bitmap payload for pattern brushes. It is never interpreted by
    /// the SVG player itself.
    pub pattern: Option<Arc<[u8]>>,
}

impl Default for Brush {
    fn default() -> Self {
        Self {
            style: brush::BS_SOLID,
            color: 0x00ff_ffff,
            hatch: 0,
            pattern: None,
        }
    }
}

#[derive(Debug, Clone)]
pub(super) struct Font {
    pub height: i16,
    pub width: i16,
    pub escapement: i16,
    pub orientation: i16,
    pub weight: u16,
    pub italic: bool,
    pub underline: bool,
    pub strike_out: bool,
    pub charset: u8,
    pub name: String,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            height: -12,
            width: 0,
            escapement: 0,
            orientation: 0,
            weight: 400,
            italic: false,
            underline: false,
            strike_out: false,
            charset: 0,
            name: "serif".to_owned(),
        }
    }
}

#[derive(Debug, Clone, Default)]
pub(super) struct Palette {
    pub entries: Vec<u32>,
}

#[derive(Debug, Clone, Default)]
pub(super) struct Region {
    pub rects: Vec<DeviceRect>,
}

#[derive(Debug, Clone)]
pub(super) enum GdiObject {
    Pen(Pen),
    Brush(Brush),
    Font(Font),
    Palette(Palette),
    Region(Region),
}

#[derive(Debug, Clone, Default)]
pub(super) struct ObjectTable {
    objects: Vec<Option<GdiObject>>,
}

impl ObjectTable {
    /// WMF requires allocation at the lowest available object-table index.
    pub(super) fn insert(&mut self, object: GdiObject) -> usize {
        if let Some(index) = self.objects.iter().position(Option::is_none) {
            self.objects[index] = Some(object);
            index
        } else {
            self.objects.push(Some(object));
            self.objects.len() - 1
        }
    }

    pub(super) fn get(&self, index: usize) -> Option<&GdiObject> {
        self.objects.get(index)?.as_ref()
    }

    pub(super) fn get_mut(&mut self, index: usize) -> Option<&mut GdiObject> {
        self.objects.get_mut(index)?.as_mut()
    }

    pub(super) fn delete(&mut self, index: usize) -> bool {
        self.objects
            .get_mut(index)
            .is_some_and(|slot| slot.take().is_some())
    }

    pub(super) fn retained_len(&self) -> usize {
        self.objects
            .iter()
            .filter(|object| object.is_some())
            .count()
    }

    pub(super) fn retained_heap_bytes(&self) -> usize {
        self.objects.iter().fold(0usize, |total, object| {
            let bytes = match object {
                Some(GdiObject::Palette(palette)) => palette.entries.len().saturating_mul(4),
                Some(GdiObject::Region(region)) => {
                    region.rects.len().saturating_mul(size_of::<DeviceRect>())
                },
                Some(GdiObject::Font(font)) => font.name.len(),
                Some(GdiObject::Brush(brush)) => {
                    brush.pattern.as_ref().map_or(0, |pattern| pattern.len())
                },
                _ => 0,
            };
            total.saturating_add(bytes)
        })
    }
}

#[derive(Debug, Clone)]
pub(super) struct GraphicsState {
    pub position: (i16, i16),
    pub pen: Pen,
    pub brush: Brush,
    pub font: Font,
    pub palette_index: Option<usize>,
    pub text_color: u32,
    pub bk_color: u32,
    pub bk_mode: u16,
    pub poly_fill_mode: u16,
    pub text_align: u16,
    pub text_char_extra: i16,
    pub break_extra: i16,
    pub break_count: i16,
    pub rop2: u16,
    pub stretch_mode: u16,
    pub mapping: MappingState,
    pub clip: ClipRegion,
    pub clip_revision: u64,
}

impl Default for GraphicsState {
    fn default() -> Self {
        Self {
            position: (0, 0),
            pen: Pen::default(),
            brush: Brush::default(),
            font: Font::default(),
            palette_index: None,
            text_color: 0,
            bk_color: 0x00ff_ffff,
            bk_mode: bk_mode::OPAQUE,
            poly_fill_mode: fill_mode::ALTERNATE,
            text_align: 0,
            text_char_extra: 0,
            break_extra: 0,
            break_count: 0,
            rop2: 13, // R2_COPYPEN
            stretch_mode: 1,
            mapping: MappingState::default(),
            clip: ClipRegion::default(),
            clip_revision: 0,
        }
    }
}

impl GraphicsState {
    pub(super) fn retained_heap_bytes(&self) -> usize {
        let included = self.clip.rects.as_ref().map_or(0, Vec::len);
        included
            .saturating_add(self.clip.excluded.len())
            .saturating_mul(size_of::<DeviceRect>())
            .saturating_add(self.font.name.len())
    }
}

impl GraphicsState {
    pub(super) fn select_stock(&mut self, handle: u16) -> bool {
        let index = handle & !stock::FLAG;
        match index {
            stock::WHITE_BRUSH => self.brush = solid_brush(0x00ff_ffff),
            stock::LTGRAY_BRUSH => self.brush = solid_brush(0x00c0_c0c0),
            stock::GRAY_BRUSH => self.brush = solid_brush(0x0080_8080),
            stock::DKGRAY_BRUSH => self.brush = solid_brush(0x0040_4040),
            stock::BLACK_BRUSH => self.brush = solid_brush(0),
            stock::NULL_BRUSH => {
                self.brush = Brush {
                    style: brush::BS_NULL,
                    ..Brush::default()
                }
            },
            stock::WHITE_PEN => {
                self.pen = Pen {
                    color: 0x00ff_ffff,
                    ..Pen::default()
                }
            },
            stock::BLACK_PEN => self.pen = Pen::default(),
            stock::NULL_PEN => {
                self.pen = Pen {
                    style: pen::PS_NULL,
                    ..Pen::default()
                }
            },
            stock::OEM_FIXED_FONT | stock::ANSI_FIXED_FONT | stock::SYSTEM_FIXED_FONT => {
                self.font = Font {
                    name: "monospace".to_owned(),
                    ..Font::default()
                }
            },
            stock::ANSI_VAR_FONT
            | stock::SYSTEM_FONT
            | stock::DEVICE_DEFAULT_FONT
            | stock::DEFAULT_GUI_FONT => self.font = Font::default(),
            stock::DEFAULT_PALETTE => self.palette_index = None,
            _ => return false,
        }
        true
    }
}

fn solid_brush(color: u32) -> Brush {
    Brush {
        color,
        ..Brush::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn object_table_reuses_lowest_free_slot() {
        let mut table = ObjectTable::default();
        assert_eq!(table.insert(GdiObject::Pen(Pen::default())), 0);
        assert_eq!(table.insert(GdiObject::Brush(Brush::default())), 1);
        assert!(table.delete(0));
        assert_eq!(table.insert(GdiObject::Font(Font::default())), 0);
    }

    #[test]
    fn mapping_honors_origins_extents_and_rtl() {
        let mapping = MappingState {
            mode: map_mode::MM_ANISOTROPIC,
            window_origin: (10.0, 20.0),
            window_extent: (100.0, 200.0),
            viewport_origin: (5.0, 7.0),
            viewport_extent: (200.0, 100.0),
            ..MappingState::default()
        };
        assert_eq!(mapping.point(60, 120), (105.0, 57.0));
        let rtl = MappingState {
            layout: layout::RTL,
            ..mapping
        };
        assert_eq!(rtl.point(60, 120), (-95.0, 57.0));
    }

    #[test]
    fn zero_mapping_extent_never_produces_non_finite_coordinates() {
        let mapping = MappingState {
            mode: map_mode::MM_ANISOTROPIC,
            window_extent: (0.0, 0.0),
            ..MappingState::default()
        };
        let point = mapping.point(i16::MIN, i16::MAX);
        assert!(point.0.is_finite() && point.1.is_finite());
    }
}
