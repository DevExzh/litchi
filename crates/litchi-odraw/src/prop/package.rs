use crate::{Container, RecordKind, Result};
use std::collections::HashMap;

use super::geometry::Geometry;
use super::model::{Array, ColorRef, Id, Prop, Value};

/// An ordered `OfficeArt` shape-property collection.
#[derive(Debug)]
pub struct Props<'data> {
    pub(super) properties: Vec<Prop<'data>>,
    pub(super) by_id: HashMap<Id, usize>,
}

impl<'data> Props<'data> {
    /// Creates an empty property collection.
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self {
            properties: Vec::new(),
            by_id: HashMap::new(),
        }
    }

    /// Parses the primary Opt child when present.
    pub fn from_container(container: &Container<'data>) -> Result<Self> {
        match container.find(RecordKind::Opt)? {
            Some(opt) => Self::parse(&opt),
            None => Ok(Self::new()),
        }
    }

    /// Returns the complete lossless descriptor for `id`.
    #[inline]
    #[must_use]
    pub fn prop(&self, id: Id) -> Option<&Prop<'data>> {
        self.by_id
            .get(&id)
            .and_then(|index| self.properties.get(*index))
    }

    /// Returns the decoded value for `id`.
    #[inline]
    pub fn get(&self, id: Id) -> Option<&Value<'data>> {
        self.prop(id).map(Prop::value)
    }

    /// Returns a simple signed value.
    #[inline]
    #[must_use]
    pub fn get_int(&self, id: Id) -> Option<i32> {
        match self.get(id) {
            Some(Value::Simple(v)) => Some(*v),
            _ => None,
        }
    }

    /// Returns a typed, lossless color reference.
    #[inline]
    #[must_use]
    pub fn get_color(&self, id: Id) -> Option<ColorRef> {
        self.get_int(id)
            .map(|value| ColorRef::from_raw(value as u32))
    }

    /// Resolves an explicitly encoded boolean without applying defaults.
    #[inline]
    #[must_use]
    pub fn get_bool(&self, id: Id) -> Option<bool> {
        let raw_id = id.raw();
        if let Some(terminal_id) = boolean_group_terminal(raw_id) {
            let terminal = Id::from(terminal_id);
            if let Some(value) = self.get_int(terminal) {
                let bit = u32::from(terminal_id - raw_id);
                let value_mask = 1u32 << bit;
                let use_mask = value_mask << 16;
                let value = value as u32;
                return (value & use_mask != 0).then_some(value & value_mask != 0);
            }
        }
        self.get_int(id).map(|value| value != 0)
    }

    /// Returns property-specific complex bytes.
    #[inline]
    #[must_use]
    pub fn get_binary(&self, id: Id) -> Option<&'data [u8]> {
        match self.get(id) {
            Some(Value::Complex(data)) => Some(data),
            _ => None,
        }
    }

    /// Returns a validated array property.
    #[inline]
    #[must_use]
    pub fn get_array(&self, id: Id) -> Option<&Array<'data>> {
        match self.get(id) {
            Some(Value::Array(array)) => Some(array),
            _ => None,
        }
    }

    /// Returns whether `id` is present.
    #[inline]
    #[must_use]
    pub fn has(&self, id: Id) -> bool {
        self.by_id.contains_key(&id)
    }

    /// Returns the number of descriptors.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Returns whether the collection has no descriptors.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    /// Iterates over lossless descriptors in their original wire order.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Prop<'data>> {
        self.properties.iter()
    }

    /// Returns direct RGB only; indirect color references return `None`.
    #[inline]
    pub fn get_rgb(&self, id: Id) -> Option<(u8, u8, u8)> {
        self.get_color(id).and_then(ColorRef::rgb)
    }

    #[inline]
    #[must_use]
    pub fn get_rotation_degrees(&self, id: Id) -> Option<f32> {
        self.get_int(id)
            .map(|fixed_point| (fixed_point as f32) / 65536.0)
    }

    #[inline]
    #[must_use]
    pub fn get_opacity(&self, id: Id) -> Option<f32> {
        self.get_int(id).map(|fixed_point| {
            let opacity = (fixed_point as f32) / 65536.0;
            opacity.clamp(0.0, 1.0)
        })
    }

    #[inline]
    #[must_use]
    pub fn get_coord(&self, id: Id) -> Option<i32> {
        self.get_int(id)
    }

    /// Returns whether a boolean is explicitly enabled, treating absence as
    /// false. Use [`Self::is_filled`] and [`Self::has_line`] for properties
    /// whose `OfficeArt` defaults are true.
    #[inline]
    #[must_use]
    pub fn is_true(&self, id: Id) -> bool {
        self.get_bool(id).unwrap_or(false)
    }

    #[inline]
    #[must_use]
    pub fn get_line_width(&self) -> Option<i32> {
        self.get_int(Id::LineWidth)
    }

    #[inline]
    #[must_use]
    pub fn get_fill_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(Id::FillColor)
    }

    #[inline]
    #[must_use]
    pub fn get_line_color(&self) -> Option<(u8, u8, u8)> {
        self.get_rgb(Id::LineColor)
    }

    /// Resolves the fill-enabled bit, whose specification default is `true`.
    #[inline]
    #[must_use]
    pub fn is_filled(&self) -> bool {
        self.get_bool(Id::Filled).unwrap_or(true)
    }

    /// Resolves the line-enabled bit, whose specification default is `true`.
    #[inline]
    #[must_use]
    pub fn has_line(&self) -> bool {
        self.get_bool(Id::AnyLine).unwrap_or(true)
    }

    #[inline]
    #[must_use]
    pub fn has_shadow(&self) -> bool {
        self.is_true(Id::Shadow)
    }

    #[inline]
    #[must_use]
    pub fn get_geometry_rect(&self) -> Option<(i32, i32, i32, i32)> {
        let left = self.get_coord(Id::GeomLeft)?;
        let top = self.get_coord(Id::GeomTop)?;
        let right = self.get_coord(Id::GeomRight)?;
        let bottom = self.get_coord(Id::GeomBottom)?;
        Some((left, top, right, bottom))
    }

    #[inline]
    #[must_use]
    pub fn get_text_margins(&self) -> Option<(i32, i32, i32, i32)> {
        const HORIZONTAL_DEFAULT: i32 = 0x0001_6530;
        const VERTICAL_DEFAULT: i32 = 0x0000_B298;
        let left = self.get_int(Id::TextLeft).unwrap_or(HORIZONTAL_DEFAULT);
        let top = self.get_int(Id::TextTop).unwrap_or(VERTICAL_DEFAULT);
        let right = self.get_int(Id::TextRight).unwrap_or(HORIZONTAL_DEFAULT);
        let bottom = self.get_int(Id::TextBottom).unwrap_or(VERTICAL_DEFAULT);
        Some((left, top, right, bottom))
    }

    #[inline]
    #[must_use]
    pub fn get_adjust(&self, id: Id) -> Option<i32> {
        self.get_int(id)
    }

    /// Decodes the custom `pVertices`/`pSegmentInfo` geometry family when it
    /// is present, retaining the underlying arrays as borrowed wire views.
    pub fn geometry(&self) -> Result<Option<Geometry<'data>>> {
        super::geometry::parse(self)
    }

    /// Decodes the optional `[MS-ODRAW]` `fillShadeColors` gradient stops.
    ///
    /// The returned view borrows the original property array and retains
    /// indirect color references exactly; it never resolves colors or renders
    /// the gradient.
    pub fn gradient_stops(&self) -> Result<Option<super::gradient::Stops<'data>>> {
        super::gradient::parse(self)
    }

    /// Decodes the optional `[MS-ODRAW]` picture-name and BLIP-flag family.
    ///
    /// The returned projection borrows a valid UTF-16 name and retains the
    /// exact raw flags, including undefined producer bits.  Use
    /// [`super::picture::Snapshot`] when an immutable record snapshot and
    /// lossless edit transaction are required.
    pub fn picture(&self) -> Result<Option<super::picture::Metadata<'data>>> {
        super::picture::parse_properties(self)
    }
}
fn boolean_group_terminal(id: u16) -> Option<u16> {
    match id {
        0x0077..=0x007F => Some(0x007F),
        0x00BB..=0x00BF => Some(0x00BF),
        0x00FA..=0x00FF => Some(0x00FF),
        0x013C..=0x013F => Some(0x013F),
        0x017A..=0x017F => Some(0x017F),
        0x01BB..=0x01BF => Some(0x01BF),
        0x01FB..=0x01FF => Some(0x01FF),
        0x023E..=0x023F => Some(0x023F),
        0x027F => Some(0x027F),
        0x02BC..=0x02BF => Some(0x02BF),
        0x033A..=0x033F => Some(0x033F),
        _ => None,
    }
}

impl Default for Props<'_> {
    fn default() -> Self {
        Self::new()
    }
}
