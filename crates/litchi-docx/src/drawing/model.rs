//! Owned semantic objects projected from a Word drawing inventory.

use litchi_core::unit::{emu_to_pt_f64, emu_to_px_96};
use litchi_drawingml::geom::Preset;

/// The host placement of a Word drawing object.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum Anchor {
    /// The object is carried by a `wp:inline` anchor.
    Inline,
    /// The object is carried by a floating `wp:anchor` anchor.
    Floating,
}

/// The checked Word 2010 identifier carried by a DrawingML anchor.
///
/// Word stores this value as eight hexadecimal ASCII digits.  The zero and
/// high-bit ranges are reserved by `[MS-DOCX]` and `[MS-ODRAWXML]`, so the
/// constructor keeps those values out of the semantic model.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct AnchorId(u32);

impl AnchorId {
    /// Construct an anchor identifier in the schema-defined range.
    #[inline]
    pub const fn new(value: u32) -> Option<Self> {
        if value != 0 && value < 0x8000_0000 {
            Some(Self(value))
        } else {
            None
        }
    }

    /// Return the numeric identifier.
    #[inline]
    pub const fn get(self) -> u32 {
        self.0
    }
}

/// The semantic family discovered for a drawing object.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum Kind {
    /// A WordprocessingML shape (`wps:wsp`).
    Shape,
    /// A shape containing a Word text-box story (`wps:txbx`).
    TextBox,
    /// An unknown or non-shape DrawingML object retained as inert inventory.
    Other,
}

/// An owned drawing object discovered inside a Word paragraph.
///
/// The object is an inventory projection, not a renderer. Unknown children,
/// unsupported DrawingML features, and external resources are deliberately
/// ignored by the inventory parser while the discovered object remains in
/// document order.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Object {
    /// Shape title from `wp:docPr@name`.
    name: String,
    /// Alternative text from `wp:docPr@descr`.
    description: String,
    /// Text collected from the nested Word text-box story, if present.
    text: String,
    /// Width in EMUs.
    width_emu: i64,
    /// Height in EMUs.
    height_emu: i64,
    /// Horizontal DrawingML offset in EMUs.
    x_emu: i64,
    /// Vertical DrawingML offset in EMUs.
    y_emu: i64,
    /// Closed DrawingML preset geometry, when a `prstGeom` was present.
    preset: Option<Preset>,
    /// Discovered semantic family.
    kind: Kind,
    /// Discovered WordprocessingML placement.
    anchor: Anchor,
    /// Checked Word 2010 anchor identifier, when authored.
    anchor_id: Option<AnchorId>,
}

impl Object {
    /// Create a shape object with inline placement and no text-box story.
    #[inline]
    pub fn new(
        name: String,
        description: String,
        width_emu: i64,
        height_emu: i64,
        preset: Preset,
    ) -> Self {
        Self::from_inventory(
            name,
            description,
            width_emu,
            height_emu,
            0,
            0,
            Some(preset),
            Kind::Shape,
            Anchor::Inline,
            None,
            String::new(),
        )
    }

    /// Build an owned object from the streaming inventory decoder.
    pub(crate) fn from_inventory(
        name: String,
        description: String,
        width_emu: i64,
        height_emu: i64,
        x_emu: i64,
        y_emu: i64,
        preset: Option<Preset>,
        kind: Kind,
        anchor: Anchor,
        anchor_id: Option<AnchorId>,
        text: String,
    ) -> Self {
        Self {
            name,
            description,
            text,
            width_emu,
            height_emu,
            x_emu,
            y_emu,
            preset,
            kind,
            anchor,
            anchor_id,
        }
    }

    /// Return the shape title.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the alternative text.
    #[inline]
    pub fn description(&self) -> &str {
        &self.description
    }

    /// Return the text collected from the nested text-box story.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Replace the owned text-box story text.
    #[inline]
    pub fn set_text(&mut self, text: String) {
        self.text = text;
    }

    /// Return the width in EMUs.
    #[inline]
    pub fn width_emu(&self) -> i64 {
        self.width_emu
    }

    /// Return the height in EMUs.
    #[inline]
    pub fn height_emu(&self) -> i64 {
        self.height_emu
    }

    /// Return the horizontal offset in EMUs.
    #[inline]
    pub fn x_emu(&self) -> i64 {
        self.x_emu
    }

    /// Return the vertical offset in EMUs.
    #[inline]
    pub fn y_emu(&self) -> i64 {
        self.y_emu
    }

    /// Set the DrawingML offset in EMUs.
    #[inline]
    pub fn set_position(&mut self, x_emu: i64, y_emu: i64) {
        self.x_emu = x_emu;
        self.y_emu = y_emu;
    }

    /// Return the width in pixels at 96 DPI.
    #[inline]
    pub fn width_px(&self) -> u32 {
        emu_to_px_96(self.width_emu)
    }

    /// Return the height in pixels at 96 DPI.
    #[inline]
    pub fn height_px(&self) -> u32 {
        emu_to_px_96(self.height_emu)
    }

    /// Return the width in points.
    #[inline]
    pub fn width_pt(&self) -> f64 {
        emu_to_pt_f64(self.width_emu)
    }

    /// Return the height in points.
    #[inline]
    pub fn height_pt(&self) -> f64 {
        emu_to_pt_f64(self.height_emu)
    }

    /// Return the closed preset geometry, when present.
    #[inline]
    pub fn preset(&self) -> Option<Preset> {
        self.preset
    }

    /// Return the discovered semantic family.
    #[inline]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the discovered WordprocessingML placement.
    #[inline]
    pub fn anchor(&self) -> Anchor {
        self.anchor
    }

    /// Return the checked Word 2010 anchor identifier.
    #[inline]
    pub fn anchor_id(&self) -> Option<AnchorId> {
        self.anchor_id
    }

    /// Return whether the object contains a text-box story.
    #[inline]
    pub fn is_text_box(&self) -> bool {
        self.kind == Kind::TextBox
    }

    /// Set whether the object contains a text-box story.
    #[inline]
    pub fn set_text_box(&mut self, is_text_box: bool) {
        if is_text_box {
            self.kind = Kind::TextBox;
        } else if self.kind == Kind::TextBox {
            self.kind = Kind::Shape;
        }
    }

    /// Return whether the object is inline rather than floating.
    #[inline]
    pub fn is_inline(&self) -> bool {
        self.anchor == Anchor::Inline
    }

    /// Set whether the object is inline rather than floating.
    #[inline]
    pub fn set_inline(&mut self, is_inline: bool) {
        self.anchor = if is_inline {
            Anchor::Inline
        } else {
            Anchor::Floating
        };
    }
}
