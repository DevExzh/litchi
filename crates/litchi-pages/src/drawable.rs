//! Drawables anchored in the Pages body: their layer order, their frame,
//! property and playback transactions, and the archive-free geometry and
//! property values shared by images, movies, audio and shapes.

//!
//! [`DrawableSelector`] picks one image, movie, audio clip or shape by kind
//! and text-flow position; [`crate::Package::edit_body_drawable_properties`]
//! stages frame, property and playback changes through [`properties::Edit`].
//!
//! ```no_run
//! use litchi_pages::drawable::FlipAxis;
//! use litchi_pages::{DrawableSelector, Package};
//!
//! # fn edit(source: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::from_bytes(source)?;
//! let mut edit = package.edit_body_drawable_properties(DrawableSelector::image(0))?;
//! edit.flip(FlipAxis::Horizontal)?;
//! let commit = edit.commit()?;
//! let restored = commit
//!     .package()
//!     .apply_body_drawable_properties(&commit.patch().inverse())?;
//! # let _ = restored;
//! # Ok(())
//! # }
//! ```

pub use crate::drawable_order::{BodyDrawableHandle, BodyDrawableSelector, DrawableLayerMove};
pub use crate::selector::DrawableSelector;
pub use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
pub use litchi_iwa_common::shape::geometry::{FlipAxis, Point, Size};

/// Exact-source frame, property and playback transactions for body
/// images, movies, audio clips and shapes.
pub mod properties {
    pub use crate::package::body_drawable_properties::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Snapshot,
    };
    pub use litchi_iwa_common::shape::line::Segment;
}

/// Exact-source lifecycle transactions (removal, duplication) for body images, movies
/// and audio clips.
pub mod lifecycle {
    pub use crate::package::body_media_lifecycle::{
        Commit, Diagnostics, Edit, Error, LimitKind, Media, Patch, Refusal,
    };
}

/// Exact-source title and caption transactions for body images, movies and
/// shapes.
pub mod caption {
    pub use crate::package::body_drawable_caption::{
        Commit, Diagnostics, Edit, Error, Kind, LimitKind, Patch, Refusal,
    };
}

/// Drawing styles of body shapes and text boxes: fill, stroke, shadow,
/// effects, text layout, text columns and line endpoints, over the shared shape-style variation engine (record 0371).
pub mod style {
    pub use crate::package::body_drawable_style::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::shape::effects::Effects;
    pub use litchi_iwa_common::shape::fill::Fill;
    pub use litchi_iwa_common::shape::line::{Endpoint, Endpoints};
    pub use litchi_iwa_common::shape::shadow::Shadow;
    pub use litchi_iwa_common::shape::stroke::Stroke;
    pub use litchi_iwa_common::text::layout::Layout;
    pub use litchi_iwa_text::columns::Columns;

    /// One drawing property of a body shape.
    #[derive(Debug, Clone, PartialEq)]
    #[non_exhaustive]
    pub enum Property {
        /// The fill (none, a color, a gradient; image fills read back but are
        /// not written by this family).
        Fill(Fill),
        /// The stroke (`None`: the hidden stroke).
        Stroke(Option<Stroke>),
        /// The shadow.
        Shadow(Shadow),
        /// Opacity and reflection.
        Effects(Effects),
        /// Vertical alignment, insets and autosizing of the text frame.
        TextLayout(Layout),
        /// The text frame's columns (count, widths and gaps).
        TextColumns(Columns),
        /// The decorations at a straight line's directed start and end.
        LineEndpoints(Endpoints),
    }

    /// Which [`Property`] a read or reset addresses.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    #[non_exhaustive]
    pub enum PropertyKind {
        /// The fill.
        Fill,
        /// The stroke.
        Stroke,
        /// The shadow.
        Shadow,
        /// Opacity and reflection.
        Effects,
        /// The text-frame layout.
        TextLayout,
        /// The text-frame columns.
        TextColumns,
        /// A straight line's endpoint decorations.
        LineEndpoints,
    }

    impl Property {
        /// The kind of this property.
        #[must_use]
        pub const fn kind(&self) -> PropertyKind {
            match self {
                Self::Fill(_) => PropertyKind::Fill,
                Self::Stroke(_) => PropertyKind::Stroke,
                Self::Shadow(_) => PropertyKind::Shadow,
                Self::Effects(_) => PropertyKind::Effects,
                Self::TextLayout(_) => PropertyKind::TextLayout,
                Self::TextColumns(_) => PropertyKind::TextColumns,
                Self::LineEndpoints(_) => PropertyKind::LineEndpoints,
            }
        }
    }
}

/// Exact-source plain-text transactions for the storage a body drawable
/// owns (ordinary shapes).
pub mod text {
    pub use crate::package::body_drawable_text::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Character and paragraph properties of the whole text a body text box
/// or shape owns, over the shared paragraph-style engine (record 0413).
pub mod text_box_text {
    pub use crate::package::text_box_text::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::{appearance, character, font, paragraph};
    pub use litchi_iwa_text_wire::paragraph_style::{Property, PropertyKind};
}

/// Paragraph lists, list levels, numbering and per-level list properties
/// of the text a body text box or shape owns, over the shared list engine
/// and storage-list publisher (record 0413).
pub mod list {
    pub use crate::package::text_box_list::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::TextPosition;
    pub use litchi_iwa_text::paragraph::list::{
        ParagraphList, ParagraphListBullet, ParagraphListBulletBaselineOffset,
        ParagraphListBulletGeometry, ParagraphListBulletScale, ParagraphListIndentation,
        ParagraphListLabelColor, ParagraphListLabelIndent, ParagraphListLevel,
        ParagraphListLevelPlacement, ParagraphListNumberFormat, ParagraphListNumberPunctuation,
        ParagraphListNumberScale, ParagraphListNumberSequence, ParagraphListNumberTiering,
        ParagraphListNumbering, ParagraphListPlacement, ParagraphListStart, ParagraphListTextGap,
    };
    pub use litchi_iwa_text_wire::list::{Property, PropertyKind};
}

/// Named paragraph styles of the text a body text box or shape owns:
/// apply, create, rename, delete and redefine, over the shared named-style
/// engine (record 0413).
pub mod named_style {
    pub use crate::package::text_box_named_style::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::paragraph::style::{
        NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName,
    };
}

/// Plain-text Drop Caps of the paragraphs of the text a body text box or
/// shape owns (record 0413).
pub mod drop_cap {
    pub use crate::package::text_box_drop_cap::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::TextPosition;
    pub use litchi_iwa_text::paragraph::drop_cap::{
        CharacterCount, DropCap, LineCount, Outdent, Padding, RaisedLines, Wrap,
    };
}

/// Highlights, text comments and replies of the text a body text box or
/// shape owns (record 0413).
pub mod annotation {
    pub use crate::package::text_box_annotation::{
        Comment, Commit, Diagnostics, Edit, Error, Highlight, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::position::{TextPosition, TextRange};
}

/// Date & Time fields of the text a body text box or shape owns (record
/// 0413).
pub mod date_time {
    pub use crate::package::text_box_date_time::{
        Commit, DateTimeField, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::date_time::{
        DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan,
    };
    pub use litchi_iwa_text::position::{TextPosition, TextRange};
}

/// Hyperlinks of the text a body text box or shape owns (record 0413).
pub mod hyperlink {
    pub use crate::package::text_box_hyperlink::{
        Commit, Diagnostics, Edit, Error, Hyperlink, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::hyperlink::TextHyperlinkTarget;
    pub use litchi_iwa_text::position::{TextPosition, TextRange};
}

/// Language runs of the text a body text box or shape owns (record 0413).
pub mod language {
    pub use crate::package::text_box_language::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::TextPosition;
    pub use litchi_iwa_text::language::{TextLanguage, TextLanguageRun};
}

/// The comment attached directly to a body drawable and its replies
/// (record 0413).
pub mod comment {
    pub use crate::package::text_box_drawable_comment::{
        Commit, Diagnostics, DrawableComment, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Exact-source transactions for the body drawables' layer order.
pub mod order {
    pub use crate::package::drawable_order::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Native layout flags stored beside a drawable's frame.
///
/// Pages keeps a small bit set with every frame. Bit 2 is the reflection
/// (horizontal mirror) bit that the Arrange Flip commands toggle; the other
/// bits are layout state Pages writes and are kept as they are.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct GeometryFlags(u32);

impl GeometryFlags {
    /// The reflection bit toggled by both Arrange Flip commands.
    pub const REFLECTED: u32 = 1 << 2;

    /// Wrap the stored flag bits unchanged.
    #[must_use]
    pub const fn from_bits(bits: u32) -> Self {
        Self(bits)
    }

    /// The stored flag bits.
    #[must_use]
    pub const fn bits(self) -> u32 {
        self.0
    }

    /// Whether the drawable is mirrored.
    #[must_use]
    pub const fn is_reflected(self) -> bool {
        self.0 & Self::REFLECTED != 0
    }

    /// These flags with the reflection bit toggled.
    #[must_use]
    pub const fn toggle_reflection(self) -> Self {
        Self(self.0 ^ Self::REFLECTED)
    }
}

/// The frame of one drawable in document points.
///
/// Each part is optional because Pages may omit it; reading a drawable keeps
/// the stored presence and writing a value with `None` removes that part.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Geometry {
    /// Top-left position.
    pub position: Option<Point>,
    /// Displayed size.
    pub size: Option<Size>,
    /// Native layout flags, including the reflection bit.
    pub flags: Option<GeometryFlags>,
    /// Clockwise rotation in degrees.
    pub angle: Option<f32>,
}

impl Geometry {
    /// Whether every present coordinate, dimension and angle is finite and
    /// every dimension is non-negative.
    #[must_use]
    pub fn is_valid(&self) -> bool {
        self.position
            .is_none_or(|position| position.x.is_finite() && position.y.is_finite())
            && self.size.is_none_or(|size| {
                size.width.is_finite()
                    && size.height.is_finite()
                    && size.width >= 0.0
                    && size.height >= 0.0
            })
            && self.angle.is_none_or(f32::is_finite)
    }

    /// The frame after one native Arrange Flip command.
    ///
    /// Both directions toggle the reflection flag; a vertical flip also turns
    /// the drawable by half a turn. `None` when the frame has no flags.
    #[must_use]
    pub fn flipped(self, axis: FlipAxis) -> Option<Self> {
        let flags = self.flags?.toggle_reflection();
        let angle = match axis {
            FlipAxis::Horizontal => self.angle,
            FlipAxis::Vertical => {
                let current = self.angle.unwrap_or_default().rem_euclid(360.0);
                let turned = (current + 180.0).rem_euclid(360.0);
                Some(if turned == 0.0 { 0.0 } else { turned })
            },
        };
        Some(Self {
            flags: Some(flags),
            angle,
            ..self
        })
    }
}

/// The user-facing properties every drawable carries.
///
/// `None` means the property is absent; `Some(false)` keeps an explicit
/// native false.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Properties {
    /// Hyperlink target URL.
    pub hyperlink: Option<String>,
    /// Whether the drawable is locked against editing.
    pub locked: Option<bool>,
    /// Whether resizing keeps the aspect ratio.
    pub aspect_ratio_locked: Option<bool>,
    /// Accessibility description read by assistive technology.
    pub accessibility_description: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn flips_toggle_reflection_and_vertical_turns_half() {
        let frame = Geometry {
            position: Some(Point { x: 1.0, y: 2.0 }),
            size: Some(Size {
                width: 3.0,
                height: 4.0,
            }),
            flags: Some(GeometryFlags::from_bits(3)),
            angle: Some(45.0),
        };
        let horizontal = frame.flipped(FlipAxis::Horizontal).unwrap();
        assert_eq!(horizontal.flags, Some(GeometryFlags::from_bits(7)));
        assert_eq!(horizontal.angle, Some(45.0));
        assert!(horizontal.flags.unwrap().is_reflected());
        let vertical = frame.flipped(FlipAxis::Vertical).unwrap();
        assert_eq!(vertical.angle, Some(225.0));
        let back = Geometry {
            angle: Some(180.0),
            ..frame
        }
        .flipped(FlipAxis::Vertical)
        .unwrap();
        assert_eq!(back.angle.map(f32::to_bits), Some(0.0_f32.to_bits()));
        assert_eq!(
            Geometry {
                flags: None,
                ..frame
            }
            .flipped(FlipAxis::Horizontal),
            None
        );
    }

    #[test]
    fn validity_rejects_non_finite_and_negative_parts() {
        assert!(Geometry::default().is_valid());
        assert!(
            !Geometry {
                angle: Some(f32::NAN),
                ..Geometry::default()
            }
            .is_valid()
        );
        assert!(
            !Geometry {
                size: Some(Size {
                    width: -1.0,
                    height: 1.0
                }),
                ..Geometry::default()
            }
            .is_valid()
        );
    }
}
