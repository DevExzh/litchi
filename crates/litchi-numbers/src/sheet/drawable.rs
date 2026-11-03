//! The drawables a Numbers sheet owns: text boxes, shapes, images, movies,
//! audio clips, charts and tables, their selector, their frame and property
//! values, and the transaction families that edit them.
//!
//! A sheet lists its drawables in one back-to-front list
//! (`TN.SheetArchive.drawable_infos`); [`crate::Package::sheet_drawable_list`]
//! reads it. A [`DrawableSelector`](crate::sheet::drawable::DrawableSelector)
//! picks one drawable by its position in that list or by its position among
//! the drawables of one [`Kind`](crate::sheet::drawable::Kind). Native
//! object identifiers, component names and protobuf payloads stay private.
//!
//! ```no_run
//! use litchi_numbers::sheet::drawable::{DrawableSelector, FlipAxis};
//! use litchi_numbers::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::open("input.numbers")?;
//! for drawable in package.sheet_drawable_list(0)?.iter() {
//!     println!("{:?} at {:?}", drawable.kind(), drawable.selector());
//! }
//! let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::image(0))?;
//! edit.flip(FlipAxis::Horizontal)?;
//! let commit = edit.commit()?;
//! let restored = commit
//!     .package()
//!     .apply_sheet_drawable_properties(&commit.patch().inverse())?;
//! # let _ = restored;
//! # Ok(())
//! # }
//! ```

use litchi_core::Position;

use crate::selector::ChartSelector;
use crate::shape::image::ImageSelector;

pub use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
pub use litchi_iwa_common::shape::geometry::{FlipAxis, Point, Size};

pub use crate::package::sheet_drawable::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
};

/// Exact-source shape and text-frame style transactions for one sheet shape
/// or text box: fill, stroke, shadow, effects, frame layout, text columns
/// and a straight line's endpoint decorations.
pub mod style {
    pub use crate::package::sheet_drawable_style::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::shape::effects::Effects;
    pub use litchi_iwa_common::shape::fill::Fill;
    pub use litchi_iwa_common::shape::line::{Endpoint, Endpoints};
    pub use litchi_iwa_common::shape::shadow::Shadow;
    pub use litchi_iwa_common::shape::stroke::Stroke;
    pub use litchi_iwa_common::text::layout::Layout;
    pub use litchi_iwa_text::columns::Columns;

    /// One semantic drawing or text-frame style value.
    #[derive(Debug, Clone, PartialEq)]
    #[non_exhaustive]
    pub enum Property {
        /// The fill (image fills are read-only in this family).
        Fill(Fill),
        /// The stroke (`None`: the hidden stroke).
        Stroke(Option<Stroke>),
        /// Opacity and reflection.
        Effects(Effects),
        /// Vertical alignment, insets and autosizing of the text frame.
        TextLayout(Layout),
        /// Uniform columns in a text box.
        Columns(Columns),
        /// The shadow of a shape.
        Shadow(Shadow),
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
        /// Opacity and reflection.
        Effects,
        /// The text-frame layout.
        TextLayout,
        /// The text-frame columns.
        Columns,
        /// The shadow.
        Shadow,
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
                Self::Effects(_) => PropertyKind::Effects,
                Self::TextLayout(_) => PropertyKind::TextLayout,
                Self::Columns(_) => PropertyKind::Columns,
                Self::Shadow(_) => PropertyKind::Shadow,
                Self::LineEndpoints(_) => PropertyKind::LineEndpoints,
            }
        }
    }
}

/// Exact-source frame, property and playback transactions for one sheet
/// drawable: [`crate::Package::edit_sheet_drawable_properties`].
pub mod properties {
    pub use crate::package::sheet_drawable_properties::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Snapshot,
    };
    pub use litchi_iwa_common::shape::line::Segment;
    pub use litchi_iwa_common::shape::path::Preset;
}

/// Exact-source title and caption transactions of a sheet shape, image,
/// movie or chart (caption only), on the shared caption graph engine:
/// [`crate::Package::edit_sheet_drawable_caption`].
pub mod caption {
    pub use crate::package::sheet_drawable_caption::{
        Commit, Diagnostics, Edit, Error, Label, LimitKind, Patch, Refusal,
    };
}

/// Exact-source Drop Cap transactions for the paragraphs of the text a
/// sheet text box or shape owns: [`crate::Package::edit_sheet_drawable_drop_caps`].
pub mod drop_cap {
    pub use crate::package::sheet_drawable_drop_cap::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::TextPosition;
    pub use litchi_iwa_text::paragraph::drop_cap::{
        CharacterCount, DropCap, LineCount, Outdent, Padding, RaisedLines, Wrap,
    };
}

/// Exact-source named paragraph-style transactions for the text a sheet
/// text box or shape owns: [`crate::Package::edit_sheet_text_box_named_styles`].
pub mod named_style {
    pub use crate::package::sheet_drawable_named_style::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::paragraph::style::{
        NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName,
    };
}

/// Exact-source transactions for the comment attached directly to a sheet
/// drawable and its replies: [`crate::Package::edit_sheet_drawable_comment`].
pub mod comment {
    pub use crate::package::sheet_drawable_comment::{
        Commit, Diagnostics, DrawableComment, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// The embedded media assets of the document and its sheets, and
/// exact-source replacement of their bytes: [`crate::Package::edit_media`].
pub mod media {
    pub use crate::package::sheet_media::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::media::Type;
    pub use litchi_iwa_common::media::embedded::{AssetId, Embedded};
}

/// Exact-source highlight, comment and reply transactions for the text a
/// sheet text box or shape owns: [`crate::Package::edit_sheet_drawable_annotations`].
pub mod annotation {
    pub use crate::package::sheet_drawable_annotation::{
        Comment, Commit, Diagnostics, Edit, Error, Highlight, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::position::{TextPosition, TextRange};
}

/// Exact-source Date & Time field transactions for the text a sheet text
/// box or shape owns: [`crate::Package::edit_sheet_drawable_date_time_fields`].
pub mod date_time {
    pub use crate::package::sheet_drawable_date_time::{
        Commit, DateTimeField, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::date_time::{
        DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan,
    };
    pub use litchi_iwa_text::position::{TextPosition, TextRange};
}

/// Exact-source hyperlink transactions for the text a sheet text box or
/// shape owns: [`crate::Package::edit_sheet_drawable_hyperlinks`].
pub mod hyperlink {
    pub use crate::package::sheet_drawable_hyperlink::{
        Commit, Diagnostics, Edit, Error, Hyperlink, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::hyperlink::TextHyperlinkTarget;
    pub use litchi_iwa_text::position::{TextPosition, TextRange};
}

/// Exact-source paragraph-list transactions for the text a sheet text box
/// or shape owns: [`crate::Package::edit_sheet_text_box_lists`].
pub mod list {
    pub use crate::package::sheet_drawable_list::{
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

/// Exact-source text-language transactions for the storage a sheet text
/// box or shape owns: [`crate::Package::edit_sheet_drawable_text_languages`].
pub mod language {
    pub use crate::package::sheet_drawable_language::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::TextPosition;
    pub use litchi_iwa_text::language::{TextLanguage, TextLanguageRun};
}

/// Exact-source plain-text transactions for the storage a sheet text box or
/// shape owns: [`crate::Package::edit_sheet_drawable_plain_text`].
pub mod plain_text {
    pub use crate::package::sheet_drawable_plain_text::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::{TextPosition, TextSpan};
}

/// Exact-source character and paragraph property transactions for the text
/// a sheet text box or shape owns: [`crate::Package::edit_sheet_text_box_text`].
pub mod text {
    pub use crate::package::sheet_drawable_text::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_text::{appearance, character, font, paragraph};
    pub use litchi_iwa_text_wire::paragraph_style::{Property, PropertyKind};
}

/// What a sheet drawable is.
///
/// ```
/// use litchi_numbers::sheet::drawable::Kind;
///
/// assert!(Kind::TextBox.has_text());
/// assert!(!Kind::Chart.has_text());
/// ```
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A free text box (`TSWP.ShapeInfoArchive` flagged as a text box).
    TextBox,
    /// An ordinary shape or line that may hold text.
    Shape,
    /// A still image.
    Image,
    /// A movie.
    Movie,
    /// An audio-only movie.
    Audio,
    /// A chart.
    Chart,
    /// A table.
    Table,
    /// A drawable of another native kind (a group, a live video, a shape
    /// without the text-box flag).
    Other,
}

impl Kind {
    /// Whether drawables of this kind own a text storage.
    #[must_use]
    pub const fn has_text(self) -> bool {
        matches!(self, Self::TextBox | Self::Shape)
    }

    /// A static lower-case word naming the kind.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::TextBox => "text box",
            Self::Shape => "shape",
            Self::Image => "image",
            Self::Movie => "movie",
            Self::Audio => "audio",
            Self::Chart => "chart",
            Self::Table => "table",
            Self::Other => "drawable",
        }
    }
}

/// Selects one drawable of a sheet.
///
/// [`Self::Position`] counts every drawable in the sheet's back-to-front
/// list (the order [`crate::Package::sheet_drawable_list`] reports);
/// [`Self::Of`] counts only the drawables of one kind in that order, so
/// `DrawableSelector::image(0)` is the sheet's back-most image, the same
/// image [`ImageSelector::index`]`(0)` selects.
///
/// ```
/// use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};
///
/// let selector = DrawableSelector::text_box(1);
/// assert_eq!(selector.kind(), Some(Kind::TextBox));
/// assert_eq!(selector.as_index(), 1);
/// assert_eq!(DrawableSelector::from(3).kind(), None);
/// ```
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DrawableSelector {
    /// The drawable at this zero-based position of the sheet's list.
    Position(Position),
    /// The drawable at this zero-based position among the sheet's drawables
    /// of `kind`.
    Of {
        /// The selected kind.
        kind: Kind,
        /// The position among drawables of that kind.
        position: Position,
    },
}

impl DrawableSelector {
    /// The drawable at a zero-based position of the sheet's list.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// The drawable at zero-based `index` of the sheet's list.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Position(Position::new(index))
    }

    /// The `index`-th drawable of `kind`.
    #[must_use]
    pub const fn of(kind: Kind, index: usize) -> Self {
        Self::Of {
            kind,
            position: Position::new(index),
        }
    }

    /// The `index`-th text box.
    #[must_use]
    pub const fn text_box(index: usize) -> Self {
        Self::of(Kind::TextBox, index)
    }

    /// The `index`-th ordinary shape.
    #[must_use]
    pub const fn shape(index: usize) -> Self {
        Self::of(Kind::Shape, index)
    }

    /// The `index`-th image.
    #[must_use]
    pub const fn image(index: usize) -> Self {
        Self::of(Kind::Image, index)
    }

    /// The `index`-th movie.
    #[must_use]
    pub const fn movie(index: usize) -> Self {
        Self::of(Kind::Movie, index)
    }

    /// The `index`-th audio clip.
    #[must_use]
    pub const fn audio(index: usize) -> Self {
        Self::of(Kind::Audio, index)
    }

    /// The `index`-th chart.
    #[must_use]
    pub const fn chart(index: usize) -> Self {
        Self::of(Kind::Chart, index)
    }

    /// The selected kind, when the selector names one.
    #[must_use]
    pub const fn kind(self) -> Option<Kind> {
        match self {
            Self::Position(_) => None,
            Self::Of { kind, .. } => Some(kind),
        }
    }

    /// The selected zero-based position (in the list, or among the kind).
    #[must_use]
    pub const fn as_position(self) -> Position {
        match self {
            Self::Position(position) | Self::Of { position, .. } => position,
        }
    }

    /// The selected zero-based index (in the list, or among the kind).
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.as_position().get()
    }
}

impl From<usize> for DrawableSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for DrawableSelector {
    fn from(position: Position) -> Self {
        Self::Position(position)
    }
}

impl From<ImageSelector> for DrawableSelector {
    fn from(selector: ImageSelector) -> Self {
        Self::Of {
            kind: Kind::Image,
            position: selector.as_position(),
        }
    }
}

impl From<ChartSelector> for DrawableSelector {
    fn from(selector: ChartSelector) -> Self {
        Self::Of {
            kind: Kind::Chart,
            position: selector.as_position(),
        }
    }
}

/// Native layout flags stored beside a drawable's frame.
///
/// Bit 2 is the reflection (horizontal mirror) bit that the Arrange Flip
/// commands toggle; the other bits are layout state Numbers writes and are
/// kept as they are.
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

/// The placement of a new sheet image (Numbers' Insert ▸ Image).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageOptions {
    /// Top-left position on the sheet, in points.
    pub position: Point,
    /// Displayed size, in points.
    pub size: Size,
    /// Untransformed media dimensions, in points.
    pub natural_size: Size,
}

impl ImageOptions {
    /// Options whose displayed and natural sizes are equal.
    #[must_use]
    pub const fn new(position: Point, size: Size) -> Self {
        Self {
            position,
            size,
            natural_size: size,
        }
    }

    /// Set the media dimensions independently of the displayed size.
    #[must_use]
    pub const fn with_natural_size(mut self, natural_size: Size) -> Self {
        self.natural_size = natural_size;
        self
    }
}

/// The placement and timing of a new sheet movie.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MovieOptions {
    /// Top-left position on the sheet, in points.
    pub position: Point,
    /// Displayed size, in points.
    pub size: Size,
    /// Untransformed media dimensions, in points.
    pub natural_size: Size,
    /// Playable duration of the movie.
    pub duration: std::time::Duration,
}

impl MovieOptions {
    /// Options whose displayed and natural sizes are equal.
    #[must_use]
    pub const fn new(position: Point, size: Size, duration: std::time::Duration) -> Self {
        Self {
            position,
            size,
            natural_size: size,
            duration,
        }
    }

    /// Set the media dimensions independently of the displayed size.
    #[must_use]
    pub const fn with_natural_size(mut self, natural_size: Size) -> Self {
        self.natural_size = natural_size;
        self
    }
}

/// The placement and timing of a new sheet audio clip (a zero-size
/// control).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AudioOptions {
    /// Position of the audio control, in points.
    pub position: Point,
    /// Playable duration of the clip.
    pub duration: std::time::Duration,
}

impl AudioOptions {
    /// Options for a clip at `position` lasting `duration`.
    #[must_use]
    pub const fn new(position: Point, duration: std::time::Duration) -> Self {
        Self { position, duration }
    }
}

/// The frame of one drawable in sheet points (`TSD.GeometryArchive`).
///
/// Each part is optional because Numbers may omit it; reading a drawable
/// keeps the stored presence and writing a value with `None` removes that
/// part.
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
    /// A frame at `position` with `size`, the default layout flags of a new
    /// Numbers drawable and no rotation.
    #[must_use]
    pub const fn new(position: Point, size: Size) -> Self {
        Self {
            position: Some(position),
            size: Some(size),
            flags: Some(GeometryFlags::from_bits(3)),
            angle: Some(0.0),
        }
    }

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

/// One drawable of a sheet's back-to-front list.
#[derive(Debug, Clone, PartialEq)]
pub struct Drawable {
    pub(crate) position: Position,
    pub(crate) kind: Kind,
    pub(crate) kind_position: Position,
    pub(crate) geometry: Option<Geometry>,
}

impl Drawable {
    /// The selector of this drawable by its list position.
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        DrawableSelector::Position(self.position)
    }

    /// The selector of this drawable by its position among its kind.
    #[must_use]
    pub const fn kind_selector(&self) -> DrawableSelector {
        DrawableSelector::Of {
            kind: self.kind,
            position: self.kind_position,
        }
    }

    /// What the drawable is.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Its zero-based layer, from the back (`0`) to the front.
    #[must_use]
    pub const fn z_position(&self) -> Position {
        self.position
    }

    /// The stored frame; `None` for a table (whose frame is its table info)
    /// or a drawable without one.
    #[must_use]
    pub const fn geometry(&self) -> Option<Geometry> {
        self.geometry
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn flips_toggle_reflection_and_vertical_turns_half() {
        let frame = Geometry {
            angle: Some(45.0),
            ..Geometry::new(
                Point { x: 1.0, y: 2.0 },
                Size {
                    width: 3.0,
                    height: 4.0,
                },
            )
        };
        let horizontal = frame.flipped(FlipAxis::Horizontal).unwrap();
        assert_eq!(horizontal.flags, Some(GeometryFlags::from_bits(7)));
        assert_eq!(horizontal.angle, Some(45.0));
        assert!(horizontal.flags.unwrap().is_reflected());
        let vertical = frame.flipped(FlipAxis::Vertical).unwrap();
        assert_eq!(vertical.angle, Some(225.0));
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
    fn selectors_convert_from_the_image_and_chart_selectors() {
        assert_eq!(
            DrawableSelector::from(ImageSelector::index(2)),
            DrawableSelector::image(2)
        );
        assert_eq!(
            DrawableSelector::from(ChartSelector::index(1)),
            DrawableSelector::chart(1)
        );
        assert_eq!(DrawableSelector::index(4).as_index(), 4);
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
    }
}
