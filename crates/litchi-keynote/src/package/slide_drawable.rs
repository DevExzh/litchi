//! Exact-source, selector-first edits of the drawables a Keynote slide owns
//! (record 0298): inventory, geometry, removal, z-order, text-box, shape and
//! image creation.
//!
//! A drawable is addressed by its [`DrawableSelector`]: its position in the
//! slide's owned-drawable list (`KN.SlideArchive` field 7), the same source
//! order the drawable-comment family uses. An [`Edit`] stages any number of
//! operations against one immutable package; selectors resolve against that
//! base snapshot and the transaction remembers native identity privately, so
//! a later operation is never shifted by an earlier one. [`Edit::commit`]
//! applies the operations in order in one structural workspace, publishes one
//! candidate, reopens and validates it, and reads every touched slide back.

mod author;
mod inventory;
mod stage;
mod table;
mod text;

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use thiserror::Error as ThisError;

use self::inventory::{ReadFailure, ReadLimits, SlideDrawables};
use super::structure::workspace::Workspace;
use super::structure::{StructureError, StructureLimit};
use super::{Package, PhysicalSource, ReadError};
use crate::{DrawableSelector, SlideSelector, SlideSelectorError};

pub use self::text::{Style, TextRun, TextSpan};
use litchi_iwa_common::shape::path::{InnerRadiusRatio, PolygonSides, StarPoints};

/// What a slide drawable is.
///
/// ```
/// use litchi_keynote::slide::drawable::Kind;
///
/// assert!(Kind::TextBox.has_text());
/// assert!(!Kind::Image.has_text());
/// ```
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A free text box.
    TextBox,
    /// A preset or custom shape (it may hold text).
    Shape,
    /// A straight or connection line.
    Line,
    /// A still image.
    Image,
    /// A movie.
    Movie,
    /// An audio-only movie.
    Audio,
    /// The slide's title placeholder.
    TitlePlaceholder,
    /// The slide's body placeholder.
    BodyPlaceholder,
    /// The slide's slide-number placeholder.
    SlideNumberPlaceholder,
    /// The slide's object (media) placeholder.
    ObjectPlaceholder,
    /// A placeholder of an unnamed native kind.
    Placeholder,
    /// A table.
    Table,
    /// A chart.
    Chart,
    /// A group of drawables.
    Group,
    /// A drawable of another native kind.
    Other,
}

impl Kind {
    /// Whether drawables of this kind own a text storage.
    #[must_use]
    pub const fn has_text(self) -> bool {
        matches!(
            self,
            Self::TextBox
                | Self::Shape
                | Self::TitlePlaceholder
                | Self::BodyPlaceholder
                | Self::SlideNumberPlaceholder
                | Self::ObjectPlaceholder
                | Self::Placeholder
        )
    }

    /// Whether this kind is a placeholder.
    #[must_use]
    pub const fn is_placeholder(self) -> bool {
        matches!(
            self,
            Self::TitlePlaceholder
                | Self::BodyPlaceholder
                | Self::SlideNumberPlaceholder
                | Self::ObjectPlaceholder
                | Self::Placeholder
        )
    }

    const fn geometry_editable(self) -> bool {
        matches!(self, Self::TextBox | Self::Shape | Self::Image) || self.is_placeholder()
    }

    const fn removable(self) -> bool {
        matches!(self, Self::TextBox | Self::Shape | Self::Image | Self::Line)
    }
}

/// A drawable's frame in slide points (`TSD.GeometryArchive`).
///
/// `rotation` is the native angle in degrees; read back, `None` means the
/// source does not record one, and passed to [`Edit::set_geometry`] it keeps
/// the current rotation.
///
/// ```
/// use litchi_keynote::slide::drawable::Geometry;
///
/// let frame = Geometry::new(72.0, 144.0, 400.0, 120.0).with_rotation(15.0);
/// assert_eq!(frame.rotation, Some(15.0));
/// ```
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Geometry {
    /// Left edge in points.
    pub x: f32,
    /// Top edge in points.
    pub y: f32,
    /// Width in points.
    pub width: f32,
    /// Height in points.
    pub height: f32,
    /// Rotation in degrees, when recorded.
    pub rotation: Option<f32>,
}

impl Geometry {
    /// A frame without rotation.
    #[must_use]
    pub const fn new(x: f32, y: f32, width: f32, height: f32) -> Self {
        Self {
            x,
            y,
            width,
            height,
            rotation: None,
        }
    }

    /// This frame with a rotation in degrees.
    #[must_use]
    pub const fn with_rotation(mut self, degrees: f32) -> Self {
        self.rotation = Some(degrees);
        self
    }

    pub(in crate::package) fn is_valid(self) -> bool {
        self.x.is_finite()
            && self.y.is_finite()
            && self.width.is_finite()
            && self.height.is_finite()
            && self.width > 0.0
            && self.height > 0.0
            && self.rotation.is_none_or(f32::is_finite)
    }
}

/// A preset path for [`Edit::add_shape`], after the legacy host's
/// `shape_path_source`.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Preset {
    /// A four-corner rectangle.
    Rectangle,
    /// A rounded rectangle with a corner radius in points (at most half the
    /// shorter side).
    RoundedRectangle {
        /// Corner radius in points.
        radius: f32,
    },
    /// An ellipse inscribed in the frame.
    Ellipse,
    /// Keynote's left-facing single arrow (native head and shaft
    /// proportions).
    LeftArrow,
    /// Keynote's right-facing single arrow.
    RightArrow,
    /// Keynote's double-headed arrow.
    DoubleArrow,
    /// A regular polygon (Keynote's Pentagon has five sides).
    RegularPolygon {
        /// The side count.
        sides: PolygonSides,
    },
    /// A star (Keynote's default has five points).
    Star {
        /// The point count.
        points: StarPoints,
        /// The inner radius as a ratio of the outer radius.
        inner_radius: InnerRadiusRatio,
    },
}

/// The raster formats [`Edit::add_image`] accepts (sniffed from the bytes).
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageFormat {
    /// PNG.
    Png,
    /// JPEG.
    Jpeg,
    /// GIF.
    Gif,
}

/// One drawable owned by a slide, as read from the package.
#[derive(Debug, Clone, PartialEq)]
pub struct Drawable {
    selector: DrawableSelector,
    kind: Kind,
    geometry: Option<Geometry>,
    text: Option<String>,
    name: Option<String>,
    locked: bool,
    z_position: Option<Position>,
}

impl Drawable {
    /// The selector that addresses this drawable (its source position).
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        self.selector
    }

    /// The drawable's kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The drawable's frame, when it records a position and size.
    #[must_use]
    pub const fn geometry(&self) -> Option<Geometry> {
        self.geometry
    }

    /// The plain text of the drawable's own text storage, if it has one.
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    /// The user-defined shape name Keynote stores on a shape path, if any.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Whether the drawable is locked against editing in Keynote.
    #[must_use]
    pub const fn is_locked(&self) -> bool {
        self.locked
    }

    /// The drawable's position in the slide's back-to-front z-order
    /// (0 is the back-most), when the slide lists it there.
    #[must_use]
    pub const fn z_position(&self) -> Option<Position> {
        self.z_position
    }
}

/// Why a drawable operation is refused.
///
/// The reasons name content classes only.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Unsupported {
    /// The operation does not support drawables of this kind (for example
    /// removing a table, chart, group, movie or placeholder, or setting the
    /// geometry of a table, chart, group, line or movie).
    Kind {
        /// The selected drawable's kind.
        kind: Kind,
    },
    /// The drawable is referenced from outside its private object graph
    /// (builds, connection lines, other slide fields or other components).
    ExternalReferences,
    /// The drawable is stored outside its slide's package component.
    ForeignComponent,
    /// The theme or stylesheet has no style preset for the new drawable.
    MissingThemeStyle,
    /// The image bytes are not PNG, JPEG or GIF.
    ImageFormat,
    /// A character style that sets nothing ([`Edit::set_text_style`]).
    TextStyling,
    /// The table holds formulas, or a formula or chart elsewhere reads it,
    /// so a value change would need a recalculation
    /// ([`Edit::set_table_cell`]).
    TableFormulas,
    /// The table's cell storage is outside the shared cell writer (a legacy
    /// tile, a missing tile, a rich-text, formula or commented cell).
    TableStorage,
    /// A native message without a reference schema holds references.
    UnmodelledContent,
    /// A package-registry record uses an unmodelled shape.
    RegistryShape,
    /// The text change would consume or introduce an inline object,
    /// footnote or attachment marker (`U+FFFC`, `U+000E`), which the plain
    /// text write cannot keep consistent.
    InlineObjects,
    /// The deck holds no slide table whose structure a new table can copy
    /// ([`Edit::add_table`]).
    TableTemplate,
}

impl fmt::Display for Unsupported {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Kind { kind } => write!(f, "a {kind:?} drawable"),
            Self::ExternalReferences => f.write_str("references from outside the drawable"),
            Self::ForeignComponent => f.write_str("a drawable outside its slide component"),
            Self::MissingThemeStyle => f.write_str("a theme without the needed style preset"),
            Self::ImageFormat => f.write_str("an image format other than PNG, JPEG or GIF"),
            Self::TextStyling => f.write_str("a character style that sets nothing"),
            Self::TableFormulas => f.write_str("a table in a formula graph"),
            Self::TableStorage => f.write_str("this table cell storage"),
            Self::UnmodelledContent => f.write_str("unmodelled native content with references"),
            Self::RegistryShape => f.write_str("an unmodelled package-registry shape"),
            Self::InlineObjects => f.write_str("text with inline object markers"),
            Self::TableTemplate => f.write_str("a deck without a slide table to copy"),
        }
    }
}

/// A finite resource governed by a slide-drawable transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Drawables owned by one slide.
    Drawables,
    /// Native objects in the package.
    Objects,
    /// The native object or data identifier space.
    Identifiers,
    /// Text bytes read from one slide.
    TextBytes,
    /// Bytes of one native payload.
    WireBytes,
    /// Fields of one native payload.
    WireFields,
    /// Nesting depth of one native payload.
    WireNesting,
    /// Rewrite work.
    Work,
    /// Bytes of one rewritten payload or member.
    OutputBytes,
    /// Bytes of one package entry, object or the package.
    EntryBytes,
    /// Slides in the presentation.
    Slides,
    /// Object references read during verification.
    References,
    /// Rich-text storages read during verification.
    TextStorages,
    /// Rich-text fragments read during verification.
    TextFragments,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::Drawables => "drawables",
            Self::Objects => "objects",
            Self::Identifiers => "identifiers",
            Self::TextBytes => "text bytes",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::Work => "work",
            Self::OutputBytes => "output bytes",
            Self::EntryBytes => "entry bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
        })
    }
}

/// A slide-drawable failure that never exposes native identifiers,
/// component names or authored content.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The package cannot be edited safely (malformed or inconsistent).
    #[error("the Keynote source cannot be edited safely")]
    InvalidSource,
    /// A resource limit was exceeded.
    #[error("Keynote slide-drawable {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exceeded resource.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// An allocation failed.
    #[error("could not allocate {amount} units for the Keynote slide-drawable transaction")]
    Allocation {
        /// The requested amount.
        amount: usize,
    },
    /// The published candidate did not verify.
    #[error("the Keynote slide-drawable candidate failed semantic verification")]
    Verification,
    /// The patch does not match the exact source package.
    #[error("the Keynote slide-drawable patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package was not opened from an exact physical source.
    #[error("this Keynote source does not support structural edits")]
    UnsupportedSource,
    /// The operation is refused for the selected content.
    #[error("the Keynote drawable operation is not supported for {reason}")]
    Unsupported {
        /// The class of content that blocks the operation.
        reason: Unsupported,
    },
    /// Several slides share the selected navigator name.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// The slide selector names an empty navigator name.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide has the selected navigator name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// No slide is at the selected position.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// The requested position.
        position: Position,
    },
    /// The slide owns no drawable at the selected position.
    #[error("the Keynote slide has no drawable at position {position:?}")]
    DrawablePositionNotFound {
        /// The requested position.
        position: Position,
    },
    /// An earlier operation of this edit removed the selected drawable.
    #[error("the selected Keynote drawable was removed earlier in this edit")]
    DrawableRemoved,
    /// The z-order destination is past the back-to-front list.
    #[error("z-order position {position:?} is outside 0..{count}")]
    ZPositionOutOfRange {
        /// The requested position.
        position: Position,
        /// The z-order length when the operation runs.
        count: usize,
    },
    /// The geometry is not finite or its size is not positive.
    #[error("the Keynote drawable geometry must be finite with a positive size")]
    InvalidGeometry,
    /// A new table's size is empty or past the native row and column
    /// identity ceiling, or a seeded cell lies outside it.
    #[error("a new Keynote table of {rows} x {columns} cells is not supported")]
    InvalidTableSize {
        /// The requested rows.
        rows: u32,
        /// The requested columns.
        columns: u32,
    },
    /// The image file name has no usable stem.
    #[error("the Keynote image needs a file name with a stem")]
    InvalidImage,
    /// The text span is empty, reversed or past the end of the text.
    #[error("the Keynote text span is outside the drawable text")]
    InvalidSpan,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. } | Self::Unsupported { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::AmbiguousSelector { .. }
            | Self::EmptySlideName { .. }
            | Self::ZPositionOutOfRange { .. }
            | Self::InvalidGeometry { .. }
            | Self::InvalidTableSize { .. }
            | Self::InvalidImage { .. }
            | Self::InvalidSpan { .. } => RefusalKind::InvalidValue,
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::DrawablePositionNotFound { .. } => RefusalKind::NotFound,
            Self::DrawableRemoved { .. } => RefusalKind::UnsupportedTarget,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One staged operation with privately remembered native identity.
#[derive(Clone)]
enum Operation {
    Geometry {
        slide: u64,
        drawable: u64,
        message_index: usize,
        message_type: u32,
        geometry: Geometry,
    },
    Remove {
        slide: u64,
        drawable: u64,
    },
    Move {
        slide: u64,
        drawable: u64,
        position: usize,
    },
    AddText {
        slide: u64,
        slide_position: Position,
        preset: Option<Preset>,
        geometry: Geometry,
        text: Box<str>,
    },
    AddImage {
        slide: u64,
        slide_position: Position,
        bytes: Arc<[u8]>,
        file_name: Box<str>,
        geometry: Geometry,
    },
    Text {
        slide: u64,
        drawable: u64,
        storage: u64,
        text: Box<str>,
    },
    TextStyle {
        slide: u64,
        drawable: u64,
        storage: u64,
        span: TextSpan,
        style: Style,
    },
    TableCell {
        slide: u64,
        drawable: u64,
        row: u32,
        column: u32,
        value: TableCellValue,
    },
    AddTable {
        slide: u64,
        slide_position: Position,
        template: u64,
        name: Box<str>,
        table: NewTable,
    },
    NotesStyle {
        slide: u64,
        note: u64,
        storage: u64,
        span: TextSpan,
        style: Style,
    },
}

/// A new slide table for [`Edit::add_table`]: its size, frame, optional
/// name and seeded cell values.
///
/// ```
/// use litchi_keynote::slide::drawable::{Geometry, NewTable, TableCellValue};
///
/// let table = NewTable::new(3, 2, Geometry::new(80.0, 120.0, 600.0, 180.0))
///     .with_name("Forecast")
///     .with_cell(0, 0, TableCellValue::Text("Region".into()));
/// assert_eq!((table.rows(), table.columns()), (3, 2));
/// assert_eq!(table.name(), Some("Forecast"));
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct NewTable {
    rows: u32,
    columns: u32,
    geometry: Geometry,
    name: Option<String>,
    cells: Vec<(u32, u32, TableCellValue)>,
}

impl NewTable {
    /// An empty table of `rows` × `columns` cells (header rows and columns
    /// included, as the template's counts allow) in `geometry`; rows and
    /// columns divide the frame evenly.
    #[must_use]
    pub const fn new(rows: u32, columns: u32, geometry: Geometry) -> Self {
        Self {
            rows,
            columns,
            geometry,
            name: None,
            cells: Vec::new(),
        }
    }

    /// This table with a name (the default is `Table N`, after the deck's
    /// tables).
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// This table with a value in one cell (zero-based); a later value of
    /// the same cell replaces an earlier one.
    #[must_use]
    pub fn with_cell(mut self, row: u32, column: u32, value: TableCellValue) -> Self {
        self.cells.retain(|(r, c, _)| (*r, *c) != (row, column));
        self.cells.push((row, column, value));
        self
    }

    /// The number of rows.
    #[must_use]
    pub const fn rows(&self) -> u32 {
        self.rows
    }

    /// The number of columns.
    #[must_use]
    pub const fn columns(&self) -> u32 {
        self.columns
    }

    /// The frame.
    #[must_use]
    pub const fn geometry(&self) -> Geometry {
        self.geometry
    }

    /// The requested name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// The seeded cells, in staging order.
    #[must_use]
    pub fn cells(&self) -> &[(u32, u32, TableCellValue)] {
        &self.cells
    }
}

/// A value [`Edit::set_table_cell`] stores in a table cell.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum TableCellValue {
    /// Plain text.
    Text(String),
    /// A finite number.
    Number(f64),
    /// A boolean.
    Boolean(bool),
    /// No value (the cell's style and format stay).
    Empty,
}

/// A drawable created by a commit: its slide and its new selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Created {
    slide: Position,
    drawable: DrawableSelector,
}

impl Created {
    /// The slide that owns the new drawable.
    #[must_use]
    pub const fn slide(self) -> Position {
        self.slide
    }

    /// The new drawable's selector on that slide.
    #[must_use]
    pub const fn drawable(self) -> DrawableSelector {
        self.drawable
    }
}

/// One bounded slide-drawable edit staged against an immutable package.
///
/// ```no_run
/// use litchi_keynote::slide::drawable::{Geometry, Preset};
/// use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let package = Package::open("deck.key")?;
/// let slide = SlideSelector::index(0);
/// let mut edit = package.edit_slide_drawables();
/// edit.set_geometry(slide, DrawableSelector::index(0), Geometry::new(40.0, 40.0, 300.0, 80.0))?
///     .add_shape(slide, Preset::Ellipse, Geometry::new(500.0, 300.0, 200.0, 200.0), "Hi")?
///     .move_to(slide, DrawableSelector::index(0), Position::new(0))?;
/// let commit = edit.commit()?;
/// assert_eq!(commit.created().len(), 1);
/// # Ok(())
/// # }
/// ```
pub struct Edit<'a> {
    source: &'a Package,
    operations: Vec<Operation>,
    inventories: BTreeMap<u64, SlideDrawables>,
    removed: BTreeSet<u64>,
    /// Net drawables added (positive) or removed per slide so far.
    deltas: BTreeMap<u64, isize>,
    /// The table template and the deck's table count, once read.
    tables: Option<(Option<u64>, usize)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

/// A selected slide of the base snapshot.
struct SelectedSlide {
    position: Position,
    identifier: u64,
}

impl<'a> Edit<'a> {
    const fn new(source: &'a Package) -> Self {
        Self {
            source,
            operations: Vec::new(),
            inventories: BTreeMap::new(),
            removed: BTreeSet::new(),
            deltas: BTreeMap::new(),
            tables: None,
        }
    }

    fn slide(&mut self, selector: SlideSelector<'_>) -> Result<SelectedSlide, Error> {
        let selected = select_slide(self.source, selector)?;
        if !self.inventories.contains_key(&selected.identifier) {
            let inventory =
                read_inventory(self.source, selected.identifier, read_limits(self.source)?)?;
            self.inventories.insert(selected.identifier, inventory);
        }
        Ok(selected)
    }

    fn drawable(
        &mut self,
        slide: SlideSelector<'_>,
        drawable: DrawableSelector,
    ) -> Result<(SelectedSlide, inventory::Item), Error> {
        let selected = self.slide(slide)?;
        let item = self
            .inventories
            .get(&selected.identifier)
            .and_then(|inventory| inventory.items.get(drawable.as_index()))
            .cloned()
            .ok_or(Error::Refused(Refusal::DrawablePositionNotFound {
                position: drawable.as_position(),
            }))?;
        if self.removed.contains(&item.identifier) {
            return Err(Error::Refused(Refusal::DrawableRemoved));
        }
        Ok((selected, item))
    }

    fn check_drawable_limit(&self, slide: u64) -> Result<(), Error> {
        let base = self
            .inventories
            .get(&slide)
            .map_or(0, |inventory| inventory.items.len());
        let observed = base.saturating_add_signed(self.deltas.get(&slide).copied().unwrap_or(0));
        let maximum = self.source.semantic_limits().max_references();
        if observed >= maximum {
            return Err(Error::Limit {
                kind: LimitKind::Drawables,
                observed: observed as u64 + 1,
                maximum: maximum as u64,
            });
        }
        Ok(())
    }

    /// Stage a new frame for one drawable.
    ///
    /// Text boxes, shapes, images and placeholders are supported; the shape
    /// path is left unchanged (Keynote scales it to the frame), as the host's
    /// `set_shape_geometry` does. A geometry without a rotation keeps the
    /// drawable's current rotation. Setting the current frame stages nothing.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::InvalidGeometry`], and
    /// [`Refusal::Unsupported`] for other kinds.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::Geometry;
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.set_geometry(
    ///     SlideSelector::index(0),
    ///     DrawableSelector::index(1),
    ///     Geometry::new(10.0, 20.0, 300.0, 200.0).with_rotation(90.0),
    /// )?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_geometry<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
        geometry: Geometry,
    ) -> Result<&mut Self, Error> {
        if !geometry.is_valid() {
            return Err(Error::Refused(Refusal::InvalidGeometry));
        }
        let (selected, item) = self.drawable(slide.into(), drawable.into())?;
        if !item.kind.geometry_editable() || item.geometry.is_none() {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        }
        let earlier = self
            .operations
            .iter()
            .rev()
            .find_map(|operation| match operation {
                Operation::Geometry {
                    drawable, geometry, ..
                } if *drawable == item.identifier => Some(*geometry),
                _ => None,
            });
        let current = earlier.or(item.geometry);
        let geometry = Geometry {
            rotation: geometry
                .rotation
                .or(current.and_then(|current| current.rotation)),
            ..geometry
        };
        if earlier.is_none() && item.geometry == Some(geometry) {
            // Already in place: staging nothing keeps the patch a no-op.
            return Ok(self);
        }
        self.operations.push(Operation::Geometry {
            slide: selected.identifier,
            drawable: item.identifier,
            message_index: item.message_index,
            message_type: item.message_type,
            geometry,
        });
        Ok(self)
    }

    /// Stage the removal of one text box, shape, line or image with its
    /// private objects (storage, stand-in captions, masks, comments).
    ///
    /// A data item whose final owner was the removed drawable is culled with
    /// its `Data/` member, and the package object watermark is lowered when
    /// the removal takes out the object holding it, as the legacy host does;
    /// shared data stays. Tables, charts, groups, movies, audio
    /// and placeholders are refused (movies and audio have
    /// `slide::media::lifecycle`); drawables referenced by builds or other
    /// objects are refused by [`Self::commit`].
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::Unsupported`].
    ///
    /// ```no_run
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// // Both selectors address the base snapshot.
    /// edit.remove(SlideSelector::index(0), DrawableSelector::index(0))?
    ///     .remove(SlideSelector::index(0), DrawableSelector::index(2))?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<&mut Self, Error> {
        let (selected, item) = self.drawable(slide.into(), drawable.into())?;
        if !item.kind.removable() {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        }
        self.removed.insert(item.identifier);
        *self.deltas.entry(selected.identifier).or_default() -= 1;
        self.operations.push(Operation::Remove {
            slide: selected.identifier,
            drawable: item.identifier,
        });
        Ok(self)
    }

    /// Stage a move of one drawable to `position` in the slide's
    /// back-to-front z-order (0 is the back-most), as Keynote's Arrange
    /// commands do. The owned-drawable order, and so every selector, is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::ZPositionOutOfRange`].
    ///
    /// ```no_run
    /// use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.move_to(SlideSelector::index(0), DrawableSelector::index(0), Position::new(2))?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn move_to<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
        position: Position,
    ) -> Result<&mut Self, Error> {
        let (selected, item) = self.drawable(slide.into(), drawable.into())?;
        let base = self
            .inventories
            .get(&selected.identifier)
            .map_or(0, |inventory| inventory.z_order.len());
        let count =
            base.saturating_add_signed(self.deltas.get(&selected.identifier).copied().unwrap_or(0));
        if position.get() >= count {
            return Err(Error::Refused(Refusal::ZPositionOutOfRange {
                position,
                count,
            }));
        }
        let untouched = !self.operations.iter().any(|operation| match operation {
            Operation::Geometry { .. }
            | Operation::Text { .. }
            | Operation::TextStyle { .. }
            | Operation::NotesStyle { .. }
            | Operation::TableCell { .. } => false,
            Operation::Remove { slide, .. }
            | Operation::Move { slide, .. }
            | Operation::AddText { slide, .. }
            | Operation::AddImage { slide, .. }
            | Operation::AddTable { slide, .. } => *slide == selected.identifier,
        });
        let current = self
            .inventories
            .get(&selected.identifier)
            .and_then(|inventory| {
                inventory
                    .z_order
                    .iter()
                    .position(|identifier| *identifier == item.identifier)
            });
        if untouched && current == Some(position.get()) {
            return Ok(self);
        }
        self.operations.push(Operation::Move {
            slide: selected.identifier,
            drawable: item.identifier,
            position: position.get(),
        });
        Ok(self)
    }

    fn add_text(
        &mut self,
        slide: SlideSelector<'_>,
        preset: Option<Preset>,
        geometry: Geometry,
        text: &str,
    ) -> Result<&mut Self, Error> {
        if !geometry.is_valid() {
            return Err(Error::Refused(Refusal::InvalidGeometry));
        }
        if let Some(Preset::RoundedRectangle { radius }) = preset
            && (!radius.is_finite()
                || radius < 0.0
                || radius > geometry.width.min(geometry.height) / 2.0)
        {
            return Err(Error::Refused(Refusal::InvalidGeometry));
        }
        let maximum = self.source.semantic_limits().max_text_bytes();
        if text.len() > maximum {
            return Err(Error::Limit {
                kind: LimitKind::TextBytes,
                observed: text.len() as u64,
                maximum: maximum as u64,
            });
        }
        let selected = self.slide(slide)?;
        self.check_drawable_limit(selected.identifier)?;
        *self.deltas.entry(selected.identifier).or_default() += 1;
        self.operations.push(Operation::AddText {
            slide: selected.identifier,
            slide_position: selected.position,
            preset,
            geometry,
            text: text.into(),
        });
        Ok(self)
    }

    /// Stage a new text box in front of the slide's drawables.
    ///
    /// The text box takes the theme's first text-box style and the body or
    /// title placeholder's paragraph, list and language styles, as the host
    /// does.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::InvalidGeometry`] and
    /// [`Error::Limit`].
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::Geometry;
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.add_text_box(SlideSelector::index(0), Geometry::new(80.0, 600.0, 640.0, 90.0), "Note")?;
    /// let commit = edit.commit()?;
    /// let created = commit.created()[0];
    /// let list = commit.package().slide_drawable_list(SlideSelector::index(0))?;
    /// assert_eq!(list[created.drawable().as_index()].text(), Some("Note"));
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_text_box<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        geometry: Geometry,
        text: &str,
    ) -> Result<&mut Self, Error> {
        self.add_text(slide.into(), None, geometry, text)
    }

    /// Stage a new preset shape with its own text, in front of the slide's
    /// drawables (styled as the host's `add_slide_shape` does).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::InvalidGeometry`] (also for a
    /// corner radius past half the shorter side) and
    /// [`Error::Limit`].
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::{Geometry, Preset};
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.add_shape(
    ///     SlideSelector::index(0),
    ///     Preset::RoundedRectangle { radius: 12.0 },
    ///     Geometry::new(100.0, 100.0, 240.0, 120.0),
    ///     "Step 1",
    /// )?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_shape<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        preset: Preset,
        geometry: Geometry,
        text: &str,
    ) -> Result<&mut Self, Error> {
        self.add_text(slide.into(), Some(preset), geometry, text)
    }

    /// Stage a new image from PNG, JPEG or GIF bytes, in front of the slide's
    /// drawables. The bytes become a new `Data/` member with a `DataInfo`
    /// record; the image takes the stylesheet's media style.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::InvalidGeometry`],
    /// [`Refusal::InvalidImage`] for a file name without a stem,
    /// [`Refusal::Unsupported`] with [`Unsupported::ImageFormat`] for other
    /// formats and [`Error::Limit`] for bytes past the source entry
    /// limit.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::Geometry;
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let logo = std::fs::read("logo.png")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.add_image(SlideSelector::index(0), &logo, "logo.png", Geometry::new(20.0, 20.0, 128.0, 128.0))?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_image<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        bytes: &[u8],
        file_name: &str,
        geometry: Geometry,
    ) -> Result<&mut Self, Error> {
        if !geometry.is_valid() {
            return Err(Error::Refused(Refusal::InvalidGeometry));
        }
        if author::sniff(bytes).is_none() {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::ImageFormat,
            }));
        }
        let maximum = self.source.source_limits().max_entry_bytes();
        if bytes.len() as u64 > maximum {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: bytes.len() as u64,
                maximum,
            });
        }
        let stem = std::path::Path::new(file_name)
            .file_stem()
            .and_then(|stem| stem.to_str());
        if stem.is_none_or(|stem| stem.is_empty() || stem.starts_with('.'))
            || file_name.contains(['/', '\\', '\0'])
        {
            return Err(Error::Refused(Refusal::InvalidImage));
        }
        let selected = self.slide(slide.into())?;
        self.check_drawable_limit(selected.identifier)?;
        *self.deltas.entry(selected.identifier).or_default() += 1;
        self.operations.push(Operation::AddImage {
            slide: selected.identifier,
            slide_position: selected.position,
            bytes: Arc::from(bytes),
            file_name: file_name.into(),
            geometry,
        });
        Ok(self)
    }

    /// The deck's first slide table (in slide and z-order) and how many
    /// tables the deck holds, read once per edit.
    fn table_template(&mut self) -> Result<(Option<u64>, usize), Error> {
        if let Some(tables) = self.tables {
            return Ok(tables);
        }
        let limits = read_limits(self.source)?;
        let slides = self.source.slides().map_err(map_read_error)?.len();
        let mut template = None;
        let mut count = 0_usize;
        for position in 0..slides {
            let record = self
                .source
                .slide_record_at(position)
                .map_err(map_read_error)?
                .ok_or(Error::InvalidSource)?;
            let inventory = read_inventory(self.source, record.slide_identifier, limits)?;
            for identifier in &inventory.z_order {
                if inventory
                    .item(*identifier)
                    .is_some_and(|item| item.kind == Kind::Table)
                {
                    template = template.or(Some(*identifier));
                    count += 1;
                }
            }
        }
        self.tables = Some((template, count));
        Ok((template, count))
    }

    /// Stage a new, empty table in front of the slide's drawables, cloned
    /// from the structure of the deck's first slide table (record 0329):
    /// its styles, header counts and storage layout, without its content,
    /// sort, filter, category or summary state. Rows and columns divide the
    /// frame evenly; seeded cells are written with [`Self::set_table_cell`]'s
    /// writer. The table gets its own formula owner.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::InvalidGeometry`],
    /// [`Refusal::InvalidTableSize`] (also for a seeded cell outside the
    /// table or a non-finite number), [`Error::Limit`], and
    /// [`Refusal::Unsupported`] with [`Unsupported::TableTemplate`] for a deck
    /// without a slide table.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::{Geometry, NewTable, TableCellValue};
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.add_table(
    ///     SlideSelector::index(0),
    ///     NewTable::new(3, 2, Geometry::new(80.0, 120.0, 600.0, 180.0))
    ///         .with_cell(0, 1, TableCellValue::Text("Q3".into())),
    /// )?;
    /// let commit = edit.commit()?;
    /// assert_eq!(commit.created().len(), 1);
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_table<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        table: NewTable,
    ) -> Result<&mut Self, Error> {
        if !table.geometry.is_valid() {
            return Err(Error::Refused(Refusal::InvalidGeometry));
        }
        let size = Error::Refused(Refusal::InvalidTableSize {
            rows: table.rows,
            columns: table.columns,
        });
        if !litchi_numbers_wire::table_create::admits_size(table.rows, table.columns)
            || table.cells.iter().any(|(row, column, value)| {
                *row >= table.rows
                    || *column >= table.columns
                    || matches!(value, TableCellValue::Number(number) if !number.is_finite())
            })
        {
            return Err(size);
        }
        let maximum = self.source.semantic_limits().max_text_bytes();
        let text_bytes = table.cells.iter().fold(0_usize, |total, (_, _, value)| {
            total.saturating_add(match value {
                TableCellValue::Text(text) => text.len(),
                _ => 0,
            })
        }) + table.name.as_deref().map_or(0, str::len);
        if text_bytes > maximum {
            return Err(Error::Limit {
                kind: LimitKind::TextBytes,
                observed: text_bytes as u64,
                maximum: maximum as u64,
            });
        }
        let selected = self.slide(slide.into())?;
        self.check_drawable_limit(selected.identifier)?;
        let (template, count) = self.table_template()?;
        let template = template.ok_or(Error::Refused(Refusal::Unsupported {
            reason: Unsupported::TableTemplate,
        }))?;
        let staged = self
            .operations
            .iter()
            .filter(|operation| matches!(operation, Operation::AddTable { .. }))
            .count();
        let name: Box<str> = match table.name.as_deref() {
            Some(name) => name.into(),
            None => format!("Table {}", count + staged + 1).into(),
        };
        *self.deltas.entry(selected.identifier).or_default() += 1;
        self.operations.push(Operation::AddTable {
            slide: selected.identifier,
            slide_position: selected.position,
            template,
            name,
            table,
        });
        Ok(self)
    }

    /// Whether a sequence of text sets of one text box or shape, each a
    /// splice within `hull` (UTF-16 code units of the drawable's text in
    /// the source package, both ends inclusive), publishes exactly what one
    /// splice from the text before them to the text after them publishes
    /// (record 0314). `false` asks the caller to commit the sets one at a
    /// time. Whether a later set follows an emptied text is the caller's to
    /// check.
    ///
    /// # Errors
    ///
    /// Typed selector errors, and [`Refusal::Unsupported`] for drawables other
    /// than text boxes and shapes.
    pub fn text_admits_combined_splice<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
        hull: std::ops::Range<usize>,
    ) -> Result<bool, Error> {
        let (_, item) = self.drawable(slide.into(), drawable.into())?;
        let (Kind::TextBox | Kind::Shape, Some(storage)) = (item.kind, item.storage) else {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        };
        stage::storage_admits_combined_splice(self.source, storage, hull)
    }

    /// Stage new plain text for one text box or shape.
    ///
    /// The commit rewrites only the changed span of the drawable's own text
    /// storage (the common prefix and suffix keep their paragraph and
    /// character attributes); paragraphs are separated by `\n`. Setting the
    /// current text stages nothing, and a later `set_text` of the same
    /// drawable in this edit replaces the earlier one. Placeholder text is
    /// written by [`Package::edit_slide_text`] instead.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::Unsupported`] for drawables
    /// other than text boxes and shapes and, with
    /// [`Unsupported::InlineObjects`], for a change that would consume or
    /// introduce an inline object marker, and [`Error::Limit`] past
    /// the semantic text limit. A storage shared with other objects is
    /// refused by [`Self::commit`] ([`Unsupported::ExternalReferences`]).
    ///
    /// ```no_run
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.set_text(SlideSelector::index(0), DrawableSelector::index(0), "Q3 review")?;
    /// let commit = edit.commit()?;
    /// let list = commit.package().slide_drawable_list(SlideSelector::index(0))?;
    /// assert_eq!(list[0].text(), Some("Q3 review"));
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_text<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
        text: &str,
    ) -> Result<&mut Self, Error> {
        let maximum = self.source.semantic_limits().max_text_bytes();
        if text.len() > maximum {
            return Err(Error::Limit {
                kind: LimitKind::TextBytes,
                observed: text.len() as u64,
                maximum: maximum as u64,
            });
        }
        let (selected, item) = self.drawable(slide.into(), drawable.into())?;
        let (Kind::TextBox | Kind::Shape, Some(storage), Some(current)) =
            (item.kind, item.storage, item.text.as_deref())
        else {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        };
        let earlier = self
            .operations
            .iter_mut()
            .find_map(|operation| match operation {
                Operation::Text { drawable, text, .. } if *drawable == item.identifier => {
                    Some(text)
                },
                _ => None,
            });
        let base = earlier.as_deref().map_or(current, |text| &**text);
        if text_change_touches_markers(base, text) {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::InlineObjects,
            }));
        }
        if let Some(earlier) = earlier {
            *earlier = text.into();
            return Ok(self);
        }
        if current == text {
            // Already in place: staging nothing keeps the patch a no-op.
            return Ok(self);
        }
        self.operations.push(Operation::Text {
            slide: selected.identifier,
            drawable: item.identifier,
            storage,
            text: text.into(),
        });
        Ok(self)
    }

    /// Stage character styling of one span (characters) of a text box's or
    /// shape's text.
    ///
    /// Each character-style run inside the span gets a new variation of its
    /// style carrying the set properties (unset ones stay inherited); the
    /// variations are registered in the storage's stylesheet. Several
    /// styles and text changes may be staged; a style follows the text as
    /// it stands when it is staged.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::InvalidSpan`],
    /// [`Refusal::Unsupported`] for drawables other than text boxes, shapes and
    /// title and body placeholders, and [`Unsupported::TextStyling`] for a
    /// style that sets nothing; a
    /// storage shared with other objects is refused by [`Self::commit`].
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::{Style, TextSpan};
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.set_text_style(
    ///     SlideSelector::index(0),
    ///     DrawableSelector::index(0),
    ///     TextSpan::new(0, 4),
    ///     Style::new().with_bold(true),
    /// )?;
    /// let commit = edit.commit()?;
    /// let runs = commit
    ///     .package()
    ///     .slide_drawable_text_runs(SlideSelector::index(0), DrawableSelector::index(0))?;
    /// assert_eq!(runs[0].style().bold, Some(true));
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_text_style<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
        span: TextSpan,
        style: Style,
    ) -> Result<&mut Self, Error> {
        let (selected, item) = self.drawable(slide.into(), drawable.into())?;
        // Title and body placeholders style their slide's own storage, as
        // text boxes do (record 0312); their text writes go through
        // `slide::text`, never through this edit.
        let (
            Kind::TextBox | Kind::Shape | Kind::TitlePlaceholder | Kind::BodyPlaceholder,
            Some(storage),
            Some(current),
        ) = (item.kind, item.storage, item.text.as_deref())
        else {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        };
        let staged_text = self
            .operations
            .iter()
            .rev()
            .find_map(|operation| match operation {
                Operation::Text { drawable, text, .. } if *drawable == item.identifier => {
                    Some(&**text)
                },
                _ => None,
            });
        let text = staged_text.unwrap_or(current);
        if span.start() >= span.end() || span.end() > text.chars().count() {
            return Err(Error::Refused(Refusal::InvalidSpan));
        }
        if style == Style::new() {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::TextStyling,
            }));
        }
        self.operations.push(Operation::TextStyle {
            slide: selected.identifier,
            drawable: item.identifier,
            storage,
            span,
            style,
        });
        Ok(self)
    }

    /// Stage character styling of one span (characters) of a slide's
    /// speaker notes (record 0329), as [`Self::set_text_style`] styles a
    /// text box: each character-style run inside the span gets a new
    /// variation of its style carrying the set properties, registered in
    /// the notes storage's stylesheet. Notes text is written by
    /// [`Package::edit_slide_notes`], never by this edit.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::InvalidSpan`],
    /// [`Refusal::Unsupported`] with [`Unsupported::Kind`] for a slide
    /// without notes and [`Unsupported::TextStyling`] for a style that sets
    /// nothing; notes stored outside their slide's component, or a storage
    /// other objects reference, are refused by [`Self::commit`].
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::{Style, TextSpan};
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.set_notes_text_style(
    ///     SlideSelector::index(0),
    ///     TextSpan::new(0, 4),
    ///     Style::new().with_italic(true),
    /// )?;
    /// let commit = edit.commit()?;
    /// let runs = commit.package().slide_notes_text_runs(SlideSelector::index(0))?;
    /// assert_eq!(runs[0].style().italic, Some(true));
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_notes_text_style<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        span: TextSpan,
        style: Style,
    ) -> Result<&mut Self, Error> {
        let selected = self.slide(slide.into())?;
        let (note, storage, text) =
            super::slide_notes::notes_identity(self.source, selected.position)
                .map_err(|_error| Error::InvalidSource)?
                .ok_or(Error::Refused(Refusal::Unsupported {
                    reason: Unsupported::Kind { kind: Kind::Other },
                }))?;
        if span.start() >= span.end() || span.end() > text.chars().count() {
            return Err(Error::Refused(Refusal::InvalidSpan));
        }
        if style == Style::new() {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::TextStyling,
            }));
        }
        self.operations.push(Operation::NotesStyle {
            slide: selected.identifier,
            note,
            storage,
            span,
            style,
        });
        Ok(self)
    }

    /// Stage a value in one cell (zero-based `row`, `column`) of a slide
    /// table.
    ///
    /// Text is interned in the table's string list; the cell keeps its
    /// style and format keys; a cleared cell without them leaves the
    /// storage. Tables in a formula graph are refused
    /// ([`Unsupported::TableFormulas`]).
    ///
    /// # Errors
    ///
    /// Typed selector errors, [`Refusal::Unsupported`] for drawables other
    /// than tables, [`Refusal::InvalidSpan`] for a cell outside the table.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::TableCellValue;
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.set_table_cell(
    ///     SlideSelector::index(0),
    ///     DrawableSelector::index(2),
    ///     1,
    ///     1,
    ///     TableCellValue::Text("Q3".into()),
    /// )?;
    /// let commit = edit.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_table_cell<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
        row: u32,
        column: u32,
        value: TableCellValue,
    ) -> Result<&mut Self, Error> {
        let (selected, item) = self.drawable(slide.into(), drawable.into())?;
        if item.kind != Kind::Table {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        }
        if let TableCellValue::Number(number) = value
            && !number.is_finite()
        {
            return Err(Error::Refused(Refusal::InvalidSpan));
        }
        self.operations.push(Operation::TableCell {
            slide: selected.identifier,
            drawable: item.identifier,
            row,
            column,
            value,
        });
        Ok(self)
    }

    /// Apply every staged operation in one workspace, publish one candidate,
    /// reopen, validate and read it back.
    ///
    /// An edit with no staged operation commits the unchanged package.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::Unsupported`] for drawables referenced from outside
    /// their private graph, [`Error::Verification`] when the candidate does
    /// not read back as staged, and typed source, limit and allocation
    /// errors.
    ///
    /// ```no_run
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.remove(SlideSelector::index(0), DrawableSelector::index(0))?;
    /// let (edited, patch) = edit.commit()?.into_parts();
    /// let restored = edited.apply_slide_drawables(&patch.inverse())?;
    /// assert_eq!(restored.package().slide_drawable_list(SlideSelector::index(0))?.len(),
    ///     package.slide_drawable_list(SlideSelector::index(0))?.len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn commit(self) -> Result<Commit, Error> {
        let source_bytes = physical_source(self.source)?;
        if self.operations.is_empty() {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    operations: 0,
                },
                diagnostics: Diagnostics::default(),
                created: Box::new([]),
            });
        }
        commit_operations(
            self.source,
            source_bytes,
            &self.operations,
            &self.inventories,
        )
    }
}

/// Exact-source-checked reversible slide-drawable patch.
///
/// The patch retains the exact source and target artifacts; applying it
/// checks the exact source bytes, and its inverse restores the source.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    operations: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("operations", &self.operations)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return whether source and target are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            operations: self.operations,
        }
    }

    /// Return how many operations the forward edit applied.
    #[must_use]
    pub const fn operations(&self) -> usize {
        self.operations
    }

    /// Return the fingerprint of the exact source artifact.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the fingerprint of the exact target artifact.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
}

/// Compact evidence for one verified slide-drawable commit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operations: usize,
    created_drawables: usize,
    removed_drawables: usize,
    removed_objects: usize,
    removed_data: usize,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return how many operations were applied.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }

    /// Return how many drawables were created.
    #[must_use]
    pub const fn created_drawables(self) -> usize {
        self.created_drawables
    }

    /// Return how many drawables were removed.
    #[must_use]
    pub const fn removed_drawables(self) -> usize {
        self.removed_drawables
    }

    /// Return how many native objects the removals took out.
    #[must_use]
    pub const fn removed_objects(self) -> usize {
        self.removed_objects
    }

    /// Return how many data items (`Data/` members with their records) the
    /// removals culled because the removed drawable was their final owner.
    #[must_use]
    pub const fn removed_data(self) -> usize {
        self.removed_data
    }

    /// Return how many package members were created or rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the candidate was fully reparsed and validated.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The verified result of one slide-drawable edit.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
    created: Box<[Created]>,
}

impl Commit {
    /// Borrow the committed package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Take the committed package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// The drawables this commit created, in staging order.
    #[must_use]
    pub fn created(&self) -> &[Created] {
        &self.created
    }

    /// Split the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    /// List the drawables a slide owns, in source order (the order
    /// [`DrawableSelector`] addresses).
    ///
    /// ```no_run
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// for drawable in package.slide_drawable_list(SlideSelector::index(0))?.iter() {
    ///     println!("{:?} {:?} {:?}", drawable.kind(), drawable.geometry(), drawable.text());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Error::Limit`] past the
    /// semantic reference or text limits and [`Error::InvalidSource`] for
    /// malformed slides.
    pub fn slide_drawable_list<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
    ) -> Result<Box<[Drawable]>, Error> {
        let selected = select_slide(self, slide.into())?;
        let inventory = read_inventory(self, selected.identifier, read_limits(self)?)?;
        Ok(public_list(&inventory))
    }

    /// Read the character-style runs of one drawable's text.
    ///
    /// Each run covers a span of characters and carries the style of the
    /// storage's character-style table where it is decodable; text without
    /// a character-style table is one run with an empty style.
    ///
    /// ```no_run
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let runs = package.slide_drawable_text_runs(SlideSelector::index(0), DrawableSelector::index(0))?;
    /// for run in runs.iter() {
    ///     println!("{:?} bold={:?}", run.span(), run.style().bold);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::Unsupported`] for
    /// drawables without text.
    pub fn slide_drawable_text_runs<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Box<[TextRun]>, Error> {
        let selected = select_slide(self, slide.into())?;
        let inventory = read_inventory(self, selected.identifier, read_limits(self)?)?;
        let selector = drawable.into();
        let item = inventory
            .items
            .get(selector.as_index())
            .ok_or(Error::Refused(Refusal::DrawablePositionNotFound {
                position: selector.as_position(),
            }))?;
        let (Some(storage), Some(text_value)) = (item.storage, item.text.as_deref()) else {
            return Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: item.kind },
            }));
        };
        let wire = self
            .semantic_wire_limits()
            .map_err(|_error| Error::InvalidSource)?;
        text::runs(self, storage, text_value, wire).map_err(lift)
    }

    /// Read the character-style runs of one slide's speaker notes (record
    /// 0329), as [`Self::slide_drawable_text_runs`] reads a text box's.
    ///
    /// ```no_run
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// for run in package.slide_notes_text_runs(SlideSelector::index(0))?.iter() {
    ///     println!("{:?} italic={:?}", run.span(), run.style().italic);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::Unsupported`] for a slide
    /// without notes.
    pub fn slide_notes_text_runs<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
    ) -> Result<Box<[TextRun]>, Error> {
        let selected = select_slide(self, slide.into())?;
        let (_, storage, text_value) = super::slide_notes::notes_identity(self, selected.position)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Kind { kind: Kind::Other },
            }))?;
        let wire = self
            .semantic_wire_limits()
            .map_err(|_error| Error::InvalidSource)?;
        text::runs(self, storage, &text_value, wire).map_err(lift)
    }

    /// Start one slide-drawable edit.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::drawable::Geometry;
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.add_text_box(SlideSelector::index(0), Geometry::new(0.0, 0.0, 200.0, 50.0), "Hi")?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_slide_drawables(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Apply a slide-drawable patch only to its exact immutable source.
    ///
    /// ```no_run
    /// use litchi_keynote::{DrawableSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_drawables();
    /// edit.remove(SlideSelector::index(0), DrawableSelector::index(0))?;
    /// let patch = edit.commit()?.patch().clone();
    /// let replayed = Package::open("deck.key")?.apply_slide_drawables(&patch)?;
    /// assert!(replayed.diagnostics().changed());
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// exact source, and typed source errors when the target cannot be
    /// reopened.
    pub fn apply_slide_drawables(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                operations: patch.operations,
                full_reparse_performed: true,
                ..Diagnostics::default()
            },
            created: Box::new([]),
        })
    }
}

fn public_list(inventory: &SlideDrawables) -> Box<[Drawable]> {
    inventory
        .items
        .iter()
        .enumerate()
        .map(|(index, item)| Drawable {
            selector: DrawableSelector::index(index),
            kind: item.kind,
            geometry: item.geometry,
            text: item.text.clone(),
            name: item.name.clone(),
            locked: item.locked,
            z_position: inventory
                .z_order
                .iter()
                .position(|identifier| *identifier == item.identifier)
                .map(Position::new),
        })
        .collect()
}

fn select_slide(package: &Package, selector: SlideSelector<'_>) -> Result<SelectedSlide, Error> {
    let show = package.show().map_err(map_read_error)?;
    let position = show
        .select_slide(selector)
        .map_err(map_selector_error)?
        .ok_or(match selector {
            SlideSelector::Name(_) => Error::Refused(Refusal::SlideNameNotFound),
            SlideSelector::Position(position) => {
                Error::Refused(Refusal::SlidePositionNotFound { position })
            },
        })?
        .position();
    let record = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource)?;
    Ok(SelectedSlide {
        position,
        identifier: record.slide_identifier,
    })
}

fn read_limits(package: &Package) -> Result<ReadLimits, Error> {
    let semantic = package.semantic_limits();
    Ok(ReadLimits {
        wire: package
            .semantic_wire_limits()
            .map_err(|_error| Error::InvalidSource)?,
        max_drawables: semantic.max_references(),
        max_text_bytes: semantic.max_text_bytes(),
    })
}

fn read_inventory(
    source: &impl inventory::Objects,
    slide: u64,
    limits: ReadLimits,
) -> Result<SlideDrawables, Error> {
    inventory::read_slide(source, slide, limits).map_err(|failure| match failure {
        ReadFailure::Structure(error) => lift(error),
        ReadFailure::Drawables { observed, maximum } => Error::Limit {
            kind: LimitKind::Drawables,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadFailure::TextBytes { observed, maximum } => Error::Limit {
            kind: LimitKind::TextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
    })
}

fn commit_operations(
    source: &Package,
    source_bytes: Arc<[u8]>,
    operations: &[Operation],
    base: &BTreeMap<u64, SlideDrawables>,
) -> Result<Commit, Error> {
    source.validate().map_err(map_read_error)?;
    let limits = read_limits(source)?;
    let mut workspace = Workspace::new(source).map_err(lift)?;
    let mut context = None;
    let mut created = Vec::new();
    let mut tables = Vec::new();
    let mut notes: Vec<(u64, String)> = Vec::new();
    let mut touched: HashMap<u64, HashSet<u64>> = HashMap::new();
    let mut diagnostics = Diagnostics {
        operations: operations.len(),
        ..Diagnostics::default()
    };
    for operation in operations {
        match operation {
            Operation::Geometry {
                slide,
                drawable,
                message_index,
                message_type,
                geometry,
            } => {
                stage::set_geometry(
                    &mut workspace,
                    *drawable,
                    *message_index,
                    *message_type,
                    *geometry,
                )?;
                let staged = read_inventory(&workspace, *slide, limits)?;
                if staged.item(*drawable).and_then(|item| item.geometry) != Some(*geometry) {
                    return Err(Error::Verification);
                }
                touched.entry(*slide).or_default().insert(*drawable);
            },
            Operation::Remove { slide, drawable } => {
                let removal = stage::remove(&mut workspace, *slide, *drawable)?;
                diagnostics.removed_drawables += 1;
                diagnostics.removed_objects += removal.objects;
                diagnostics.removed_data += removal.data;
                touched.entry(*slide).or_default().insert(*drawable);
            },
            Operation::Move {
                slide,
                drawable,
                position,
            } => {
                stage::move_to(&mut workspace, *slide, *drawable, *position)?;
                let staged = read_inventory(&workspace, *slide, limits)?;
                if staged.z_order.get(*position) != Some(drawable) {
                    return Err(Error::Verification);
                }
                touched.entry(*slide).or_default();
            },
            Operation::AddText {
                slide,
                slide_position,
                preset,
                geometry,
                text,
            } => {
                if context.is_none() {
                    context = Some(author::Context::read(&workspace).map_err(lift)?);
                }
                let context = context.as_ref().ok_or(Error::InvalidSource)?;
                let drawable = author::add_text_shape(
                    &mut workspace,
                    context,
                    *slide,
                    *preset,
                    *geometry,
                    text,
                )?;
                let staged = read_inventory(&workspace, *slide, limits)?;
                let item = staged.item(drawable).ok_or(Error::Verification)?;
                let expected_kind = if preset.is_some() {
                    Kind::Shape
                } else {
                    Kind::TextBox
                };
                if item.kind != expected_kind
                    || item.text.as_deref() != Some(&**text)
                    || item.geometry
                        != Some(Geometry {
                            rotation: Some(geometry.rotation.unwrap_or(0.0)),
                            ..*geometry
                        })
                {
                    return Err(Error::Verification);
                }
                created.push((*slide_position, *slide, drawable));
                touched.entry(*slide).or_default().insert(drawable);
            },
            Operation::Text {
                slide,
                drawable,
                storage,
                text,
            } => {
                stage::set_text(&mut workspace, *slide, *drawable, *storage, text, source)?;
                let staged = read_inventory(&workspace, *slide, limits)?;
                if staged.item(*drawable).and_then(|item| item.text.as_deref()) != Some(&**text) {
                    return Err(Error::Verification);
                }
                touched.entry(*slide).or_default().insert(*drawable);
            },
            Operation::TextStyle {
                slide,
                drawable,
                storage,
                span,
                style,
            } => {
                let before = read_inventory(&workspace, *slide, limits)?
                    .item(*drawable)
                    .and_then(|item| item.text.clone());
                stage::set_text_style(&mut workspace, *slide, *drawable, *storage, *span, style)?;
                let staged = read_inventory(&workspace, *slide, limits)?;
                if staged.item(*drawable).and_then(|item| item.text.clone()) != before {
                    return Err(Error::Verification);
                }
                touched.entry(*slide).or_default().insert(*drawable);
            },
            Operation::TableCell {
                slide,
                drawable,
                row,
                column,
                value,
            } => {
                stage::set_table_cell(&mut workspace, *drawable, *row, *column, value)?;
                touched.entry(*slide).or_default().insert(*drawable);
            },
            Operation::NotesStyle {
                slide,
                note,
                storage,
                span,
                style,
            } => {
                let wire = workspace.wire;
                let before = inventory::storage_text(&workspace, *storage, wire).map_err(lift)?;
                stage::set_text_style(&mut workspace, *slide, *note, *storage, *span, style)?;
                let after = inventory::storage_text(&workspace, *storage, wire).map_err(lift)?;
                if before != after {
                    return Err(Error::Verification);
                }
                notes.push((*slide, before));
                touched.entry(*slide).or_default();
            },
            Operation::AddTable {
                slide,
                slide_position,
                template,
                name,
                table,
            } => {
                let drawable = table::add_table(
                    &mut workspace,
                    &table::Request {
                        slide: *slide,
                        template: *template,
                        name,
                        rows: table.rows,
                        columns: table.columns,
                        geometry: table.geometry,
                    },
                )?;
                for (row, column, value) in &table.cells {
                    stage::set_table_cell(&mut workspace, drawable, *row, *column, value)?;
                }
                let staged = read_inventory(&workspace, *slide, limits)?;
                let item = staged.item(drawable).ok_or(Error::Verification)?;
                if item.kind != Kind::Table
                    || item
                        .geometry
                        .map(|frame| (frame.x, frame.y, frame.width, frame.height))
                        != Some((
                            table.geometry.x,
                            table.geometry.y,
                            table.geometry.width,
                            table.geometry.height,
                        ))
                {
                    return Err(Error::Verification);
                }
                tables.push((*slide_position, drawable, name.clone(), table.clone()));
                created.push((*slide_position, *slide, drawable));
                touched.entry(*slide).or_default().insert(drawable);
            },
            Operation::AddImage {
                slide,
                slide_position,
                bytes,
                file_name,
                geometry,
            } => {
                if context.is_none() {
                    context = Some(author::Context::read(&workspace).map_err(lift)?);
                }
                let context = context.as_ref().ok_or(Error::InvalidSource)?;
                let drawable = author::add_image(
                    &mut workspace,
                    context,
                    *slide,
                    bytes,
                    file_name,
                    *geometry,
                )?;
                let staged = read_inventory(&workspace, *slide, limits)?;
                let item = staged.item(drawable).ok_or(Error::Verification)?;
                if item.kind != Kind::Image {
                    return Err(Error::Verification);
                }
                created.push((*slide_position, *slide, drawable));
                touched.entry(*slide).or_default().insert(drawable);
            },
        }
    }
    check_object_limit(source, &workspace)?;
    let mut expected = Vec::with_capacity(touched.len());
    for (slide, changed) in &touched {
        let staged = read_inventory(&workspace, *slide, limits)?;
        // Every drawable no operation touched reads back unchanged.
        if let Some(before) = base.get(slide) {
            for item in &before.items {
                if changed.contains(&item.identifier) {
                    if staged.item(item.identifier).is_some()
                        && operations.iter().any(|operation| {
                            matches!(operation, Operation::Remove { drawable, .. }
                                if *drawable == item.identifier)
                        })
                    {
                        return Err(Error::Verification);
                    }
                    continue;
                }
                if staged.item(item.identifier) != Some(item) {
                    return Err(Error::Verification);
                }
            }
        }
        expected.push(staged);
    }
    diagnostics.touched_components = workspace.touched_members();
    if diagnostics.removed_drawables > 0 {
        workspace.release_removed_identifier_suffix();
    }
    let bytes = workspace.publish().map_err(lift)?;
    let candidate = source.structure_reopen(bytes).map_err(map_read_error)?;
    for staged in &expected {
        if read_inventory(&candidate, staged.slide, limits)? != *staged {
            return Err(Error::Verification);
        }
    }
    for (slide_position, drawable, name, table) in &tables {
        verify_table(
            &candidate,
            &expected,
            *slide_position,
            *drawable,
            name,
            table,
        )?;
    }
    // Styled notes keep their text.
    for (slide, text) in &notes {
        let mut position = None;
        for index in 0.. {
            match candidate.slide_record_at(index).map_err(map_read_error)? {
                Some(record) if record.slide_identifier == *slide => {
                    position = Some(Position::new(index));
                    break;
                },
                Some(_) => {},
                None => break,
            }
        }
        let position = position.ok_or(Error::Verification)?;
        let read = super::slide_notes::notes_identity(&candidate, position)
            .map_err(|_error| Error::Verification)?
            .ok_or(Error::Verification)?;
        if read.2 != *text {
            return Err(Error::Verification);
        }
    }
    let mut public_created = Vec::with_capacity(created.len());
    for (slide_position, slide, drawable) in created {
        let inventory = expected
            .iter()
            .find(|inventory| inventory.slide == slide)
            .ok_or(Error::Verification)?;
        let index = inventory
            .items
            .iter()
            .position(|item| item.identifier == drawable)
            .ok_or(Error::Verification)?;
        public_created.push(Created {
            slide: slide_position,
            drawable: DrawableSelector::index(index),
        });
    }
    let target = physical_source(&candidate)?;
    let patch = Patch {
        artifacts: ExactArtifacts::new(source_bytes, target),
        operations: operations.len(),
    };
    diagnostics.changed = !patch.is_noop();
    diagnostics.created_drawables = public_created.len();
    diagnostics.full_reparse_performed = true;
    Ok(Commit {
        package: candidate,
        patch,
        diagnostics,
        created: public_created.into_boxed_slice(),
    })
}

/// Read a created table back through the slide-table reader: its name,
/// size and seeded values.
fn verify_table(
    candidate: &Package,
    expected: &[SlideDrawables],
    slide: Position,
    drawable: u64,
    name: &str,
    table: &NewTable,
) -> Result<(), Error> {
    use crate::slide::table::cells::{At, Value};
    let inventory = expected
        .iter()
        .find(|inventory| inventory.item(drawable).is_some())
        .ok_or(Error::Verification)?;
    let ordinal = inventory
        .z_order
        .iter()
        .take_while(|identifier| **identifier != drawable)
        .filter(|identifier| {
            inventory
                .item(**identifier)
                .is_some_and(|item| item.kind == Kind::Table)
        })
        .count();
    let read = candidate
        .slide_table_cells(
            SlideSelector::position(slide),
            crate::TableSelector::index(ordinal),
        )
        .map_err(|_error| Error::Verification)?;
    let read = read.table();
    let dimensions = read.dimensions();
    if read.name() != name
        || (dimensions.rows(), dimensions.columns()) != (table.rows, table.columns)
    {
        return Err(Error::Verification);
    }
    for (row, column, value) in &table.cells {
        let stored = read.get(At::new(*row, *column));
        let matches = match (value, stored) {
            (TableCellValue::Text(text), Some(Value::Text(stored))) => stored == text,
            (TableCellValue::Number(number), Some(Value::Number(stored))) => {
                stored.get() == *number
            },
            (TableCellValue::Boolean(flag), Some(Value::Boolean(stored))) => stored == flag,
            (TableCellValue::Empty, None | Some(Value::Empty)) => true,
            _ => false,
        };
        if !matches {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

/// Whether replacing `before` with `after` would consume or introduce an
/// inline object marker: the changed middle (after the common prefix and
/// suffix) of either side holds one.
fn text_change_touches_markers(before: &str, after: &str) -> bool {
    let (before, after) = stage::changed_middle(before, after);
    [before, after]
        .iter()
        .any(|text| text.contains(['\u{fffc}', '\u{000e}']))
}

fn check_object_limit(source: &Package, workspace: &Workspace<'_>) -> Result<(), Error> {
    let maximum = source.semantic_limits().max_objects();
    let observed = source
        .structure_object_count()
        .saturating_add(workspace.created_objects());
    if observed > maximum {
        return Err(Error::Limit {
            kind: LimitKind::Objects,
            observed: observed as u64,
            maximum: maximum as u64,
        });
    }
    Ok(())
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) if source.source_is_exact() => Ok(source.shared_source()),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            Err(Error::Refused(Refusal::UnsupportedSource))
        },
    }
}

/// Lift a kernel error into this family's vocabulary.
fn lift(error: StructureError) -> Error {
    match error {
        StructureError::UnsupportedContent(_) => Error::Refused(Refusal::Unsupported {
            reason: Unsupported::UnmodelledContent,
        }),
        StructureError::UnsupportedRegistry => Error::Refused(Refusal::Unsupported {
            reason: Unsupported::RegistryShape,
        }),
        StructureError::IdentifierOverflow => Error::Limit {
            kind: LimitKind::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        },
        StructureError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                StructureLimit::WireBytes => LimitKind::WireBytes,
                StructureLimit::WireFields => LimitKind::WireFields,
                StructureLimit::WireNesting => LimitKind::WireNesting,
                StructureLimit::Work => LimitKind::Work,
                StructureLimit::OutputBytes => LimitKind::OutputBytes,
                StructureLimit::EntryBytes => LimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        StructureError::Allocation(amount) => Error::Allocation { amount },
        StructureError::MissingObject
        | StructureError::Registry(_)
        | StructureError::Unexpected(_)
        | StructureError::Malformed
        | StructureError::Physical => Error::InvalidSource,
    }
}

fn map_selector_error(error: SlideSelectorError) -> Error {
    match error {
        SlideSelectorError::EmptySlideName => Error::Refused(Refusal::EmptySlideName),
        SlideSelectorError::DuplicateSlideName { .. } => Error::Refused(Refusal::AmbiguousSelector),
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            path: _,
        } => Error::Limit {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Objects,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            path: _,
        } => Error::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::Work,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// Edit the nested message of `payload` at `path` (shared with other
/// structural families).
pub(in crate::package) fn edit_nested_payload(
    payload: &[u8],
    path: &[u32],
    wire: litchi_iwa_common::WireLimits,
    edit: &mut dyn FnMut(&mut super::structure::wire::Fields) -> Result<(), StructureError>,
) -> Result<Vec<u8>, StructureError> {
    stage::edit_nested(payload, path, wire, edit)
}

/// Patch a `TSD.GeometryArchive` field list in place.
pub(in crate::package) fn patch_geometry_fields(
    fields: &mut super::structure::wire::Fields,
    geometry: Geometry,
    wire: litchi_iwa_common::WireLimits,
) -> Result<(), StructureError> {
    stage::patch_geometry(fields, geometry, wire)
}

/// Read a `TSD.DrawableArchive`'s frame.
pub(in crate::package) fn read_geometry(
    drawable: &[u8],
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Geometry>, StructureError> {
    let fields = super::structure::wire::Fields::parse(drawable, wire)?;
    match fields.unique(1)? {
        Some(geometry) => inventory::read_geometry(geometry.message()?, wire),
        None => Ok(None),
    }
}

/// Replace the unique `message_type` payload of `identifier`, keeping the
/// header.
pub(in crate::package) fn replace_message_payload(
    workspace: &mut Workspace<'_>,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
) -> Result<(), StructureError> {
    let index = workspace
        .require(identifier)?
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(StructureError::Unexpected("message type"))?;
    stage::replace_payload(workspace, identifier, index, message_type, data)
}

/// Remove one slide-owned drawable and its private closure (shared with the
/// chart lifecycle). Returns the removed object count and culled data.
pub(in crate::package) fn remove_drawable_closure(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
) -> Result<(usize, usize), Error> {
    let removal = stage::remove(workspace, slide, drawable)?;
    Ok((removal.objects, removal.data))
}

/// Every object reference one staged object holds (headers and payloads).
pub(in crate::package) fn outgoing_references(
    workspace: &Workspace<'_>,
    identifier: u64,
) -> Result<Vec<u64>, StructureError> {
    stage::outgoing(workspace, identifier)
}
