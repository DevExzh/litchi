//! Detached, bounded WordprocessingML Ink host authoring.
//!
//! This module only writes the story-side host.  Package parts, relationship
//! allocation, content types, and fallback image publication remain the
//! responsibility of the source-bound transaction.  The values below contain
//! no package or source identity, so a prepared host can be validated before
//! it is attached to an OPC graph.

#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "the host grammar is kept beside its detached values"
)]
use std::fmt::Write as FmtWrite;
use std::sync::Arc;

use litchi_core::xml::escape_xml;

use super::codec::{Form, scan};
use super::{Limits, host};
use crate::package::story::StoryDialect;
use crate::{Error, Result};

const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const TRANSITIONAL_DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const TRANSITIONAL_WORDPROCESSING_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const STRICT_WORDPROCESSING_DRAWING: &str =
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const MARKUP_COMPATIBILITY: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WORD_2010_WORDML: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const WORDPROCESSING_INK: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
const WORDPROCESSING_CANVAS: &str =
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const WORDPROCESSING_GROUP: &str =
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const VML_NAMESPACE: &str = "urn:schemas-microsoft-com:vml";
const OFFICE_NAMESPACE: &str = "urn:schemas-microsoft-com:office:office";

/// Maximum fallback image bytes retained by one detached value.
const MAX_FALLBACK_IMAGE_BYTES: usize = 16 * 1024 * 1024;
/// Maximum pixel dimension accepted by the bounded fallback image subset.
const MAX_FALLBACK_IMAGE_DIMENSION: u32 = 1_000_000;
/// Maximum encoded host fragment emitted by this module.
const MAX_HOST_BYTES: usize = 1024 * 1024;

const MAX_RELATIONSHIP_ID_BYTES: usize = 1024;
const MAX_HOST_NODES: usize = 4096;
const MAX_HOST_DEPTH: usize = 64;
const MAX_EMU: u64 = 27_273_042_316_900;
// ECMA-376 `ST_CoordinateUnqualified` is asymmetric in the local strict
// schema: minInclusive=-27273042329600 and maxInclusive=27273042316900.
const MIN_COORDINATE: i64 = -27_273_042_329_600;
const MAX_COORDINATE: i64 = 27_273_042_316_900;
const MAX_FALLBACK_IMAGE_PIXELS: u64 = 64 * 1024 * 1024;
const MAX_FALLBACK_EXPANDED_BYTES: u64 = 256 * 1024 * 1024;

/// Direct WordprocessingML `contentPart` target profile.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum BaseProfile {
    /// Word's direct-base variation, whose target is `text/xml`.
    WordTextXml,
    /// The ODRAWXML Ink Content Part variation, whose target is
    /// `application/inkml+xml`.
    InkContent,
}

impl BaseProfile {
    /// Content type required for a newly created direct-base target.
    #[must_use]
    pub const fn content_type(self) -> &'static str {
        match self {
            Self::WordTextXml => "text/xml",
            Self::InkContent => "application/inkml+xml",
        }
    }
}

/// Validated image media type for a VML fallback.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum ImageType {
    /// ISO/IEC 15948 PNG.
    Png,
    /// ISO/IEC 10918 JPEG.
    Jpeg,
}

impl ImageType {
    /// OPC media type used for the fallback image part.
    #[must_use]
    pub const fn content_type(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
        }
    }

    const fn extension(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
        }
    }
}

/// Positive pixel dimensions supplied with a fallback image.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct ImageDimensions {
    width: u32,
    height: u32,
}

impl ImageDimensions {
    /// Construct bounded, non-zero image dimensions.
    pub fn new(width: u32, height: u32) -> Result<Self> {
        let pixels = u64::from(width).saturating_mul(u64::from(height));
        if width == 0
            || height == 0
            || width > MAX_FALLBACK_IMAGE_DIMENSION
            || height > MAX_FALLBACK_IMAGE_DIMENSION
            || pixels > MAX_FALLBACK_IMAGE_PIXELS
        {
            return Err(invalid(
                "DOCX Ink fallback image dimensions must be positive and bounded",
            ));
        }
        Ok(Self { width, height })
    }

    /// Pixel width.
    #[must_use]
    pub const fn width(self) -> u32 {
        self.width
    }

    /// Pixel height.
    #[must_use]
    pub const fn height(self) -> u32 {
        self.height
    }
}

/// Caller-supplied, validated PNG or JPEG fallback image.
///
/// The bytes are retained in one shared immutable allocation so package graph
/// code can publish the media part without copying it.  XML is never accepted
/// as a fallback input and no renderer is inferred from InkML traces.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct FallbackImage {
    bytes: Arc<Vec<u8>>,
    media_type: ImageType,
    dimensions: ImageDimensions,
}

impl FallbackImage {
    /// Validate image bytes against the caller-supplied media type and size.
    pub fn new(
        bytes: impl Into<Arc<Vec<u8>>>,
        media_type: ImageType,
        dimensions: ImageDimensions,
    ) -> Result<Self> {
        let bytes = bytes.into();
        validate_image(bytes.as_slice(), media_type, dimensions)?;
        Ok(Self {
            bytes,
            media_type,
            dimensions,
        })
    }

    /// Detect the supported media type and dimensions from bounded bytes.
    pub fn from_bytes(bytes: impl Into<Arc<Vec<u8>>>) -> Result<Self> {
        let bytes = bytes.into();
        if bytes.len() > MAX_FALLBACK_IMAGE_BYTES {
            return Err(image_limit(bytes.len()));
        }
        let (media_type, dimensions) = detect_image(bytes.as_slice())?;
        Self::new(bytes, media_type, dimensions)
    }

    /// Borrow the exact image bytes for media-part publication.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Clone the shared image allocation for package publication.
    #[must_use]
    pub fn shared_source(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.bytes)
    }

    /// Image media type.
    #[must_use]
    pub const fn media_type(&self) -> ImageType {
        self.media_type
    }

    /// Image OPC content type.
    #[must_use]
    pub const fn content_type(&self) -> &'static str {
        self.media_type.content_type()
    }

    /// Image dimensions.
    #[must_use]
    pub const fn dimensions(&self) -> ImageDimensions {
        self.dimensions
    }

    /// Image file extension used by package URI allocation.
    #[must_use]
    pub const fn extension(&self) -> &'static str {
        self.media_type.extension()
    }
}

/// Positive DrawingML extent in EMUs.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct Geometry {
    width: u64,
    height: u64,
}

impl Geometry {
    /// Construct a positive bounded extent.
    pub fn new(width: u64, height: u64) -> Result<Self> {
        if width == 0 || height == 0 || width > MAX_EMU || height > MAX_EMU {
            return Err(invalid(
                "DOCX Ink drawing extent must be positive and schema-bounded",
            ));
        }
        Ok(Self { width, height })
    }

    /// Width in EMUs.
    #[must_use]
    pub const fn width_emu(self) -> u64 {
        self.width
    }

    /// Height in EMUs.
    #[must_use]
    pub const fn height_emu(self) -> u64 {
        self.height
    }
}

/// A checked DrawingML point used by floating anchors and group transforms.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct Point {
    x: i64,
    y: i64,
}

impl Point {
    /// Construct a point in the schema coordinate range.
    pub fn new(x: i64, y: i64) -> Result<Self> {
        if !(MIN_COORDINATE..=MAX_COORDINATE).contains(&x)
            || !(MIN_COORDINATE..=MAX_COORDINATE).contains(&y)
        {
            return Err(invalid(
                "DOCX Ink anchor point exceeds the DrawingML coordinate range",
            ));
        }
        Ok(Self { x, y })
    }

    const fn origin() -> Self {
        Self { x: 0, y: 0 }
    }

    /// Horizontal coordinate in EMUs.
    #[must_use]
    pub const fn x(self) -> i64 {
        self.x
    }

    /// Vertical coordinate in EMUs.
    #[must_use]
    pub const fn y(self) -> i64 {
        self.y
    }
}

/// Horizontal positioning base from `wp:positionH/@relativeFrom`.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum HorizontalRelativeFrom {
    Margin,
    Page,
    Column,
    Character,
    LeftMargin,
    RightMargin,
    InsideMargin,
    OutsideMargin,
}

/// Horizontal alignment from `wp:positionH/wp:align`.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum HorizontalAlignment {
    Left,
    Right,
    Center,
    Inside,
    Outside,
}

/// Vertical positioning base from `wp:positionV/@relativeFrom`.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum VerticalRelativeFrom {
    Margin,
    Page,
    Paragraph,
    Line,
    TopMargin,
    BottomMargin,
    InsideMargin,
    OutsideMargin,
}

/// Vertical alignment from `wp:positionV/wp:align`.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum VerticalAlignment {
    Top,
    Bottom,
    Center,
    Inside,
    Outside,
}

/// Checked horizontal anchor position.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum HorizontalPosition {
    Align {
        relative_from: HorizontalRelativeFrom,
        alignment: HorizontalAlignment,
    },
    Offset {
        relative_from: HorizontalRelativeFrom,
        offset: i32,
    },
}

impl HorizontalPosition {
    /// Construct an alignment-based horizontal position.
    #[must_use]
    pub const fn align(
        relative_from: HorizontalRelativeFrom,
        alignment: HorizontalAlignment,
    ) -> Self {
        Self::Align {
            relative_from,
            alignment,
        }
    }

    /// Construct an offset-based horizontal position.
    #[must_use]
    pub const fn offset(relative_from: HorizontalRelativeFrom, offset: i32) -> Self {
        Self::Offset {
            relative_from,
            offset,
        }
    }
}

/// Checked vertical anchor position.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum VerticalPosition {
    Align {
        relative_from: VerticalRelativeFrom,
        alignment: VerticalAlignment,
    },
    Offset {
        relative_from: VerticalRelativeFrom,
        offset: i32,
    },
}

impl VerticalPosition {
    /// Construct an alignment-based vertical position.
    #[must_use]
    pub const fn align(relative_from: VerticalRelativeFrom, alignment: VerticalAlignment) -> Self {
        Self::Align {
            relative_from,
            alignment,
        }
    }

    /// Construct an offset-based vertical position.
    #[must_use]
    pub const fn offset(relative_from: VerticalRelativeFrom, offset: i32) -> Self {
        Self::Offset {
            relative_from,
            offset,
        }
    }
}

/// The bounded WordprocessingML wrapping subset emitted for a floating host.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum WrapMode {
    /// `wp:wrapNone`, which has no additional required children.
    None,
}

/// All required values for a legal `wp:anchor`.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct AnchorGeometry {
    extent: Geometry,
    simple_position: Point,
    horizontal: HorizontalPosition,
    vertical: VerticalPosition,
    wrap: WrapMode,
    relative_height: u32,
    behind_document: bool,
    locked: bool,
    layout_in_cell: bool,
    allow_overlap: bool,
}

impl AnchorGeometry {
    /// Construct an anchor with schema-complete, deterministic defaults.
    ///
    /// The defaults use page-relative left/top alignment, `wrapNone`, and the
    /// required boolean attributes. Callers can replace each placement value
    /// with the consuming setters below.
    #[must_use]
    pub const fn new(extent: Geometry) -> Self {
        Self {
            extent,
            simple_position: Point::origin(),
            horizontal: HorizontalPosition::Align {
                relative_from: HorizontalRelativeFrom::Page,
                alignment: HorizontalAlignment::Left,
            },
            vertical: VerticalPosition::Align {
                relative_from: VerticalRelativeFrom::Page,
                alignment: VerticalAlignment::Top,
            },
            wrap: WrapMode::None,
            relative_height: 0,
            behind_document: false,
            locked: false,
            layout_in_cell: true,
            allow_overlap: true,
        }
    }

    /// Replace the `wp:simplePos` point.
    #[must_use]
    pub const fn with_simple_position(mut self, point: Point) -> Self {
        self.simple_position = point;
        self
    }

    /// Replace the required horizontal positioning element.
    #[must_use]
    pub const fn with_horizontal(mut self, position: HorizontalPosition) -> Self {
        self.horizontal = position;
        self
    }

    /// Replace the required vertical positioning element.
    #[must_use]
    pub const fn with_vertical(mut self, position: VerticalPosition) -> Self {
        self.vertical = position;
        self
    }

    /// Set the required relative height.
    #[must_use]
    pub const fn with_relative_height(mut self, value: u32) -> Self {
        self.relative_height = value;
        self
    }

    /// Set the required `behindDoc` flag.
    #[must_use]
    pub const fn with_behind_document(mut self, value: bool) -> Self {
        self.behind_document = value;
        self
    }

    /// Set the required `locked` flag.
    #[must_use]
    pub const fn with_locked(mut self, value: bool) -> Self {
        self.locked = value;
        self
    }

    /// Set the required `layoutInCell` flag.
    #[must_use]
    pub const fn with_layout_in_cell(mut self, value: bool) -> Self {
        self.layout_in_cell = value;
        self
    }

    /// Set the required `allowOverlap` flag.
    #[must_use]
    pub const fn with_allow_overlap(mut self, value: bool) -> Self {
        self.allow_overlap = value;
        self
    }

    const fn extent(self) -> Geometry {
        self.extent
    }

    pub(crate) const fn durable_extent(self) -> Geometry {
        self.extent
    }

    pub(crate) const fn durable_simple_position(self) -> Point {
        self.simple_position
    }

    pub(crate) const fn durable_horizontal(self) -> HorizontalPosition {
        self.horizontal
    }

    pub(crate) const fn durable_vertical(self) -> VerticalPosition {
        self.vertical
    }

    pub(crate) const fn durable_wrap(self) -> WrapMode {
        self.wrap
    }

    pub(crate) const fn durable_relative_height(self) -> u32 {
        self.relative_height
    }

    pub(crate) const fn durable_behind_document(self) -> bool {
        self.behind_document
    }

    pub(crate) const fn durable_locked(self) -> bool {
        self.locked
    }

    pub(crate) const fn durable_layout_in_cell(self) -> bool {
        self.layout_in_cell
    }

    pub(crate) const fn durable_allow_overlap(self) -> bool {
        self.allow_overlap
    }
}

/// Placement shape shared by Drawing Ink, Canvas, and Group hosts.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum Placement {
    /// A `wp:inline` placement.
    Inline(Geometry),
    /// A schema-complete `wp:anchor` placement.
    ///
    /// The generated legacy VML fallback carries the horizontal and vertical
    /// position modes in its Office style properties.  Word-specific anchor
    /// flags such as `behindDoc` and `allowOverlap` have no portable VML
    /// equivalent and are intentionally not projected into the fallback.
    Anchor(AnchorGeometry),
}

impl Placement {
    /// Construct an inline placement.
    #[must_use]
    pub const fn inline(extent: Geometry) -> Self {
        Self::Inline(extent)
    }

    /// Construct an anchored placement.
    #[must_use]
    pub const fn anchor(geometry: AnchorGeometry) -> Self {
        Self::Anchor(geometry)
    }

    const fn extent(self) -> Geometry {
        match self {
            Self::Inline(extent) => extent,
            Self::Anchor(geometry) => geometry.extent(),
        }
    }
}

/// Detached host style and target profile.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum Style {
    /// A direct WordprocessingML `w:contentPart`.
    Base(BaseProfile),
    /// A Wordprocessing Ink DrawingML host.
    Drawing {
        placement: Placement,
        fallback: FallbackImage,
    },
    /// A one-slot Wordprocessing Canvas content part.
    Canvas {
        placement: Placement,
        fallback: FallbackImage,
    },
    /// A one-slot Wordprocessing Group content part.
    Group {
        placement: Placement,
        fallback: FallbackImage,
    },
}

impl Style {
    /// Construct a direct-base style.
    #[must_use]
    pub const fn base(profile: BaseProfile) -> Self {
        Self::Base(profile)
    }

    /// Construct a Drawing Ink style with its caller-supplied fallback.
    #[must_use]
    pub fn drawing(placement: Placement, fallback: FallbackImage) -> Self {
        Self::Drawing {
            placement,
            fallback,
        }
    }

    /// Construct a one-slot Canvas style with its complete fallback image.
    #[must_use]
    pub fn canvas(placement: Placement, fallback: FallbackImage) -> Self {
        Self::Canvas {
            placement,
            fallback,
        }
    }

    /// Construct a one-slot Group style with its complete fallback image.
    #[must_use]
    pub fn group(placement: Placement, fallback: FallbackImage) -> Self {
        Self::Group {
            placement,
            fallback,
        }
    }

    /// Return the detached placement, when this style has one.
    #[must_use]
    pub fn placement(&self) -> Option<&Placement> {
        match self {
            Self::Base(_) => None,
            Self::Drawing { placement, .. }
            | Self::Canvas { placement, .. }
            | Self::Group { placement, .. } => Some(placement),
        }
    }

    /// Return the target content type selected by this style.
    pub(crate) const fn content_type(&self) -> &'static str {
        match self {
            Self::Base(profile) => profile.content_type(),
            Self::Drawing { .. } | Self::Canvas { .. } | Self::Group { .. } => {
                "application/inkml+xml"
            },
        }
    }

    /// Return the complete caller-supplied fallback, when this style needs it.
    pub(crate) fn fallback(&self) -> Option<&FallbackImage> {
        match self {
            Self::Base(_) => None,
            Self::Drawing { fallback, .. }
            | Self::Canvas { fallback, .. }
            | Self::Group { fallback, .. } => Some(fallback),
        }
    }

    pub(crate) const fn durable_kind(&self) -> u8 {
        match self {
            Self::Base(_) => 0,
            Self::Drawing { .. } => 1,
            Self::Canvas { .. } => 2,
            Self::Group { .. } => 3,
        }
    }

    pub(crate) const fn base_profile(&self) -> Option<BaseProfile> {
        match self {
            Self::Base(profile) => Some(*profile),
            Self::Drawing { .. } | Self::Canvas { .. } | Self::Group { .. } => None,
        }
    }

    const fn form(&self) -> Form {
        match self {
            Self::Base(_) => Form::Base,
            Self::Drawing { .. } => Form::Drawing,
            Self::Canvas { .. } | Self::Group { .. } => Form::GenericDrawing,
        }
    }
}

/// Render one complete host fragment suitable for insertion under `w:r`.
///
/// The Ink relationship is always required.  The image relationship is
/// required for Drawing, Canvas, and Group styles and ignored for Base.  A
/// Strict story can author Base hosts, but Word 2010 extension hosts remain a
/// typed refusal because the local conformance policy has no Strict alias.
pub(crate) fn render(
    style: &Style,
    dialect: StoryDialect,
    ink_relationship_id: &str,
    image_relationship_id: Option<&str>,
    unique_drawing_id: u32,
) -> Result<Vec<u8>> {
    validate_relationship_id(ink_relationship_id)?;
    if matches!(dialect, StoryDialect::Strict) && !matches!(style, Style::Base(_)) {
        return Err(strict_extension_refusal());
    }
    if !matches!(style, Style::Base(_)) {
        validate_drawing_id(unique_drawing_id)?;
        validate_relationship_id(image_relationship_id.ok_or_else(|| {
            invalid("DOCX Ink Drawing, Canvas, or Group style requires an image relationship")
        })?)?;
    }

    let mut output = Output::new(4 * 1024)?;
    match style {
        Style::Base(_) => {
            output.append(b"<w:contentPart xmlns:w=\"")?;
            output.append_str(word_namespace(dialect))?;
            output.append(b"\" xmlns:r=\"")?;
            output.append_str(relationship_namespace(dialect))?;
            output.append(b"\" r:id=\"")?;
            output.append_escaped(ink_relationship_id)?;
            output.append(b"\"/>")?;
        },
        Style::Drawing {
            placement,
            fallback: _,
        }
        | Style::Canvas {
            placement,
            fallback: _,
        }
        | Style::Group {
            placement,
            fallback: _,
        } => {
            let Some(image_relationship_id) = image_relationship_id else {
                return Err(invalid(
                    "DOCX Ink extension host is missing an image relationship",
                ));
            };
            let (choice_prefix, choice_uri) = match style {
                Style::Drawing { .. } => ("wpi", WORDPROCESSING_INK),
                Style::Canvas { .. } => ("wpc", WORDPROCESSING_CANVAS),
                Style::Group { .. } => ("wpg", WORDPROCESSING_GROUP),
                Style::Base(_) => unreachable!("base handled above"),
            };
            append_alternate_content_start(&mut output, dialect, choice_prefix, choice_uri)?;
            output.append(b"<mc:Choice Requires=\"")?;
            output.append(choice_prefix.as_bytes())?;
            output.append(b"\"><w:drawing>")?;
            append_placement(
                &mut output,
                placement,
                choice_uri,
                style,
                ink_relationship_id,
                unique_drawing_id,
            )?;
            output.append(b"</w:drawing></mc:Choice><mc:Fallback>")?;
            append_fallback_contents(
                &mut output,
                image_relationship_id,
                unique_drawing_id,
                placement.extent(),
                *placement,
            )?;
            output.append(b"</mc:Fallback></mc:AlternateContent>")?;
        },
    }
    let bytes = output.finish();
    validate_host_readback(&bytes, dialect, style.form())?;
    Ok(bytes)
}

/// Render a complete, self-contained MCE fallback for replacement.
#[cfg(test)]
fn render_fallback(
    image: &FallbackImage,
    dialect: StoryDialect,
    image_relationship_id: &str,
    unique_shape_id: u32,
) -> Result<Vec<u8>> {
    validate_relationship_id(image_relationship_id)?;
    validate_drawing_id(unique_shape_id)?;
    let mut output = Output::new(2 * 1024)?;
    output.append(b"<mc:Fallback xmlns:mc=\"")?;
    output.append_str(MARKUP_COMPATIBILITY)?;
    output.append(b"\" xmlns:w=\"")?;
    output.append_str(word_namespace(dialect))?;
    output.append(b"\" xmlns:r=\"")?;
    output.append_str(relationship_namespace(dialect))?;
    output.append(b"\" xmlns:v=\"")?;
    output.append_str(VML_NAMESPACE)?;
    output.append(b"\" xmlns:o=\"")?;
    output.append_str(OFFICE_NAMESPACE)?;
    output.append(b"\">")?;
    let geometry = image_geometry(image)?;
    append_fallback_contents(
        &mut output,
        image_relationship_id,
        unique_shape_id,
        geometry,
        Placement::Inline(geometry),
    )?;
    output.append(b"</mc:Fallback>")?;
    Ok(output.finish())
}

fn append_alternate_content_start(
    output: &mut Output,
    dialect: StoryDialect,
    choice_prefix: &str,
    choice_uri: &str,
) -> Result<()> {
    output.append(b"<mc:AlternateContent xmlns:mc=\"")?;
    output.append_str(MARKUP_COMPATIBILITY)?;
    output.append(b"\" xmlns:w=\"")?;
    output.append_str(word_namespace(dialect))?;
    output.append(b"\" xmlns:wp=\"")?;
    output.append_str(wordprocessing_drawing_namespace(dialect))?;
    output.append(b"\" xmlns:a=\"")?;
    output.append_str(drawingml_namespace(dialect))?;
    output.append(b"\" xmlns:w14=\"")?;
    output.append_str(WORD_2010_WORDML)?;
    output.append(b"\" xmlns:r=\"")?;
    // The Word 2010 extension relationship is specified only in the
    // Transitional namespace by the local ODRAWXML profile.
    output.append_str(TRANSITIONAL_RELATIONSHIPS)?;
    output.append(b"\" xmlns:v=\"")?;
    output.append_str(VML_NAMESPACE)?;
    output.append(b"\" xmlns:o=\"")?;
    output.append_str(OFFICE_NAMESPACE)?;
    output.append(b"\" xmlns:")?;
    output.append(choice_prefix.as_bytes())?;
    output.append(b"=\"")?;
    output.append_str(choice_uri)?;
    output.append(b"\">")?;
    Ok(())
}

fn append_placement(
    output: &mut Output,
    placement: &Placement,
    graphic_data_uri: &str,
    style: &Style,
    ink_relationship_id: &str,
    unique_drawing_id: u32,
) -> Result<()> {
    match placement {
        Placement::Inline(extent) => {
            output.append(b"<wp:inline><wp:extent cx=\"")?;
            output.append_u64(extent.width_emu())?;
            output.append(b"\" cy=\"")?;
            output.append_u64(extent.height_emu())?;
            output.append(b"\"/><wp:docPr id=\"")?;
            output.append_u64(u64::from(unique_drawing_id))?;
            output.append(b"\" name=\"litchiInk")?;
            output.append_u64(u64::from(unique_drawing_id))?;
            output.append(b"\"/>")?;
            append_graphic(
                output,
                graphic_data_uri,
                *extent,
                style,
                ink_relationship_id,
            )?;
            output.append(b"</wp:inline>")?;
        },
        Placement::Anchor(geometry) => {
            append_anchor_start(output, geometry, unique_drawing_id)?;
            append_graphic(
                output,
                graphic_data_uri,
                geometry.extent(),
                style,
                ink_relationship_id,
            )?;
            output.append(b"</wp:anchor>")?;
        },
    }
    Ok(())
}

fn append_graphic(
    output: &mut Output,
    graphic_data_uri: &str,
    extent: Geometry,
    style: &Style,
    ink_relationship_id: &str,
) -> Result<()> {
    output.append(b"<a:graphic><a:graphicData uri=\"")?;
    output.append_escaped(graphic_data_uri)?;
    output.append(b"\">")?;
    append_graphic_data_child(output, extent, style, ink_relationship_id)?;
    output.append(b"</a:graphicData></a:graphic>")?;
    Ok(())
}

fn append_anchor_start(
    output: &mut Output,
    geometry: &AnchorGeometry,
    unique_drawing_id: u32,
) -> Result<()> {
    output.append(b"<wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" relativeHeight=\"")?;
    output.append_u64(u64::from(geometry.relative_height))?;
    output.append(b"\" behindDoc=\"")?;
    output.append_bool(geometry.behind_document)?;
    output.append(b"\" locked=\"")?;
    output.append_bool(geometry.locked)?;
    output.append(b"\" layoutInCell=\"")?;
    output.append_bool(geometry.layout_in_cell)?;
    output.append(b"\" allowOverlap=\"")?;
    output.append_bool(geometry.allow_overlap)?;
    output.append(b"\"><wp:simplePos x=\"")?;
    output.append_i64(geometry.simple_position.x())?;
    output.append(b"\" y=\"")?;
    output.append_i64(geometry.simple_position.y())?;
    output.append(b"\"/>")?;
    append_horizontal(output, geometry.horizontal)?;
    append_vertical(output, geometry.vertical)?;
    output.append(b"<wp:extent cx=\"")?;
    output.append_u64(geometry.extent.width_emu())?;
    output.append(b"\" cy=\"")?;
    output.append_u64(geometry.extent.height_emu())?;
    output.append(b"\"/>")?;
    match geometry.wrap {
        WrapMode::None => output.append(b"<wp:wrapNone/>")?,
    }
    output.append(b"<wp:docPr id=\"")?;
    output.append_u64(u64::from(unique_drawing_id))?;
    output.append(b"\" name=\"litchiInk")?;
    output.append_u64(u64::from(unique_drawing_id))?;
    output.append(b"\"/>")?;
    Ok(())
}

fn append_horizontal(output: &mut Output, position: HorizontalPosition) -> Result<()> {
    match position {
        HorizontalPosition::Align {
            relative_from,
            alignment,
        } => {
            output.append(b"<wp:positionH relativeFrom=\"")?;
            output.append(relative_from.lexical().as_bytes())?;
            output.append(b"\"><wp:align>")?;
            output.append(alignment.lexical().as_bytes())?;
            output.append(b"</wp:align></wp:positionH>")?;
        },
        HorizontalPosition::Offset {
            relative_from,
            offset,
        } => {
            output.append(b"<wp:positionH relativeFrom=\"")?;
            output.append(relative_from.lexical().as_bytes())?;
            output.append(b"\"><wp:posOffset>")?;
            output.append_i64(i64::from(offset))?;
            output.append(b"</wp:posOffset></wp:positionH>")?;
        },
    }
    Ok(())
}

fn append_vertical(output: &mut Output, position: VerticalPosition) -> Result<()> {
    match position {
        VerticalPosition::Align {
            relative_from,
            alignment,
        } => {
            output.append(b"<wp:positionV relativeFrom=\"")?;
            output.append(relative_from.lexical().as_bytes())?;
            output.append(b"\"><wp:align>")?;
            output.append(alignment.lexical().as_bytes())?;
            output.append(b"</wp:align></wp:positionV>")?;
        },
        VerticalPosition::Offset {
            relative_from,
            offset,
        } => {
            output.append(b"<wp:positionV relativeFrom=\"")?;
            output.append(relative_from.lexical().as_bytes())?;
            output.append(b"\"><wp:posOffset>")?;
            output.append_i64(i64::from(offset))?;
            output.append(b"</wp:posOffset></wp:positionV>")?;
        },
    }
    Ok(())
}

fn append_graphic_data_child(
    output: &mut Output,
    extent: Geometry,
    style: &Style,
    ink_relationship_id: &str,
) -> Result<()> {
    match style {
        Style::Drawing { .. } => {
            output.append(b"<w14:contentPart r:id=\"")?;
            output.append_escaped(ink_relationship_id)?;
            output.append(b"\"/>")?;
        },
        Style::Canvas { .. } => {
            output.append(b"<wpc:wpc><w14:contentPart r:id=\"")?;
            output.append_escaped(ink_relationship_id)?;
            output.append(b"\"/></wpc:wpc>")?;
        },
        Style::Group { .. } => {
            output.append(b"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"")?;
            output.append_u64(extent.width_emu())?;
            output.append(b"\" cy=\"")?;
            output.append_u64(extent.height_emu())?;
            output.append(b"\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"")?;
            output.append_u64(extent.width_emu())?;
            output.append(b"\" cy=\"")?;
            output.append_u64(extent.height_emu())?;
            output.append(b"\"/></a:xfrm></wpg:grpSpPr><w14:contentPart r:id=\"")?;
            output.append_escaped(ink_relationship_id)?;
            output.append(b"\"/></wpg:wgp>")?;
        },
        Style::Base(_) => return Err(invalid("base style has no DrawingML graphic data")),
    }
    Ok(())
}

fn append_fallback_contents(
    output: &mut Output,
    image_relationship_id: &str,
    unique_shape_id: u32,
    geometry: Geometry,
    placement: Placement,
) -> Result<()> {
    output.append(b"<w:pict><v:shape id=\"litchiInk")?;
    output.append_u64(u64::from(unique_shape_id))?;
    output.append(
        b"\" coordsize=\"21600,21600\" path=\"m,l,21600r21600,l21600,xe\" \
          o:spt=\"1\" stroked=\"f\" filled=\"f\" style=\"width:",
    )?;
    append_vml_length(output, geometry.width_emu())?;
    output.append(b";height:")?;
    append_vml_length(output, geometry.height_emu())?;
    output.append(b";z-index:0")?;
    append_vml_position(output, placement)?;
    output.append(b"\"><v:imagedata r:id=\"")?;
    output.append_escaped(image_relationship_id)?;
    output.append(b"\" o:title=\"Ink\"/></v:shape></w:pict>")?;
    Ok(())
}

#[cfg(test)]
fn image_geometry(image: &FallbackImage) -> Result<Geometry> {
    const EMU_PER_PIXEL: u64 = 9_525;
    let width = u64::from(image.dimensions.width())
        .checked_mul(EMU_PER_PIXEL)
        .ok_or_else(|| invalid("DOCX Ink fallback width overflowed EMU bounds"))?;
    let height = u64::from(image.dimensions.height())
        .checked_mul(EMU_PER_PIXEL)
        .ok_or_else(|| invalid("DOCX Ink fallback height overflowed EMU bounds"))?;
    Geometry::new(width, height)
}

/// Write an EMU extent as a bounded, deterministic point length.
///
/// VML consumes physical CSS-style lengths.  One point is 12,700 EMUs, so
/// the conversion is performed entirely with integer arithmetic and carries
/// six fractional decimal places.  The output is rounded half-up at that
/// fixed precision; no binary floating-point value participates in the
/// generated XML.
fn append_vml_length(output: &mut Output, emu: u64) -> Result<()> {
    const EMU_PER_POINT: u64 = 12_700;
    const MAX_VML_POINTS: u64 = 169_093;
    const SCALE: u64 = 1_000_000;
    if emu > MAX_VML_POINTS * EMU_PER_POINT {
        return Err(invalid(
            "DOCX Ink VML length exceeds the Office point-unit bound",
        ));
    }
    let mut whole = emu / EMU_PER_POINT;
    let remainder = emu % EMU_PER_POINT;
    let mut fraction = remainder
        .checked_mul(SCALE)
        .and_then(|value| value.checked_add(EMU_PER_POINT / 2))
        .ok_or_else(|| invalid("DOCX Ink VML length conversion overflowed"))?
        / EMU_PER_POINT;
    if fraction == SCALE {
        whole = whole
            .checked_add(1)
            .ok_or_else(|| invalid("DOCX Ink VML length conversion overflowed"))?;
        fraction = 0;
    }
    output.append_u64(whole)?;
    if fraction != 0 {
        let mut digits = [b'0'; 6];
        let mut value = fraction;
        for index in (0..digits.len()).rev() {
            digits[index] = b'0' + u8::try_from(value % 10).unwrap_or(0);
            value /= 10;
        }
        let mut end = digits.len();
        while end > 0 && digits[end - 1] == b'0' {
            end -= 1;
        }
        output.append(b".")?;
        output.append(&digits[..end])?;
    }
    output.append(b"pt")?;
    Ok(())
}

fn append_vml_signed_length(output: &mut Output, emu: i32) -> Result<()> {
    if emu < 0 {
        output.append(b"-")?;
        append_vml_length(output, i64::from(emu).unsigned_abs())?;
    } else {
        append_vml_length(
            output,
            u64::try_from(emu)
                .map_err(|_| invalid("DOCX Ink VML signed length conversion overflowed"))?,
        )?;
    }
    Ok(())
}

fn append_vml_position(output: &mut Output, placement: Placement) -> Result<()> {
    let Placement::Anchor(geometry) = placement else {
        return Ok(());
    };
    output.append(b";position:absolute")?;
    match geometry.horizontal {
        HorizontalPosition::Align {
            relative_from,
            alignment,
        } => {
            output.append(b";mso-position-horizontal:")?;
            output.append(alignment.lexical().as_bytes())?;
            output.append(b";mso-position-horizontal-relative:")?;
            output.append(vml_horizontal_relative(relative_from)?.as_bytes())?;
        },
        HorizontalPosition::Offset {
            relative_from,
            offset,
        } => {
            output.append(b";margin-left:")?;
            append_vml_signed_length(output, offset)?;
            output
                .append(b";mso-position-horizontal:absolute;mso-position-horizontal-relative:")?;
            output.append(vml_horizontal_relative(relative_from)?.as_bytes())?;
        },
    }
    match geometry.vertical {
        VerticalPosition::Align {
            relative_from,
            alignment,
        } => {
            output.append(b";mso-position-vertical:")?;
            output.append(alignment.lexical().as_bytes())?;
            output.append(b";mso-position-vertical-relative:")?;
            output.append(vml_vertical_relative(relative_from).as_bytes())?;
        },
        VerticalPosition::Offset {
            relative_from,
            offset,
        } => {
            output.append(b";margin-top:")?;
            append_vml_signed_length(output, offset)?;
            output.append(b";mso-position-vertical:absolute;mso-position-vertical-relative:")?;
            output.append(vml_vertical_relative(relative_from).as_bytes())?;
        },
    }
    Ok(())
}

fn vml_horizontal_relative(value: HorizontalRelativeFrom) -> Result<&'static str> {
    match value {
        HorizontalRelativeFrom::Margin => Ok("margin"),
        HorizontalRelativeFrom::Page => Ok("page"),
        HorizontalRelativeFrom::Column => Err(invalid(
            "DOCX Ink VML fallback cannot represent column-relative horizontal placement",
        )),
        HorizontalRelativeFrom::Character => Ok("char"),
        HorizontalRelativeFrom::LeftMargin => Ok("left-margin-area"),
        HorizontalRelativeFrom::RightMargin => Ok("right-margin-area"),
        HorizontalRelativeFrom::InsideMargin => Ok("inner-margin-area"),
        HorizontalRelativeFrom::OutsideMargin => Ok("outer-margin-area"),
    }
}

const fn vml_vertical_relative(value: VerticalRelativeFrom) -> &'static str {
    match value {
        VerticalRelativeFrom::Margin => "margin",
        VerticalRelativeFrom::Page => "page",
        VerticalRelativeFrom::Paragraph => "text",
        VerticalRelativeFrom::Line => "line",
        VerticalRelativeFrom::TopMargin => "top-margin-area",
        VerticalRelativeFrom::BottomMargin => "bottom-margin-area",
        VerticalRelativeFrom::InsideMargin => "inner-margin-area",
        VerticalRelativeFrom::OutsideMargin => "outer-margin-area",
    }
}

fn validate_host_readback(bytes: &[u8], dialect: StoryDialect, expected: Form) -> Result<()> {
    let mut wrapper = Output::new(bytes.len().saturating_add(256))?;
    wrapper.append(b"<w:document xmlns:w=\"")?;
    wrapper.append_str(word_namespace(dialect))?;
    wrapper.append(b"\" xmlns:r=\"")?;
    wrapper.append_str(relationship_namespace(dialect))?;
    wrapper.append(b"\"><w:body><w:p><w:r>")?;
    wrapper.append(bytes)?;
    wrapper.append(b"</w:r></w:p></w:body></w:document>")?;
    let anchors = scan(wrapper.bytes(), dialect, MAX_HOST_NODES, MAX_HOST_DEPTH, 2)?;
    if anchors.len() != 1 || anchors[0].form != expected {
        return Err(invalid(
            "generated DOCX Ink host failed namespace-aware scanner readback",
        ));
    }
    let hosts = host::capture(wrapper.bytes(), dialect, Limits::default())?;
    if hosts.len() != 1 || !hosts[0].removable {
        return Err(invalid(
            "generated DOCX Ink host failed conservative removal readback",
        ));
    }
    Ok(())
}

fn validate_image(bytes: &[u8], media_type: ImageType, expected: ImageDimensions) -> Result<()> {
    if bytes.len() > MAX_FALLBACK_IMAGE_BYTES {
        return Err(image_limit(bytes.len()));
    }
    let actual = match media_type {
        ImageType::Png => parse_png_dimensions(bytes)?,
        ImageType::Jpeg => parse_jpeg_dimensions(bytes)?,
    };
    if actual != expected {
        return Err(invalid(
            "DOCX Ink fallback image dimensions do not match the encoded image",
        ));
    }
    Ok(())
}

fn detect_image(bytes: &[u8]) -> Result<(ImageType, ImageDimensions)> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        let dimensions = parse_png_dimensions(bytes)?;
        return Ok((ImageType::Png, dimensions));
    }
    if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
        let dimensions = parse_jpeg_dimensions(bytes)?;
        return Ok((ImageType::Jpeg, dimensions));
    }
    Err(invalid(
        "DOCX Ink fallback image is not bounded PNG or JPEG",
    ))
}

fn parse_png_dimensions(bytes: &[u8]) -> Result<ImageDimensions> {
    if bytes.len() < 33 || !bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Err(invalid("DOCX Ink PNG fallback has an invalid signature"));
    }
    let mut dimensions = None;
    let mut color_type = None;
    let mut saw_ihdr = false;
    let mut saw_plte = false;
    let mut saw_idat = false;
    let mut saw_iend = false;
    let mut idat_bytes = 0usize;
    let mut idat_header = [0u8; 2];
    let mut idat_header_bytes = 0usize;
    let mut offset = 8usize;
    while offset < bytes.len() {
        let end_header = offset
            .checked_add(8)
            .ok_or_else(|| invalid("DOCX Ink PNG chunk offset overflow"))?;
        if end_header > bytes.len() {
            return Err(invalid("DOCX Ink PNG fallback has a truncated chunk"));
        }
        let length = usize::try_from(u32::from_be_bytes(
            bytes[offset..offset + 4]
                .try_into()
                .map_err(|_| invalid("DOCX Ink PNG chunk length is truncated"))?,
        ))
        .map_err(|_| invalid("DOCX Ink PNG chunk length overflows usize"))?;
        let end = end_header
            .checked_add(length)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| invalid("DOCX Ink PNG chunk size overflow"))?;
        if end > bytes.len() {
            return Err(invalid("DOCX Ink PNG fallback has a truncated chunk"));
        }
        let chunk = &bytes[offset + 4..offset + 8];
        if !chunk.iter().all(u8::is_ascii_alphabetic) || chunk[2].is_ascii_lowercase() {
            return Err(invalid("DOCX Ink PNG fallback has an invalid chunk type"));
        }
        if chunk[0].is_ascii_uppercase() && !matches!(chunk, b"IHDR" | b"PLTE" | b"IDAT" | b"IEND")
        {
            return Err(invalid(
                "DOCX Ink PNG fallback has an unsupported critical chunk",
            ));
        }
        let data = &bytes[offset + 8..offset + 8 + length];
        let expected_crc = u32::from_be_bytes(
            bytes[end - 4..end]
                .try_into()
                .map_err(|_| invalid("DOCX Ink PNG chunk CRC is truncated"))?,
        );
        if png_crc32(chunk, data) != expected_crc {
            return Err(invalid("DOCX Ink PNG fallback has a bad chunk CRC"));
        }
        match chunk {
            b"IHDR" if !saw_ihdr && offset == 8 && length == 13 => {
                let width = u32::from_be_bytes(
                    data[0..4]
                        .try_into()
                        .map_err(|_| invalid("DOCX Ink PNG width is truncated"))?,
                );
                let height = u32::from_be_bytes(
                    data[4..8]
                        .try_into()
                        .map_err(|_| invalid("DOCX Ink PNG height is truncated"))?,
                );
                let image_dimensions = ImageDimensions::new(width, height)?;
                let bit_depth = data[8];
                let kind = data[9];
                if data[10] != 0 || data[11] != 0 || data[12] > 1 {
                    return Err(invalid(
                        "DOCX Ink PNG fallback has unsupported compression, filter, or interlace",
                    ));
                }
                let bits_per_pixel = png_bits_per_pixel(bit_depth, kind)?;
                let scanline_bits = u64::from(width)
                    .checked_mul(u64::from(bits_per_pixel))
                    .ok_or_else(|| invalid("DOCX Ink PNG scanline size overflows"))?;
                let scanline_bytes = scanline_bits
                    .checked_add(7)
                    .map(|value| value / 8 + 1)
                    .ok_or_else(|| invalid("DOCX Ink PNG scanline size overflows"))?;
                let expanded = scanline_bytes
                    .checked_mul(u64::from(height))
                    .ok_or_else(|| invalid("DOCX Ink PNG expanded size overflows"))?;
                if expanded > MAX_FALLBACK_EXPANDED_BYTES {
                    return Err(invalid(
                        "DOCX Ink PNG fallback exceeds the expanded byte bound",
                    ));
                }
                color_type = Some(kind);
                dimensions = Some(image_dimensions);
                saw_ihdr = true;
            },
            b"IHDR" => return Err(invalid("DOCX Ink PNG fallback repeats IHDR")),
            b"PLTE" => {
                if !saw_ihdr || data.is_empty() || data.len() % 3 != 0 || data.len() > 768 {
                    return Err(invalid("DOCX Ink PNG fallback has an invalid palette"));
                }
                saw_plte = true;
            },
            b"IDAT" => {
                if !saw_ihdr || data.is_empty() || (color_type == Some(3) && !saw_plte) {
                    return Err(invalid("DOCX Ink PNG fallback has an empty IDAT"));
                }
                saw_idat = true;
                idat_bytes = idat_bytes
                    .checked_add(data.len())
                    .ok_or_else(|| invalid("DOCX Ink PNG IDAT size overflows"))?;
                for byte in data {
                    if idat_header_bytes < idat_header.len() {
                        idat_header[idat_header_bytes] = *byte;
                        idat_header_bytes += 1;
                    }
                }
            },
            b"IEND" if data.is_empty() => {
                saw_iend = true;
                if end != bytes.len() {
                    return Err(invalid("DOCX Ink PNG fallback has data after IEND"));
                }
                break;
            },
            b"IEND" => return Err(invalid("DOCX Ink PNG IEND has data")),
            _ => {},
        }
        offset = end;
    }
    let dimensions = dimensions.ok_or_else(|| invalid("DOCX Ink PNG fallback has no IHDR"))?;
    if color_type == Some(3) && !saw_plte {
        return Err(invalid("DOCX Ink indexed PNG fallback has no palette"));
    }
    if !saw_idat || idat_bytes < 6 || idat_header_bytes != 2 {
        return Err(invalid("DOCX Ink PNG fallback has no complete IDAT stream"));
    }
    let cmf = idat_header[0];
    let flg = idat_header[1];
    if cmf & 0x0f != 8
        || cmf >> 4 > 7
        || (u16::from(cmf) << 8 | u16::from(flg)) % 31 != 0
        || flg & 0x20 != 0
    {
        return Err(invalid(
            "DOCX Ink PNG fallback has an invalid zlib stream header",
        ));
    }
    if !saw_ihdr || !saw_iend {
        return Err(invalid("DOCX Ink PNG fallback is missing IEND"));
    }
    Ok(dimensions)
}

fn parse_jpeg_dimensions(bytes: &[u8]) -> Result<ImageDimensions> {
    if bytes.len() < 4 || bytes[0..2] != [0xff, 0xd8] {
        return Err(invalid("DOCX Ink JPEG fallback has an invalid SOI"));
    }
    let mut offset = 2usize;
    let mut dimensions = None;
    while offset < bytes.len() {
        if bytes[offset] != 0xff {
            return Err(invalid("DOCX Ink JPEG fallback has data outside a marker"));
        }
        while offset < bytes.len() && bytes[offset] == 0xff {
            offset += 1;
        }
        if offset >= bytes.len() {
            break;
        }
        let marker = bytes[offset];
        offset += 1;
        if marker == 0x00 || marker == 0xd8 {
            return Err(invalid("DOCX Ink JPEG fallback has an invalid marker"));
        }
        if marker == 0xd9 {
            if offset != bytes.len() {
                return Err(invalid("DOCX Ink JPEG fallback has data after EOI"));
            }
            return dimensions
                .ok_or_else(|| invalid("DOCX Ink JPEG fallback has no supported SOF dimensions"));
        }
        if marker == 0xda {
            let length_end = offset
                .checked_add(2)
                .ok_or_else(|| invalid("DOCX Ink JPEG SOS offset overflow"))?;
            if length_end > bytes.len() {
                return Err(invalid("DOCX Ink JPEG SOS length is truncated"));
            }
            let length = usize::from(u16::from_be_bytes(
                bytes[offset..length_end]
                    .try_into()
                    .map_err(|_| invalid("DOCX Ink JPEG SOS length is truncated"))?,
            ));
            if length < 2 {
                return Err(invalid("DOCX Ink JPEG SOS segment is too short"));
            }
            let segment_end = offset
                .checked_add(length)
                .ok_or_else(|| invalid("DOCX Ink JPEG SOS size overflows"))?;
            if segment_end > bytes.len() {
                return Err(invalid("DOCX Ink JPEG SOS segment is truncated"));
            }
            offset = segment_end;
            let mut entropy = offset;
            let mut next_marker = None;
            while entropy < bytes.len() {
                if bytes[entropy] != 0xff {
                    entropy += 1;
                    continue;
                }
                let marker_start = entropy;
                while entropy < bytes.len() && bytes[entropy] == 0xff {
                    entropy += 1;
                }
                if entropy >= bytes.len() {
                    return Err(invalid("DOCX Ink JPEG fallback is missing EOI"));
                }
                let code = bytes[entropy];
                entropy += 1;
                if code == 0x00 || (0xd0..=0xd7).contains(&code) {
                    continue;
                }
                if code == 0xd9 {
                    if entropy != bytes.len() {
                        return Err(invalid("DOCX Ink JPEG fallback has data after EOI"));
                    }
                    return dimensions.ok_or_else(|| {
                        invalid("DOCX Ink JPEG fallback has no supported SOF dimensions")
                    });
                }
                next_marker = Some(marker_start);
                break;
            }
            offset = next_marker.unwrap_or(entropy).min(bytes.len());
            continue;
        }
        if marker == 0x01 || (0xd0..=0xd7).contains(&marker) {
            continue;
        }
        let length_end = offset
            .checked_add(2)
            .ok_or_else(|| invalid("DOCX Ink JPEG segment offset overflow"))?;
        if length_end > bytes.len() {
            return Err(invalid("DOCX Ink JPEG fallback has a truncated segment"));
        }
        let length = usize::from(u16::from_be_bytes(
            bytes[offset..length_end]
                .try_into()
                .map_err(|_| invalid("DOCX Ink JPEG segment length is truncated"))?,
        ));
        if length < 2 {
            return Err(invalid(
                "DOCX Ink JPEG fallback has an invalid segment length",
            ));
        }
        let segment_end = offset
            .checked_add(length)
            .ok_or_else(|| invalid("DOCX Ink JPEG segment size overflow"))?;
        if segment_end > bytes.len() {
            return Err(invalid("DOCX Ink JPEG fallback has a truncated segment"));
        }
        if is_jpeg_sof(marker) {
            if length < 8 || dimensions.is_some() {
                return Err(invalid("DOCX Ink JPEG SOF segment is invalid"));
            }
            let height = u32::from(u16::from_be_bytes(
                bytes[offset + 3..offset + 5]
                    .try_into()
                    .map_err(|_| invalid("DOCX Ink JPEG height is truncated"))?,
            ));
            let width = u32::from(u16::from_be_bytes(
                bytes[offset + 5..offset + 7]
                    .try_into()
                    .map_err(|_| invalid("DOCX Ink JPEG width is truncated"))?,
            ));
            dimensions = Some(ImageDimensions::new(width, height)?);
        }
        offset = segment_end;
    }
    Err(invalid(
        "DOCX Ink JPEG fallback is missing EOI or supported SOF dimensions",
    ))
}

fn png_bits_per_pixel(bit_depth: u8, color_type: u8) -> Result<u8> {
    let valid = match color_type {
        0 => matches!(bit_depth, 1 | 2 | 4 | 8 | 16),
        2 => matches!(bit_depth, 8 | 16),
        3 => matches!(bit_depth, 1 | 2 | 4 | 8),
        4 => matches!(bit_depth, 8 | 16),
        6 => matches!(bit_depth, 8 | 16),
        _ => false,
    };
    if !valid {
        return Err(invalid("DOCX Ink PNG fallback has an invalid color depth"));
    }
    Ok(match color_type {
        0 => bit_depth,
        2 => bit_depth.saturating_mul(3),
        3 => bit_depth,
        4 => bit_depth.saturating_mul(2),
        6 => bit_depth.saturating_mul(4),
        _ => unreachable!("color type validated above"),
    })
}

fn png_crc32(chunk: &[u8], data: &[u8]) -> u32 {
    let mut crc = 0xffff_ffffu32;
    for byte in chunk.iter().chain(data) {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            crc = if crc & 1 != 0 {
                (crc >> 1) ^ 0xedb8_8320
            } else {
                crc >> 1
            };
        }
    }
    !crc
}

fn is_jpeg_sof(marker: u8) -> bool {
    matches!(marker, 0xc0..=0xc3 | 0xc5..=0xc7 | 0xc9..=0xcb | 0xcd..=0xcf)
}

fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid("DOCX Ink relationship ID is empty or oversized"));
    }
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(invalid("DOCX Ink relationship ID is empty"));
    };
    if !(first == '_' || first.is_ascii_alphabetic())
        || chars.any(|character| {
            !(character == '_'
                || character == '-'
                || character == '.'
                || character.is_ascii_alphanumeric())
        })
    {
        return Err(invalid("DOCX Ink relationship ID is not an XML ID token"));
    }
    Ok(())
}

fn validate_drawing_id(value: u32) -> Result<()> {
    if value == 0 {
        return Err(invalid("DOCX Ink drawing identifier must be positive"));
    }
    Ok(())
}

fn word_namespace(dialect: StoryDialect) -> &'static str {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_WORD,
        StoryDialect::Strict => STRICT_WORD,
    }
}

fn relationship_namespace(dialect: StoryDialect) -> &'static str {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_RELATIONSHIPS,
        StoryDialect::Strict => STRICT_RELATIONSHIPS,
    }
}

fn drawingml_namespace(dialect: StoryDialect) -> &'static str {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_DRAWINGML,
        StoryDialect::Strict => STRICT_DRAWINGML,
    }
}

fn wordprocessing_drawing_namespace(dialect: StoryDialect) -> &'static str {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_WORDPROCESSING_DRAWING,
        StoryDialect::Strict => STRICT_WORDPROCESSING_DRAWING,
    }
}

fn strict_extension_refusal() -> Error {
    Error::UnsafeEdit {
        format: "DOCX",
        operation: "author_ink_host",
        reason: "Strict WordprocessingML has no accepted Word 2010 Ink extension relationship policy",
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn image_limit(actual: usize) -> Error {
    Error::InkLimit {
        resource: "fallback image bytes",
        actual,
        maximum: MAX_FALLBACK_IMAGE_BYTES,
    }
}

struct Output {
    bytes: Vec<u8>,
}

impl Output {
    fn new(capacity: usize) -> Result<Self> {
        let capacity = capacity.min(MAX_HOST_BYTES);
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink host output",
                source,
            })?;
        Ok(Self { bytes })
    }

    fn append(&mut self, value: &[u8]) -> Result<()> {
        let new_len = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid("DOCX Ink host output size overflow"))?;
        if new_len > MAX_HOST_BYTES {
            return Err(Error::InkLimit {
                resource: "host output bytes",
                actual: new_len,
                maximum: MAX_HOST_BYTES,
            });
        }
        self.bytes
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink host output",
                source,
            })?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn append_str(&mut self, value: &str) -> Result<()> {
        self.append(value.as_bytes())
    }

    fn append_escaped(&mut self, value: &str) -> Result<()> {
        let escaped = escape_xml(value);
        self.append(escaped.as_bytes())
    }

    fn append_u64(&mut self, value: u64) -> Result<()> {
        let mut text = String::new();
        text.try_reserve(20).map_err(|source| Error::Allocation {
            resource: "DOCX Ink numeric host value",
            source,
        })?;
        write!(&mut text, "{value}").map_err(|error| Error::Xml(error.to_string()))?;
        self.append_str(&text)
    }

    fn append_i64(&mut self, value: i64) -> Result<()> {
        let mut text = String::new();
        text.try_reserve(21).map_err(|source| Error::Allocation {
            resource: "DOCX Ink numeric host value",
            source,
        })?;
        write!(&mut text, "{value}").map_err(|error| Error::Xml(error.to_string()))?;
        self.append_str(&text)
    }

    fn append_bool(&mut self, value: bool) -> Result<()> {
        self.append(if value { b"1" } else { b"0" })
    }

    fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

impl HorizontalRelativeFrom {
    const fn lexical(self) -> &'static str {
        match self {
            Self::Margin => "margin",
            Self::Page => "page",
            Self::Column => "column",
            Self::Character => "character",
            Self::LeftMargin => "leftMargin",
            Self::RightMargin => "rightMargin",
            Self::InsideMargin => "insideMargin",
            Self::OutsideMargin => "outsideMargin",
        }
    }
}

impl HorizontalAlignment {
    const fn lexical(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Right => "right",
            Self::Center => "center",
            Self::Inside => "inside",
            Self::Outside => "outside",
        }
    }
}

impl VerticalRelativeFrom {
    const fn lexical(self) -> &'static str {
        match self {
            Self::Margin => "margin",
            Self::Page => "page",
            Self::Paragraph => "paragraph",
            Self::Line => "line",
            Self::TopMargin => "topMargin",
            Self::BottomMargin => "bottomMargin",
            Self::InsideMargin => "insideMargin",
            Self::OutsideMargin => "outsideMargin",
        }
    }
}

impl VerticalAlignment {
    const fn lexical(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Bottom => "bottom",
            Self::Center => "center",
            Self::Inside => "inside",
            Self::Outside => "outside",
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "authoring tests use direct fixture assertions"
)]
mod tests {
    use super::*;

    const PNG_1X1: &[u8] = b"\x89PNG\r\n\x1a\n\x00\x00\x00\x0dIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\x0dIDATx\x9cc\xf8\xcf\xc0\xf0\x1f\x00\x05\x00\x01\xff\x89\x99=\x1d\x00\x00\x00\x00IEND\xaeB`\x82";
    const PNG_WITHOUT_IDAT: &[u8] = b"\x89PNG\r\n\x1a\n\x00\x00\x00\x0dIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\x00IEND\xaeB`\x82";

    fn image() -> FallbackImage {
        FallbackImage::new(
            PNG_1X1.to_vec(),
            ImageType::Png,
            ImageDimensions::new(1, 1).expect("dimensions"),
        )
        .expect("PNG")
    }

    fn extent() -> Geometry {
        Geometry::new(127_000, 127_000).expect("extent")
    }

    fn wrapper(host: &[u8], dialect: StoryDialect) -> Vec<u8> {
        let mut output = Vec::new();
        output.extend_from_slice(b"<w:document xmlns:w=\"");
        output.extend_from_slice(word_namespace(dialect).as_bytes());
        output.extend_from_slice(b"\"><w:body><w:p><w:r>");
        output.extend_from_slice(host);
        output.extend_from_slice(b"</w:r></w:p></w:body></w:document>");
        output
    }

    fn count_bytes(haystack: &[u8], needle: &[u8]) -> usize {
        if needle.is_empty() {
            return 0;
        }
        haystack
            .windows(needle.len())
            .filter(|window| *window == needle)
            .count()
    }

    #[test]
    fn base_profiles_render_and_read_back_in_both_dialects() {
        for dialect in [StoryDialect::Transitional, StoryDialect::Strict] {
            for profile in [BaseProfile::WordTextXml, BaseProfile::InkContent] {
                let style = Style::base(profile);
                let host = render(&style, dialect, "rIdInk", None, 0).expect("base host");
                let anchors =
                    scan(&wrapper(&host, dialect), dialect, 128, 32, 2).expect("base scanner");
                assert_eq!(anchors.len(), 1);
                assert_eq!(anchors[0].form, Form::Base);
                assert!(style.fallback().is_none());
                assert_eq!(style.content_type(), profile.content_type());
            }
        }
    }

    #[test]
    fn drawing_inline_and_anchor_render_complete_mce_hosts() {
        for placement in [
            Placement::inline(extent()),
            Placement::anchor(AnchorGeometry::new(extent())),
        ] {
            let style = Style::drawing(placement, image());
            let host = render(
                &style,
                StoryDialect::Transitional,
                "rIdInk",
                Some("rIdImage"),
                42,
            )
            .expect("drawing host");
            let anchors = scan(
                &wrapper(&host, StoryDialect::Transitional),
                StoryDialect::Transitional,
                512,
                64,
                2,
            )
            .expect("drawing scanner");
            assert_eq!(
                anchors,
                vec![super::super::codec::Anchor {
                    relationship_id: "rIdInk".into(),
                    form: Form::Drawing,
                }]
            );
            assert_eq!(count_bytes(&host, b"<mc:Choice Requires=\"wpi\">"), 1);
            assert!(count_bytes(&host, b"<wp:wrapNone/>") <= 1);
            assert_eq!(count_bytes(&host, b"<mc:Fallback>"), 1);
            assert_eq!(count_bytes(&host, b"</mc:Fallback>"), 1);
            assert_eq!(count_bytes(&host, b"coordsize=\"21600,21600\""), 1);
            assert_eq!(count_bytes(&host, b"path=\"m,l,21600r21600,l21600,xe\""), 1);
            assert_eq!(count_bytes(&host, b"width:10pt;height:10pt"), 1);
            let captured = host::capture(
                &wrapper(&host, StoryDialect::Transitional),
                StoryDialect::Transitional,
                Limits::default(),
            )
            .expect("drawing host capture");
            assert_eq!(captured.len(), 1);
            assert!(captured[0].removable);
        }
    }

    #[test]
    fn canvas_and_group_render_one_schema_valid_content_slot() {
        for style in [
            Style::canvas(Placement::inline(extent()), image()),
            Style::group(Placement::anchor(AnchorGeometry::new(extent())), image()),
        ] {
            let host = render(
                &style,
                StoryDialect::Transitional,
                "inkId",
                Some("imageId"),
                9,
            )
            .expect("generic host");
            let anchors = scan(
                &wrapper(&host, StoryDialect::Transitional),
                StoryDialect::Transitional,
                1024,
                64,
                2,
            )
            .expect("generic scanner");
            assert_eq!(anchors.len(), 1);
            assert_eq!(anchors[0].relationship_id, "inkId");
            assert_eq!(anchors[0].form, Form::GenericDrawing);
        }
    }

    #[test]
    fn fallback_uses_placement_extent_and_anchor_position_modes() {
        let inch = Geometry::new(914_400, 914_400).expect("inch");
        let anchor = AnchorGeometry::new(inch)
            .with_horizontal(HorizontalPosition::offset(
                HorizontalRelativeFrom::Margin,
                127_000,
            ))
            .with_vertical(VerticalPosition::offset(
                VerticalRelativeFrom::Margin,
                -63_500,
            ));
        let host = render(
            &Style::drawing(Placement::anchor(anchor), image()),
            StoryDialect::Transitional,
            "inkId",
            Some("imageId"),
            10,
        )
        .expect("anchored host");
        assert_eq!(count_bytes(&host, b"width:72pt;height:72pt"), 1);
        assert_eq!(count_bytes(&host, b"margin-left:10pt"), 1);
        assert_eq!(count_bytes(&host, b"margin-top:-5pt"), 1);
        assert_eq!(
            count_bytes(
                &host,
                b"mso-position-horizontal:absolute;mso-position-horizontal-relative:margin"
            ),
            1
        );
        assert_eq!(
            count_bytes(
                &host,
                b"mso-position-vertical:absolute;mso-position-vertical-relative:margin"
            ),
            1
        );
    }

    #[test]
    fn strict_extension_is_refused_but_fallback_can_be_rendered() {
        let style = Style::drawing(Placement::inline(extent()), image());
        assert!(matches!(
            render(&style, StoryDialect::Strict, "ink", Some("image"), 1),
            Err(Error::UnsafeEdit { .. })
        ));
        let fallback = render_fallback(&image(), StoryDialect::Strict, "image", 1)
            .expect("standalone fallback");
        assert!(fallback.starts_with(b"<mc:Fallback"));
        assert_eq!(
            count_bytes(
                &fallback,
                b"xmlns:r=\"http://purl.oclc.org/ooxml/officeDocument/relationships\""
            ),
            1
        );
    }

    #[test]
    fn fallback_image_bounds_and_signatures_are_checked() {
        assert!(
            FallbackImage::new(
                vec![1, 2, 3],
                ImageType::Png,
                ImageDimensions::new(1, 1).expect("dimensions"),
            )
            .is_err()
        );
        assert!(
            FallbackImage::new(
                PNG_WITHOUT_IDAT.to_vec(),
                ImageType::Png,
                ImageDimensions::new(1, 1).expect("dimensions"),
            )
            .is_err()
        );
        let mut bad_crc = PNG_1X1.to_vec();
        bad_crc[42] ^= 1;
        assert!(
            FallbackImage::new(
                bad_crc,
                ImageType::Png,
                ImageDimensions::new(1, 1).expect("dimensions"),
            )
            .is_err()
        );
        assert!(ImageDimensions::new(0, 1).is_err());
        assert!(ImageDimensions::new(1_000_000, 1_000_000).is_err());
        assert!(
            FallbackImage::new(
                PNG_1X1.to_vec(),
                ImageType::Png,
                ImageDimensions::new(2, 1).expect("dimensions"),
            )
            .is_err()
        );
        let detected = FallbackImage::from_bytes(PNG_1X1.to_vec()).expect("detected PNG");
        assert_eq!(detected.media_type(), ImageType::Png);
        assert_eq!(
            detected.dimensions(),
            ImageDimensions::new(1, 1).expect("dimensions")
        );
    }

    #[test]
    fn anchor_geometry_rejects_out_of_range_values() {
        assert!(Geometry::new(MAX_EMU + 1, 1).is_err());
        assert!(Point::new(MIN_COORDINATE, 0).is_ok());
        assert!(Point::new(MAX_COORDINATE, 0).is_ok());
        assert!(Point::new(MIN_COORDINATE - 1, 0).is_err());
        assert!(Point::new(MAX_COORDINATE + 1, 0).is_err());
        assert!(
            render(
                &Style::base(BaseProfile::InkContent),
                StoryDialect::Transitional,
                "bad id",
                None,
                1,
            )
            .is_err()
        );
    }
}
