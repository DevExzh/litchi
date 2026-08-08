use super::super::package::Error;
/// Base shape trait and common shape functionality.
///
/// This module defines the core Shape trait that all shape types implement,
/// along with common properties and methods for working with shapes.
use std::borrow::Cow;
use std::fmt;

/// Semantic shape mutation that requires a persistence-capable edit transaction.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mutation {
    /// Replace shape text.
    Text,
    /// Replace the shape fill.
    Fill,
    /// Replace the shape outline.
    Line,
    /// Change shape geometry or adjustment data.
    Geometry,
    /// Change text layout or character formatting.
    Formatting,
    /// Change placeholder identity or sizing.
    Placeholder,
    /// Change picture identity, framing, bounds, or retained OfficeArt data.
    Picture,
    /// Change a shape's child structure.
    Structure,
    /// Obtain unrestricted mutable access to common shape properties.
    Properties,
}

/// Refusal returned when a parsed shape cannot be mutated losslessly in place.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MutationError {
    /// The shape belongs to an opened presentation, but no source-checked
    /// transaction can faithfully publish this mutation yet.
    SourceBound { mutation: Mutation },
    /// A one-based picture-store index was zero or otherwise invalid.
    InvalidPictureIndex { index: u32 },
}

impl MutationError {
    pub(crate) const fn source_bound(mutation: Mutation) -> Self {
        Self::SourceBound { mutation }
    }

    /// Return the refused semantic mutation.
    pub const fn mutation(self) -> Mutation {
        match self {
            Self::SourceBound { mutation } => mutation,
            Self::InvalidPictureIndex { .. } => Mutation::Picture,
        }
    }
}

impl fmt::Display for MutationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::SourceBound { mutation } => write!(
                formatter,
                "source-bound PPT shape mutation {mutation:?} requires a persistence transaction"
            ),
            Self::InvalidPictureIndex { index } => {
                write!(formatter, "invalid one-based PPT picture index {index}")
            },
        }
    }
}

impl std::error::Error for MutationError {}

/// Types of shapes in PowerPoint presentations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    /// Text box shape
    TextBox,
    /// Placeholder shape (title, content, etc.)
    Placeholder,
    /// Auto shape (rectangle, oval, etc.)
    AutoShape,
    /// Picture shape
    Picture,
    /// Group shape (container for other shapes)
    Group,
    /// Line shape
    Line,
    /// Connector shape
    Connector,
    /// Object shape (embedded objects)
    Object,
    /// Media shape (audio or video preview frame)
    Media,
    /// Table shape
    Table,
    /// Unknown shape type
    Unknown(u16),
}

impl From<u16> for ShapeType {
    fn from(value: u16) -> Self {
        match value {
            1 => ShapeType::TextBox,
            2 => ShapeType::Placeholder,
            3 => ShapeType::AutoShape,
            4 => ShapeType::Picture,
            5 => ShapeType::Group,
            6 => ShapeType::Line,
            7 => ShapeType::Connector,
            8 => ShapeType::Object,
            9 => ShapeType::Table,
            10 => ShapeType::Media,
            other => ShapeType::Unknown(other),
        }
    }
}

impl fmt::Display for ShapeType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ShapeType::TextBox => write!(f, "TextBox"),
            ShapeType::Placeholder => write!(f, "Placeholder"),
            ShapeType::AutoShape => write!(f, "AutoShape"),
            ShapeType::Picture => write!(f, "Picture"),
            ShapeType::Group => write!(f, "Group"),
            ShapeType::Line => write!(f, "Line"),
            ShapeType::Connector => write!(f, "Connector"),
            ShapeType::Object => write!(f, "Object"),
            ShapeType::Media => write!(f, "Media"),
            ShapeType::Table => write!(f, "Table"),
            ShapeType::Unknown(id) => write!(f, "Unknown({})", id),
        }
    }
}

/// Common properties shared by all shape types.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeProperties {
    /// Shape ID
    pub id: u32,
    /// Shape type
    pub shape_type: ShapeType,
    /// X position in EMUs (English Metric Units)
    pub x: i32,
    /// Y position in EMUs
    pub y: i32,
    /// Width in EMUs
    pub width: i32,
    /// Height in EMUs
    pub height: i32,
    /// Rotation angle (0-360 degrees)
    pub rotation: u16,
    /// Fill color (RGB)
    pub fill_color: Option<u32>,
    /// Line color (RGB)
    pub line_color: Option<u32>,
    /// Line width in points
    pub line_width: Option<u16>,
    /// Is the shape hidden?
    pub hidden: bool,
    /// Z-order (drawing order)
    pub z_order: u16,
    /// Inert PowerPoint 12 placeholder compatibility metadata.
    pub powerpoint12_shape_metadata: Option<crate::slide_extension::ShapeMetadata>,
}

impl Default for ShapeProperties {
    fn default() -> Self {
        Self {
            id: 0,
            shape_type: ShapeType::Unknown(0),
            x: 0,
            y: 0,
            width: 0,
            height: 0,
            rotation: 0,
            fill_color: None,
            line_color: None,
            line_width: None,
            hidden: false,
            z_order: 0,
            powerpoint12_shape_metadata: None,
        }
    }
}

/// Base trait for all shape types in PowerPoint presentations.
///
/// This trait defines the common interface that all shape implementations
/// must provide, including access to properties and basic operations.
pub trait Shape: std::any::Any {
    /// Get the shape's properties.
    fn properties(&self) -> &ShapeProperties;

    /// Get the shape's properties for a detached semantic value.
    ///
    /// Parsed presentation shapes return [`MutationError::SourceBound`]
    /// before exposing mutable state.
    fn properties_mut(&mut self) -> Result<&mut ShapeProperties, MutationError>;

    /// Replace the fill color on a detached shape.
    fn set_fill_color(&mut self, color: Option<u32>) -> Result<(), MutationError> {
        let properties = self
            .properties_mut()
            .map_err(|_| MutationError::source_bound(Mutation::Fill))?;
        properties.fill_color = color;
        Ok(())
    }

    /// Replace the line color and width on a detached shape.
    fn set_line(&mut self, color: Option<u32>, width: Option<u16>) -> Result<(), MutationError> {
        let properties = self
            .properties_mut()
            .map_err(|_| MutationError::source_bound(Mutation::Line))?;
        properties.line_color = color;
        properties.line_width = width;
        Ok(())
    }

    /// Get the shape type.
    fn shape_type(&self) -> ShapeType {
        self.properties().shape_type
    }

    /// Get the shape ID.
    fn id(&self) -> u32 {
        self.properties().id
    }

    /// Get the shape's text content, if any.
    fn text(&self) -> Result<String, Error>;

    /// Get the shape's position and size.
    fn bounds(&self) -> (i32, i32, i32, i32) {
        let props = self.properties();
        (props.x, props.y, props.width, props.height)
    }

    /// Check if the shape is a placeholder.
    fn is_placeholder(&self) -> bool {
        matches!(self.shape_type(), ShapeType::Placeholder)
    }

    /// Check if the shape has text content.
    fn has_text(&self) -> bool;

    /// Clone the shape as a boxed trait object.
    fn clone_box(&self) -> Box<dyn Shape>;

    /// Get the shape as an Any reference for downcasting.
    fn as_any(&self) -> &dyn std::any::Any;
}

impl Clone for Box<dyn Shape> {
    fn clone(&self) -> Self {
        self.clone_box()
    }
}

/// Shape container that holds shape data and provides common operations.
///
/// This container includes Escher text properties extracted from the shape's
/// OfficeArtFOPT records, following Apache POI's EscherProperties model.
///
/// Uses `Cow<'a, [u8]>` for zero-copy optimization when parsing shapes,
/// allowing the raw data to be borrowed from the original source when possible.
///
/// Note: Child shapes use `'static` lifetime to maintain trait object safety,
/// while raw_data can be borrowed with lifetime `'a`.
#[derive(Clone)]
pub struct ShapeContainer<'a> {
    /// Shape properties
    pub(crate) properties: ShapeProperties,
    /// Raw shape data (for parsing) - uses Cow to avoid unnecessary clones
    pub(crate) raw_data: Cow<'a, [u8]>,
    /// Text content (if applicable)
    pub(crate) text_content: Option<String>,
    /// Child shapes (for group shapes) - must have 'static lifetime for trait object safety
    pub(crate) children: Vec<Box<dyn Shape>>,

    // Escher text properties (from OfficeArtFOPT records)
    /// Text left margin in EMUs
    /// Property ID: 0x0081 (TEXT_LEFT)
    pub(crate) text_left: Option<i32>,

    /// Text top margin in EMUs
    /// Property ID: 0x0082 (TEXT_TOP)
    pub(crate) text_top: Option<i32>,

    /// Text right margin in EMUs
    /// Property ID: 0x0083 (TEXT_RIGHT)
    pub(crate) text_right: Option<i32>,

    /// Text bottom margin in EMUs
    /// Property ID: 0x0084 (TEXT_BOTTOM)
    pub(crate) text_bottom: Option<i32>,

    /// Text flow direction
    /// Property ID: 0x0088 (TEXT_FLOW)
    /// Values: 0=horizontal, 1=vertical, 2=vertical rotated, 3=word art vertical
    pub(crate) text_flow: Option<u16>,

    /// Text wrapping mode (`MSOWRAPMODE`)
    /// Property ID: 0x0085 (WRAP_TEXT)
    pub(crate) wrap_text: Option<u16>,

    /// Text anchor (vertical alignment)
    /// Property ID: 0x0087 (ANCHOR_TEXT)
    /// Values: 0=top, 1=middle, 2=bottom, 3=top centered, 4=middle centered,
    ///         5=bottom centered, 6=top baseline, 7=bottom baseline, 8=top centered baseline
    pub(crate) anchor_text: Option<u16>,

    /// Text ID (identifier for the text)
    /// Property ID: 0x0080 (TEXT_ID)
    pub(crate) text_id: Option<i32>,

    /// Size shape to fit text content
    /// Packed Boolean property ID: 0x00BE (fFitShapeToText)
    pub(crate) size_shape_to_fit_text: Option<bool>,

    /// Font rotation (`MSOCDIR`)
    /// Property ID: 0x0089 (cdirFont)
    pub(crate) font_rotation: Option<u16>,

    /// Text direction (`MSOTXDIR`)
    /// Property ID: 0x008B (txdir)
    pub(crate) text_direction: Option<u16>,

    /// Use automatic default text margins
    /// Packed Boolean property ID: 0x00BC (fAutoTextMargin)
    pub(crate) auto_text_margin: Option<bool>,

    /// Enter text editing mode when the contained text area is clicked
    /// Packed Boolean property ID: 0x00BB (fSelectText)
    pub(crate) select_text: Option<bool>,

    /// ID of next shape in sequence
    /// Property ID: 0x008A (hspNext)
    pub(crate) id_of_next_shape: Option<u32>,

    source_bound: bool,
}

impl<'a> fmt::Debug for ShapeContainer<'a> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut debug = f.debug_struct("ShapeContainer");
        debug
            .field("properties", &self.properties)
            .field("raw_data_len", &self.raw_data.len())
            .field("text_content", &self.text_content)
            .field("children_count", &self.children.len());

        // Only show Escher text properties if they're set
        if self.text_left.is_some()
            || self.text_top.is_some()
            || self.text_right.is_some()
            || self.text_bottom.is_some()
        {
            debug.field(
                "text_margins",
                &format_args!(
                    "L:{:?} T:{:?} R:{:?} B:{:?}",
                    self.text_left, self.text_top, self.text_right, self.text_bottom
                ),
            );
        }
        if let Some(flow) = self.text_flow {
            debug.field("text_flow", &flow);
        }
        if let Some(anchor) = self.anchor_text {
            debug.field("anchor_text", &anchor);
        }
        if let Some(wrap) = self.wrap_text {
            debug.field("wrap_text", &wrap);
        }

        debug.finish()
    }
}

impl<'a> ShapeContainer<'a> {
    /// Return the immutable common shape properties.
    pub const fn properties(&self) -> &ShapeProperties {
        &self.properties
    }

    /// Return the retained OfficeArt bytes.
    pub fn raw_data(&self) -> &[u8] {
        self.raw_data.as_ref()
    }

    /// Return the decoded text, when present.
    pub fn text_content(&self) -> Option<&str> {
        self.text_content.as_deref()
    }

    /// Return the explicit left text margin.
    pub const fn text_left(&self) -> Option<i32> {
        self.text_left
    }

    /// Return the explicit top text margin.
    pub const fn text_top(&self) -> Option<i32> {
        self.text_top
    }

    /// Return the explicit right text margin.
    pub const fn text_right(&self) -> Option<i32> {
        self.text_right
    }

    /// Return the explicit bottom text margin.
    pub const fn text_bottom(&self) -> Option<i32> {
        self.text_bottom
    }

    /// Return the raw text-flow value.
    pub const fn text_flow(&self) -> Option<u16> {
        self.text_flow
    }

    /// Return the raw text-wrapping value.
    pub const fn wrap_text(&self) -> Option<u16> {
        self.wrap_text
    }

    /// Return the raw text-anchor value.
    pub const fn anchor_text(&self) -> Option<u16> {
        self.anchor_text
    }

    /// Return the retained text identifier.
    pub const fn text_id(&self) -> Option<i32> {
        self.text_id
    }

    /// Return the explicit fit-shape-to-text flag.
    pub const fn size_shape_to_fit_text(&self) -> Option<bool> {
        self.size_shape_to_fit_text
    }

    /// Return the raw font-rotation value.
    pub const fn font_rotation(&self) -> Option<u16> {
        self.font_rotation
    }

    /// Return the raw text-direction value.
    pub const fn text_direction(&self) -> Option<u16> {
        self.text_direction
    }

    /// Return the explicit automatic-margin flag.
    pub const fn auto_text_margin(&self) -> Option<bool> {
        self.auto_text_margin
    }

    /// Return the explicit select-text flag.
    pub const fn select_text(&self) -> Option<bool> {
        self.select_text
    }

    /// Return the next shape identifier in the sequence.
    pub const fn id_of_next_shape(&self) -> Option<u32> {
        self.id_of_next_shape
    }

    /// Create a new shape container with owned data.
    ///
    /// All Escher text properties are initialized to `None` and can be
    /// populated later by parsing OfficeArtFOPT records.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let container = ShapeContainer::new(properties, data.to_vec());
    /// ```
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            properties,
            raw_data: Cow::Owned(raw_data),
            text_content: None,
            children: Vec::new(),
            // Initialize all Escher text properties to None
            text_left: None,
            text_top: None,
            text_right: None,
            text_bottom: None,
            text_flow: None,
            wrap_text: None,
            anchor_text: None,
            text_id: None,
            size_shape_to_fit_text: None,
            font_rotation: None,
            text_direction: None,
            auto_text_margin: None,
            select_text: None,
            id_of_next_shape: None,
            source_bound: false,
        }
    }

    /// Create a new shape container with borrowed data (zero-copy).
    ///
    /// This constructor enables zero-copy parsing by borrowing the raw data
    /// instead of copying it. Use this when the shape data lifetime is tied
    /// to a larger buffer that will outlive the container.
    ///
    /// All Escher text properties are initialized to `None` and can be
    /// populated later by parsing OfficeArtFOPT records.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let container = ShapeContainer::new_borrowed(properties, &data_slice);
    /// ```
    pub fn new_borrowed(properties: ShapeProperties, raw_data: &'a [u8]) -> Self {
        Self {
            properties,
            raw_data: Cow::Borrowed(raw_data),
            text_content: None,
            children: Vec::new(),
            // Initialize all Escher text properties to None
            text_left: None,
            text_top: None,
            text_right: None,
            text_bottom: None,
            text_flow: None,
            wrap_text: None,
            anchor_text: None,
            text_id: None,
            size_shape_to_fit_text: None,
            font_rotation: None,
            text_direction: None,
            auto_text_margin: None,
            select_text: None,
            id_of_next_shape: None,
            source_bound: false,
        }
    }

    /// Add a child shape to this shape (for group shapes).
    pub fn add_child(&mut self, shape: Box<dyn Shape>) -> Result<(), MutationError> {
        self.ensure_mutable(Mutation::Structure)?;
        self.children.push(shape);
        Ok(())
    }

    /// Get all child shapes.
    pub fn children(&self) -> &[Box<dyn Shape>] {
        &self.children
    }

    /// Set the text content of a detached shape.
    pub fn set_text(&mut self, text: String) -> Result<(), MutationError> {
        self.ensure_mutable(Mutation::Text)?;
        self.set_decoded_text(text);
        Ok(())
    }

    pub(crate) fn set_decoded_text(&mut self, text: String) {
        self.text_content = Some(text);
    }

    pub(crate) fn ensure_mutable(&self, mutation: Mutation) -> Result<(), MutationError> {
        if self.source_bound {
            Err(MutationError::source_bound(mutation))
        } else {
            Ok(())
        }
    }

    pub(crate) fn properties_mut_checked(&mut self) -> Result<&mut ShapeProperties, MutationError> {
        self.ensure_mutable(Mutation::Properties)?;
        Ok(&mut self.properties)
    }

    pub(crate) fn mark_source_bound(&mut self) {
        self.source_bound = true;
    }

    /// Set text margins from a 4-value tuple (left, top, right, bottom).
    ///
    /// # Arguments
    ///
    /// * `margins` - Tuple of (left, top, right, bottom) margins in EMUs
    ///
    /// # Example
    ///
    /// ```ignore
    /// container.set_text_margins(Some((91440, 45720, 91440, 45720)));
    /// ```
    pub fn set_text_margins(
        &mut self,
        margins: Option<(i32, i32, i32, i32)>,
    ) -> Result<(), MutationError> {
        self.ensure_mutable(Mutation::Formatting)?;
        if let Some((left, top, right, bottom)) = margins {
            self.text_left = Some(left);
            self.text_top = Some(top);
            self.text_right = Some(right);
            self.text_bottom = Some(bottom);
        }
        Ok(())
    }

    /// Get text margins as a 4-value tuple.
    ///
    /// # Returns
    ///
    /// `Some((left, top, right, bottom))` if all four margins are set, `None` otherwise
    pub fn text_margins(&self) -> Option<(i32, i32, i32, i32)> {
        match (
            self.text_left,
            self.text_top,
            self.text_right,
            self.text_bottom,
        ) {
            (Some(l), Some(t), Some(r), Some(b)) => Some((l, t, r, b)),
            _ => None,
        }
    }

    /// Get text margins with the MS-ODRAW defaults applied to absent values.
    pub fn effective_text_margins(&self) -> (i32, i32, i32, i32) {
        const HORIZONTAL_DEFAULT: i32 = 0x0001_6530;
        const VERTICAL_DEFAULT: i32 = 0x0000_B298;
        (
            self.text_left.unwrap_or(HORIZONTAL_DEFAULT),
            self.text_top.unwrap_or(VERTICAL_DEFAULT),
            self.text_right.unwrap_or(HORIZONTAL_DEFAULT),
            self.text_bottom.unwrap_or(VERTICAL_DEFAULT),
        )
    }

    /// Set text flow direction.
    ///
    /// # Values
    ///
    /// - 0: Horizontal (left to right)
    /// - 1: Vertical (top to bottom)
    /// - 2: Vertical rotated
    /// - 3: Word art vertical
    pub fn set_text_flow(&mut self, flow: Option<u16>) -> Result<(), MutationError> {
        self.ensure_mutable(Mutation::Formatting)?;
        self.text_flow = flow;
        Ok(())
    }

    /// Set text anchor (vertical alignment).
    ///
    /// # Values
    ///
    /// - 0: Top
    /// - 1: Middle
    /// - 2: Bottom
    /// - 3: Top centered
    /// - 4: Middle centered
    /// - 5: Bottom centered
    /// - 6: Top baseline
    /// - 7: Bottom baseline
    /// - 8: Top centered baseline
    pub fn set_anchor_text(&mut self, anchor: Option<u16>) -> Result<(), MutationError> {
        self.ensure_mutable(Mutation::Formatting)?;
        self.anchor_text = anchor;
        Ok(())
    }

    /// Set the raw `MSOWRAPMODE` text wrapping value.
    pub fn set_wrap_text(&mut self, wrap: Option<u16>) -> Result<(), MutationError> {
        self.ensure_mutable(Mutation::Formatting)?;
        self.wrap_text = wrap;
        Ok(())
    }

    /// Whether the wrapping mode allows wrapping within the shape.
    ///
    /// `msowrapNone` is encoded as 2; every other defined mode wraps text.
    pub fn word_wrap_enabled(&self) -> Option<bool> {
        self.wrap_text.map(|mode| mode != 2)
    }

    /// Set the raw `MSOCDIR` font rotation value.
    pub fn set_font_rotation(&mut self, rotation: Option<u16>) -> Result<(), MutationError> {
        self.ensure_mutable(Mutation::Formatting)?;
        self.font_rotation = rotation;
        Ok(())
    }
}

impl<'a> Shape for ShapeContainer<'a>
where
    'a: 'static,
{
    fn properties(&self) -> &ShapeProperties {
        &self.properties
    }

    fn properties_mut(&mut self) -> Result<&mut ShapeProperties, MutationError> {
        self.properties_mut_checked()
    }

    fn text(&self) -> Result<String, Error> {
        Ok(self.text_content.clone().unwrap_or_default())
    }

    fn has_text(&self) -> bool {
        self.text_content.is_some()
    }

    fn clone_box(&self) -> Box<dyn Shape> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

#[cfg(test)]
mod persistence_tests {
    use super::{Mutation, MutationError, Shape, ShapeContainer, ShapeProperties};

    #[derive(Clone)]
    struct ExternalShape {
        properties: ShapeProperties,
    }

    impl Shape for ExternalShape {
        fn properties(&self) -> &ShapeProperties {
            &self.properties
        }

        fn properties_mut(&mut self) -> Result<&mut ShapeProperties, MutationError> {
            Ok(&mut self.properties)
        }

        fn text(&self) -> crate::package::Result<String> {
            Ok(String::new())
        }

        fn has_text(&self) -> bool {
            false
        }

        fn clone_box(&self) -> Box<dyn Shape> {
            Box::new(self.clone())
        }

        fn as_any(&self) -> &dyn std::any::Any {
            self
        }
    }

    #[test]
    fn external_shape_implementers_have_no_lineage_contract() {
        let mut shape = ExternalShape {
            properties: ShapeProperties::default(),
        };
        assert!(Shape::set_fill_color(&mut shape, Some(0x0011_2233)).is_ok());
        assert_eq!(shape.properties.fill_color, Some(0x0011_2233));
    }

    #[test]
    fn source_bound_text_fill_and_line_mutations_are_atomic_refusals() {
        let mut shape = ShapeContainer::new(ShapeProperties::default(), Vec::new());
        assert!(shape.set_text("before".to_string()).is_ok());
        assert!(Shape::set_fill_color(&mut shape, Some(0x0011_2233)).is_ok());
        assert!(Shape::set_line(&mut shape, Some(0x0044_5566), Some(2)).is_ok());
        shape.mark_source_bound();
        let before = shape.clone();

        assert_eq!(
            shape.set_text("after".to_string()),
            Err(MutationError::SourceBound {
                mutation: Mutation::Text,
            })
        );
        assert_eq!(
            Shape::set_fill_color(&mut shape, None),
            Err(MutationError::SourceBound {
                mutation: Mutation::Fill,
            })
        );
        assert_eq!(
            Shape::set_line(&mut shape, None, None),
            Err(MutationError::SourceBound {
                mutation: Mutation::Line,
            })
        );
        assert_eq!(shape.text_content, before.text_content);
        assert_eq!(shape.properties, before.properties);
    }

    #[test]
    fn source_bound_properties_never_expose_a_mutable_reference() {
        let mut shape = ShapeContainer::new(ShapeProperties::default(), Vec::new());
        shape.mark_source_bound();
        assert!(matches!(
            Shape::properties_mut(&mut shape),
            Err(MutationError::SourceBound {
                mutation: Mutation::Properties,
            })
        ));
    }

    #[test]
    fn source_bound_container_structure_and_formatting_are_atomic_refusals() {
        let mut shape = ShapeContainer::new(ShapeProperties::default(), Vec::new());
        assert!(shape.set_text_margins(Some((1, 2, 3, 4))).is_ok());
        assert!(shape.set_text_flow(Some(1)).is_ok());
        assert!(shape.set_anchor_text(Some(2)).is_ok());
        assert!(shape.set_wrap_text(Some(1)).is_ok());
        assert!(shape.set_font_rotation(Some(2)).is_ok());
        shape.mark_source_bound();
        let before = shape.clone();

        assert_eq!(
            shape.add_child(Box::new(ShapeContainer::new(
                ShapeProperties::default(),
                Vec::new(),
            ))),
            Err(MutationError::SourceBound {
                mutation: Mutation::Structure,
            })
        );
        for refusal in [
            shape.set_text_margins(Some((5, 6, 7, 8))),
            shape.set_text_flow(None),
            shape.set_anchor_text(None),
            shape.set_wrap_text(None),
            shape.set_font_rotation(None),
        ] {
            assert_eq!(
                refusal,
                Err(MutationError::SourceBound {
                    mutation: Mutation::Formatting,
                })
            );
        }
        assert_eq!(shape.children.len(), before.children.len());
        assert_eq!(shape.text_margins(), before.text_margins());
        assert_eq!(shape.text_flow, before.text_flow);
        assert_eq!(shape.anchor_text, before.anchor_text);
        assert_eq!(shape.wrap_text, before.wrap_text);
        assert_eq!(shape.font_rotation, before.font_rotation);
    }
}
