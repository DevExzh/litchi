//! Context retained for inert drawing resources.

/// XML part containing a drawing-resource occurrence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Part {
    /// The package `content.xml` part.
    Content,
    /// The package `styles.xml` part.
    Styles,
    /// A flat OpenDocument XML document.
    FlatDocument,
}

/// Drawing-frame context for an image or embedded-object occurrence.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Frame {
    /// The optional `draw:name` value.
    pub name: Option<String>,
    /// The optional `xml:id` value.
    pub xml_id: Option<String>,
    /// Short alternative title from a direct `svg:title` child.
    pub title: Option<String>,
    /// Prose alternative description from a direct `svg:desc` child.
    pub description: Option<String>,
    /// The optional `text:anchor-type` value.
    pub anchor_type: Option<String>,
    /// The optional `svg:x` value.
    pub x: Option<String>,
    /// The optional `svg:y` value.
    pub y: Option<String>,
    /// The optional `svg:width` value.
    pub width: Option<String>,
    /// The optional `svg:height` value.
    pub height: Option<String>,
    /// Cell address anchoring the frame's bottom-right corner
    /// (`table:end-cell-address`, spreadsheets only).
    pub end_cell_address: Option<String>,
    /// The containing drawing page name, if any.
    pub page_name: Option<String>,
    /// The containing spreadsheet sheet name, if any.
    pub sheet_name: Option<String>,
    /// Whether this frame is a direct child of a spreadsheet `table:shapes`
    /// container.
    pub sheet_shape: bool,
}
