/// Common shape types and enumerations.
///
/// This module provides unified shape types used by both legacy (.ppt) and
/// modern (.pptx) presentation formats.

use std::fmt;

/// Types of shapes in presentations.
///
/// This enumeration is used for both legacy .ppt and modern .pptx formats,
/// providing a unified interface for shape type identification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    /// Text box or text shape
    TextBox,
    /// Placeholder shape (title, content, footer, etc.)
    Placeholder,
    /// Auto shape (rectangle, oval, arrow, etc.)
    AutoShape,
    /// Picture/image shape
    Picture,
    /// Group shape (container for other shapes)
    Group,
    /// Line shape
    Line,
    /// Connector shape
    Connector,
    /// Table shape
    Table,
    /// Graphic frame (chart, SmartArt, etc.)
    GraphicFrame,
    /// Unknown or unsupported shape type
    Unknown,
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
            ShapeType::Table => write!(f, "Table"),
            ShapeType::GraphicFrame => write!(f, "GraphicFrame"),
            ShapeType::Unknown => write!(f, "Unknown"),
        }
    }
}

/// Placeholder types in presentations.
///
/// Defines the semantic role of a placeholder shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PlaceholderType {
    /// Title placeholder
    Title,
    /// Body/content placeholder
    Body,
    /// Center title placeholder
    CenteredTitle,
    /// Subtitle placeholder
    Subtitle,
    /// Date placeholder
    Date,
    /// Slide number placeholder
    SlideNumber,
    /// Footer placeholder
    Footer,
    /// Header placeholder
    Header,
    /// Object placeholder (chart, table, etc.)
    Object,
    /// Chart placeholder
    Chart,
    /// Table placeholder
    Table,
    /// Clip art placeholder
    ClipArt,
    /// Diagram/organization chart placeholder
    Diagram,
    /// Media placeholder (audio, video)
    Media,
    /// Picture placeholder
    Picture,
    /// Unknown placeholder type
    Unknown,
}

impl fmt::Display for PlaceholderType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            PlaceholderType::Title => write!(f, "Title"),
            PlaceholderType::Body => write!(f, "Body"),
            PlaceholderType::CenteredTitle => write!(f, "CenteredTitle"),
            PlaceholderType::Subtitle => write!(f, "Subtitle"),
            PlaceholderType::Date => write!(f, "Date"),
            PlaceholderType::SlideNumber => write!(f, "SlideNumber"),
            PlaceholderType::Footer => write!(f, "Footer"),
            PlaceholderType::Header => write!(f, "Header"),
            PlaceholderType::Object => write!(f, "Object"),
            PlaceholderType::Chart => write!(f, "Chart"),
            PlaceholderType::Table => write!(f, "Table"),
            PlaceholderType::ClipArt => write!(f, "ClipArt"),
            PlaceholderType::Diagram => write!(f, "Diagram"),
            PlaceholderType::Media => write!(f, "Media"),
            PlaceholderType::Picture => write!(f, "Picture"),
            PlaceholderType::Unknown => write!(f, "Unknown"),
        }
    }
}

