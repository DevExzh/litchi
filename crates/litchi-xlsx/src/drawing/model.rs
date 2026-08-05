//! Semantic SpreadsheetDrawing objects.
//!
//! The worksheet/chartsheet owner keeps only the small amount of information
//! needed to resolve drawing relationships. Full shape and text semantics stay
//! with [`crate::shapes`] and reuse the neutral DrawingML text-body model.

use super::super::chart::Anchor;

/// One bounded worksheet or chartsheet drawing part.
#[derive(Debug, Clone, Default)]
pub struct Drawing {
    objects: Vec<Object>,
}

impl Drawing {
    /// Return anchored objects in source order.
    #[must_use]
    pub fn objects(&self) -> &[Object] {
        &self.objects
    }

    /// Iterate over picture objects in source order.
    pub fn pictures(&self) -> impl Iterator<Item = &Picture> {
        self.objects.iter().filter_map(Object::as_picture)
    }

    /// Iterate over chart objects in source order.
    pub fn charts(&self) -> impl Iterator<Item = &Chart> {
        self.objects.iter().filter_map(Object::as_chart)
    }

    /// Iterate over unsupported or inert objects in source order.
    pub fn unknown(&self) -> impl Iterator<Item = &Unknown> {
        self.objects.iter().filter_map(Object::as_unknown)
    }

    /// Return whether the drawing contains no anchored objects.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.objects.is_empty()
    }

    /// Return the number of anchored objects.
    #[must_use]
    pub fn len(&self) -> usize {
        self.objects.len()
    }

    pub(super) fn push(&mut self, object: Object) {
        self.objects.push(object);
    }
}

/// One anchored SpreadsheetDrawing object.
#[derive(Debug, Clone)]
pub enum Object {
    /// A picture whose image relationship remains owned by the package graph.
    Picture(Picture),
    /// A chart frame whose chart relationship remains owned by the package graph.
    Chart(Chart),
    /// An unsupported or inert object retained as a structural fallback.
    Unknown(Unknown),
}

impl Object {
    fn as_picture(&self) -> Option<&Picture> {
        match self {
            Self::Picture(picture) => Some(picture),
            Self::Chart(_) | Self::Unknown(_) => None,
        }
    }

    fn as_chart(&self) -> Option<&Chart> {
        match self {
            Self::Chart(chart) => Some(chart),
            Self::Picture(_) | Self::Unknown(_) => None,
        }
    }

    fn as_unknown(&self) -> Option<&Unknown> {
        match self {
            Self::Unknown(unknown) => Some(unknown),
            Self::Picture(_) | Self::Chart(_) => None,
        }
    }
}

/// A SpreadsheetDrawing picture object.
#[derive(Debug, Clone)]
pub struct Picture {
    /// Cell-based anchor used by the chart integration model.
    pub anchor: Anchor,
    /// Relationship ID for the picture image.
    pub relationship_id: String,
    /// Optional accessibility description from `xdr:cNvPr@descr`.
    pub description: Option<String>,
}

/// A SpreadsheetDrawing chart object.
#[derive(Debug, Clone)]
pub struct Chart {
    /// Cell-based anchor used by the chart integration model.
    pub anchor: Anchor,
    /// Relationship ID for the chart part.
    pub relationship_id: String,
}

/// The supported structural class of an unknown anchored object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum UnknownKind {
    /// A DrawingML shape (`xdr:sp`).
    Shape,
    /// A DrawingML group (`xdr:grpSp`).
    Group,
    /// A DrawingML connector (`xdr:cxnSp`).
    Connection,
    /// An embedded or otherwise inert content part (`xdr:contentPart`).
    ContentPart,
    /// A producer extension or object not recognized by this owner.
    Other,
}

/// An unsupported or inert anchored object.
#[derive(Debug, Clone)]
pub struct Unknown {
    /// Cell-based anchor retained for structural navigation.
    pub anchor: Anchor,
    /// Optional accessibility description from `xdr:cNvPr@descr`.
    pub description: Option<String>,
    /// Best-effort structural classification; payload remains inert.
    pub kind: UnknownKind,
}

/// Shared DrawingML text-body vocabulary used by SpreadsheetDrawing shapes.
pub mod text {
    pub use litchi_drawingml::text::body::{Body, Insets, Paragraph, Properties, Run};
    pub use litchi_drawingml::text::{
        Anchor as VerticalAnchor, Autofit, Columns, Coordinate32, Direction, TextSize, Underline,
        Wrap,
    };
}
