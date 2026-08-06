//! Typed annotation models and bounded scanner state.

pub use litchi_odf_common::annotation::Annotation;
use litchi_odf_common::annotation::Builder;

/// A schema location to which an annotation is attached.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum AnnotationPosition {
    TextParagraph {
        paragraph_index: usize,
    },
    SpreadsheetCell {
        sheet_index: usize,
        row: usize,
        column: usize,
    },
    PresentationPage {
        page_index: usize,
    },
    PresentationShape {
        page_index: usize,
        shape_name: String,
    },
    AnnotationBody {
        annotation_index: usize,
    },
}

/// Start position and optional named-range end position.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AnnotationAnchor {
    pub start: AnnotationPosition,
    pub end: Option<AnnotationPosition>,
}

impl AnnotationAnchor {
    pub fn point(start: AnnotationPosition) -> Self {
        Self { start, end: None }
    }

    pub fn range(start: AnnotationPosition, end: AnnotationPosition) -> Self {
        Self {
            start,
            end: Some(end),
        }
    }
}

/// One annotation in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AnnotationInfo {
    pub index: usize,
    pub annotation: Annotation,
    pub anchor: AnnotationAnchor,
}

/// Partial typed metadata update. `None` retains a value; `Some(None)` clears it.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct AnnotationUpdate {
    pub creator: Option<Option<String>>,
    pub date: Option<Option<String>>,
    pub date_string: Option<Option<String>>,
    pub initials: Option<Option<String>>,
    pub display: Option<Option<bool>>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum AnnotationHost {
    Text,
    Spreadsheet,
    Presentation,
}

#[derive(Clone)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) close_start: Option<usize>,
    pub(crate) qname: String,
}

#[derive(Clone)]
pub(crate) struct Site {
    pub(crate) position: AnnotationPosition,
    pub(crate) span: Span,
}

pub(crate) struct Record {
    pub(crate) span: Span,
    pub(crate) parent_start: usize,
    pub(crate) annotation: Option<Annotation>,
    pub(crate) start_position: AnnotationPosition,
    pub(crate) end: Option<(Span, AnnotationPosition)>,
}

pub(crate) struct EndMarker {
    pub(crate) span: Span,
    pub(crate) name: String,
    pub(crate) position: AnnotationPosition,
}

pub(crate) struct Scan {
    pub(crate) records: Vec<Record>,
    pub(crate) sites: Vec<Site>,
}

pub(crate) enum FrameKind {
    Table {
        sheet: usize,
        next_row: usize,
    },
    Row {
        sheet: usize,
        row: usize,
        next_column: usize,
    },
    Cell {
        site: usize,
    },
    Page {
        site: usize,
        page: usize,
    },
    Shape {
        site: usize,
    },
    Paragraph {
        site: Option<usize>,
    },
    Annotation {
        record: usize,
    },
    Other,
}

pub(crate) struct Frame {
    pub(crate) start: usize,
    pub(crate) kind: FrameKind,
    pub(crate) namespace_changes: Vec<(String, Option<String>)>,
}

pub(crate) struct ActiveBuilder {
    pub(crate) record: usize,
    pub(crate) builder: Builder,
}
