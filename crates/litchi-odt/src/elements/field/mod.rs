//! Layered semantic owner for ODF field models and XML codecs.
//!
//! The historical `crate::elements::field` path remains the public facade.
//! Typed field values live in `model`, namespace-aware XML mechanics in
//! `codec`, document-content parsing in `package`, and focused regressions
//! in `tests`.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(super) const MAX_FIELD_DEPTH: usize = 4_096;
pub(super) const MAX_FIELDS: usize = 1_000_000;
pub(super) const TEXT_DATABASE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const FORM_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
pub(super) const STYLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
pub(super) const MAX_DATABASE_VALUE: usize = 65_536;
pub(super) const MAX_DATABASE_AGGREGATE: usize = 16 * 1_048_576;
pub(super) const MAX_DATABASE_INTEGER_DIGITS: usize = 4_096;
pub(super) const MAX_DYNAMIC_FIELD_VALUE: usize = 65_536;
pub(super) const MAX_DYNAMIC_FIELD_AGGREGATE: usize = 1_048_576;
pub(super) const MAX_DROP_DOWN_LABELS: usize = 65_536;
pub(super) const MAX_META_FIELD_XML_BYTES: usize = 64 * 1_048_576;
pub(super) const MAX_META_FIELD_DEPTH: usize = 256;
pub(super) const MAX_META_FIELD_NODES: usize = 100_000;
pub(super) const MAX_META_FIELD_ATTRIBUTES: usize = 256;
pub(super) const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
pub(super) const DRAW_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(super) const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(super) const PRESENTATION_NAMESPACE: &str =
    "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
pub(super) const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(super) const FO_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
pub(super) const NUMBER_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
pub(super) const META_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
pub(super) const DC_NAMESPACE: &str = "http://purl.org/dc/elements/1.1/";
pub(super) const XHTML_NAMESPACE: &str = "http://www.w3.org/1999/xhtml";
pub(super) const DR3D_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
pub(super) const SCRIPT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";

pub use model::{
    CalculatedFieldValue, ChapterDisplay, CrossReferenceFormat, DatabaseConnectionResource,
    DatabaseField, DatabaseFieldKind, DatabaseSource, DatabaseTableType, DateField, DateValueKind,
    DropDownLabel, DynamicTextField, Field, FieldDateValue, FieldDuration, FieldTimeValue,
    FieldValueType, FileNameDisplay, FormulaFieldDisplay, IdentityFieldKind, MeasureKind,
    MetaFieldAttribute, MetaFieldContent, MetaFieldElement, MetaFieldNode, MetadataFieldKind,
    MetadataFieldValue, NonNegativeInteger, NoteBodyContent, NoteReferenceClass,
    NoteReferenceFormat, PageContinuationSelection, PageNumberField, PageSelection,
    PlaceholderType, ReferenceField, SenderFieldKind, SequenceNumberFormat,
    SequenceReferenceFormat, StatisticKind, TemplateNameDisplay, TimeValueKind,
    UserDefinedMetadataValues, UserFieldDisplay, VariableSetDisplay,
};

pub use package::FieldParser;
pub(crate) use package::parse_note_body_contents;
