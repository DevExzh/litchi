//! Typed Word field models grouped by their semantic field families.
//!
//! This module is the stable facade used by `parts::fields`. The child
//! modules own the wire descriptors, legacy controls, indexing fields,
//! references, numbering, document metadata, links, and mail-merge models;
//! their public types are re-exported here so the existing API remains flat at
//! the facade boundary.

mod core;
mod document;
mod indexing;
mod legacy;
mod links;
mod mail_merge;
mod mail_merge_control;
mod numbering;
mod references;

pub use core::{
    Field, FieldBoundary, FieldDescriptor, FieldEndFlags, FieldMarker, FieldMarkerValue,
    FieldStory, FieldText, FieldType,
};
pub use document::{
    DocumentContextField, DocumentContextFieldKind, DocumentInformationField,
    DocumentInformationFieldKind, DocumentPropertyField, DocumentVariableField, InfoField,
};
pub use indexing::{
    IndexEntryField, IndexEntryOption, NonPlcfFields, PrivateField, ReferencedDocumentField,
    TableOfAuthoritiesEntryField, TableOfAuthoritiesEntryOption, TableOfContentsEntryField,
    TableOfContentsEntryOption, TableOfContentsField, TableOfContentsOption,
};
pub use legacy::{
    ActiveContentField, ActiveContentFieldKind, BarcodeField, BidiOutlineField, EmbedField,
    GoToButtonField, LegacyFormField, LegacyFormFieldKind, MacroButtonField, PrintField,
    ShapeField,
};
pub use links::{
    DdeField, DdeFieldKind, DdeRepresentation, ExternalIncludeField, ExternalIncludeOption,
    IncludeFieldKind, LinkField, LinkFormatting, LinkResultOption,
};
pub use mail_merge::{MailMergeDataField, MergeField, MergeFieldSwitch};
pub use mail_merge_control::{
    AddressBlockCountryInclusion, AdvanceField, AdvanceFieldAdjustment, AdvanceFieldOperation,
    CompareField, IfField, MailMergeConditionalControlField, MailMergeConditionalControlKind,
    MailMergeCounterField, MailMergeCounterKind, MailMergeNextField, MailMergeRecipientField,
    MailMergeRecipientFieldKind, PromptField, PromptFieldKind, UserIdentityField,
    UserIdentityFieldKind, UserIdentityFormatting,
};
pub use numbering::{
    AutoNumberField, AutoNumberFieldKind, AutoTextField, AutoTextFieldKind, AutoTextListField,
    AutoTextListOption, ListNumberField, SequenceField, StyleReferenceField,
    StyleReferenceFieldOption,
};
pub use references::{
    EquationField, FormulaField, HyperlinkField, IndexField, IndexOption, QuoteField,
    ReferenceField, ReferenceFieldKind, ReferenceFieldOption, SetField, SymbolField,
    TableOfAuthoritiesField, TableOfAuthoritiesOption,
};

pub(crate) use core::non_plcf_field_texts;
pub(in crate::parts::fields) use core::{CP_SIZE, FLD_SIZE, MAX_FIELD_MARKERS, MAX_PLCFLD_BYTES};

use crate::package::Error as PackageError;

pub(super) fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
