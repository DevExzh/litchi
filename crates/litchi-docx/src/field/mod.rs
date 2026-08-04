//! Layered WordprocessingML field models and bounded codecs.
//!
//! The public module remains the historical `litchi_docx::field` entry point.
//! Field values and typed instruction metadata live in [`model`], while
//! bounded Word instruction and document-XML parsing lives in [`codec`].

mod codec;
mod model;

#[cfg(test)]
mod tests;

use crate::error::{Error, Result};

const MAX_FIELD_SWITCHES: usize = 64;
const MAX_FORMULA_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_QUOTE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_SYMBOL_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_SET_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_EMBED_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_BARCODE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_SHAPE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_PRIVATE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DATABASE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_INFO_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_REFERENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_EQUATION_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

pub use model::{
    ActiveContent, ActiveContentKind, Advance, AdvanceAdjustment, AdvanceOperation, AutoNumber,
    AutoNumberKind, AutoText, AutoTextKind, AutoTextList, AutoTextListOption, Barcode,
    Bibliography, BidiOutline, Citation, Compare, Context, ContextKind, CountryInclusion, Database,
    Dde, DdeFormat, DdeKind, Embed, Equation, Field, Formula, GoToButton, Hyperlink, If, Include,
    IncludeKind, IncludeOption, Index, IndexEntry, IndexOrder, Info, Information, InformationKind,
    LegacyForm, LegacyFormKind, Link, LinkFormat, LinkResult, ListNumber, MacroButton, Merge,
    MergeControl, MergeControlKind, MergeCounter, MergeCounterKind, MergeData, MergeNext, Print,
    Private, Prompt, PromptKind, Property, Quote, Recipient, RecipientKind, Reference,
    ReferenceKind, ReferenceOption, Sequence, Set, Shape, StyleOption, StyleReference, SubDocument,
    Switch, Symbol, Toa, ToaEntry, Toc, TocEntry, TocLevelRange, TocSwitch, UserIdentity,
    UserIdentityFormat, UserIdentityKind, Variable,
};

// Historical names remain aliases so existing callers keep the exact public
// `litchi_docx::field` and root-facade API while the canonical vocabulary is
// contextual to this module.
pub type ActiveContentField = ActiveContent;
pub type ActiveContentFieldKind = ActiveContentKind;
pub type AddressBlockCountryInclusion = CountryInclusion;
pub type AdvanceField = Advance;
pub type AdvanceFieldAdjustment = AdvanceAdjustment;
pub type AdvanceFieldOperation = AdvanceOperation;
pub type AutoNumberField = AutoNumber;
pub type AutoNumberFieldKind = AutoNumberKind;
pub type AutoTextField = AutoText;
pub type AutoTextFieldKind = AutoTextKind;
pub type AutoTextListField = AutoTextList;
pub type BarcodeField = Barcode;
pub type BibliographyField = Bibliography;
pub type BidiOutlineField = BidiOutline;
pub type CitationField = Citation;
pub type CompareField = Compare;
pub type DatabaseField = Database;
pub type DdeField = Dde;
pub type DdeFieldKind = DdeKind;
pub type DdeRepresentation = DdeFormat;
pub type DocumentContextField = Context;
pub type DocumentContextFieldKind = ContextKind;
pub type DocumentInformationField = Information;
pub type DocumentInformationFieldKind = InformationKind;
pub type DocumentPropertyField = Property;
pub type DocumentVariableField = Variable;
pub type EmbedField = Embed;
pub type EquationField = Equation;
pub type ExternalIncludeField = Include;
pub type ExternalIncludeOption = IncludeOption;
pub type FieldSwitch = Switch;
pub type FormulaField = Formula;
pub type GoToButtonField = GoToButton;
pub type HyperlinkField = Hyperlink;
pub type IfField = If;
pub type IncludeFieldKind = IncludeKind;
pub type IndexEntryField = IndexEntry;
pub type IndexField = Index;
pub type IndexSortOrder = IndexOrder;
pub type InfoField = Info;
pub type LegacyFormField = LegacyForm;
pub type LegacyFormFieldKind = LegacyFormKind;
pub type LinkField = Link;
pub type LinkFormatting = LinkFormat;
pub type LinkResultOption = LinkResult;
pub type ListNumberField = ListNumber;
pub type MacroButtonField = MacroButton;
pub type MailMergeConditionalControlField = MergeControl;
pub type MailMergeConditionalControlKind = MergeControlKind;
pub type MailMergeCounterField = MergeCounter;
pub type MailMergeCounterKind = MergeCounterKind;
pub type MailMergeDataField = MergeData;
pub type MailMergeNextField = MergeNext;
pub type MailMergeRecipientField = Recipient;
pub type MailMergeRecipientFieldKind = RecipientKind;
pub type MergeField = Merge;
pub type PrintField = Print;
pub type PrivateField = Private;
pub type PromptField = Prompt;
pub type PromptFieldKind = PromptKind;
pub type QuoteField = Quote;
pub type ReferenceField = Reference;
pub type ReferenceFieldKind = ReferenceKind;
pub type ReferenceFieldOption = ReferenceOption;
pub type ReferencedDocumentField = SubDocument;
pub type SequenceField = Sequence;
pub type SetField = Set;
pub type ShapeField = Shape;
pub type StyleReferenceField = StyleReference;
pub type StyleReferenceFieldOption = StyleOption;
pub type SymbolField = Symbol;
pub type TableOfAuthoritiesEntryField = ToaEntry;
pub type TableOfAuthoritiesField = Toa;
pub type TableOfContentsEntryField = TocEntry;
pub type TableOfContentsField = Toc;
pub type TableOfContentsLevelRange = TocLevelRange;
pub type TableOfContentsSwitch = TocSwitch;
pub type UserIdentityField = UserIdentity;
pub type UserIdentityFieldKind = UserIdentityKind;
pub type UserIdentityFormatting = UserIdentityFormat;
