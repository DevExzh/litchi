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
    Switch, Symbol, Toa, ToaEntry, Toc, TocEntry, TocLevelRange, UserIdentity,
    UserIdentityFormat, UserIdentityKind, Variable,
};
