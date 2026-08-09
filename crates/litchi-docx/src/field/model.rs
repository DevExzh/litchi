//! Contextual, inert Word field values and typed instruction models.

mod base;
mod bibliography;
mod building_blocks;
mod buttons;
mod conditionals;
mod document;
mod external;
mod formatting;
mod generated;
mod identity;
mod mail_merge;
mod opaque;
mod reference;

pub use base::Field;
pub use bibliography::{Bibliography, Citation};
pub use building_blocks::{AutoText, AutoTextKind, AutoTextList, AutoTextListOption};
pub use buttons::{GoToButton, MacroButton};
pub use conditionals::{Compare, Equation, Formula, If, Sequence, Set};
pub use document::{Context, ContextKind, Info, Information, InformationKind, Property, Variable};
pub use external::{
    Dde, DdeFormat, DdeKind, Include, IncludeKind, IncludeOption, Link, LinkFormat, LinkResult,
    SubDocument,
};
pub use formatting::{
    AutoNumber, AutoNumberKind, ListNumber, Quote, StyleOption, StyleReference, Symbol,
};
pub use generated::{Index, IndexEntry, IndexOrder, Toa, ToaEntry, Toc, TocEntry, TocLevelRange};
pub use identity::{
    Advance, AdvanceAdjustment, AdvanceOperation, UserIdentity, UserIdentityFormat,
    UserIdentityKind,
};
pub use mail_merge::{
    CountryInclusion, Merge, MergeControl, MergeControlKind, MergeCounter, MergeCounterKind,
    MergeData, MergeNext, Prompt, PromptKind, Recipient, RecipientKind,
};
pub use opaque::{
    ActiveContent, ActiveContentKind, Barcode, BidiOutline, Database, Embed, LegacyForm,
    LegacyFormKind, Print, Private, Shape,
};
pub use reference::{Hyperlink, Reference, ReferenceKind, ReferenceOption};

/// One lexical switch in a Word field instruction.
///
/// Switch names are normalized to ASCII lowercase. Quoted and unquoted
/// arguments are decoded into their logical text. Typed field models retain the
/// complete original instruction alongside these values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Switch {
    pub(super) name: char,
    pub(super) argument: Option<String>,
}

impl Switch {
    /// Return the switch character, without its leading backslash.
    #[must_use]
    pub fn name(&self) -> char {
        self.name
    }

    /// Return the optional argument supplied to this switch.
    #[must_use]
    pub fn argument(&self) -> Option<&str> {
        self.argument.as_deref()
    }
}
