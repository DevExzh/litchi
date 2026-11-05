//! The property names [`super::SummaryProperties::set`] recognizes.
//!
//! Lookup compares ASCII-case-insensitively and skips ASCII spaces in the
//! caller's name, without allocating. The keys below are therefore written in
//! lowercase without spaces.

/// A `SummaryInformation` text property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SummaryText {
    Title,
    Subject,
    Author,
    Keywords,
    Comments,
    Template,
    LastAuthor,
    RevisionNumber,
    ApplicationName,
}

/// A `SummaryInformation` FILETIME timestamp.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SummaryTime {
    Created,
    LastSaved,
    LastPrinted,
}

/// A `DocumentSummaryInformation` text property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DocumentSummaryText {
    Category,
    PresentationFormat,
    Manager,
    Company,
    ContentType,
    ContentStatus,
    Language,
    DocumentVersion,
}

impl SummaryText {
    /// The PIDSI identifier this property is stored under.
    pub(super) const fn identifier(self) -> u32 {
        match self {
            Self::Title => 2,
            Self::Subject => 3,
            Self::Author => 4,
            Self::Keywords => 5,
            Self::Comments => 6,
            Self::Template => 7,
            Self::LastAuthor => 8,
            Self::RevisionNumber => 9,
            Self::ApplicationName => 18,
        }
    }
}

impl SummaryTime {
    /// The PIDSI identifier this timestamp is stored under.
    pub(super) const fn identifier(self) -> u32 {
        match self {
            Self::LastPrinted => 11,
            Self::Created => 12,
            Self::LastSaved => 13,
        }
    }
}

impl DocumentSummaryText {
    /// The PIDDSI identifier this property is stored under.
    pub(super) const fn identifier(self) -> u32 {
        match self {
            Self::Category => 2,
            Self::PresentationFormat => 3,
            Self::Manager => 14,
            Self::Company => 15,
            Self::ContentType => 26,
            Self::ContentStatus => 27,
            Self::Language => 28,
            Self::DocumentVersion => 29,
        }
    }
}

/// What a recognized name sets.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Name {
    Summary(SummaryText),
    SummaryTime(SummaryTime),
    DocumentSummary(DocumentSummaryText),
    /// An Office built-in property that is not text (a count, a duration,
    /// the security flags) or not a plain value (the hyperlink base).
    NotText,
}

const NAMES: &[(&str, Name)] = &[
    ("title", Name::Summary(SummaryText::Title)),
    ("subject", Name::Summary(SummaryText::Subject)),
    ("author", Name::Summary(SummaryText::Author)),
    ("keywords", Name::Summary(SummaryText::Keywords)),
    ("comments", Name::Summary(SummaryText::Comments)),
    ("description", Name::Summary(SummaryText::Comments)),
    ("template", Name::Summary(SummaryText::Template)),
    ("lastauthor", Name::Summary(SummaryText::LastAuthor)),
    ("lastmodifiedby", Name::Summary(SummaryText::LastAuthor)),
    ("revisionnumber", Name::Summary(SummaryText::RevisionNumber)),
    (
        "applicationname",
        Name::Summary(SummaryText::ApplicationName),
    ),
    ("creationdate", Name::SummaryTime(SummaryTime::Created)),
    ("created", Name::SummaryTime(SummaryTime::Created)),
    ("lastsavetime", Name::SummaryTime(SummaryTime::LastSaved)),
    ("modified", Name::SummaryTime(SummaryTime::LastSaved)),
    ("lastprintdate", Name::SummaryTime(SummaryTime::LastPrinted)),
    ("lastprinted", Name::SummaryTime(SummaryTime::LastPrinted)),
    (
        "category",
        Name::DocumentSummary(DocumentSummaryText::Category),
    ),
    (
        "format",
        Name::DocumentSummary(DocumentSummaryText::PresentationFormat),
    ),
    (
        "presentationformat",
        Name::DocumentSummary(DocumentSummaryText::PresentationFormat),
    ),
    (
        "manager",
        Name::DocumentSummary(DocumentSummaryText::Manager),
    ),
    (
        "company",
        Name::DocumentSummary(DocumentSummaryText::Company),
    ),
    (
        "contenttype",
        Name::DocumentSummary(DocumentSummaryText::ContentType),
    ),
    (
        "contentstatus",
        Name::DocumentSummary(DocumentSummaryText::ContentStatus),
    ),
    (
        "language",
        Name::DocumentSummary(DocumentSummaryText::Language),
    ),
    (
        "documentversion",
        Name::DocumentSummary(DocumentSummaryText::DocumentVersion),
    ),
    // Office's remaining built-in names (the `BuiltinDocumentProperties`
    // collection): their values are not text, so they are set through the
    // typed edits instead.
    ("totaleditingtime", Name::NotText),
    ("numberofpages", Name::NotText),
    ("numberofwords", Name::NotText),
    ("numberofcharacters", Name::NotText),
    ("security", Name::NotText),
    ("numberofbytes", Name::NotText),
    ("numberoflines", Name::NotText),
    ("numberofparagraphs", Name::NotText),
    ("numberofslides", Name::NotText),
    ("numberofnotes", Name::NotText),
    ("numberofhiddenslides", Name::NotText),
    ("numberofmultimediaclips", Name::NotText),
    ("hyperlinkbase", Name::NotText),
    ("numberofcharacters(withspaces)", Name::NotText),
];

/// Resolves a caller's property name, or `None` for a user-defined name.
pub(super) fn lookup(name: &str) -> Option<Name> {
    NAMES
        .iter()
        .find(|(key, _)| matches_key(name, key))
        .map(|(_, target)| *target)
}

/// Whether `name` equals the lowercase, space-free `key` once ASCII case and
/// ASCII spaces are ignored.
fn matches_key(name: &str, key: &str) -> bool {
    let mut candidate = name.bytes().filter(|byte| *byte != b' ');
    let mut expected = key.bytes();
    loop {
        match (candidate.next(), expected.next()) {
            (None, None) => return true,
            (Some(left), Some(right)) if left.to_ascii_lowercase() == right => {},
            _ => return false,
        }
    }
}
