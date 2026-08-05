//! Semantic ODF variable declaration model.

use crate::datatype::DurationValue;
use chrono::{DateTime, FixedOffset, NaiveDate};

/// XML part containing a declaration group.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Part {
    Content,
    Styles,
    Flat,
}

/// Standard body family containing declarations.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Body {
    Text,
    Spreadsheet,
    Presentation,
    Drawing,
    Chart,
}

/// Header or footer variant containing declarations.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum HeaderFooter {
    Header,
    HeaderFirst,
    HeaderLeft,
    Footer,
    FooterFirst,
    FooterLeft,
}

/// Structural scope in which a declaration group occurs.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum Scope {
    Body(Body),
    HeaderFooter {
        kind: HeaderFooter,
        master_page_name: Option<String>,
    },
}

/// One of the three variable classes defined by ODF 1.3 section 7.4.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    Simple,
    User,
    Sequence,
}

/// Declared ODF value type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ValueType {
    Float,
    Percentage,
    Currency,
    Date,
    Time,
    Boolean,
    String,
    Void,
}

/// Typed date or date-time value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DateValue {
    Date(NaiveDate),
    DateTime(DateTime<FixedOffset>),
}

/// Typed user-field value retaining its exact lexical representation.
#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Float {
        value: f64,
        lexical: String,
    },
    Percentage {
        value: f64,
        lexical: String,
    },
    Currency {
        value: f64,
        lexical: String,
        currency: String,
    },
    Date {
        value: DateValue,
        lexical: String,
    },
    Time {
        value: DurationValue,
        lexical: String,
    },
    Boolean {
        value: bool,
        lexical: String,
    },
    String {
        value: String,
    },
    Void,
}

impl Value {
    pub fn value_type(&self) -> ValueType {
        match self {
            Self::Float { .. } => ValueType::Float,
            Self::Percentage { .. } => ValueType::Percentage,
            Self::Currency { .. } => ValueType::Currency,
            Self::Date { .. } => ValueType::Date,
            Self::Time { .. } => ValueType::Time,
            Self::Boolean { .. } => ValueType::Boolean,
            Self::String { .. } => ValueType::String,
            Self::Void => ValueType::Void,
        }
    }

    pub fn lexical(&self) -> &str {
        match self {
            Self::Float { lexical, .. }
            | Self::Percentage { lexical, .. }
            | Self::Currency { lexical, .. }
            | Self::Date { lexical, .. }
            | Self::Time { lexical, .. }
            | Self::Boolean { lexical, .. } => lexical,
            Self::String { value } => value,
            Self::Void => "",
        }
    }
}

/// One variable declaration.
#[derive(Clone, Debug, PartialEq)]
#[allow(clippy::large_enum_variant)] // public API; boxing would break callers
pub enum Declaration {
    Simple {
        name: String,
        value_type: ValueType,
    },
    User {
        name: String,
        value: Option<Value>,
        formula: Option<String>,
    },
    Sequence {
        name: String,
        display_outline_level: u8,
        separation_character: Option<char>,
    },
}

impl Declaration {
    pub fn kind(&self) -> Kind {
        match self {
            Self::Simple { .. } => Kind::Simple,
            Self::User { .. } => Kind::User,
            Self::Sequence { .. } => Kind::Sequence,
        }
    }

    pub fn name(&self) -> &str {
        match self {
            Self::Simple { name, .. } | Self::User { name, .. } | Self::Sequence { name, .. } => {
                name
            },
        }
    }

    /// Effective separator. Nonzero sequence levels default to `.`.
    pub fn effective_separation_character(&self) -> Option<char> {
        match self {
            Self::Sequence {
                display_outline_level: 1..=10,
                separation_character,
                ..
            } => Some(separation_character.unwrap_or('.')),
            _ => None,
        }
    }
}

/// One declaration container in source order.
#[derive(Clone, Debug, PartialEq)]
pub struct Group {
    pub kind: Kind,
    pub part: Part,
    pub scope: Scope,
    pub declarations: Vec<Declaration>,
}

/// Ordered declaration groups from all scanned XML parts.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Declarations {
    pub groups: Vec<Group>,
    /// Inert DDE source declarations in document order.
    pub dde_connections: Vec<crate::DdeConnectionDeclaration>,
    /// Validated references to DDE declarations in document order.
    pub dde_connection_uses: Vec<crate::DdeConnectionUse>,
    /// Optional document-wide bibliography formatting and sorting policy from styles metadata.
    pub bibliography_configuration: Option<crate::BibliographyConfiguration>,
    /// Inert `text:alphabetical-index-auto-mark-file` references in document order.
    pub auto_mark_files: Vec<crate::AlphabeticalIndexAutoMarkFile>,
}

impl Declarations {
    pub fn declarations(&self) -> impl Iterator<Item = &Declaration> {
        self.groups
            .iter()
            .flat_map(|group| group.declarations.iter())
    }

    pub fn find(&self, kind: Kind, name: &str) -> Option<&Declaration> {
        self.declarations()
            .find(|declaration| declaration.kind() == kind && declaration.name() == name)
    }

    pub fn find_dde_connection(&self, name: &str) -> Option<&crate::DdeConnectionDeclaration> {
        self.dde_connections
            .iter()
            .find(|connection| connection.name == name)
    }
}
