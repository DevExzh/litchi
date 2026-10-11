//! The small, source-backed DOC vocabulary currently exposed by the unified
//! adapter.

use crate::node::spec::{ElementSpec, FormatSpec, PropSpec, check_format_spec};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};

/// A DOC node kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The document root (`/`).
    Document,
    /// The main story (`/body`).
    Body,
    /// An ordinary main-story paragraph (`/body/p[N]`).
    Paragraph,
}

impl Kind {
    /// The OfficeCLI type name.
    #[must_use]
    pub const fn type_name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "paragraph",
        }
    }
}

impl ElementKind for Kind {
    const VOCABULARY: &'static Vocabulary = &vocabulary::DOC;
    const ROOT: Self = Self::Document;

    fn from_name(name: &str) -> Option<Self> {
        match name {
            "document" => Some(Self::Document),
            "body" => Some(Self::Body),
            "p" | "paragraph" => Some(Self::Paragraph),
            _ => None,
        }
    }

    fn name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "p",
        }
    }

    fn ordinal(self) -> u8 {
        self as u8
    }
}

const DOCUMENT: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Document,
    name: "document",
    aliases: &["document"],
    parents: &[],
    children: &[Kind::Body],
    identities: &[],
    paths: &["/"],
    props: &[],
    gap: None,
};

const BODY: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Body,
    name: "body",
    aliases: &["body"],
    parents: &[Kind::Document],
    children: &[Kind::Paragraph],
    identities: &[],
    paths: &["/body"],
    props: &[],
    gap: None,
};

const PARAGRAPH: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Paragraph,
    name: "paragraph",
    aliases: &["p", "paragraph"],
    parents: &[Kind::Body],
    children: &[],
    identities: &[],
    paths: &["/body/p[N]"],
    props: &[writable_text()],
    gap: None,
};

const fn writable_text() -> PropSpec {
    let mut spec = PropSpec::new("text", crate::value::ValueType::Text);
    spec.access.set = true;
    spec
}

/// The one static table shared by path resolution, query compilation and the
/// node-view projection.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[DOCUMENT, BODY, PARAGRAPH],
    gaps: &[],
};

check_format_spec!(SPEC);

const _: () = assert!(
    SPEC.elements.len() <= 128,
    "a query kind set holds 128 kinds"
);
