//! The small, source-backed RTF vocabulary currently exposed by the unified
//! adapter.

use crate::node::spec::{
    Access, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::Tokens;

/// A RTF node kind.
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
    const VOCABULARY: &'static Vocabulary = &vocabulary::RTF;
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

pub(crate) static ALIGNMENT: Tokens =
    Tokens::new("alignment", &["left", "right", "center", "justify"], &[]);

pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &[
    PropSpec {
        name: "text",
        ty: crate::value::ValueType::Text,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
    },
    PropSpec {
        name: "align",
        ty: crate::value::ValueType::Token(&ALIGNMENT),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
];

pub(crate) const fn alignment_name(alignment: litchi_rtf::Alignment) -> &'static str {
    match alignment {
        litchi_rtf::Alignment::Left => "left",
        litchi_rtf::Alignment::Right => "right",
        litchi_rtf::Alignment::Center => "center",
        litchi_rtf::Alignment::Justify => "justify",
    }
}

const PARAGRAPH: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Paragraph,
    name: "paragraph",
    aliases: &["p", "paragraph"],
    parents: &[Kind::Body],
    children: &[],
    identities: &[],
    paths: &["/body/p[N]"],
    props: PARAGRAPH_PROPS,
    gap: None,
};

/// The one static table shared by path resolution, query compilation and the
/// node-view projection.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[DOCUMENT, BODY, PARAGRAPH],
    gaps: &[
        read_gap("run"),
        read_gap("table"),
        read_gap("row"),
        read_gap("cell"),
        read_gap("header"),
        read_gap("footer"),
        read_gap("footnote"),
        read_gap("endnote"),
        read_gap("comment"),
        read_gap("picture"),
        read_gap("shape"),
        read_gap("hyperlink"),
        read_gap("style"),
        read_gap("numbering"),
    ],
};

check_format_spec!(SPEC);

const _: () = assert!(
    SPEC.elements.len() <= 128,
    "a query kind set holds 128 kinds"
);

const fn read_gap(element: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: crate::Code::UnsupportedCapability,
        follow_up: "unified RTF retained-node projection",
        reason: "this native RTF concept is not yet integrated into generic paths and queries",
    }
}
