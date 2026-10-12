//! The small, source-backed RTF vocabulary currently exposed by the unified
//! adapter.

use crate::node::spec::{
    Access, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

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
    /// A contiguous character-format run (`/body/p[N]/r[N]`).
    Run,
}

impl Kind {
    /// The OfficeCLI type name.
    #[must_use]
    pub const fn type_name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "paragraph",
            Self::Run => "run",
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
            "r" | "run" => Some(Self::Run),
            _ => None,
        }
    }

    fn name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "p",
            Self::Run => "r",
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

/// The complete local underline vocabulary retained by the native RTF
/// formatter.  Every native variant has its own token so a read never
/// collapses a dotted, dashed, or wave underline into the generic `single`
/// value used by some other document families.
pub(crate) static RUN_UNDERLINE: Tokens = Tokens::new(
    "underline",
    &[
        "none",
        "single",
        "double",
        "dotted",
        "dashed",
        "dashDot",
        "dashDotDot",
        "words",
        "thick",
        "wave",
        "hairline",
        "thickDotted",
        "thickDashed",
        "thickDashDot",
        "thickDashDotDot",
        "thickLongDash",
        "longDash",
        "heavyWave",
        "doubleWave",
    ],
    &[],
);

/// Nonnegative paragraph spacing, read as Word twips and printed in points.
pub(crate) const SPACING: LengthSpec =
    LengthSpec::new(LengthContext::WordSpacing, LengthStyle::Points);

/// Signed paragraph offsets, read with the ordinary signed length coercion
/// and printed in points.
pub(crate) const SIGNED_POINTS: LengthSpec =
    LengthSpec::new(LengthContext::Emu, LengthStyle::Points);

pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &[
    PropSpec {
        name: "text",
        ty: ValueType::Text,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
    },
    PropSpec {
        name: "align",
        ty: ValueType::Token(&ALIGNMENT),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "spaceBefore",
        ty: ValueType::Length(SPACING),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "spaceAfter",
        ty: ValueType::Length(SPACING),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "indent",
        ty: ValueType::Length(SIGNED_POINTS),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "rightIndent",
        ty: ValueType::Length(SIGNED_POINTS),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "firstLineIndent",
        ty: ValueType::Length(SIGNED_POINTS),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "keepNext",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "keepLines",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "pageBreakBefore",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
];

/// Local character properties for a native RTF run.  The scalar types follow
/// the common word-processing run registry, while the native RTF projection
/// keeps its own exact underline token vocabulary.  The ten format facets are
/// writable through the native omission-preserving run patch; text, font and
/// foreground colour remain read-only until their source-preserving editors
/// exist.
pub(crate) const RUN_PROPS: &[PropSpec] = &[
    PropSpec {
        name: "text",
        ty: ValueType::Text,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "font",
        ty: ValueType::Text,
        access: Access::READ,
        raw: false,
    },
    PropSpec {
        name: "size",
        ty: ValueType::Length(LengthSpec::FONT_SIZE),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "bold",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "italic",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "color",
        ty: ValueType::Color(ColorContext::Word),
        access: Access::READ,
        raw: false,
    },
    PropSpec {
        name: "underline",
        ty: ValueType::Token(&RUN_UNDERLINE),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "strike",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "dstrike",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "vanish",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "caps",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "smallcaps",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
    },
    PropSpec {
        name: "outline",
        ty: ValueType::Bool,
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

pub(crate) const fn underline_name(underline: litchi_rtf::text::Underline) -> &'static str {
    match underline {
        litchi_rtf::text::Underline::None => "none",
        litchi_rtf::text::Underline::Single => "single",
        litchi_rtf::text::Underline::Double => "double",
        litchi_rtf::text::Underline::Dotted => "dotted",
        litchi_rtf::text::Underline::Dashed => "dashed",
        litchi_rtf::text::Underline::DashDot => "dashDot",
        litchi_rtf::text::Underline::DashDotDot => "dashDotDot",
        litchi_rtf::text::Underline::Words => "words",
        litchi_rtf::text::Underline::Thick => "thick",
        litchi_rtf::text::Underline::Wave => "wave",
        litchi_rtf::text::Underline::Hairline => "hairline",
        litchi_rtf::text::Underline::ThickDotted => "thickDotted",
        litchi_rtf::text::Underline::ThickDashed => "thickDashed",
        litchi_rtf::text::Underline::ThickDashDot => "thickDashDot",
        litchi_rtf::text::Underline::ThickDashDotDot => "thickDashDotDot",
        litchi_rtf::text::Underline::ThickLongDash => "thickLongDash",
        litchi_rtf::text::Underline::LongDash => "longDash",
        litchi_rtf::text::Underline::HeavyWave => "heavyWave",
        litchi_rtf::text::Underline::DoubleWave => "doubleWave",
    }
}

const PARAGRAPH: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Paragraph,
    name: "paragraph",
    aliases: &["p", "paragraph"],
    parents: &[Kind::Body],
    children: &[Kind::Run],
    identities: &[],
    paths: &["/body/p[N]"],
    props: PARAGRAPH_PROPS,
    gap: None,
};

const RUN: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Run,
    name: "run",
    aliases: &["r", "run"],
    parents: &[Kind::Paragraph],
    children: &[],
    identities: &[],
    paths: &["/body/p[N]/r[N]"],
    props: RUN_PROPS,
    gap: None,
};

/// The one static table shared by path resolution, query compilation and the
/// node-view projection.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[DOCUMENT, BODY, PARAGRAPH, RUN],
    gaps: &[
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
