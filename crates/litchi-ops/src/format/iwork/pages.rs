//! Apple Pages documents (`.pages`): the body, its paragraphs and the
//! sections, read from `litchi_pages::Package` and written through its
//! section-text transaction (record 0296).
//!
//! A paragraph is one line of its section's body text: Pages separates
//! paragraphs with a newline in the section's text storage. Paragraph text
//! is written by replacing that line's UTF-16 span in the section, so the
//! section's other paragraphs, its section break and every unmodeled native
//! record stay as the leaf preserves them.

use litchi_core::FileFormat;
use litchi_pages::{Package, SectionSelector, TextSpan};

use super::tree::{count, text};
use super::write::{
    Chain, Commit, LeafPatch, bound, element_refusal, finish, key_refusal, leaf_refusal, refusal,
    resolve,
};
use super::{Flavor, Iwork, LeafError, Node, Owner, TreeBuilder, element, writable};
use crate::Result;
use crate::error::Code;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::ValueType;

kinds! {
    /// The element kinds of a Pages document.
    Kind: &crate::path::vocabulary::WORD, root Document {
        Document => "document",
        Body => "body",
        Paragraph => "p",
        Section => "section",
    }
}

const DOCUMENT_PROPS: &[PropSpec] = &[
    PropSpec::new("sectionCount", ValueType::Int),
    PropSpec::new("paragraphCount", ValueType::Int),
];

const BODY_PROPS: &[PropSpec] = &[PropSpec::new("paragraphCount", ValueType::Int)];

pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &[
    writable("text", ValueType::Text),
    PropSpec::new("section", ValueType::Int),
];

pub(crate) const SECTION_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    writable("text", ValueType::Text),
    PropSpec::new("paragraphCount", ValueType::Int),
];

static GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Element("r"),
        code: Code::UnsupportedCapability,
        follow_up: "0297 Pages text runs (IWORK_API_REQUESTS.md request 13)",
        reason: "the Pages leaf exposes body text as plain section text; character runs and their formatting have no public model yet",
    },
    GapSpec {
        subject: Subject::Element("tbl"),
        code: Code::UnsupportedCapability,
        follow_up: "0297 Pages body tables",
        reason: "the Pages leaf's body-table catalog and cell APIs are not wired into the unified tree yet",
    },
    GapSpec {
        subject: Subject::Element("metadata"),
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 14",
        reason: "Pages document properties are read through File::metadata(); the leaf has no metadata write transaction",
    },
];

static OPERATION_GAPS: [GapSpec; 4] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "p",
            context: "a new paragraph",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 12",
        reason: "the Pages leaf has no paragraph insertion transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "p",
            context: "a paragraph",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 12",
        reason: "the Pages leaf has no paragraph removal transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "section",
            context: "a new section",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 11",
        reason: "the Pages leaf has no section lifecycle transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "a document from scratch",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 1",
        reason: "the Pages leaf has no from-scratch package builder",
    },
];

/// The Pages spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[
        element(
            Kind::Document,
            "document",
            &["document"],
            &[],
            &[Kind::Body, Kind::Section],
            &["/"],
            DOCUMENT_PROPS,
        ),
        element(
            Kind::Body,
            "body",
            &["body"],
            &[Kind::Document],
            &[Kind::Paragraph],
            &["/body"],
            BODY_PROPS,
        ),
        element(
            Kind::Paragraph,
            "paragraph",
            &["p", "paragraph"],
            &[Kind::Body],
            &[],
            &["/body/p[N]"],
            PARAGRAPH_PROPS,
        ),
        element(
            Kind::Section,
            "section",
            &["section"],
            &[Kind::Document],
            &[],
            &["/section[N]"],
            SECTION_PROPS,
        ),
    ],
    gaps: &GAPS,
};

crate::node::spec::check_format_spec!(SPEC);

/// The Pages flavor.
pub(crate) struct Pages;

impl Flavor for Pages {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Pages;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Body];
    const SINGLETONS: &'static [Kind] = &[Kind::Body];
}

/// The elements `set` writes, `add` creates and `remove` deletes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["p", "section"],
    addable: &[],
    removable: &[],
    movable: &[],
    gaps: &OPERATION_GAPS,
};

/// The text of section `index`: the section's own body text when the leaf
/// can address it, else its rendered plain text.
fn section_text(package: &Package, index: usize) -> String {
    match package.section_text(SectionSelector::index(index)) {
        Ok(text) => text.to_owned(),
        Err(_) => package
            .sections()
            .get(index)
            .map(litchi_pages::section::Section::plain_text)
            .unwrap_or_default(),
    }
}

/// Build the tree of `package`.
pub(crate) fn build(package: &Package) -> super::Tree {
    let mut tree = TreeBuilder::new(SPEC.elements, Pages::SINGLETONS);
    let body = tree.push(super::tree::ROOT, Kind::Body).done();
    let mut paragraphs = 0_usize;
    let sections = package.sections();
    let mut section_items = Vec::with_capacity(sections.len());
    for (index, section) in sections.iter().enumerate() {
        let value = section_text(package, index);
        let mut lines = 0_usize;
        for (line, paragraph) in value.split('\n').enumerate() {
            tree.push(body, Kind::Paragraph)
                .prop("text", text(Some(paragraph)))
                .prop("section", count(index + 1))
                .text(paragraph)
                .leaf([index, line, 0, 0])
                .done();
            lines += 1;
        }
        paragraphs += lines;
        section_items.push((index, section.name().map(str::to_owned), value, lines));
    }
    for (index, name, value, lines) in section_items {
        tree.push(super::tree::ROOT, Kind::Section)
            .prop("name", text(name.as_deref()))
            .prop("text", text(Some(&value)))
            .prop("paragraphCount", count(lines))
            .label(name.as_deref())
            .text(value)
            .leaf([index, 0, 0, 0])
            .done();
    }
    tree.root()
        .prop("sectionCount", count(sections.len()))
        .prop("paragraphCount", count(paragraphs))
        .done();
    tree.draft(body)
        .prop("paragraphCount", count(paragraphs))
        .done();
    tree.build()
}

/// The UTF-16 span of line `line` of `value`, without its newline.
fn line_span(value: &str, line: usize) -> Option<TextSpan> {
    let mut start = 0_usize;
    for (index, part) in value.split('\n').enumerate() {
        let length = part.encode_utf16().count();
        if index == line {
            return TextSpan::from_utf16_indexes(start, start + length).ok();
        }
        start += length + 1;
    }
    None
}

/// Set paragraph or section text on every target.
pub(crate) fn set(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Pages>(file, target, limits.max_expanded_targets())?;
    let mut chain = Chain::new(package.clone());
    let mut operations = 0_usize;
    for node in &nodes {
        for (key, value) in props.iter() {
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            let key = key.as_str();
            if key != "text" {
                return Err(key_refusal(node, key));
            }
            let value = super::write::text(value, "text")?.unwrap_or_default();
            stage(&mut chain, node, &value)?;
        }
    }
    finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        operations,
        limits,
    )
}

/// Commit one text write on the chain's current package.
fn stage(chain: &mut Chain<Package>, node: &Node<'_>, value: &str) -> Result<()> {
    let [section, line, _, _] = node.item().leaf().map(|value| value as usize);
    let selector = SectionSelector::index(section);
    let error = |error| leaf_refusal(LeafError::PagesText(error), FileFormat::Pages);
    let commit = match node.item().type_name() {
        "paragraph" => {
            if value.contains('\n') {
                return Err(refusal(
                    Code::InvalidValue,
                    "a paragraph's text cannot contain a newline, which separates Pages paragraphs",
                )
                .with_property("text".to_owned())
                .with_suggestion("set each paragraph separately, or set the section's `text`"));
            }
            let current = chain.package.section_text(selector).map_err(error)?;
            let span = line_span(current, line).ok_or_else(|| {
                refusal(
                    Code::PathNotFound,
                    "the paragraph is no longer in its section",
                )
            })?;
            let mut edit = chain.package.edit_section_text(selector).map_err(error)?;
            edit.replace(span, value).map_err(error)?;
            edit.commit().map_err(error)?
        },
        "section" => chain
            .package
            .set_section_text(selector, value)
            .map_err(error)?,
        _ => return Err(element_refusal("set", node)),
    };
    let patch = commit.patch().clone();
    let changed = !patch.is_noop();
    chain.push(commit.into_package(), LeafPatch::PagesText(patch), changed);
    Ok(())
}
