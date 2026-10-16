//! Apple Pages documents (`.pages`): the body's paragraphs and tables, the
//! sections with their headers and footers, and the body footnotes, read
//! from `litchi_pages::Package` and written through its focused
//! transactions (records 0296 and 0299).
//!
//! A paragraph is one line of its section's body text: Pages separates
//! paragraphs with a newline in the section's text storage. The paragraph
//! writes (`set` of `text`, `add`, `remove`, `move`, `swap`) run the leaf's
//! paragraph transaction (`litchi_pages::section::paragraphs`), which turns
//! them into exact-source section-text splices: one splice for a whole run
//! of changes when it consumes no object anchor. A paragraph's `text`
//! excludes the anchor characters of the objects it holds (an inline table,
//! chart, image, or footnote mark, counted by `objects`); setting it keeps
//! those objects. Removing or moving a paragraph that holds an object is
//! refused: the object's own capability deletes or moves it.
//!
//! Tables are read from the leaf's body-table catalog and cell reader
//! (`/body/tbl[N]/tr[R]/tc[C]`) and removed through its deletion
//! transaction; cell writes are a typed gap. Section names, header and
//! footer text, footnote text and the document properties
//! (`set("/metadata", ..)`) write through their leaf families.

use std::collections::BTreeMap;

use litchi_core::{FileFormat, Position};
use litchi_pages::header_footer::{HeaderFooterSelector, Kind as HeaderKind, Template};
use litchi_pages::{BodyTableSelector, Package};

use super::tree::{ROOT, count, text};
use super::write::{
    Chain, Commit, LeafPatch, bound, element_refusal, finish, key_refusal, leaf_refusal, refusal,
    resolve,
};
use super::{Flavor, Iwork, LeafError, Node, Owner, TreeBuilder, element, writable};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::{Path, PathArg};
use crate::props::{PropValue, Props};
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
        Table => "tbl",
        Row => "tr",
        Cell => "tc",
        Header => "header",
        Footer => "footer",
        Footnotes => "footnotes",
        Footnote => "footnote",
    }
}

const DOCUMENT_PROPS: &[PropSpec] = &[
    PropSpec::new("sectionCount", ValueType::Int),
    PropSpec::new("paragraphCount", ValueType::Int),
    PropSpec::new("tableCount", ValueType::Int),
];

const BODY_PROPS: &[PropSpec] = &[PropSpec::new("paragraphCount", ValueType::Int)];

/// A property `set` writes and `add` accepts.
const fn addable(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = writable(name, ty);
    spec.access.add = true;
    spec
}

pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &[
    addable("text", ValueType::Text),
    PropSpec::new("section", ValueType::Int),
    PropSpec::new("objects", ValueType::Int),
];

pub(crate) const SECTION_PROPS: &[PropSpec] = &[
    writable("name", ValueType::Text),
    writable("text", ValueType::Text),
    PropSpec::new("paragraphCount", ValueType::Int),
];

const TABLE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("rows", ValueType::Int),
    PropSpec::new("columns", ValueType::Int),
];

const ROW_PROPS: &[PropSpec] = &[PropSpec::new("cellCount", ValueType::Int)];

const CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("text", ValueType::Text),
    PropSpec::new("type", ValueType::Text),
];

const REGION_PROPS: &[PropSpec] = &[
    writable("text", ValueType::Text),
    PropSpec::new("template", ValueType::Text),
    PropSpec::new("slot", ValueType::Int),
];

const FOOTNOTES_PROPS: &[PropSpec] = &[PropSpec::new("count", ValueType::Int)];

const FOOTNOTE_PROPS: &[PropSpec] = &[
    addable("text", ValueType::Text),
    PropSpec::new("mark", ValueType::Text),
];

static GAPS: [GapSpec; 2] = [
    GapSpec {
        subject: Subject::Element("r"),
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 13 (Pages text runs and formatting)",
        reason: "the Pages leaf reads and writes body text as plain section text; character runs and their styles have no transaction yet",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "get",
            element: "metadata",
            context: "path-read",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0299 Pages metadata node",
        reason: "`set(\"/metadata\", ..)` writes the document properties through the leaf metadata family; the path is set-only, and reads go through File::metadata()",
    },
];

static OPERATION_GAPS: [GapSpec; 6] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "section",
            context: "a new section",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 11",
        reason: "the Pages leaf has no section lifecycle transaction (a section break plus a new native section object)",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "section",
            context: "a section",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 11",
        reason: "the Pages leaf has no section lifecycle transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "tc",
            context: "a body-table cell",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0299 follow-up: Pages body-table cell writes",
        reason: "the Pages leaf reads body-table cells but has no cell-write transaction (the Numbers cell writer is not shared with Pages)",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "tbl",
            context: "a new body table",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0299 follow-up: Pages body-table creation",
        reason: "the Pages leaf deletes body tables but has no table-creation transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "p",
            context: "a paragraph that anchors an object",
        },
        code: Code::UnsupportedCapability,
        follow_up: "remove the anchored object (`tbl`) itself",
        reason: "removing or moving the paragraph would delete or move the inline table, chart, image or footnote it anchors",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "a document with a fresh identity",
        },
        code: Code::UnsupportedCapability,
        follow_up: "litchi_pages::builder::Builder with Identity::generate",
        reason: "`create` returns the deterministic minimal Pages document (`Package::new`); a fresh document identity is an explicit leaf builder input",
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
            &[Kind::Body, Kind::Section, Kind::Footnotes],
            &["/"],
            DOCUMENT_PROPS,
        ),
        element(
            Kind::Body,
            "body",
            &["body"],
            &[Kind::Document],
            &[Kind::Paragraph, Kind::Table],
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
            &[Kind::Header, Kind::Footer],
            &["/section[N]"],
            SECTION_PROPS,
        ),
        element(
            Kind::Table,
            "table",
            &["table", "tbl"],
            &[Kind::Body],
            &[Kind::Row],
            &["/body/tbl[N]"],
            TABLE_PROPS,
        ),
        element(
            Kind::Row,
            "row",
            &["row", "tr"],
            &[Kind::Table],
            &[Kind::Cell],
            &["/body/tbl[N]/tr[N]"],
            ROW_PROPS,
        ),
        element(
            Kind::Cell,
            "cell",
            &["cell", "tc"],
            &[Kind::Row],
            &[],
            &["/body/tbl[N]/tr[N]/tc[N]"],
            CELL_PROPS,
        ),
        element(
            Kind::Header,
            "header",
            &["header"],
            &[Kind::Section],
            &[],
            &["/section[N]/header[N]"],
            REGION_PROPS,
        ),
        element(
            Kind::Footer,
            "footer",
            &["footer"],
            &[Kind::Section],
            &[],
            &["/section[N]/footer[N]"],
            REGION_PROPS,
        ),
        element(
            Kind::Footnotes,
            "footnotes",
            &["footnotes"],
            &[Kind::Document],
            &[Kind::Footnote],
            &["/footnotes"],
            FOOTNOTES_PROPS,
        ),
        element(
            Kind::Footnote,
            "footnote",
            &["footnote"],
            &[Kind::Footnotes],
            &[],
            &["/footnotes/footnote[N]"],
            FOOTNOTE_PROPS,
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
    const SINGLETONS: &'static [Kind] = &[Kind::Body, Kind::Footnotes];
}

/// The elements `set` writes, `add` creates and `remove` deletes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["p", "section", "header", "footer", "footnote"],
    addable: &["p", "footnote"],
    removable: &["p", "tbl", "footnote"],
    movable: &["p"],
    gaps: &OPERATION_GAPS,
};

/// The native anchors a paragraph's text can hold: an inline object marker
/// and a footnote anchor.
const fn is_anchor(character: char) -> bool {
    matches!(character, '\u{fffc}' | '\u{000e}')
}

/// The leaf position of a paragraph: its section and line.
const SECTION: usize = 0;
const LINE: usize = 1;

/// The proof of one Pages leaf transaction.
#[derive(Debug)]
#[allow(
    clippy::large_enum_variant,
    reason = "a write holds its proofs once; boxing would add an allocation per commit"
)]
pub(crate) enum Patch {
    Text(litchi_pages::section::text::Patch),
    Paragraphs(litchi_pages::section::paragraphs::Patch),
    Name(litchi_pages::section::name::Patch),
    HeaderFooter(litchi_pages::header_footer::text::Patch),
    Footnote(litchi_pages::footnote::text::Patch),
    Footnotes(litchi_pages::footnote::body::transaction::Patch),
    Metadata(litchi_pages::metadata::Patch),
    TableDeletion(
        #[allow(
            dead_code,
            reason = "retained as the leaf's reversible proof; a committed deletion is never a no-op"
        )]
        litchi_pages::table::deletion::Patch,
    ),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Text(patch) => patch.is_noop(),
            Self::Paragraphs(patch) => patch.is_noop(),
            Self::Name(patch) => patch.is_noop(),
            Self::HeaderFooter(patch) => patch.is_noop(),
            Self::Footnote(patch) => patch.is_noop(),
            Self::Footnotes(patch) => patch.is_noop(),
            Self::Metadata(patch) => patch.is_noop(),
            Self::TableDeletion(_) => false,
        }
    }
}

/// Record one committed Pages transaction on the chain.
fn push(chain: &mut Chain<Package>, package: Package, patch: Patch) {
    let changed = !patch.is_noop();
    chain.push(package, LeafPatch::Pages(patch), changed);
}

/// The leaf failure of a Pages transaction.
fn failure(error: LeafError) -> crate::Error {
    leaf_refusal(error, FileFormat::Pages)
}

/// The text of section `index`: the section's own body text when the leaf
/// can address it, else its rendered plain text.
fn section_text(package: &Package, index: usize) -> String {
    match package.section_text(index) {
        Ok(text) => text.to_owned(),
        Err(_) => package
            .sections()
            .get(index)
            .map(litchi_pages::section::Section::plain_text)
            .unwrap_or_default(),
    }
}

/// A paragraph's visible text (without anchors) and its anchor count.
fn visible(paragraph: &str) -> (String, usize) {
    let objects = paragraph.chars().filter(|c| is_anchor(*c)).count();
    if objects == 0 {
        return (paragraph.to_owned(), 0);
    }
    (
        paragraph.chars().filter(|c| !is_anchor(*c)).collect(),
        objects,
    )
}

fn template_name(template: Template) -> &'static str {
    match template {
        Template::First => "first",
        Template::Even => "even",
        Template::Odd => "odd",
    }
}

const fn template_ordinal(template: Template) -> usize {
    template as usize
}

fn template_of(ordinal: u32) -> Template {
    match ordinal {
        0 => Template::First,
        1 => Template::Even,
        _ => Template::Odd,
    }
}

/// Build the tree of `package` without the section headers and footers:
/// reading them decodes every page template, so they join the tree the
/// first time a read reaches a section's children ([`build_full`]).
pub(crate) fn build(package: &Package) -> super::Tree {
    build_with(package, false)
}

/// Build the complete tree of `package`, headers and footers included.
/// Every item of [`build`]'s tree keeps its index: the regions are pushed
/// last.
pub(crate) fn build_full(package: &Package) -> super::Tree {
    build_with(package, true)
}

fn build_with(package: &Package, regions: bool) -> super::Tree {
    let mut tree = TreeBuilder::new(SPEC.elements, Pages::SINGLETONS);
    let body = tree.push(ROOT, Kind::Body).done();
    let mut paragraphs = 0_usize;
    let sections = package.sections();
    let mut section_items = Vec::with_capacity(sections.len());
    for (index, section) in sections.iter().enumerate() {
        let value = section_text(package, index);
        let mut lines = 0_usize;
        for (line, paragraph) in value.split('\n').enumerate() {
            let (shown, objects) = visible(paragraph);
            tree.push(body, Kind::Paragraph)
                .prop("text", text(Some(&shown)))
                .prop("section", count(index + 1))
                .prop("objects", count(objects))
                .text(shown)
                .leaf([index, line, 0, 0])
                .done();
            lines += 1;
        }
        paragraphs += lines;
        section_items.push((index, section.name().map(str::to_owned), value, lines));
    }
    let tables = build_tables(package, &mut tree, body);
    let mut section_indices = Vec::with_capacity(section_items.len());
    for (index, name, value, lines) in section_items {
        let section = tree
            .push(ROOT, Kind::Section)
            .prop("name", text(name.as_deref()))
            .prop("text", text(Some(&value)))
            .prop("paragraphCount", count(lines))
            .label(name.as_deref())
            .text(value)
            .leaf([index, 0, 0, 0])
            .done();
        section_indices.push(section);
    }
    if let Ok(footnotes) = package.body_footnotes() {
        let container = tree
            .push(ROOT, Kind::Footnotes)
            .prop("count", count(footnotes.len()))
            .done();
        for (index, footnote) in footnotes.iter().enumerate() {
            tree.push(container, Kind::Footnote)
                .prop("text", text(Some(&footnote.text)))
                .prop("mark", text(footnote.custom_mark.as_deref()))
                .text(footnote.text.as_ref())
                .leaf([index, 0, 0, 0])
                .done();
        }
    }
    tree.root()
        .prop("sectionCount", count(sections.len()))
        .prop("paragraphCount", count(paragraphs))
        .prop("tableCount", count(tables))
        .done();
    tree.draft(body)
        .prop("paragraphCount", count(paragraphs))
        .done();
    if regions {
        for region in package.header_footers().unwrap_or_default() {
            let index = region.section_position().get();
            let Some(&section) = section_indices.get(index) else {
                continue;
            };
            let kind = match region.kind() {
                HeaderKind::Header => Kind::Header,
                HeaderKind::Footer => Kind::Footer,
            };
            tree.push(section, kind)
                .prop("text", text(Some(region.text())))
                .prop("template", text(Some(template_name(region.template()))))
                .prop("slot", count(region.slot_index()))
                .text(region.text())
                .leaf([
                    index,
                    template_ordinal(region.template()),
                    usize::from(region.kind() == HeaderKind::Footer),
                    region.slot_index(),
                ])
                .done();
        }
    }
    tree.build()
}

/// The body tables, their rows and cells; returns the table count.
fn build_tables(package: &Package, tree: &mut TreeBuilder<Kind>, body: u32) -> usize {
    let Ok(catalog) = package.body_tables() else {
        return 0;
    };
    for (index, snapshot) in catalog.iter().enumerate() {
        let Ok(read) = package.body_table_cells(BodyTableSelector::index(index)) else {
            tree.push(body, Kind::Table)
                .prop("name", text(Some(snapshot.name())))
                .label(Some(snapshot.name()))
                .leaf([index, 0, 0, 0])
                .done();
            continue;
        };
        let rows = usize::try_from(read.row_count()).unwrap_or(usize::MAX);
        let columns = usize::try_from(read.column_count()).unwrap_or(usize::MAX);
        let mut values: BTreeMap<(usize, usize), (String, &'static str)> = BTreeMap::new();
        for cell in read.iter_cells() {
            let position = cell.position();
            values.insert(
                (position.row() as usize, position.column() as usize),
                (cell.value().to_string(), value_type(cell.value())),
            );
        }
        let table = tree
            .push(body, Kind::Table)
            .prop("name", text(Some(snapshot.name())))
            .prop("rows", count(rows))
            .prop("columns", count(columns))
            .label(Some(snapshot.name()))
            .leaf([index, 0, 0, 0])
            .done();
        for row in 0..rows {
            let item = tree
                .push(table, Kind::Row)
                .prop("cellCount", count(columns))
                .leaf([index, row, 0, 0])
                .done();
            for column in 0..columns {
                let (value, kind) = values
                    .get(&(row, column))
                    .map_or((String::new(), "empty"), |(value, kind)| {
                        (value.clone(), *kind)
                    });
                tree.push(item, Kind::Cell)
                    .prop("text", text(Some(&value)))
                    .prop("type", text(Some(kind)))
                    .text(value)
                    .leaf([index, row, column, 0])
                    .done();
            }
        }
    }
    catalog.len()
}

/// The unified type name of a cell value.
fn value_type(value: &litchi_pages::table::read::Value) -> &'static str {
    use litchi_pages::table::read::Value;
    match value {
        Value::Empty => "empty",
        Value::Text(_) => "text",
        Value::Number(_) => "number",
        Value::Boolean(_) => "boolean",
        Value::Date(_) => "date",
        Value::Duration(_) => "duration",
        Value::Formula(_) => "formula",
        Value::Error(_) => "error",
    }
}

// ---------------------------------------------------------------------------
// Writes
// ---------------------------------------------------------------------------

/// Paragraph text sets waiting for one paragraph transaction per section:
/// section → `(line, text)` in staging order. A batch stages in line order,
/// so a list append is cheaper than a map insert per set (record 0302);
/// [`Pending::flush`] orders each section's lines and keeps the last text
/// of a repeated line, which is what a map keyed by line held.
#[derive(Debug, Default)]
struct Pending(BTreeMap<usize, Vec<(usize, String)>>);

impl Pending {
    fn stage(&mut self, section: usize, line: usize, value: String) {
        self.0.entry(section).or_default().push((line, value));
    }

    /// Commit the staged sets: one leaf paragraph transaction per section.
    fn flush(&mut self, chain: &mut Chain<Package>) -> Result<()> {
        let error = |error| failure(LeafError::PagesParagraphs(error));
        for (section, mut lines) in std::mem::take(&mut self.0) {
            // Stable: the sets of one line stay in staging order.
            lines.sort_by_key(|(line, _)| *line);
            let mut edit = chain
                .package
                .edit_section_paragraphs(section)
                .map_err(error)?;
            for (at, (line, value)) in lines.iter().enumerate() {
                if lines.get(at + 1).is_some_and(|(next, _)| next == line) {
                    // A later set of the same line replaces this one.
                    continue;
                }
                edit.set(Position::new(*line), value).map_err(error)?;
            }
            let commit = edit.commit().map_err(error)?;
            let (package, patch) = commit.into_parts();
            push(chain, package, Patch::Paragraphs(patch));
        }
        Ok(())
    }
}

/// The text of a paragraph `text` value: newlines are refused (they would
/// split the paragraph).
fn paragraph_text(value: &PropValue<'_>) -> Result<String> {
    let value = super::write::text(value, "text")?.unwrap_or_default();
    if value.contains('\n') {
        return Err(refusal(
            Code::InvalidValue,
            "a paragraph's text cannot contain a newline, which separates Pages paragraphs",
        )
        .with_property("text".to_owned())
        .with_suggestion("add each paragraph separately, or set the section's `text`"));
    }
    Ok(value)
}

/// Set properties on every target.
pub(crate) fn set(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    if is_metadata(target) {
        return set_metadata(file, package, props, limits);
    }
    let nodes = resolve::<Pages>(file, target, limits.max_expanded_targets())?;
    let mut chain = Chain::new(package.clone());
    let mut pending = Pending::default();
    let mut operations = 0_usize;
    for node in &nodes {
        for (key, value) in props.iter() {
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            let key = key.as_str();
            let leaf = node.item().leaf().map(|value| value as usize);
            match (node.item().type_name(), key) {
                ("paragraph", "text") => {
                    pending.stage(leaf[SECTION], leaf[LINE], paragraph_text(value)?);
                },
                ("section", "text") => {
                    pending.flush(&mut chain)?;
                    let value = super::write::text(value, "text")?.unwrap_or_default();
                    let error = |error| failure(LeafError::PagesText(error));
                    let commit = chain
                        .package
                        .set_section_text(leaf[SECTION], &value)
                        .map_err(error)?;
                    let patch = commit.patch().clone();
                    push(&mut chain, commit.into_package(), Patch::Text(patch));
                },
                ("section", "name") => {
                    pending.flush(&mut chain)?;
                    let value = super::write::text(value, "name")?;
                    let error = |error| failure(LeafError::PagesName(error));
                    let mut edit = chain
                        .package
                        .edit_section_name(leaf[SECTION])
                        .map_err(error)?;
                    edit.set_name(value.as_deref()).map_err(error)?;
                    let commit = edit.commit().map_err(error)?;
                    let patch = commit.patch().clone();
                    push(&mut chain, commit.into_package(), Patch::Name(patch));
                },
                ("header" | "footer", "text") => {
                    pending.flush(&mut chain)?;
                    let value = super::write::text(value, "text")?.unwrap_or_default();
                    let kind = if leaf[2] == 0 {
                        HeaderKind::Header
                    } else {
                        HeaderKind::Footer
                    };
                    let selector = HeaderFooterSelector::index(
                        leaf[SECTION],
                        template_of(node.item().leaf()[1]),
                        kind,
                        leaf[3],
                    );
                    let error = |error| failure(LeafError::PagesHeaderFooter(error));
                    let mut edit = chain
                        .package
                        .edit_header_footer_text(selector)
                        .map_err(error)?;
                    edit.set(&value).map_err(error)?;
                    let commit = edit.commit().map_err(error)?;
                    let patch = commit.patch().clone();
                    push(
                        &mut chain,
                        commit.into_package(),
                        Patch::HeaderFooter(patch),
                    );
                },
                ("footnote", "text") => {
                    pending.flush(&mut chain)?;
                    let value = super::write::text(value, "text")?.unwrap_or_default();
                    let error = |error| failure(LeafError::PagesFootnote(error));
                    let mut edit = chain
                        .package
                        .edit_body_footnote_text(litchi_pages::footnote::body::Selector::index(
                            leaf[0],
                        ))
                        .map_err(error)?;
                    edit.set(&value).map_err(error)?;
                    let commit = edit.commit().map_err(error)?;
                    let patch = commit.patch().clone();
                    push(&mut chain, commit.into_package(), Patch::Footnote(patch));
                },
                ("cell", _) => {
                    return Err(refusal(
                        Code::UnsupportedCapability,
                        "the Pages leaf reads body-table cells but has no cell-write transaction",
                    )
                    .with_property(key.to_owned())
                    .with_suggestion("follow-up: Pages body-table cell writes (record 0299)"));
                },
                ("paragraph" | "section" | "header" | "footer" | "footnote", _) => {
                    return Err(key_refusal(node, key));
                },
                _ => return Err(element_refusal("set", node)),
            }
        }
    }
    pending.flush(&mut chain)?;
    finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        operations,
        limits,
    )
}

/// Whether `target` is exactly the `/metadata` path.
fn is_metadata(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        let path = path.path();
        path.is_absolute()
            && matches!(path.steps(), [step]
                if step.predicates().is_empty() && step.name().eq_ignore_ascii_case("metadata"))
    })
}

/// The `/metadata` keys Pages stores and the root-metadata keys of other
/// formats it has no place for.
const METADATA_UNOWNED: [&str; 18] = [
    "application",
    "category",
    "company",
    "contentStatus",
    "contentType",
    "created",
    "extended.company",
    "extended.manager",
    "extended.template",
    "language",
    "lastModifiedBy",
    "lastPrinted",
    "manager",
    "modified",
    "revisionNumber",
    "subject",
    "template",
    "version",
];

/// `set("/metadata", props)`: title, author (`creator`), keywords and
/// description in one leaf transaction; `null` removes a property.
fn set_metadata(
    file: &Iwork,
    package: &Package,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    bound(LimitKind::Operations, props.len(), limits.max_operations())?;
    let error = |error| failure(LeafError::PagesMetadata(error));
    let mut edit = package.edit_metadata();
    for (key, value) in props.iter() {
        let name = key.as_str();
        let value = match value {
            PropValue::Null => None,
            other => Some(other.as_text().ok_or_else(|| {
                refusal(Code::InvalidValue, "a metadata value must be text or null")
                    .with_property(name.to_owned())
            })?),
        };
        match name.to_ascii_lowercase().as_str() {
            "title" => edit.set_title(value),
            "author" | "creator" => edit.set_author(value),
            "keywords" => edit.set_keywords(value),
            "description" => edit.set_description(value),
            _ if METADATA_UNOWNED
                .iter()
                .any(|known| known.eq_ignore_ascii_case(name))
                || name.to_ascii_lowercase().starts_with("custom.") =>
            {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "Pages stores title, author, keywords and description",
                )
                .with_property(name.to_owned())
                .with_suggestion("set title, author, keywords or description"));
            },
            _ => {
                return Err(refusal(
                    Code::UnsupportedProperty,
                    "the metadata key is not a Pages document property",
                )
                .with_property(name.to_owned())
                .with_suggestion("use title, author, keywords or description"));
            },
        }
        .map_err(error)?;
    }
    let commit = edit.commit().map_err(error)?;
    let (next, patch) = commit.into_parts();
    let mut chain = Chain::new(package.clone());
    push(&mut chain, next, Patch::Metadata(patch));
    finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        props.len(),
        limits,
    )
}

/// Every paragraph's `(section, line)`, in body order.
fn paragraph_positions(file: &Iwork) -> Vec<(usize, usize)> {
    let tree = file.tree();
    tree.items()
        .iter()
        .filter(|item| item.type_name() == "paragraph")
        .map(|item| {
            let leaf = item.leaf();
            (leaf[0] as usize, leaf[1] as usize)
        })
        .collect()
}

/// The paragraph node at `path`.
fn paragraph_at<'f>(file: &'f Iwork, path: &PathArg<'_>) -> Result<Node<'f>> {
    let node = super::write::node_at::<Pages>(file, path)?;
    if node.item().type_name() == "paragraph" {
        Ok(node)
    } else {
        Err(refusal(
            Code::InvalidArgument,
            "a Pages paragraph is placed relative to another paragraph",
        ))
    }
}

/// The `(section, line)` a paragraph inserted at global position `index`
/// lands at: before the paragraph now at `index`, or at the end of the last
/// section.
fn insertion_point(positions: &[(usize, usize)], index: usize) -> Result<(usize, usize)> {
    if let Some(&(section, line)) = positions.get(index) {
        return Ok((section, line));
    }
    if index == positions.len() {
        let (section, line) = positions.last().copied().unwrap_or((0, 0));
        return Ok((section, line + 1));
    }
    Err(refusal(
        Code::OutOfRange,
        "the insertion position is outside the body's paragraphs",
    ))
}

/// Add one paragraph to the body.
pub(crate) fn add(
    file: &Iwork,
    package: &Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    file.checked_tree()?;
    let parent_node = super::write::node_at::<Pages>(file, parent)?;
    let kind = node.kind().to_ascii_lowercase();
    if kind == "footnote" {
        return add_footnote(file, package, &parent_node, node, limits);
    }
    if !matches!(kind.as_str(), "p" | "paragraph") {
        return Err(refusal(
            Code::UnsupportedCapability,
            match kind.as_str() {
                "section" => {
                    "adding a Pages section needs a leaf section lifecycle transaction (IWORK_API_REQUESTS.md request 11)"
                },
                "tbl" | "table" => {
                    "adding a Pages body table needs a leaf table-creation transaction (record 0299 follow-up)"
                },
                _ => "a Pages body accepts only paragraphs",
            },
        ));
    }
    let mut value = String::new();
    for (key, prop) in node.props().iter() {
        if key.as_str() == "text" {
            value = paragraph_text(prop)?;
        } else {
            return Err(super::write::unwritable("paragraph", key.as_str()));
        }
    }
    let positions = paragraph_positions(file);
    let (section, line) = match parent_node.item().type_name() {
        "body" => {
            use crate::Placement;
            let index = match node.placement() {
                None | Some(Placement::End) => positions.len(),
                Some(Placement::Start) => 0,
                Some(Placement::Position(position)) | Some(Placement::Into(_, Some(position))) => {
                    position.get()
                },
                Some(Placement::Into(_, None)) => positions.len(),
                Some(Placement::Before(path)) => {
                    let anchor = paragraph_at(file, path)?;
                    let leaf = anchor.item().leaf();
                    return add_at(
                        file,
                        package,
                        leaf[0] as usize,
                        leaf[1] as usize,
                        &value,
                        limits,
                    );
                },
                Some(Placement::After(path)) => {
                    let anchor = paragraph_at(file, path)?;
                    let leaf = anchor.item().leaf();
                    return add_at(
                        file,
                        package,
                        leaf[0] as usize,
                        leaf[1] as usize + 1,
                        &value,
                        limits,
                    );
                },
            };
            insertion_point(&positions, index)?
        },
        "section" => {
            let section = parent_node.item().leaf()[0] as usize;
            let lines = positions.iter().filter(|(s, _)| *s == section).count();
            (section, lines)
        },
        _ => return Err(element_refusal("add", &parent_node)),
    };
    add_at(file, package, section, line, &value, limits)
}

/// Insert a paragraph at `(section, line)` and name its new path.
fn add_at(
    file: &Iwork,
    package: &Package,
    section: usize,
    line: usize,
    value: &str,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let error = |error| failure(LeafError::PagesParagraphs(error));
    let mut edit = package.edit_section_paragraphs(section).map_err(error)?;
    edit.insert(Position::new(line), value).map_err(error)?;
    let commit = edit.commit().map_err(error)?;
    let (next, patch) = commit.into_parts();
    let mut chain = Chain::new(package.clone());
    push(&mut chain, next, Patch::Paragraphs(patch));
    let commit = finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        1,
        limits,
    )?;
    let index = paragraph_positions(&commit.after)
        .iter()
        .position(|position| *position == (section, line))
        .ok_or_else(|| refusal(Code::InternalError, "the added paragraph is missing"))?;
    let path = Path::parse(&format!("/body/p[{}]", index + 1))?.into_owned();
    Ok((commit, path))
}

/// The UTF-16 body position at the end of paragraph `(section, line)`:
/// the sections' texts are joined by one section-break unit in the body
/// storage.
fn paragraph_end(package: &Package, section: usize, line: usize) -> Result<usize> {
    let mut offset = 0_usize;
    for index in 0..section {
        offset += section_text(package, index).encode_utf16().count() + 1;
    }
    let value = section_text(package, section);
    for (index, paragraph) in value.split('\n').enumerate() {
        offset += paragraph.encode_utf16().count();
        if index == line {
            return Ok(offset);
        }
        offset += 1;
    }
    Err(refusal(
        Code::PathNotFound,
        "the paragraph is no longer in its section",
    ))
}

/// Add a body footnote anchored at the end of the parent paragraph (or of
/// the last paragraph for `/footnotes`).
fn add_footnote(
    file: &Iwork,
    package: &Package,
    parent: &Node<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    // The anchor: the end of the parent paragraph, or of the body's last
    // paragraph when the footnote is added to `/footnotes`.
    let leaf = match parent.item().type_name() {
        "paragraph" => parent.item().leaf().map(|value| value as usize),
        "footnotes" => {
            let (section, line) = paragraph_positions(file).last().copied().unwrap_or((0, 0));
            [section, line, 0, 0]
        },
        _ => {
            return Err(refusal(
                Code::InvalidArgument,
                "a Pages footnote is added to the paragraph whose end anchors it",
            )
            .with_suggestion("add(\"/body/p[N]\", footnote)"));
        },
    };
    let mut value = String::new();
    let mut mark = None;
    for (key, prop) in node.props().iter() {
        match key.as_str() {
            "text" => value = super::write::text(prop, "text")?.unwrap_or_default(),
            "mark" => mark = super::write::text(prop, "mark")?,
            other => return Err(super::write::unwritable("footnote", other)),
        }
    }
    let end = paragraph_end(package, leaf[SECTION], leaf[LINE])?;
    let position = litchi_pages::footnote::body::Position::from_utf16_index(end)
        .map_err(|_| refusal(Code::OutOfRange, "the footnote anchor is outside the body"))?;
    let error = |error| failure(LeafError::PagesFootnotes(error));
    let commit = package
        .insert_body_footnote(position, &value, mark.as_deref())
        .map_err(error)?;
    let patch = commit.patch().clone();
    let next = commit.into_package();
    let ordinal = next
        .body_footnotes()
        .map_err(|error| failure(LeafError::Pages(error)))?
        .iter()
        .position(|footnote| footnote.position == position)
        .ok_or_else(|| refusal(Code::InternalError, "the added footnote is missing"))?;
    let mut chain = Chain::new(package.clone());
    push(&mut chain, next, Patch::Footnotes(patch));
    let commit = finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        1,
        limits,
    )?;
    let path = Path::parse(&format!("/footnotes/footnote[{}]", ordinal + 1))?.into_owned();
    Ok((commit, path))
}

/// Remove paragraphs and body tables.
pub(crate) fn remove(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Pages>(file, target, limits.max_expanded_targets())?;
    let mut tables = Vec::new();
    let mut footnotes = Vec::new();
    let mut lines: BTreeMap<usize, Vec<usize>> = BTreeMap::new();
    for node in &nodes {
        let leaf = node.item().leaf().map(|value| value as usize);
        match node.item().type_name() {
            "paragraph" => lines.entry(leaf[SECTION]).or_default().push(leaf[LINE]),
            "table" => tables.push(leaf[0]),
            "footnote" => footnotes.push(leaf[0]),
            _ => return Err(element_refusal("remove", node)),
        }
    }
    bound(LimitKind::Operations, nodes.len(), limits.max_operations())?;
    tables.sort_unstable();
    tables.dedup();
    let mut chain = Chain::new(package.clone());
    // Tables first, from the last: each deletion leaves the earlier
    // tables' catalog positions and every paragraph's line in place.
    for table in tables.iter().rev() {
        let error = |error| failure(LeafError::PagesTableDeletion(error));
        let commit = chain
            .package
            .remove_body_table(BodyTableSelector::index(*table))
            .map_err(error)?;
        let patch = commit.patch().clone();
        let next = commit.into_package();
        push(&mut chain, next, Patch::TableDeletion(patch));
    }
    // Footnotes from the last: each removal leaves the earlier ones'
    // positions in place.
    footnotes.sort_unstable();
    footnotes.dedup();
    for footnote in footnotes.iter().rev() {
        let error = |error| failure(LeafError::PagesFootnotes(error));
        let edit = chain
            .package
            .edit_body_footnote(litchi_pages::footnote::body::Selector::index(*footnote))
            .map_err(error)?;
        let mut edit = edit;
        edit.clear();
        let commit = edit.commit().map_err(error)?;
        let patch = commit.patch().clone();
        push(&mut chain, commit.into_package(), Patch::Footnotes(patch));
    }
    for (section, mut targets) in lines {
        targets.sort_unstable();
        targets.dedup();
        let error = |error| failure(LeafError::PagesParagraphs(error));
        let mut edit = chain
            .package
            .edit_section_paragraphs(section)
            .map_err(error)?;
        for line in targets.iter().rev() {
            edit.remove(Position::new(*line)).map_err(error)?;
        }
        let commit = edit.commit().map_err(error)?;
        let (next, patch) = commit.into_parts();
        push(&mut chain, next, Patch::Paragraphs(patch));
    }
    finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        nodes.len(),
        limits,
    )
}

/// The raw text (with anchors) of the paragraph at `(section, line)`.
fn raw_paragraph(package: &Package, section: usize, line: usize) -> Result<String> {
    let paragraphs = package
        .section_paragraphs(section)
        .map_err(|error| failure(LeafError::PagesParagraphs(error)))?;
    paragraphs
        .get(line)
        .map(|value| (*value).to_owned())
        .ok_or_else(|| {
            refusal(
                Code::PathNotFound,
                "the paragraph is no longer in its section",
            )
        })
}

/// The refusal of moving a paragraph that anchors an object across
/// sections.
fn anchored() -> crate::Error {
    failure(LeafError::PagesParagraphs(
        litchi_pages::section::paragraphs::Error::AnchoredParagraph,
    ))
}

/// Move one paragraph among the body's paragraphs.
pub(crate) fn move_to(
    file: &Iwork,
    package: &Package,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let node = paragraph_at(file, path)?;
    let leaf = node.item().leaf().map(|value| value as usize);
    let (section, line) = (leaf[SECTION], leaf[LINE]);
    let positions = paragraph_positions(file);
    let from = positions
        .iter()
        .position(|position| *position == (section, line))
        .ok_or_else(|| refusal(Code::InternalError, "the paragraph is missing"))?;
    let to = super::write::destination(from, positions.len(), placement, |anchor| {
        let anchor = paragraph_at(file, anchor)?;
        let leaf = anchor.item().leaf();
        positions
            .iter()
            .position(|position| *position == (leaf[0] as usize, leaf[1] as usize))
            .ok_or_else(|| refusal(Code::InternalError, "the anchor paragraph is missing"))
    })?;
    let mut rest = positions.clone();
    rest.remove(from);
    let (target_section, target_line) = match rest.get(to) {
        Some(&(other, other_line)) => {
            let adjusted = if other == section && other_line > line {
                other_line - 1
            } else {
                other_line
            };
            (other, adjusted)
        },
        None => {
            let last = rest.last().map_or(section, |position| position.0);
            let lines = rest.iter().filter(|(s, _)| *s == last).count();
            (last, lines)
        },
    };
    let error = |error| failure(LeafError::PagesParagraphs(error));
    let mut chain = Chain::new(package.clone());
    if target_section == section {
        let mut edit = package.edit_section_paragraphs(section).map_err(error)?;
        edit.relocate(Position::new(line), Position::new(target_line))
            .map_err(error)?;
        let (next, patch) = edit.commit().map_err(error)?.into_parts();
        push(&mut chain, next, Patch::Paragraphs(patch));
    } else {
        let value = raw_paragraph(package, section, line)?;
        if value.contains(is_anchor) {
            return Err(anchored());
        }
        let mut edit = package.edit_section_paragraphs(section).map_err(error)?;
        edit.remove(Position::new(line)).map_err(error)?;
        let (next, patch) = edit.commit().map_err(error)?.into_parts();
        push(&mut chain, next, Patch::Paragraphs(patch));
        let mut edit = chain
            .package
            .edit_section_paragraphs(target_section)
            .map_err(error)?;
        edit.insert(Position::new(target_line), &value)
            .map_err(error)?;
        let (next, patch) = edit.commit().map_err(error)?.into_parts();
        push(&mut chain, next, Patch::Paragraphs(patch));
    }
    finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        1,
        limits,
    )
}

/// Exchange two paragraphs.
pub(crate) fn swap(
    file: &Iwork,
    package: &Package,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let left = paragraph_at(file, left)?
        .item()
        .leaf()
        .map(|value| value as usize);
    let right = paragraph_at(file, right)?
        .item()
        .leaf()
        .map(|value| value as usize);
    let error = |error| failure(LeafError::PagesParagraphs(error));
    let mut chain = Chain::new(package.clone());
    if left[SECTION] == right[SECTION] {
        let mut edit = package
            .edit_section_paragraphs(left[SECTION])
            .map_err(error)?;
        edit.swap(Position::new(left[LINE]), Position::new(right[LINE]))
            .map_err(error)?;
        let (next, patch) = edit.commit().map_err(error)?.into_parts();
        push(&mut chain, next, Patch::Paragraphs(patch));
    } else {
        let first = raw_paragraph(package, left[SECTION], left[LINE])?;
        let second = raw_paragraph(package, right[SECTION], right[LINE])?;
        if first.contains(is_anchor) || second.contains(is_anchor) {
            return Err(anchored());
        }
        let mut pending = Pending::default();
        pending.stage(left[SECTION], left[LINE], second);
        pending.stage(right[SECTION], right[LINE], first);
        pending.flush(&mut chain)?;
    }
    finish(
        file,
        chain.changed,
        Owner::Pages(chain.package),
        chain.patches,
        1,
        limits,
    )
}

// ---------------------------------------------------------------------------
// Batch coalescing
// ---------------------------------------------------------------------------

/// One planned paragraph text set: its section, line and text.
#[derive(Debug)]
pub(crate) struct Planned {
    section: usize,
    line: usize,
    text: String,
}

/// Plan `set(target, props)` for a coalesced run: a direct path to a
/// paragraph and exactly one `text` property. Anything else is `None` (it
/// runs on its own).
///
/// # Errors
///
/// A refused paragraph text (it holds a newline).
pub(crate) fn plan(
    file: &Iwork,
    target: &TargetArg<'_>,
    props: &Props<'_>,
) -> Result<Option<Planned>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    let mut props_iter = props.iter();
    let (Some((key, value)), None) = (props_iter.next(), props_iter.next()) else {
        return Ok(None);
    };
    if key.as_str() != "text" || is_metadata(target) {
        return Ok(None);
    }
    if let Some(found) = direct_paragraph(file, path.path()) {
        let Some((section, line)) = found else {
            // Past the last paragraph: the ordinary path reports it.
            return Ok(None);
        };
        return Ok(Some(Planned {
            section,
            line,
            text: paragraph_text(value)?,
        }));
    }
    if file.checked_tree().is_err() {
        return Ok(None);
    }
    let Ok(node) = crate::path::resolve::get(&super::Nav::<Pages>::new(file), path) else {
        return Ok(None);
    };
    if node.item().type_name() != "paragraph" {
        return Ok(None);
    }
    let leaf = node.item().leaf().map(|value| value as usize);
    Ok(Some(Planned {
        section: leaf[SECTION],
        line: leaf[LINE],
        text: paragraph_text(value)?,
    }))
}

/// The `(section, line)` of the paragraph `/body/p[N]` names, from each
/// section's paragraph count rather than the object tree (record 0302): a
/// coalesced run of sets on a large document otherwise builds the whole
/// tree, with a text copy and properties per paragraph, before its first
/// set. `None` for any other spelling (it resolves through the tree);
/// `Some(None)` when the document has fewer than N paragraphs.
///
/// The tree lists each section's lines (its text split at `\n`) in section
/// order under the body, ahead of the tables, so the N-th paragraph child
/// of the body is the N-th line overall; the test
/// `direct_paragraphs_match_the_tree` compares the two.
fn direct_paragraph(file: &Iwork, path: &Path<'_>) -> Option<Option<(usize, usize)>> {
    let [body, paragraph] = path.steps() else {
        return None;
    };
    if !path.is_absolute()
        || body.name() != "body"
        || !body.predicates().is_empty()
        || paragraph.name() != "p"
    {
        return None;
    }
    let index = paragraph.index()?.get();
    let package = file.pages()?;
    let starts = file.paragraph_starts(|| paragraph_starts(package));
    let (&total, sections) = starts.split_last()?;
    if index >= total {
        return Some(None);
    }
    // The last section whose first paragraph is at or before `index`.
    let section = sections
        .partition_point(|&start| start <= index)
        .checked_sub(1)?;
    Some(Some((section, index - sections[section])))
}

/// The global index of each section's first paragraph, then the total.
fn paragraph_starts(package: &Package) -> Box<[usize]> {
    let sections = package.sections().len();
    let mut starts = Vec::with_capacity(sections + 1);
    let mut total = 0_usize;
    for index in 0..sections {
        starts.push(total);
        let lines = match package.section_text(index) {
            Ok(text) => line_count(text),
            Err(_) => line_count(&section_text(package, index)),
        };
        total = total.saturating_add(lines);
    }
    starts.push(total);
    starts.into_boxed_slice()
}

/// The paragraphs of a section text: its `\n`-separated lines.
fn line_count(text: &str) -> usize {
    text.bytes().filter(|&byte| byte == b'\n').count() + 1
}

/// Consecutive paragraph text sets of one transaction, published as one
/// leaf paragraph transaction per section (a single splice when the change
/// consumes no object anchor).
#[derive(Debug)]
pub(crate) struct Run {
    base: Iwork,
    pending: Pending,
    limits: SemanticLimits,
    first_index: Option<usize>,
    operations: usize,
}

impl Run {
    /// A run over the immutable handle `base`.
    pub(crate) fn new(base: &Iwork, limits: &SemanticLimits) -> Self {
        Self {
            base: base.clone(),
            pending: Pending::default(),
            limits: *limits,
            first_index: None,
            operations: 0,
        }
    }

    /// The transaction item index of the run's first set.
    pub(crate) const fn first_index(&self) -> Option<usize> {
        self.first_index
    }

    /// Stage one planned set.
    ///
    /// # Errors
    ///
    /// The transaction operation limit.
    pub(crate) fn stage(&mut self, plan: Planned, index: usize) -> Result<bool> {
        let operations = self.operations.saturating_add(1);
        bound(
            LimitKind::Operations,
            operations,
            self.limits.max_operations(),
        )?;
        self.pending.stage(plan.section, plan.line, plan.text);
        self.operations = operations;
        self.first_index.get_or_insert(index);
        Ok(true)
    }

    /// Commit the run: one leaf transaction per touched section.
    ///
    /// # Errors
    ///
    /// The leaf's refusal or the snapshot limit.
    pub(crate) fn commit(mut self) -> Result<(Iwork, Commit)> {
        let Some(package) = self.base.pages().cloned() else {
            return Err(refusal(
                Code::InternalError,
                "the Pages run lost its package",
            ));
        };
        let mut chain = Chain::new(package);
        self.pending.flush(&mut chain)?;
        let commit = finish(
            &self.base,
            chain.changed,
            Owner::Pages(chain.package),
            chain.patches,
            self.operations,
            &self.limits,
        )?;
        Ok((self.base, commit))
    }
}

#[cfg(test)]
mod direct_tests {
    use super::{Pages, direct_paragraph};
    use crate::File;
    use crate::path::{IntoPath, Path};

    /// Every Pages fixture: `/body/p[N]` found from the section counts names
    /// the paragraph the tree resolves, for every N up to two past the end.
    #[test]
    fn direct_paragraphs_match_the_tree() {
        let directory =
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages");
        let mut checked = 0;
        let mut entries: Vec<_> = std::fs::read_dir(&directory)
            .expect("fixtures")
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| path.extension().is_some_and(|ext| ext == "pages"))
            .collect();
        entries.sort();
        for fixture in entries {
            let Ok(bytes) = std::fs::read(&fixture) else {
                continue;
            };
            let Ok(File::Pages(file)) = crate::open_bytes(bytes) else {
                continue;
            };
            let mut index = 1;
            loop {
                let text = format!("/body/p[{index}]");
                let path = Path::parse(&text).expect("path");
                let direct = direct_paragraph(&file, &path).expect("a direct spelling");
                let arg = text.as_str().into_path().expect("path");
                let tree = crate::path::resolve::get(&super::super::Nav::<Pages>::new(&file), &arg)
                    .ok()
                    .filter(|node| node.item().type_name() == "paragraph")
                    .map(|node| {
                        let leaf = node.item().leaf();
                        (leaf[0] as usize, leaf[1] as usize)
                    });
                assert_eq!(direct, tree, "{} {text}", fixture.display());
                checked += 1;
                if direct.is_none() {
                    break;
                }
                index += 1;
            }
            for other in ["/body/paragraph[1]", "/body/p", "/body[1]/p[1]", "/p[1]"] {
                let path = Path::parse(other).expect("path");
                assert!(direct_paragraph(&file, &path).is_none(), "{other}");
            }
        }
        assert!(checked > 20, "{checked}");
    }
}
