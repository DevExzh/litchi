//! `File::dump`: a document written out as an OfficeCLI v2 batch, the
//! inverse of [`File::batch`] (record 0336).
//!
//! The dump walks the node view and emits the commands that rebuild it on
//! a [`create`](crate::create)d document of the same format: `add` for the
//! elements the target lacks, `set` for the ones `create` (or an earlier
//! `add`) already made, and `remove` for the ones the source does not have.
//! Each command carries the node's readable properties that the registry
//! marks as writable by that command, so a dump never names a property the
//! replay would refuse. Content the commands cannot rebuild (pictures,
//! charts, SmartArt, comments, fields, resources such as styles and
//! masters) is skipped and reported in [`Dump::warnings`].

use std::borrow::Cow;
use std::fmt;

use litchi_core::FileFormat;
use serde::ser::{Serialize, SerializeSeq, Serializer};

use super::BatchItem;
use crate::capabilities::{Capabilities, Support};
use crate::node::{Node, Value, ViewOptions};
use crate::{Code, Error, File, PropValue, Props, Result};

/// A document as OfficeCLI v2 batch items (record 0336).
///
/// The first item is `{"command":"meta","dumpVersion":2}`: soft line
/// breaks in run text are `\u{b}` (OfficeCLI's `\v`), and a `\n` in a
/// shape's `text` separates paragraphs. Serialized, a dump is the JSON
/// array [`File::batch`] and `litchi batch` read.
///
/// ```
/// # #[cfg(feature = "docx")]
/// # {
/// use litchi_core::FileFormat;
/// use litchi_ops::{BatchMode, props};
///
/// let file = litchi_ops::create(FileFormat::Docx)?;
/// let file = file.add("/body", litchi_ops::NewNode::new("paragraph"))?.file;
/// let file = file
///     .add("/body/p[1]", litchi_ops::NewNode::new("run").with_props(props! { "text" => "Hi", "bold" => true })?)?
///     .file;
/// let dump = file.dump()?;
/// assert!(dump.warnings().is_empty());
///
/// let replay = litchi_ops::create(FileFormat::Docx)?.batch(dump.items(), BatchMode::Atomic)?;
/// assert_eq!(replay.file().text()?, file.text()?);
/// assert_eq!(replay.file().dump()?.items(), dump.items());
/// # }
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Clone)]
pub struct Dump {
    items: Vec<BatchItem<'static>>,
    warnings: Vec<Error>,
}

impl Dump {
    /// The items, the `meta` item first.
    #[must_use]
    pub fn items(&self) -> &[BatchItem<'static>] {
        &self.items
    }

    /// The items, without a copy.
    #[must_use]
    pub fn into_items(self) -> Vec<BatchItem<'static>> {
        self.items
    }

    /// What the dump skipped, one `unsupported_element` error per node:
    /// its message names the element kind and its positional path.
    #[must_use]
    pub fn warnings(&self) -> &[Error] {
        &self.warnings
    }
}

impl fmt::Debug for Dump {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Dump")
            .field("items", &self.items.len())
            .field("warnings", &self.warnings.len())
            .finish()
    }
}

impl Serialize for Dump {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(Some(self.items.len()))?;
        for item in &self.items {
            seq.serialize_element(item)?;
        }
        seq.end()
    }
}

/// The families the dump rebuilds, and the formats in each.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Plan {
    Document,
    Workbook,
    Presentation,
}

const fn plan(format: FileFormat) -> Option<Plan> {
    match format {
        FileFormat::Docx | FileFormat::Odt | FileFormat::Rtf | FileFormat::Doc => {
            Some(Plan::Document)
        },
        FileFormat::Xlsx | FileFormat::Xlsb | FileFormat::Ods | FileFormat::Xls => {
            Some(Plan::Workbook)
        },
        FileFormat::Pptx => Some(Plan::Presentation),
        _ => None,
    }
}

/// Whether `format` can be dumped, for `capabilities()`.
pub(crate) const fn support(format: FileFormat) -> Support {
    if plan(format).is_some() {
        Support::Supported
    } else {
        Support::Unsupported {
            code: Code::UnsupportedCapability,
            follow_up: "record 0336 follow-up: dump plans for the remaining formats",
            reason: "this format's write verbs cannot yet rebuild its node tree on a created document",
        }
    }
}

impl File {
    /// Write the document out as an OfficeCLI v2 batch that rebuilds it on
    /// a [`create`](crate::create)d document of the same format (OfficeCLI
    /// `dump`; record 0336). See [`Dump`].
    ///
    /// Supported: DOCX, ODT, RTF and DOC body paragraphs, runs (with
    /// hyperlinks, breaks and tabs) and tables; XLSX, XLSB, ODS and XLS
    /// sheets, cells and row layout; PPTX slides (by layout name),
    /// placeholders, shapes and notes.
    ///
    /// # Errors
    /// Returns `unsupported_capability` naming the format for the other
    /// formats, `document_protected` for a locked file, and read errors.
    pub fn dump(&self) -> Result<Dump> {
        let Some(format) = self.format() else {
            return Err(refused(self));
        };
        if let File::Locked(locked) = self {
            return Err(locked.read_error());
        }
        let Some(plan) = plan(format) else {
            return Err(refused(self));
        };
        let mut out = Out::new(format);
        out.items.push(meta());
        let root = self.root()?;
        match plan {
            Plan::Document => document(&mut out, &root, format)?,
            Plan::Workbook => workbook(&mut out, &root, format)?,
            Plan::Presentation => presentation(&mut out, &root, format)?,
        }
        Ok(Dump {
            items: out.items,
            warnings: out.warnings,
        })
    }
}

fn refused(file: &File) -> Error {
    let error = Error::new(
        Code::UnsupportedCapability,
        "dump is not available for this format",
    )
    .with_suggestion(
        "dump covers DOCX, ODT, RTF, DOC, XLSX, XLSB, ODS, XLS and PPTX (record 0336)",
    );
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}

fn meta() -> BatchItem<'static> {
    let mut item = BatchItem::new("meta");
    item.dump_version = Some(2);
    item
}

struct Out {
    format: FileFormat,
    capabilities: Capabilities,
    /// The registry's writable properties per element kind, looked up once
    /// per kind instead of once per node.
    writable: Vec<(String, Option<Writable>)>,
    items: Vec<BatchItem<'static>>,
    warnings: Vec<Error>,
}

/// One element's properties a dump may write, by key.
struct Writable {
    props: std::collections::HashMap<&'static str, Writes>,
}

#[derive(Clone, Copy)]
struct Writes {
    add: bool,
    set: bool,
    ty: crate::value::ValueType,
}

/// Which command a node's properties are filtered for.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Command {
    Add,
    Set,
}

impl Out {
    fn new(format: FileFormat) -> Self {
        Self {
            format,
            capabilities: Capabilities::of(format),
            writable: Vec::new(),
            items: Vec::new(),
            warnings: Vec::new(),
        }
    }

    fn skip(&mut self, node: &Node<'_>) {
        self.warnings.push(Error::new(
            Code::UnsupportedElement,
            format!(
                "dump skipped the {} at {}",
                node.kind(),
                node.positional_path()
            ),
        ));
    }

    /// The node's readable properties that `command` can write to its
    /// element, in property-table order, minus `exclude`.
    fn props(
        &mut self,
        node: &Node<'_>,
        element: &str,
        command: Command,
        exclude: &[&str],
    ) -> Result<Props<'static>> {
        let mut props = Props::new();
        let index = match self.writable.iter().position(|(kind, _)| kind == element) {
            Some(index) => index,
            None => {
                let writable = self.capabilities.element(element).map(|element| Writable {
                    props: element
                        .properties()
                        .filter(|property| property.write_gap().is_none())
                        .map(|property| {
                            (
                                property.name(),
                                Writes {
                                    add: property.add(),
                                    set: property.set(),
                                    ty: property.value_type(),
                                },
                            )
                        })
                        .collect(),
                });
                self.writable.push((element.to_owned(), writable));
                self.writable.len() - 1
            },
        };
        let Some(writable) = &self.writable[index].1 else {
            return Ok(props);
        };
        let mut dropped: Vec<&'static str> = Vec::new();
        for prop in node.view().format()? {
            let key = prop.key();
            if exclude.contains(&key) {
                continue;
            }
            let Some(property) = writable.props.get(key).copied() else {
                continue;
            };
            let allowed = match command {
                Command::Add => property.add,
                Command::Set => property.set,
            };
            if !allowed {
                continue;
            }
            let value = match prop.value() {
                Value::Bool(value) => PropValue::Bool(*value),
                Value::Int(value) => PropValue::Int(*value),
                Value::Number(value) => PropValue::Number(*value),
                Value::List(items) => PropValue::List(Cow::Owned(items.to_vec())),
                _ => {
                    let text = prop.to_string();
                    // A read form the write coercion does not accept (a
                    // colour with its theme annotation) is dropped and
                    // reported rather than failing the replay. PPTX writes
                    // opaque RGB colours only, not scheme colours.
                    let scheme_colour = self.format == FileFormat::Pptx
                        && matches!(property.ty, crate::value::ValueType::Color(_))
                        && !text.starts_with('#');
                    if scheme_colour || property.ty.coerce(&text).is_err() {
                        dropped.push(key);
                        continue;
                    }
                    PropValue::Text(Cow::Owned(text))
                },
            };
            props.push(key, value);
        }
        for key in dropped {
            self.warnings.push(
                Error::new(
                    Code::UnsupportedProperty,
                    format!(
                        "dump dropped a property of the {} at {}: its read form is not writable",
                        node.kind(),
                        node.positional_path()
                    ),
                )
                .with_property(key),
            );
        }
        Ok(props)
    }

    /// `props` without the style references (`style`, `styleId`, `rStyle`)
    /// that name a style the replay target will not have, each reported.
    fn known_styles(
        &mut self,
        node: &Node<'_>,
        props: Props<'static>,
        styles: &[String],
    ) -> Props<'static> {
        let reference = |key: &str| matches!(key, "style" | "styleId" | "rStyle");
        if !props.iter().any(|(key, value)| {
            reference(key.as_str())
                && value
                    .as_text()
                    .is_some_and(|id| !styles.iter().any(|style| style == id))
        }) {
            return props;
        }
        let mut kept = Props::new();
        for (key, value) in props.iter() {
            let unknown = reference(key.as_str())
                && value
                    .as_text()
                    .is_some_and(|id| !styles.iter().any(|style| style == id));
            if unknown {
                self.warnings.push(
                    Error::new(
                        Code::UnsupportedProperty,
                        format!(
                            "dump dropped a style reference of the {} at {}: the replay target has no such style",
                            node.kind(),
                            node.positional_path()
                        ),
                    )
                    .with_property(key.as_str().to_owned()),
                );
            } else {
                kept.push(key.clone(), value.clone());
            }
        }
        kept
    }

    /// The properties `add` takes, and the remaining writable ones as the
    /// `set` that follows it.
    fn split(
        &mut self,
        node: &Node<'_>,
        element: &str,
        exclude: &[&str],
    ) -> Result<(Props<'static>, Props<'static>)> {
        let before = self.warnings.len();
        let add = self.props(node, element, Command::Add, exclude)?;
        let middle = self.warnings.len();
        let set = self.props(node, element, Command::Set, exclude)?;
        // A read form dropped by both filters is reported once.
        let reported: Vec<Option<String>> = self.warnings[before..middle]
            .iter()
            .map(|warning| warning.property().map(str::to_owned))
            .collect();
        let mut index = middle;
        while index < self.warnings.len() {
            let property = self.warnings[index].property().map(str::to_owned);
            if reported.contains(&property) {
                self.warnings.remove(index);
            } else {
                index += 1;
            }
        }
        let mut rest = Props::new();
        for (key, value) in set.iter() {
            if add.get(key.as_str()).is_none() {
                rest.push(key.clone(), value.clone());
            }
        }
        Ok((add, rest))
    }

    fn add(&mut self, parent: &str, kind: &str, props: Props<'static>) {
        let mut item = BatchItem::new("add");
        item.parent = Some(Cow::Owned(parent.to_owned()));
        item.kind = Some(Cow::Owned(kind.to_owned()));
        if !props.is_empty() {
            item.props = Some(props);
        }
        self.items.push(item);
    }

    fn set(&mut self, path: &str, props: Props<'static>) {
        if props.is_empty() {
            return;
        }
        let mut item = BatchItem::new("set");
        item.path = Some(Cow::Owned(path.to_owned()));
        item.props = Some(props);
        self.items.push(item);
    }

    fn remove(&mut self, path: &str) {
        let mut item = BatchItem::new("remove");
        item.path = Some(Cow::Owned(path.to_owned()));
        self.items.push(item);
    }
}

/// A node's children, each projected with its own format: some adapters
/// leave the format of a child the view lists at depth 0 unread (PPTX
/// slides), so the walk asks for one more level.
fn children<'n, 'f>(node: &'n Node<'f>) -> crate::node::Children<'n, 'f> {
    node.view_with(ViewOptions::DEFAULT.depth(2)).children()
}

fn text(node: &Node<'_>) -> Result<String> {
    Ok(node.view().text()?.map(Cow::into_owned).unwrap_or_default())
}

// ---------------------------------------------------------------------------
// Word processing: styles, body paragraphs, runs and tables.

/// What the format's write verbs can rebuild of a word-processing body.
struct Doc {
    /// Runs are added one by one (DOCX, ODT, RTF). Otherwise a paragraph is
    /// added with its text and its first run takes the run formatting
    /// (DOC).
    runs: bool,
    /// Paragraphs the created document's body already holds (DOC: one).
    existing: usize,
    /// The style ids a replay can reference: the created document's and
    /// the ones the dump adds.
    styles: Vec<String>,
    /// Whether a paragraph takes a `set` right after its `add` (RTF's
    /// created body has no literal source to splice into, so it does not).
    paragraph_sets: bool,
    /// Without those sets: the properties a freshly added paragraph reads
    /// back with, to report any paragraph whose formatting differs.
    paragraph_defaults: Props<'static>,
}

fn document(out: &mut Out, root: &Node<'_>, format: FileFormat) -> Result<()> {
    let blank = crate::create(format)?;
    let existing = match blank.get("/body") {
        Ok(body) => {
            let mut count = 0;
            for child in children(&body) {
                if child?.kind() == "paragraph" {
                    count += 1;
                }
            }
            count
        },
        Err(_) => 0,
    };
    let mut doc = Doc {
        runs: out
            .capabilities
            .element("run")
            .is_some_and(|run| run.operations().add),
        existing,
        styles: Vec::new(),
        paragraph_sets: format != FileFormat::Rtf,
        paragraph_defaults: Props::new(),
    };
    if !doc.paragraph_sets {
        if let Ok(added) = blank.add("/body", crate::NewNode::new("paragraph")) {
            let probe = added.file.get(&added.path)?;
            let excluded = paragraph_excluded(out);
            doc.paragraph_defaults = out.split(&probe, "paragraph", excluded)?.1;
        }
    }
    if let Ok(node) = blank.get("/styles") {
        for style in children(&node) {
            doc.styles.push(style_id(&style?)?);
        }
    }
    // Styles first, so the body's style references resolve on replay.
    for child in children(root) {
        let child = child?;
        if child.kind() == "styles" {
            styles(out, &child, &mut doc.styles)?;
        }
    }
    for child in children(root) {
        let child = child?;
        if child.kind() == "body" {
            body(out, &child, "/body", &doc)?;
        }
        // Other resources and stories (numbering, headers, notes,
        // comments) are not rebuilt; the created document has its own.
    }
    Ok(())
}

/// The style definitions the created document lacks. Styles it already
/// has keep the created definitions: `set` cannot clear the properties the
/// source style leaves unset, so rewriting them would not reproduce the
/// source either.
fn styles(out: &mut Out, styles: &Node<'_>, defaults: &mut Vec<String>) -> Result<()> {
    let style_add = out
        .capabilities
        .element("style")
        .is_some_and(|style| style.operations().add);
    if !style_add {
        return Ok(());
    }
    for style in children(styles) {
        let style = style?;
        if style.kind() != "style" {
            continue;
        }
        let id = style_id(&style)?;
        if id.is_empty() {
            out.skip(&style);
        } else if !defaults.contains(&id) {
            let mut props = out.props(&style, "style", Command::Add, &[])?;
            if props.get("name").is_none() {
                // `add` names an unnamed style after its id, right after it.
                let mut named = Props::new();
                for (key, value) in props.iter() {
                    named.push(key.clone(), value.clone());
                    if key.as_str() == "id" {
                        named.push("name", id.clone());
                    }
                }
                props = named;
            }
            out.add("/styles", "style", props);
            defaults.push(id);
        }
    }
    Ok(())
}

fn style_id(style: &Node<'_>) -> Result<String> {
    Ok(style
        .view()
        .format()?
        .find(|prop| prop.key() == "id")
        .map(|prop| prop.to_string())
        .unwrap_or_default())
}

/// Paragraph keys the dump does not write: the text (runs carry it) and
/// the two other spellings of the style reference (`styleId` carries it,
/// or `style` where the format has no `styleId`).
fn paragraph_excluded(out: &Out) -> &'static [&'static str] {
    let has_id = out
        .capabilities
        .element("paragraph")
        .and_then(|paragraph| paragraph.property("styleId"))
        .is_some_and(|property| property.add() || property.set());
    // List numbering refers to native numbering-instance ids the created
    // document does not share (ADR 0034); `numbering_warning` reports it.
    if has_id {
        &[
            "text",
            "style",
            "styleName",
            "numId",
            "numLevel",
            "listStyle",
            "start",
        ]
    } else {
        &[
            "text",
            "styleName",
            "numId",
            "numLevel",
            "listStyle",
            "start",
        ]
    }
}

fn body(out: &mut Out, body: &Node<'_>, path: &str, doc: &Doc) -> Result<()> {
    let mut paragraphs = 0_usize;
    let mut tables = 0_usize;
    for child in children(body) {
        let child = child?;
        match child.kind() {
            "paragraph" => {
                paragraphs += 1;
                let paragraph_path = format!("{path}/p[{paragraphs}]");
                paragraph(
                    out,
                    &child,
                    path,
                    &paragraph_path,
                    paragraphs <= doc.existing,
                    doc,
                )?;
            },
            "table" => {
                tables += 1;
                table(out, &child, path, tables, doc)?;
            },
            // Every created document has its final section already.
            "section" => {},
            _ => out.skip(&child),
        }
    }
    Ok(())
}

/// One paragraph: `add` under `parent` (or `set`, when the target already
/// holds it), the properties `add` cannot take as a `set`, then its
/// inline content.
fn paragraph(
    out: &mut Out,
    node: &Node<'_>,
    parent: &str,
    path: &str,
    existing: bool,
    doc: &Doc,
) -> Result<()> {
    let excluded = paragraph_excluded(out);
    if node
        .view()
        .format()?
        .any(|prop| matches!(prop.key(), "numId" | "listStyle" | "start"))
    {
        out.warnings.push(Error::new(
            Code::UnsupportedProperty,
            format!(
                "dump dropped the list numbering of the {} at {}: numbering ids are native identities of the source",
                node.kind(),
                node.positional_path()
            ),
        ).with_property("numId"));
    }
    let text = if doc.runs {
        None
    } else {
        Some(text(node)?).filter(|text| !text.is_empty())
    };
    if existing {
        let mut props = out.props(node, "paragraph", Command::Set, excluded)?;
        props = out.known_styles(node, props, &doc.styles);
        if let Some(text) = text {
            props.push("text", text.replace('\n', "\u{b}"));
        }
        out.set(path, props);
    } else {
        let (add, rest) = out.split(node, "paragraph", excluded)?;
        let mut add = out.known_styles(node, add, &doc.styles);
        let rest = out.known_styles(node, rest, &doc.styles);
        if let Some(text) = text {
            add.push("text", text.replace('\n', "\u{b}"));
        }
        out.add(parent, "paragraph", add);
        if doc.paragraph_sets {
            out.set(path, rest);
        } else if rest != doc.paragraph_defaults {
            out.warnings.push(Error::new(
                Code::UnsupportedProperty,
                format!(
                    "dump dropped the paragraph formatting of the {} at {}: the format cannot set it on an added paragraph",
                    node.kind(),
                    node.positional_path()
                ),
            ));
        }
    }
    if doc.runs {
        inlines(out, node, path, doc)
    } else {
        first_run(out, node, path)
    }
}

/// The runs, hyperlinks, breaks and tabs of the paragraph at `path`.
fn inlines(out: &mut Out, paragraph: &Node<'_>, path: &str, doc: &Doc) -> Result<()> {
    let mut any = false;
    let mut runs = 0_usize;
    let mut rebuilt = String::new();
    for child in children(paragraph) {
        let child = child?;
        any = true;
        match child.kind() {
            kind @ ("run" | "hyperlink") => {
                let (add, rest) = out.split(&child, kind, &["text"])?;
                let mut add = out.known_styles(&child, add, &doc.styles);
                let rest = out.known_styles(&child, rest, &doc.styles);
                let text = text(&child)?;
                rebuilt.push_str(&text);
                if !text.is_empty() {
                    // v2: a soft line break inside a run is `\v`.
                    add.push("text", text.replace('\n', "\u{b}"));
                }
                out.add(path, kind, add);
                if kind == "run" {
                    runs += 1;
                    out.set(&format!("{path}/r[{runs}]"), rest);
                }
            },
            kind @ ("break" | "tab") => {
                rebuilt.push_str(&text(&child)?);
                let props = out.props(&child, kind, Command::Add, &[])?;
                out.add(path, kind, props);
            },
            _ => out.skip(&child),
        }
    }
    if any && rebuilt != text(paragraph)? {
        // Text without a child node of its own (a field result, a content
        // control's runs) is not rebuilt.
        out.warnings.push(Error::new(
            Code::UnsupportedElement,
            format!(
                "dump skipped content of the {} at {} that has no node of its own",
                paragraph.kind(),
                paragraph.positional_path()
            ),
        ));
    }
    if !any {
        // A format without run nodes keeps the paragraph's text itself.
        let text = text(paragraph)?;
        if !text.is_empty() {
            out.add(
                path,
                "run",
                Props::new().with("text", text.replace('\n', "\u{b}")),
            );
        }
    }
    Ok(())
}

/// A format whose paragraphs are added with their text (DOC): the first
/// run takes the run formatting; a paragraph whose runs differ loses the
/// formatting of the others, reported.
fn first_run(out: &mut Out, paragraph: &Node<'_>, path: &str) -> Result<()> {
    let mut first: Option<Props<'static>> = None;
    for child in children(paragraph) {
        let child = child?;
        if child.kind() != "run" {
            out.skip(&child);
            continue;
        }
        let props = out.props(&child, "run", Command::Set, &["text"])?;
        match &first {
            None => first = Some(props),
            Some(first) if *first == props => {},
            Some(_) => out.warnings.push(Error::new(
                Code::UnsupportedProperty,
                format!(
                    "dump kept the first run's formatting only for the {} at {}",
                    paragraph.kind(),
                    paragraph.positional_path()
                ),
            )),
        }
    }
    // An empty paragraph is added without a run to format.
    if let Some(props) = first.filter(|_| !text(paragraph).unwrap_or_default().is_empty()) {
        out.set(&format!("{path}/r[1]"), props);
    }
    Ok(())
}

fn table(out: &mut Out, table: &Node<'_>, parent: &str, index: usize, doc: &Doc) -> Result<()> {
    if !out
        .capabilities
        .element("table")
        .is_some_and(|table| table.operations().add)
    {
        out.skip(table);
        return Ok(());
    }
    let mut rows: Vec<Node<'_>> = Vec::new();
    for child in children(table) {
        let child = child?;
        if child.kind() == "row" {
            rows.push(child);
        } else {
            out.skip(&child);
        }
    }
    let mut cells: Vec<Vec<Node<'_>>> = Vec::with_capacity(rows.len());
    for row in &rows {
        let mut row_cells = Vec::new();
        for cell in children(row) {
            let cell = cell?;
            if cell.kind() == "cell" {
                row_cells.push(cell);
            } else {
                out.skip(&cell);
            }
        }
        cells.push(row_cells);
    }
    let columns = cells.iter().map(Vec::len).max().unwrap_or(0);
    if rows.is_empty() || columns == 0 {
        out.skip(table);
        return Ok(());
    }
    let path = format!("{parent}/tbl[{index}]");
    let (mut add, rest) = out.split(table, "table", &["rows", "cols", "data", "colWidths"])?;
    add.push("rows", i64::try_from(rows.len()).unwrap_or(i64::MAX));
    add.push("cols", i64::try_from(columns).unwrap_or(i64::MAX));
    out.add(parent, "table", add);
    out.set(&path, rest);
    for (r, row_cells) in cells.iter().enumerate() {
        for (c, cell) in row_cells.iter().enumerate() {
            let cell_path = format!("{path}/tr[{}]/tc[{}]", r + 1, c + 1);
            let mut paragraphs = 0_usize;
            for child in children(cell) {
                let child = child?;
                if child.kind() != "paragraph" {
                    out.skip(&child);
                    continue;
                }
                paragraphs += 1;
                let paragraph_path = format!("{cell_path}/p[{paragraphs}]");
                // The created cell holds one empty paragraph.
                paragraph(
                    out,
                    &child,
                    &cell_path,
                    &paragraph_path,
                    paragraphs == 1,
                    doc,
                )?;
            }
        }
        // A short row: the created table has a cell the source lacks.
        for c in (row_cells.len()..columns).rev() {
            out.remove(&format!("{path}/tr[{}]/tc[{}]", r + 1, c + 1));
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Workbooks: sheets, cells and row layout.

fn workbook(out: &mut Out, root: &Node<'_>, format: FileFormat) -> Result<()> {
    // XLS sets a cell's formatting apart from its value or formula.
    let separate = format == FileFormat::Xls;
    let mut sheets = 0_usize;
    let mut hidden: Vec<usize> = Vec::new();
    for child in children(root) {
        let child = child?;
        if child.kind() != "sheet" {
            out.skip(&child);
            continue;
        }
        sheets += 1;
        let props = out.props(&child, "sheet", Command::Add, &["hidden"])?;
        if sheets == 1 {
            // The created workbook has its first sheet already.
            out.set("/sheet[1]", props);
        } else {
            out.add("/", "sheet", props);
        }
        if child
            .view()
            .format()?
            .any(|prop| prop.key() == "hidden" && matches!(prop.value(), Value::Bool(true)))
        {
            hidden.push(sheets);
        }
        sheet(out, &child, sheets, separate)?;
    }
    for index in hidden {
        out.set(
            &format!("/sheet[{index}]"),
            Props::new().with("hidden", true),
        );
    }
    Ok(())
}

fn sheet(out: &mut Out, sheet: &Node<'_>, index: usize, separate: bool) -> Result<()> {
    for child in children(sheet) {
        let child = child?;
        match child.kind() {
            "row" => {
                let row_path = format!("/sheet[{index}]/{}", last_step(&child));
                let props = out.props(&child, "row", Command::Set, &[])?;
                out.set(&row_path, props);
                for cell in children(&child) {
                    let cell = cell?;
                    if cell.kind() == "cell" {
                        cell_item(out, &cell, index, separate)?;
                    } else {
                        out.skip(&cell);
                    }
                }
            },
            "cell" => cell_item(out, &child, index, separate)?,
            _ => out.skip(&child),
        }
    }
    Ok(())
}

/// The last step of a node's positional path (`A1`, `row[3]`).
fn last_step(node: &Node<'_>) -> String {
    let path = node.positional_path().to_string();
    path.rsplit('/').next().unwrap_or_default().to_owned()
}

fn cell_item(out: &mut Out, cell: &Node<'_>, sheet: usize, separate: bool) -> Result<()> {
    let path = format!("/sheet[{sheet}]/{}", last_step(cell));
    let mut kind = None;
    let mut formula = None;
    for prop in cell.view().format()? {
        match prop.key() {
            "type" => kind = Some(prop.to_string()),
            "formula" => formula = Some(prop.to_string()),
            _ => {},
        }
    }
    let mut props = out.props(
        cell,
        "cell",
        Command::Set,
        &["value", "formula", "type", "clear"],
    )?;
    if props
        .get("numberformat")
        .and_then(PropValue::as_text)
        .is_some_and(|format| format.eq_ignore_ascii_case("General"))
    {
        // The default format reads back as no format at all.
        props = props
            .iter()
            .filter(|(key, _)| key.as_str() != "numberformat")
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect();
    }
    let text = text(cell)?;
    let mut content = Props::new();
    if let Some(formula) = formula.filter(|formula| !formula.is_empty()) {
        content.push("formula", formula);
    } else if !text.is_empty() {
        let value = match kind.as_deref() {
            Some("Number" | "number" | "float" | "currency" | "percentage") => text
                .parse::<f64>()
                .map_or_else(|_| PropValue::from(text.clone()), PropValue::Number),
            Some("Boolean" | "boolean" | "Bool") => match text.as_str() {
                "TRUE" | "true" | "1" => PropValue::Bool(true),
                "FALSE" | "false" | "0" => PropValue::Bool(false),
                _ => PropValue::from(text.clone()),
            },
            _ => PropValue::from(text.clone()),
        };
        content.push("value", value);
    }
    if separate {
        out.set(&path, content);
        out.set(&path, props);
    } else {
        for (key, value) in content.iter() {
            props.push(key.clone(), value.clone());
        }
        out.set(&path, props);
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Presentations: slides by layout name, placeholders, shapes and notes.

fn presentation(out: &mut Out, root: &Node<'_>, format: FileFormat) -> Result<()> {
    // The replay target's layouts and the placeholders each instantiates.
    let blank = crate::create(format)?;
    let mut layouts: Vec<(String, Vec<String>)> = Vec::new();
    let mut slides = 0_usize;
    for child in children(root) {
        let child = child?;
        if child.kind() != "slide" {
            // Masters, layouts and the theme come with the created deck.
            continue;
        }
        slides += 1;
        let layout = child
            .view()
            .format()?
            .find(|prop| prop.key() == "layout")
            .map(|prop| prop.to_string())
            .unwrap_or_default();
        let placeholders = match layouts.iter().find(|(name, _)| *name == layout) {
            Some((_, placeholders)) => Some(placeholders.clone()),
            None => {
                let probe = blank.add(
                    "/",
                    crate::NewNode::new("slide")
                        .with_props(Props::new().with("layout", layout.as_str()))?,
                );
                match probe {
                    Ok(added) => {
                        let slide = added.file.get(&added.path)?;
                        let mut types = Vec::new();
                        for shape in children(&slide) {
                            let shape = shape?;
                            if matches!(shape.kind(), "title" | "placeholder") {
                                types.push(placeholder_type(&shape)?.unwrap_or_default());
                            }
                        }
                        layouts.push((layout.clone(), types.clone()));
                        Some(types)
                    },
                    Err(_) => None,
                }
            },
        };
        let Some(mut placeholders) = placeholders.map(|types| {
            types
                .into_iter()
                .map(|kind| (kind, false))
                .collect::<Vec<(String, bool)>>()
        }) else {
            out.skip(&child);
            slides -= 1;
            continue;
        };
        let mut props = out.props(&child, "slide", Command::Add, &["layout", "name", "hidden"])?;
        props.push("layout", layout);
        out.add("/", "slide", props);
        let path = format!("/slide[{slides}]");
        let mut slide_props = out.props(
            &child,
            "slide",
            Command::Set,
            &["layout", "notes", "transition"],
        )?;
        if slide_props.get("name").is_none() {
            // `add` names a new slide; an unnamed source slide gets the
            // empty name, which reads back the same way on every replay.
            slide_props.push("name", "");
        }
        out.set(&path, slide_props);
        let mut added = Vec::new();
        for shape in children(&child) {
            let shape = shape?;
            match shape.kind() {
                "title" | "placeholder" => {
                    let kind = placeholder_type(&shape)?.unwrap_or_default();
                    let slot = placeholders
                        .iter()
                        .position(|(existing, used)| !*used && *existing == kind);
                    let Some(slot) = slot else {
                        out.skip(&shape);
                        continue;
                    };
                    placeholders[slot].1 = true;
                    let mut props = out.props(&shape, "placeholder", Command::Set, &["text"])?;
                    // A placeholder inherits its transform from the layout;
                    // the PPTX writer sets geometry only on a complete local
                    // transform.
                    const GEOMETRY: [&str; 4] = ["x", "y", "width", "height"];
                    if props
                        .iter()
                        .any(|(key, _)| GEOMETRY.contains(&key.as_str()))
                    {
                        props = props
                            .iter()
                            .filter(|(key, _)| !GEOMETRY.contains(&key.as_str()))
                            .map(|(key, value)| (key.clone(), value.clone()))
                            .collect();
                        out.warnings.push(Error::new(
                            Code::UnsupportedProperty,
                            format!(
                                "dump dropped the geometry of the {} at {}: a placeholder's geometry is not writable",
                                shape.kind(),
                                shape.positional_path()
                            ),
                        ).with_property("x"));
                    }
                    let text = text(&shape)?;
                    if !text.is_empty() {
                        props.push("text", text);
                    }
                    out.set(&format!("{path}/shape[{}]", slot + 1), props);
                },
                "shape" => added.push(shape),
                "notes" => {
                    let text = text(&shape)?;
                    if !text.is_empty() {
                        out.set(&format!("{path}/notes"), Props::new().with("text", text));
                    }
                },
                _ => out.skip(&shape),
            }
        }
        // The created slide's placeholders the source slide lacks.
        for (slot, (_, used)) in placeholders.iter().enumerate().rev() {
            if !used {
                out.remove(&format!("{path}/shape[{}]", slot + 1));
            }
        }
        let mut count = placeholders.iter().filter(|(_, used)| *used).count();
        for shape in added {
            let props = out.props(&shape, "shape", Command::Add, &["text"])?;
            out.add(&path, "shape", props);
            count += 1;
            // Text goes through `set`, whose `\n` separates paragraphs.
            let text = text(&shape)?;
            if !text.is_empty() {
                out.set(
                    &format!("{path}/shape[{count}]"),
                    Props::new().with("text", text),
                );
            }
        }
    }
    Ok(())
}

/// A placeholder's `phType`, `None` for other shapes.
fn placeholder_type(node: &Node<'_>) -> Result<Option<String>> {
    Ok(node
        .view()
        .format()?
        .find(|prop| prop.key() == "phType")
        .map(|prop| prop.to_string()))
}
