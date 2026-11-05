//! The ODT parts of record 0328: master-page headers and footers, list
//! styles as numbering definitions, instances and levels, body text fields,
//! tables of contents, tab stops, drawing shapes, formula objects and table
//! columns.
//!
//! Every part reads through the leaf part index
//! ([`litchi_odt::body::Parts`]), built once per handle on the first part
//! read, so the paths, queries and views of the structural kinds never pay
//! for it. Parts are reached by path and by queries that name them; the
//! views of the root, the body, paragraphs and tables list the structural
//! kinds only (record 0324 kept DOC's views the same way).

use std::borrow::Cow;
use std::ops::Range;

use litchi_odt::body::parts::{ListLevel, ListLevelKind, RegionKind, Shape, TabStop};
use litchi_odt::body::{Body, Kind as LeafKind, NodeId, Parts, PartsView, TextField};

use super::node::{Node, Part, Step, Target};
use super::spec::Kind;
use super::{Odt, core_error};
use crate::error::Result;
use crate::node::Value;
use crate::path::navigate::KindSet;
use crate::value::{Length, LengthContext};

/// The part index of one handle: the leaf index, the body's text fields
/// (by their byte ranges) and its formula objects.
pub(crate) struct Index {
    parts: Parts,
    fields: Vec<Range<usize>>,
    equations: Vec<Equation>,
}

/// One formula object: its index among the leaf objects and its text.
struct Equation {
    object: usize,
    formula: String,
}

impl Index {
    /// Index the parts of `odt`'s body; formula objects stored as package
    /// subdocuments are read here, once.
    pub(super) fn build(odt: &Odt) -> Result<Self> {
        let body = odt.body()?;
        let parts = Parts::build(body).map_err(core_error)?;
        let fields = litchi_odt::body::fields(body)
            .map_err(core_error)?
            .iter()
            .map(TextField::range)
            .collect();
        let mut equations = Vec::new();
        for (object, embedded) in parts.view(body).objects().enumerate() {
            let formula = match embedded.inline_math() {
                Some(math) => litchi_odt::body::parts::formula_text(math),
                None => match embedded.href().as_deref().and_then(subdocument) {
                    Some(path) => odt
                        .member(&path)?
                        .and_then(|bytes| String::from_utf8(bytes).ok())
                        .and_then(|xml| litchi_odt::body::parts::formula_text(&xml)),
                    None => None,
                },
            };
            if let Some(formula) = formula {
                equations.push(Equation { object, formula });
            }
        }
        Ok(Self {
            parts,
            fields,
            equations,
        })
    }

    fn view<'a>(&'a self, body: Body<'a>) -> PartsView<'a> {
        self.parts.view(body)
    }
}

/// The content part of the package subdocument an object's `xlink:href`
/// names (`./Object 1` → `Object 1/content.xml`), or `None` for a link
/// outside the package.
fn subdocument(href: &str) -> Option<String> {
    let path = href
        .strip_prefix("./")
        .unwrap_or(href)
        .trim_end_matches('/');
    if path.is_empty() || path.contains(':') || path.starts_with('/') || path.starts_with("..") {
        return None;
    }
    Some(format!("{path}/content.xml"))
}

/// The body and part view of a handle.
fn view(odt: &Odt) -> Result<(Body<'_>, PartsView<'_>, &Index)> {
    let body = odt.body()?;
    let index = odt.parts()?;
    Ok((body, index.view(body), index))
}

fn count(value: usize) -> u32 {
    u32::try_from(value).unwrap_or(u32::MAX)
}

/// The header (or footer) regions, as indexes among every region.
pub(super) fn regions(odt: &Odt, header: bool) -> Result<Vec<u32>> {
    let (_, view, _) = view(odt)?;
    Ok(view
        .regions()
        .filter(|region| region.kind().is_header() == header)
        .map(|region| count(region.index()))
        .collect())
}

/// The DOCX `type` of a region: `default`, `even` or `first`.
const fn region_type(kind: RegionKind) -> &'static str {
    match kind {
        RegionKind::HeaderLeft | RegionKind::FooterLeft => "even",
        RegionKind::HeaderFirst | RegionKind::FooterFirst => "first",
        _ => "default",
    }
}

/// The type of region `index`.
pub(super) fn region_type_of(odt: &Odt, index: u32) -> Result<Option<&'static str>> {
    let (_, view, _) = view(odt)?;
    Ok(view
        .regions()
        .nth(index as usize)
        .map(|region| region_type(region.kind())))
}

/// The number of list styles.
pub(super) fn list_count(odt: &Odt) -> Result<usize> {
    let (_, view, _) = view(odt)?;
    Ok(view.list_styles().len())
}

/// The number of levels list style `list` declares.
pub(super) fn level_count(odt: &Odt, list: u32) -> Result<usize> {
    let (_, view, _) = view(odt)?;
    Ok(view
        .list_styles()
        .nth(list as usize)
        .map_or(0, |style| style.levels().len()))
}

/// The level of list style `list` whose `text:level` is `ilvl + 1`.
pub(super) fn level_by_ilvl(odt: &Odt, list: u32, ilvl: u16) -> Result<Option<usize>> {
    let (_, view, _) = view(odt)?;
    Ok(view.list_styles().nth(list as usize).and_then(|style| {
        style
            .levels()
            .position(|level| level.level() == ilvl.saturating_add(1))
    }))
}

/// The number of body text fields.
pub(super) fn field_count(odt: &Odt) -> Result<usize> {
    Ok(odt.parts()?.fields.len())
}

/// The number of tables of contents.
pub(super) fn toc_count(odt: &Odt) -> Result<usize> {
    let (_, view, _) = view(odt)?;
    Ok(view.tables_of_contents().len())
}

/// The number of shapes.
pub(super) fn shape_count(odt: &Odt) -> Result<usize> {
    let (_, view, _) = view(odt)?;
    Ok(view.shapes().len())
}

/// The number of equations.
pub(super) fn equation_count(odt: &Odt) -> Result<usize> {
    Ok(odt.parts()?.equations.len())
}

/// The number of tab stops of a paragraph or heading.
pub(super) fn tab_count(odt: &Odt, paragraph: NodeId) -> Result<usize> {
    let body = odt.body()?;
    Ok(body
        .node(paragraph)
        .map_or(0, |node| litchi_odt::body::parts::tab_stops(node).len()))
}

/// The number of declared columns of a table.
pub(super) fn column_count(odt: &Odt, table: NodeId) -> Result<usize> {
    let (body, view, _) = view(odt)?;
    Ok(body
        .node(table)
        .map_or(0, |node| view.columns(node).count()))
}

/// A part node under `parent`, `ordinal`th of its kind there.
pub(super) fn node<'f>(parent: &Node<'f>, kind: Kind, part: Part, ordinal: usize) -> Node<'f> {
    Node::new(
        parent.odt,
        kind,
        Target::Part(part),
        parent.address().with(Step::new(kind, count(ordinal))),
    )
}

/// The parts listed under `parent` whose kinds are in `kinds`, in the
/// order paths number them.
pub(super) fn children<'f>(parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Vec<Node<'f>>> {
    let odt = parent.odt;
    let mut out = Vec::new();
    let wants = |kind: Kind| kinds.contains(kind);
    match (parent.kind, parent.target) {
        (Kind::Document, _) => {
            for (kind, header) in [(Kind::Header, true), (Kind::Footer, false)] {
                if wants(kind) {
                    for (ordinal, region) in regions(odt, header)?.into_iter().enumerate() {
                        out.push(node(parent, kind, Part::Region(region), ordinal));
                    }
                }
            }
            if wants(Kind::Numbering) && list_count(odt)? > 0 {
                out.push(node(parent, Kind::Numbering, Part::Numbering, 0));
            }
            if wants(Kind::Field) {
                for ordinal in 0..field_count(odt)? {
                    out.push(node(
                        parent,
                        Kind::Field,
                        Part::Field(count(ordinal)),
                        ordinal,
                    ));
                }
            }
            if wants(Kind::Toc) {
                for ordinal in 0..toc_count(odt)? {
                    out.push(node(parent, Kind::Toc, Part::Toc(count(ordinal)), ordinal));
                }
            }
        },
        (Kind::Body, _) => {
            if wants(Kind::Shape) {
                for ordinal in 0..shape_count(odt)? {
                    out.push(node(
                        parent,
                        Kind::Shape,
                        Part::Shape(count(ordinal)),
                        ordinal,
                    ));
                }
            }
            if wants(Kind::Equation) {
                for ordinal in 0..equation_count(odt)? {
                    out.push(node(
                        parent,
                        Kind::Equation,
                        Part::Equation(count(ordinal)),
                        ordinal,
                    ));
                }
            }
        },
        (Kind::Paragraph | Kind::Heading, Target::Body(_, id)) if wants(Kind::Tab) => {
            for ordinal in 0..tab_count(odt, id)? {
                out.push(node(
                    parent,
                    Kind::Tab,
                    Part::Tab(id, count(ordinal)),
                    ordinal,
                ));
            }
        },
        (Kind::Table, Target::Body(_, id)) if wants(Kind::Column) => {
            for ordinal in 0..column_count(odt, id)? {
                out.push(node(
                    parent,
                    Kind::Column,
                    Part::Column(id, count(ordinal)),
                    ordinal,
                ));
            }
        },
        (Kind::Numbering, _) => {
            let lists = list_count(odt)?;
            for kind in [Kind::AbstractNum, Kind::Num] {
                if wants(kind) {
                    for ordinal in 0..lists {
                        out.push(node(parent, kind, Part::List(count(ordinal)), ordinal));
                    }
                }
            }
        },
        (Kind::AbstractNum, Target::Part(Part::List(list))) if wants(Kind::Level) => {
            for ordinal in 0..level_count(odt, list)? {
                out.push(node(
                    parent,
                    Kind::Level,
                    Part::Level(list, count(ordinal)),
                    ordinal,
                ));
            }
        },
        _ => {},
    }
    Ok(out)
}

/// Whether a node of `kind` lists parts of a kind in `kinds`.
pub(super) fn has_part_children(kind: Kind, kinds: KindSet<Kind>) -> bool {
    match kind {
        Kind::Document => [
            Kind::Header,
            Kind::Footer,
            Kind::Numbering,
            Kind::Field,
            Kind::Toc,
        ]
        .iter()
        .any(|kind| kinds.contains(*kind)),
        Kind::Body => kinds.contains(Kind::Shape) || kinds.contains(Kind::Equation),
        Kind::Paragraph | Kind::Heading => kinds.contains(Kind::Tab),
        Kind::Table => kinds.contains(Kind::Column),
        Kind::Numbering => kinds.contains(Kind::AbstractNum) || kinds.contains(Kind::Num),
        Kind::AbstractNum => kinds.contains(Kind::Level),
        _ => false,
    }
}

/// The text of a part node: a region's, a field's displayed text, a table
/// of contents' index body, a shape's or an equation's formula.
pub(super) fn text<'f>(odt: &'f Odt, part: Part) -> Result<Option<Cow<'f, str>>> {
    let (body, view, index) = view(odt)?;
    Ok(match part {
        Part::Region(region) => match view.regions().nth(region as usize) {
            Some(region) => Some(Cow::Owned(region.text().map_err(core_error)?)),
            None => None,
        },
        Part::Field(field) => index
            .fields
            .get(field as usize)
            .and_then(|range| TextField::at(body, range.clone()))
            .map(|field| Cow::Owned(field.text())),
        Part::Toc(toc) => match view.tables_of_contents().nth(toc as usize) {
            Some(toc) => Some(Cow::Owned(toc.text().map_err(core_error)?)),
            None => None,
        },
        Part::Shape(shape) => match view.shapes().nth(shape as usize) {
            Some(shape) => Some(Cow::Owned(shape.text().map_err(core_error)?)),
            None => None,
        },
        Part::Equation(equation) => index
            .equations
            .get(equation as usize)
            .map(|equation| Cow::Borrowed(equation.formula.as_str())),
        Part::Numbering | Part::List(_) | Part::Level(..) | Part::Tab(..) | Part::Column(..) => {
            None
        },
    })
}

fn text_value<'a>(value: Option<Cow<'a, str>>) -> Option<Value<'a>> {
    value.map(Value::Text)
}

fn int(value: usize) -> Value<'static> {
    Value::Int(i64::try_from(value).unwrap_or(i64::MAX))
}

/// An ODF length as a unified length.
fn length(raw: Option<&str>) -> Option<Value<'static>> {
    Length::parse(raw?).ok().map(Value::Length)
}

/// An ODF length in integer twips (the DOCX level and tab-stop unit).
fn twips(raw: &str) -> Option<i64> {
    let emu = Length::parse(raw).ok()?.emu();
    Some((emu + emu.signum() * 317) / 635)
}

fn color(raw: Option<&str>) -> Option<Value<'static>> {
    super::props::color(raw)
}

/// The value at `index` of a part's property table.
pub(super) fn value<'f>(odt: &'f Odt, kind: Kind, part: Part, index: usize) -> Option<Value<'f>> {
    let (body, view, parts) = view(odt).ok()?;
    match (kind, part) {
        (Kind::Header | Kind::Footer, Part::Region(region)) => {
            let region = view.regions().nth(region as usize)?;
            match index {
                0 => region.text().ok().map(|text| Value::Text(Cow::Owned(text))),
                1 => Some(Value::Token(region_type(region.kind()))),
                2 => text_value(region.master_name()),
                _ => None,
            }
        },
        (Kind::Numbering, _) => match index {
            0 | 1 => Some(int(view.list_styles().len())),
            _ => None,
        },
        (Kind::AbstractNum, Part::List(list)) => {
            let style = view.list_styles().nth(list as usize)?;
            match index {
                0 => Some(int(list as usize)),
                1 => Some(Value::Token(if style.levels().len() > 1 {
                    "multilevel"
                } else {
                    "singleLevel"
                })),
                2 => text_value(style.display_name().or_else(|| style.name())),
                3 => text_value(style.name()),
                _ => None,
            }
        },
        (Kind::Num, Part::List(list)) => match index {
            0 | 1 => Some(int(list as usize)),
            _ => None,
        },
        (Kind::Level, Part::Level(list, level)) => {
            let style = view.list_styles().nth(list as usize)?;
            level_value(style.levels().nth(level as usize)?, index)
        },
        (Kind::Field, Part::Field(field)) => {
            let field = TextField::at(body, parts.fields.get(field as usize)?.clone())?;
            field_value(&field, index)
        },
        (Kind::Toc, Part::Toc(toc)) => {
            let toc = view.tables_of_contents().nth(toc as usize)?;
            match index {
                0 => toc
                    .outline_level()
                    .map(|level| Value::Text(Cow::Owned(format!("1-{level}")))),
                1 => text_value(toc.title()),
                2 => Some(Value::Bool(toc.links())),
                3 => Some(Value::Bool(toc.page_numbers())),
                4 => text_value(toc.name()),
                _ => None,
            }
        },
        (Kind::Tab, Part::Tab(paragraph, tab)) => {
            let node = body.node(paragraph)?;
            let stops = litchi_odt::body::parts::tab_stops(node);
            tab_value(stops.get(tab as usize)?, index)
        },
        (Kind::Shape, Part::Shape(shape)) => shape_value(view.shapes().nth(shape as usize)?, index),
        (Kind::Equation, Part::Equation(equation)) => {
            let equation = parts.equations.get(equation as usize)?;
            let object = view.objects().nth(equation.object)?;
            match index {
                0 => Some(Value::Text(Cow::Borrowed(equation.formula.as_str()))),
                1 => Some(Value::Token(
                    if object.anchor_type().as_deref() == Some("as-char") {
                        "inline"
                    } else {
                        "display"
                    },
                )),
                2 => text_value(object.name()),
                3 => length(object.width().as_deref()),
                4 => length(object.height().as_deref()),
                _ => None,
            }
        },
        (Kind::Column, Part::Column(table, column)) => {
            let table = body.node(table)?;
            if table.kind() != LeafKind::Table {
                return None;
            }
            let column = view.columns(table).nth(column as usize)?;
            match index {
                0 => length(column.width().as_deref()),
                _ => None,
            }
        },
        _ => None,
    }
}

/// The DOCX `numFmt` of an ODF `style:num-format`.
fn num_format(format: &str) -> Cow<'static, str> {
    Cow::Borrowed(match format {
        "1" => "decimal",
        "a" => "lowerLetter",
        "A" => "upperLetter",
        "i" => "lowerRoman",
        "I" => "upperRoman",
        "" => "none",
        other => return Cow::Owned(other.to_owned()),
    })
}

fn level_value<'a>(level: ListLevel<'a>, index: usize) -> Option<Value<'a>> {
    let number = level.kind() == ListLevelKind::Number;
    match index {
        0 => Some(int(usize::from(level.level().saturating_sub(1)))),
        1 => number.then(|| Value::Int(i64::from(level.start_value().unwrap_or(1)))),
        2 => Some(Value::Text(if number {
            num_format(level.num_format().as_deref().unwrap_or(""))
        } else {
            Cow::Borrowed("bullet")
        })),
        3 => level_text(level).map(|text| Value::Text(Cow::Owned(text))),
        4 => level.text_align().and_then(|align| {
            Some(Value::Token(match &*align {
                "start" | "left" => "left",
                "center" => "center",
                "end" | "right" => "right",
                _ => return None,
            }))
        }),
        5 => level.label_followed_by().and_then(|follow| {
            Some(Value::Token(match &*follow {
                "listtab" => "tab",
                "space" => "space",
                "nothing" => "nothing",
                _ => return None,
            }))
        }),
        6 => match level.margin_left() {
            Some(margin) => twips(&margin).map(Value::Int),
            None => {
                let before = level.space_before().and_then(|value| twips(&value));
                let width = level.min_label_width().and_then(|value| twips(&value));
                match (before, width) {
                    (None, None) => None,
                    (before, width) => Some(Value::Int(
                        before.unwrap_or(0).saturating_add(width.unwrap_or(0)),
                    )),
                }
            },
        },
        7 => match level.text_indent() {
            Some(indent) => twips(&indent).map(|value| Value::Int(-value)),
            None => level
                .min_label_width()
                .and_then(|value| twips(&value))
                .map(Value::Int),
        },
        8 => text_value(level.font()),
        9 => level.font_size().and_then(|size| {
            Length::parse_in(&size, LengthContext::FontSize)
                .or_else(|_| Length::parse(&size))
                .ok()
                .map(Value::Length)
        }),
        10 => color(level.color().as_deref()),
        11 => level.bold().map(Value::Bool),
        12 => level.italic().map(Value::Bool),
        _ => None,
    }
}

/// A level's label in the DOCX `lvlText` form: `%N` for level N's number,
/// the levels a label shows joined by `.`, between the prefix and the
/// suffix; a bullet level's character.
fn level_text(level: ListLevel<'_>) -> Option<String> {
    let mut out = String::new();
    if let Some(prefix) = level.prefix() {
        out.push_str(&prefix);
    }
    match level.kind() {
        ListLevelKind::Number => {
            if level.num_format().as_deref() != Some("") {
                let own = level.level().max(1);
                let shown = level.display_levels().unwrap_or(1).clamp(1, own);
                for (position, number) in (own - shown + 1..=own).enumerate() {
                    if position > 0 {
                        out.push('.');
                    }
                    out.push('%');
                    out.push_str(&number.to_string());
                }
            }
        },
        ListLevelKind::Bullet => out.push_str(&level.bullet_char()?),
        _ => return None,
    }
    if let Some(suffix) = level.suffix() {
        out.push_str(&suffix);
    }
    Some(out)
}

/// The Word field type of an ODF field element, or the element name when
/// Word has no such field.
fn field_type<'a>(field: &TextField<'a>) -> Cow<'a, str> {
    let name = field.name();
    Cow::Borrowed(match name {
        "page-number" => "page",
        "page-count" => "numpages",
        "date" => "date",
        "time" => "time",
        "title" => "title",
        "subject" => "subject",
        "description" => "comments",
        "keywords" => "keywords",
        "initial-creator" | "author-name" => "author",
        "creator" => "lastsavedby",
        "creation-date" | "creation-time" => "createdate",
        "modification-date" | "modification-time" => "savedate",
        "print-date" | "print-time" => "printdate",
        "editing-duration" => "edittime",
        "editing-cycles" => "revnum",
        "word-count" => "numwords",
        "character-count" => "numchars",
        "file-name" => "filename",
        "template-name" => "template",
        "sequence" => "seq",
        "note-ref" => "noteref",
        "user-defined" => "docproperty",
        "database-display" => "mergefield",
        "conditional-text" => "if",
        "chapter" => "styleref",
        "bookmark-ref" | "reference-ref" | "sequence-ref" => {
            if field.attribute("reference-format").as_deref() == Some("page") {
                "pageref"
            } else {
                "ref"
            }
        },
        other => other,
    })
}

fn field_value<'a>(field: &TextField<'a>, index: usize) -> Option<Value<'a>> {
    match index {
        0 => Some(Value::Text(field_type(field))),
        1 => text_value(
            field
                .attribute("name")
                .or_else(|| field.attribute("ref-name")),
        ),
        2 => text_value(
            field
                .attribute("num-format")
                .or_else(|| field.attribute("data-style-name")),
        ),
        3 => Some(Value::Bool(
            field.attribute("fixed").as_deref() == Some("true"),
        )),
        4 => Some(Value::Bool(!field.text().is_empty())),
        5 => Some(Value::Text(Cow::Borrowed(field.name()))),
        _ => None,
    }
}

fn tab_value(stop: &TabStop<'_>, index: usize) -> Option<Value<'static>> {
    match index {
        0 => twips(stop.position()).map(Value::Int),
        1 => Some(Value::Token(match stop.kind() {
            "center" => "center",
            "right" => "right",
            "char" => "decimal",
            _ => "left",
        })),
        2 => Some(Value::Token(
            match (stop.leader_text(), stop.leader_style()) {
                (Some("."), _) => "dot",
                (Some("-"), _) => "hyphen",
                (Some("_"), _) => "underscore",
                (Some("·"), _) => "middleDot",
                (_, Some("dotted")) => "dot",
                (_, Some("dash" | "long-dash" | "dot-dash" | "dot-dot-dash")) => "hyphen",
                (_, Some("solid" | "wave")) => "underscore",
                _ => "none",
            },
        )),
        _ => None,
    }
}

/// The DOCX preset name of an ODF geometry.
fn geometry(raw: &str) -> Cow<'static, str> {
    Cow::Borrowed(match raw {
        "rect" | "rectangle" => "rect",
        "round-rectangle" => "roundRect",
        "ellipse" | "circle" => "ellipse",
        "isosceles-triangle" => "triangle",
        "right-triangle" => "rtTriangle",
        "diamond" => "diamond",
        "pentagon" => "pentagon",
        "hexagon" => "hexagon",
        "octagon" => "octagon",
        "parallelogram" => "parallelogram",
        "trapezoid" => "trapezoid",
        "line" => "line",
        "can" => "can",
        "cube" => "cube",
        "heart" => "heart",
        "smiley" => "smileyFace",
        "sun" => "sun",
        "moon" => "moon",
        "star4" => "star4",
        "star5" => "star5",
        "star6" => "star6",
        "star8" => "star8",
        "star12" => "star12",
        "star24" => "star24",
        "right-arrow" => "rightArrow",
        "left-arrow" => "leftArrow",
        "up-arrow" => "upArrow",
        "down-arrow" => "downArrow",
        other => return Cow::Owned(other.to_owned()),
    })
}

fn shape_value<'a>(shape: Shape<'a>, index: usize) -> Option<Value<'a>> {
    match index {
        0 => Some(Value::Text(geometry(&shape.geometry()))),
        1 => text_value(shape.name()),
        2 => length(shape.width().as_deref()),
        3 => length(shape.height().as_deref()),
        4 => length(shape.x().as_deref()),
        5 => length(shape.y().as_deref()),
        6 => shape.anchor_type().and_then(|anchor| {
            Some(Value::Token(match &*anchor {
                "paragraph" => "paragraph",
                "char" => "char",
                "as-char" => "as-char",
                "page" => "page",
                "frame" => "frame",
                _ => return None,
            }))
        }),
        7 => shape.wrap().and_then(|wrap| {
            Some(Value::Token(match &*wrap {
                "none" => "topAndBottom",
                "run-through" => "through",
                "parallel" | "dynamic" | "left" | "right" | "biggest" => "square",
                _ => return None,
            }))
        }),
        8 => {
            if shape.fill().as_deref() == Some("none") {
                return None;
            }
            color(shape.fill_color().as_deref())
        },
        9 => {
            if shape.stroke().as_deref() == Some("none") {
                return None;
            }
            color(shape.stroke_color().as_deref())
        },
        10 => length(shape.stroke_width().as_deref()),
        _ => None,
    }
}

/// Region `index` of the handle's part index.
pub(super) fn region(odt: &Odt, index: u32) -> Result<Option<litchi_odt::body::parts::Region<'_>>> {
    let (_, view, _) = view(odt)?;
    Ok(view.regions().nth(index as usize))
}

/// The tab stops of a paragraph or heading as leaf edits, in declaration
/// order, with each stop's position in twips.
pub(super) fn tab_edits(
    odt: &Odt,
    paragraph: NodeId,
) -> Result<Vec<(Option<i64>, litchi_odt::body::TabStopSpec)>> {
    let body = odt.body()?;
    let Some(node) = body.node(paragraph) else {
        return Ok(Vec::new());
    };
    Ok(litchi_odt::body::parts::tab_stops(node)
        .iter()
        .map(|stop| {
            let mut edit = litchi_odt::body::TabStopSpec::new(stop.position());
            edit.kind = stop.kind_attribute().map(str::to_owned);
            edit.leader_style = stop.leader_style().map(str::to_owned);
            edit.leader_text = stop.leader_text().map(str::to_owned);
            edit.character = stop.character().map(str::to_owned);
            (twips(stop.position()), edit)
        })
        .collect())
}
