//! Document parts beyond the body index (record 0328): list styles and
//! their levels, master-page headers and footers, text fields, tables of
//! contents, drawing shapes, embedded objects and table columns.
//!
//! [`Parts`] is built on demand, in one pass over each of `content.xml` and
//! `styles.xml`, from the same bytes as the body [`super::Tree`]; it keeps
//! offsets only, and its views read names, attributes and text from the
//! source bytes when asked. Building it never changes the body index, so
//! every body read and edit is unaffected.
//!
//! ```no_run
//! use litchi_odt::Package;
//! use litchi_odt::body::Parts;
//!
//! # fn main() -> litchi_core::Result<()> {
//! let document = Package::open("report.odt")?.document()?;
//! let body = document.body()?;
//! let parts = Parts::build(body)?;
//! for region in parts.view(body).regions() {
//!     println!("{:?}: {}", region.kind(), region.text()?);
//! }
//! # Ok(())
//! # }
//! ```

use std::borrow::Cow;
use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::xml::{self, Ns, Prefixes};
use super::{Body, Kind, Node, StylePart};

/// The most entries of one kind a [`Parts`] index holds.
const MAX_ENTRIES: usize = 1_000_000;
/// The deepest nesting the scans follow.
const MAX_DEPTH: usize = 4_096;

/// Offsets of one element: start, end of start tag, start of end tag, end.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct Span {
    start: u32,
    open_end: u32,
    close_start: u32,
    end: u32,
}

impl Span {
    fn new(start: u32, open_end: u32, empty: bool) -> Self {
        Self {
            start,
            open_end,
            close_start: if empty { open_end } else { u32::MAX },
            end: if empty { open_end } else { u32::MAX },
        }
    }

    fn close(&mut self, start: u32, end: u32) {
        self.close_start = start;
        self.end = end;
    }

    fn range(self) -> Range<usize> {
        self.start as usize..self.end as usize
    }

    fn tag(self) -> Range<usize> {
        self.start as usize..self.open_end as usize
    }

    fn content(self) -> Range<usize> {
        self.open_end as usize..self.close_start.max(self.open_end) as usize
    }
}

#[derive(Clone, Debug)]
struct ListData {
    part: StylePart,
    span: Span,
    levels: Vec<LevelData>,
}

#[derive(Clone, Debug)]
struct LevelData {
    span: Span,
    kind: ListLevelKind,
    level: u16,
    /// `style:list-level-properties` start tag.
    properties: Option<Range<u32>>,
    /// `style:list-level-label-alignment` start tag.
    alignment: Option<Range<u32>>,
    /// `style:text-properties` start tag.
    text: Option<Range<u32>>,
}

#[derive(Clone, Debug)]
struct MasterData {
    span: Span,
}

#[derive(Clone, Debug)]
struct RegionData {
    master: u32,
    kind: RegionKind,
    span: Span,
}

#[derive(Clone, Debug)]
struct TocData {
    span: Span,
    /// `text:table-of-content-source` start tag.
    source: Option<Range<u32>>,
    /// `text:index-title-template` element.
    title: Option<Span>,
    /// `text:index-body` element.
    index_body: Option<Span>,
    links: bool,
    page_numbers: bool,
}

#[derive(Clone, Debug)]
struct ShapeData {
    span: Span,
    /// `draw:enhanced-geometry` start tag of a custom shape.
    geometry: Option<Range<u32>>,
}

#[derive(Clone, Debug)]
struct ObjectData {
    /// The `draw:frame` holding the object.
    frame: Span,
    /// The `draw:object` or `draw:object-ole` element.
    object: Span,
    ole: bool,
    /// An inline `math:math` payload.
    math: Option<Span>,
}

#[derive(Clone, Debug)]
struct ColumnData {
    /// Start of the owning `table:table` element.
    table: u32,
    span: Span,
    repeat: u32,
}

/// The part index of one document: list styles, master pages and their
/// regions, tables of contents, shapes, embedded objects and table columns.
/// Built once per document by its owner; see the module documentation.
#[derive(Debug, Default)]
pub struct Parts {
    style_prefixes: Prefixes,
    lists: Vec<ListData>,
    masters: Vec<MasterData>,
    regions: Vec<RegionData>,
    tocs: Vec<TocData>,
    shapes: Vec<ShapeData>,
    objects: Vec<ObjectData>,
    columns: Vec<ColumnData>,
}

/// A borrowed view of [`Parts`] over the parts it indexes.
#[derive(Clone, Copy)]
pub struct PartsView<'a> {
    parts: &'a Parts,
    body: Body<'a>,
}

impl std::fmt::Debug for PartsView<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PartsView")
            .field("lists", &self.parts.lists.len())
            .field("regions", &self.parts.regions.len())
            .finish_non_exhaustive()
    }
}

fn offset(value: u64) -> Result<u32> {
    u32::try_from(value).map_err(|_error| Error::InvalidFormat("ODT XML offset overflow".into()))
}

fn bounded<T>(list: &[T], what: &str) -> Result<()> {
    if list.len() >= MAX_ENTRIES {
        return Err(Error::InvalidFormat(format!(
            "ODT document holds more than {MAX_ENTRIES} {what}"
        )));
    }
    Ok(())
}

impl Parts {
    /// Index the parts of the document `body` was built from.
    ///
    /// # Errors
    ///
    /// Returns an error when `content.xml` or `styles.xml` is not
    /// well-formed, or holds more entries than the index admits.
    pub fn build(body: Body<'_>) -> Result<Self> {
        let mut parts = Self {
            style_prefixes: Prefixes::of_root(body.styles),
            ..Self::default()
        };
        if !body.styles.is_empty() {
            let prefixes = parts.style_prefixes.clone();
            scan_styles(body.styles, &prefixes, &mut parts)?;
        }
        scan_content(body.content, body.prefixes(), &mut parts)?;
        Ok(parts)
    }

    /// A borrowed view of this index over `body`, the body it was built
    /// from.
    #[must_use]
    pub const fn view<'a>(&'a self, body: Body<'a>) -> PartsView<'a> {
        PartsView { parts: self, body }
    }
}

/// Where a list style is declared, by scan.
fn scan_styles(xml: &str, prefixes: &Prefixes, parts: &mut Parts) -> Result<()> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    // Open elements: (namespace, local name) of every open element.
    let mut stack: Vec<Option<(Ns, Vec<u8>)>> = Vec::new();
    let mut list: Option<ListData> = None;
    let mut level: Option<LevelData> = None;
    let mut master: Option<u32> = None;
    let mut region: Option<RegionData> = None;
    loop {
        let start = offset(reader.buffer_position())?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("ODT styles.xml: {error}")))?;
        let end = offset(reader.buffer_position())?;
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let empty = matches!(xml.as_bytes().get(end as usize - 2), Some(b'/'));
                let name = tag.name();
                let resolved = prefixes
                    .resolve(name.as_ref())
                    .map(|(ns, local)| (ns, local.to_vec()));
                let depth = stack.len();
                let parent = stack.last().cloned().flatten();
                let span = Span::new(start, end, empty);
                match (&resolved, &parent) {
                    (Some((Ns::Text, local)), Some((Ns::Office, container)))
                        if local == b"list-style"
                            && container == b"styles"
                            && depth == 2
                            && list.is_none() =>
                    {
                        list = Some(ListData {
                            part: StylePart::Common,
                            span,
                            levels: Vec::new(),
                        });
                    },
                    (Some((Ns::Text, local)), Some((Ns::Text, container)))
                        if container == b"list-style" && list.is_some() && level.is_none() =>
                    {
                        if let Some(kind) = ListLevelKind::of(local) {
                            level = Some(level_data(&xml[span.tag()], prefixes, span, kind));
                        }
                    },
                    (Some((Ns::Style, local)), _) if level.is_some() => {
                        record_level_property(level.as_mut(), local, start, end);
                    },
                    (Some((Ns::Style, local)), Some((Ns::Office, container)))
                        if local == b"master-page"
                            && container == b"master-styles"
                            && master.is_none() =>
                    {
                        bounded(&parts.masters, "master pages")?;
                        parts.masters.push(MasterData { span });
                        master = Some((parts.masters.len() - 1) as u32);
                    },
                    (Some((ns, local)), Some((Ns::Style, container)))
                        if container == b"master-page"
                            && region.is_none()
                            && matches!(ns, Ns::Style | Ns::Loext) =>
                    {
                        if let (Some(index), Some(kind)) = (master, RegionKind::of(local)) {
                            region = Some(RegionData {
                                master: index,
                                kind,
                                span,
                            });
                        }
                    },
                    _ => {},
                }
                if empty {
                    finish_open(
                        &resolved,
                        depth,
                        &mut list,
                        &mut level,
                        &mut master,
                        &mut region,
                        start,
                        end,
                        parts,
                    )?;
                } else {
                    if stack.len() >= MAX_DEPTH {
                        return Err(Error::InvalidFormat("ODT styles.xml nests too deep".into()));
                    }
                    stack.push(resolved);
                }
            },
            Event::End(_) => {
                let closed = stack.pop().flatten();
                let depth = stack.len();
                finish_open(
                    &closed,
                    depth,
                    &mut list,
                    &mut level,
                    &mut master,
                    &mut region,
                    start,
                    end,
                    parts,
                )?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

/// Close whichever open entry the element `closed` at `depth` ends.
#[allow(
    clippy::too_many_arguments,
    reason = "the scan's open entries, closed in one place for empty and end tags alike"
)]
fn finish_open(
    closed: &Option<(Ns, Vec<u8>)>,
    depth: usize,
    list: &mut Option<ListData>,
    level: &mut Option<LevelData>,
    master: &mut Option<u32>,
    region: &mut Option<RegionData>,
    start: u32,
    end: u32,
    parts: &mut Parts,
) -> Result<()> {
    let Some((ns, local)) = closed else {
        return Ok(());
    };
    if depth == 3
        && region
            .as_ref()
            .is_some_and(|open| RegionKind::of(local) == Some(open.kind))
        && let Some(mut done) = region.take()
    {
        if done.span.end == u32::MAX {
            done.span.close(start, end);
        }
        bounded(&parts.regions, "header and footer regions")?;
        parts.regions.push(done);
        return Ok(());
    }
    if *ns == Ns::Style && local == b"master-page" && depth == 2 {
        if let Some(index) = master.take() {
            let span = &mut parts.masters[index as usize].span;
            if span.end == u32::MAX {
                span.close(start, end);
            }
        }
        return Ok(());
    }
    if *ns == Ns::Text && ListLevelKind::of(local).is_some() {
        if let (Some(mut done), Some(open)) = (level.take(), list.as_mut()) {
            if done.span.end == u32::MAX {
                done.span.close(start, end);
            }
            bounded(&open.levels, "list levels")?;
            open.levels.push(done);
        }
        return Ok(());
    }
    if *ns == Ns::Text && local == b"list-style" {
        if let Some(mut done) = list.take() {
            if done.span.end == u32::MAX {
                done.span.close(start, end);
            }
            bounded(&parts.lists, "list styles")?;
            parts.lists.push(done);
        }
    }
    Ok(())
}

fn level_data(tag: &str, prefixes: &Prefixes, span: Span, kind: ListLevelKind) -> LevelData {
    let level = xml::attributes(tag)
        .find_map(|(name, value)| {
            prefixes
                .is(name.as_bytes(), Ns::Text, "level")
                .then(|| value.parse::<u16>().ok())
                .flatten()
        })
        .unwrap_or(1);
    LevelData {
        span,
        kind,
        level,
        properties: None,
        alignment: None,
        text: None,
    }
}

fn record_level_property(level: Option<&mut LevelData>, local: &[u8], start: u32, end: u32) {
    let Some(level) = level else {
        return;
    };
    match local {
        b"list-level-properties" => level.properties = Some(start..end),
        b"list-level-label-alignment" => level.alignment = Some(start..end),
        b"text-properties" => level.text = Some(start..end),
        _ => {},
    }
}

/// Shapes the drawing scan lists, by `draw:` local name.
const SHAPES: [&[u8]; 13] = [
    b"rect",
    b"ellipse",
    b"circle",
    b"line",
    b"polyline",
    b"polygon",
    b"regular-polygon",
    b"path",
    b"custom-shape",
    b"connector",
    b"caption",
    b"measure",
    b"g",
];

/// What an open element of `content.xml` stands for in the scan.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Open {
    Plain,
    /// An `office:automatic-styles` container.
    AutomaticStyles,
    /// A list style of the automatic styles.
    List,
    /// A level of that list style.
    Level,
    /// A table of contents (index into `tocs`).
    Toc(u32),
    /// `text:index-title-template` of the open table of contents.
    TocTitle,
    /// `text:index-body` of the open table of contents.
    TocBody,
    /// A listed shape (index into `shapes`).
    Shape(u32),
    /// A shape nested in a listed shape.
    NestedShape,
    /// A `draw:frame` (start offsets).
    Frame(Span),
    /// A `draw:object` directly in a frame (index into `objects`).
    Object(u32),
    /// An inline `math:math` of an object.
    Math(u32),
    /// A `table:table` (start offset).
    Table(u32),
    /// A `table:table-columns`, `table-header-columns` or
    /// `table-column-group` of the innermost table.
    ColumnGroup,
    /// A column (index into `columns`).
    Column(u32),
    /// A subtree the scan skips (tracked changes).
    Skipped,
}

#[allow(
    clippy::too_many_lines,
    reason = "one pass over content.xml for every part kind, kept together so the open-element bookkeeping reads in one place"
)]
fn scan_content(xml: &str, prefixes: &Prefixes, parts: &mut Parts) -> Result<()> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut stack: Vec<Open> = Vec::new();
    let mut list: Option<ListData> = None;
    let mut level: Option<LevelData> = None;
    let mut in_body = false;
    loop {
        let start = offset(reader.buffer_position())?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("ODT content.xml: {error}")))?;
        let end = offset(reader.buffer_position())?;
        let (tag, empty) = match event {
            Event::Start(tag) => (tag, false),
            Event::Empty(tag) => (tag, true),
            Event::End(_) => {
                let open = stack.pop().unwrap_or(Open::Plain);
                close_content(open, start, end, &mut list, &mut level, parts)?;
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        let span = Span::new(start, end, empty);
        let name = tag.name();
        let name = name.as_ref();
        let resolved = prefixes.resolve(name);
        let parent = stack.last().copied().unwrap_or(Open::Plain);
        let skipping = stack.contains(&Open::Skipped);
        let open = if skipping {
            Open::Skipped
        } else {
            match resolved {
                Some((Ns::Office, b"automatic-styles")) if stack.len() == 1 => {
                    Open::AutomaticStyles
                },
                Some((Ns::Office, b"body")) => {
                    in_body = true;
                    Open::Plain
                },
                Some((Ns::Text, b"list-style")) if parent == Open::AutomaticStyles => {
                    list = Some(ListData {
                        part: StylePart::Automatic,
                        span,
                        levels: Vec::new(),
                    });
                    Open::List
                },
                Some((Ns::Text, local)) if parent == Open::List => match ListLevelKind::of(local) {
                    Some(kind) if level.is_none() => {
                        level = Some(level_data(&xml[span.tag()], prefixes, span, kind));
                        Open::Level
                    },
                    _ => Open::Plain,
                },
                Some((Ns::Style, local)) if parent == Open::Level => {
                    record_level_property(level.as_mut(), local, start, end);
                    Open::Plain
                },
                _ if !in_body => Open::Plain,
                Some((Ns::Text, b"tracked-changes")) => Open::Skipped,
                Some((Ns::Text, b"table-of-content")) => {
                    bounded(&parts.tocs, "tables of contents")?;
                    parts.tocs.push(TocData {
                        span,
                        source: None,
                        title: None,
                        index_body: None,
                        links: false,
                        page_numbers: false,
                    });
                    Open::Toc((parts.tocs.len() - 1) as u32)
                },
                Some((Ns::Text, local)) if innermost_toc(&stack).is_some() => {
                    let index = innermost_toc(&stack).unwrap_or(0) as usize;
                    let toc = &mut parts.tocs[index];
                    match local {
                        b"table-of-content-source" => {
                            toc.source = Some(start..end);
                            Open::Plain
                        },
                        b"index-title-template" => {
                            toc.title = Some(span);
                            Open::TocTitle
                        },
                        b"index-body" => {
                            toc.index_body = Some(span);
                            Open::TocBody
                        },
                        b"index-entry-link-start" if !stack.contains(&Open::TocBody) => {
                            toc.links = true;
                            Open::Plain
                        },
                        b"index-entry-page-number" if !stack.contains(&Open::TocBody) => {
                            toc.page_numbers = true;
                            Open::Plain
                        },
                        _ => Open::Plain,
                    }
                },
                Some((Ns::Draw, local)) if SHAPES.contains(&local) => {
                    if stack
                        .iter()
                        .any(|open| matches!(open, Open::Shape(_) | Open::NestedShape))
                    {
                        Open::NestedShape
                    } else {
                        bounded(&parts.shapes, "shapes")?;
                        parts.shapes.push(ShapeData {
                            span,
                            geometry: None,
                        });
                        Open::Shape((parts.shapes.len() - 1) as u32)
                    }
                },
                Some((Ns::Draw, b"enhanced-geometry")) => {
                    if let Open::Shape(index) = parent {
                        parts.shapes[index as usize].geometry = Some(start..end);
                    }
                    Open::Plain
                },
                Some((Ns::Draw, b"frame")) => Open::Frame(span),
                Some((Ns::Draw, local @ (b"object" | b"object-ole"))) => match parent {
                    Open::Frame(frame) => {
                        bounded(&parts.objects, "embedded objects")?;
                        parts.objects.push(ObjectData {
                            frame,
                            object: span,
                            ole: local == b"object-ole",
                            math: None,
                        });
                        Open::Object((parts.objects.len() - 1) as u32)
                    },
                    _ => Open::Plain,
                },
                Some((Ns::Table, b"table")) => Open::Table(start),
                Some((
                    Ns::Table,
                    b"table-columns" | b"table-header-columns" | b"table-column-group",
                )) if matches!(parent, Open::Table(_) | Open::ColumnGroup) => Open::ColumnGroup,
                Some((Ns::Table, b"table-column"))
                    if matches!(parent, Open::Table(_) | Open::ColumnGroup) =>
                {
                    let table = innermost_table(&stack).unwrap_or(0);
                    let repeat = xml::attributes(&xml[span.tag()])
                        .find_map(|(name, value)| {
                            prefixes
                                .is(name.as_bytes(), Ns::Table, "number-columns-repeated")
                                .then(|| value.parse::<u32>().ok())
                                .flatten()
                        })
                        .unwrap_or(1)
                        .clamp(1, 16_384);
                    bounded(&parts.columns, "table columns")?;
                    parts.columns.push(ColumnData {
                        table,
                        span,
                        repeat,
                    });
                    Open::Column((parts.columns.len() - 1) as u32)
                },
                _ if xml::tag_name_local(name) == b"math" => match parent {
                    Open::Object(index) => {
                        parts.objects[index as usize].math = Some(span);
                        Open::Math(index)
                    },
                    _ => Open::Plain,
                },
                _ => Open::Plain,
            }
        };
        if empty {
            close_content(open, start, end, &mut list, &mut level, parts)?;
        } else {
            if stack.len() >= MAX_DEPTH {
                return Err(Error::InvalidFormat(
                    "ODT content.xml nests too deep".into(),
                ));
            }
            stack.push(open);
        }
    }
    Ok(())
}

fn innermost_toc(stack: &[Open]) -> Option<u32> {
    stack.iter().rev().find_map(|open| match open {
        Open::Toc(index) => Some(*index),
        _ => None,
    })
}

fn innermost_table(stack: &[Open]) -> Option<u32> {
    stack.iter().rev().find_map(|open| match open {
        Open::Table(start) => Some(*start),
        _ => None,
    })
}

fn close_content(
    open: Open,
    start: u32,
    end: u32,
    list: &mut Option<ListData>,
    level: &mut Option<LevelData>,
    parts: &mut Parts,
) -> Result<()> {
    let close = |span: &mut Span| {
        if span.end == u32::MAX {
            span.close(start, end);
        }
    };
    match open {
        Open::List => {
            if let Some(mut done) = list.take() {
                close(&mut done.span);
                bounded(&parts.lists, "list styles")?;
                parts.lists.push(done);
            }
        },
        Open::Level => {
            if let (Some(mut done), Some(open)) = (level.take(), list.as_mut()) {
                close(&mut done.span);
                bounded(&open.levels, "list levels")?;
                open.levels.push(done);
            }
        },
        Open::Toc(index) => close(&mut parts.tocs[index as usize].span),
        Open::TocTitle => {
            if let Some(toc) = parts.tocs.last_mut()
                && let Some(title) = toc.title.as_mut()
            {
                close(title);
            }
        },
        Open::TocBody => {
            if let Some(toc) = parts.tocs.last_mut()
                && let Some(body) = toc.index_body.as_mut()
            {
                close(body);
            }
        },
        Open::Shape(index) => close(&mut parts.shapes[index as usize].span),
        Open::Object(index) => {
            let object = &mut parts.objects[index as usize];
            close(&mut object.object);
        },
        Open::Math(index) => {
            if let Some(math) = parts.objects[index as usize].math.as_mut() {
                close(math);
            }
        },
        Open::Frame(_) => {
            // A frame closes after its objects: record its end on them.
            for object in parts.objects.iter_mut().rev() {
                if object.frame.end != u32::MAX {
                    break;
                }
                object.frame.close(start, end);
            }
        },
        Open::Column(index) => close(&mut parts.columns[index as usize].span),
        Open::Plain
        | Open::AutomaticStyles
        | Open::NestedShape
        | Open::Table(_)
        | Open::ColumnGroup
        | Open::Skipped => {},
    }
    Ok(())
}

/// The decoded value of `ns:local` in `tag`.
fn attribute<'a>(tag: &'a str, prefixes: &Prefixes, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
    xml::attributes(tag).find_map(|(name, value)| {
        prefixes
            .is(name.as_bytes(), ns, local)
            .then(|| xml::unescape(value))
    })
}

impl<'a> PartsView<'a> {
    fn styles_xml(self) -> &'a str {
        self.body.styles
    }

    fn content_xml(self) -> &'a str {
        self.body.content
    }

    /// The list styles: the common ones of `styles.xml`, then the
    /// automatic ones of `content.xml`, each in declaration order.
    pub fn list_styles(self) -> impl ExactSizeIterator<Item = ListStyle<'a>> + 'a {
        (0..self.parts.lists.len()).map(move |index| ListStyle { view: self, index })
    }

    /// The list style named `name`, automatic styles first (the lookup
    /// order of a list in `content.xml`).
    #[must_use]
    pub fn list_style(self, name: &str) -> Option<ListStyle<'a>> {
        let mut found = None;
        for style in self.list_styles() {
            if style.name().as_deref() == Some(name) {
                if style.part() == StylePart::Automatic {
                    return Some(style);
                }
                found.get_or_insert(style);
            }
        }
        found
    }

    /// The header and footer regions of every master page, in declaration
    /// order.
    pub fn regions(self) -> impl ExactSizeIterator<Item = Region<'a>> + 'a {
        (0..self.parts.regions.len()).map(move |index| Region { view: self, index })
    }

    /// The tables of contents of the body, in document order.
    pub fn tables_of_contents(self) -> impl ExactSizeIterator<Item = TableOfContents<'a>> + 'a {
        (0..self.parts.tocs.len()).map(move |index| TableOfContents { view: self, index })
    }

    /// The drawing shapes of the body that no other shape holds, in
    /// document order.
    pub fn shapes(self) -> impl ExactSizeIterator<Item = Shape<'a>> + 'a {
        (0..self.parts.shapes.len()).map(move |index| Shape { view: self, index })
    }

    /// The embedded objects (`draw:object`, `draw:object-ole`) directly in
    /// a frame of the body, in document order.
    pub fn objects(self) -> impl ExactSizeIterator<Item = EmbeddedObject<'a>> + 'a {
        (0..self.parts.objects.len()).map(move |index| EmbeddedObject { view: self, index })
    }

    /// The columns a body table declares (`table:table-column`, a
    /// repeated declaration once per column it stands for); empty for a
    /// node that is not a table.
    pub fn columns(self, table: Node<'a>) -> impl Iterator<Item = Column<'a>> + 'a {
        let start = if table.kind() == Kind::Table {
            u32::try_from(table.range().start).ok()
        } else {
            None
        };
        self.parts
            .columns
            .iter()
            .enumerate()
            .filter(move |(_, column)| Some(column.table) == start)
            .flat_map(move |(index, column)| {
                (0..column.repeat).map(move |_| Column { view: self, index })
            })
    }
}

/// The kind of one list level declaration.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ListLevelKind {
    /// `text:list-level-style-number`.
    Number,
    /// `text:list-level-style-bullet`.
    Bullet,
    /// `text:list-level-style-image`.
    Image,
}

impl ListLevelKind {
    fn of(local: &[u8]) -> Option<Self> {
        Some(match local {
            b"list-level-style-number" => Self::Number,
            b"list-level-style-bullet" => Self::Bullet,
            b"list-level-style-image" => Self::Image,
            _ => return None,
        })
    }
}

/// One `text:list-style` declaration.
#[derive(Clone, Copy)]
pub struct ListStyle<'a> {
    view: PartsView<'a>,
    index: usize,
}

impl std::fmt::Debug for ListStyle<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ListStyle")
            .field("index", &self.index)
            .finish_non_exhaustive()
    }
}

impl<'a> ListStyle<'a> {
    fn data(self) -> &'a ListData {
        &self.view.parts.lists[self.index]
    }

    fn source(self) -> (&'a str, &'a Prefixes) {
        match self.data().part {
            StylePart::Common => (self.view.styles_xml(), &self.view.parts.style_prefixes),
            StylePart::Automatic => (self.view.content_xml(), self.view.body.prefixes()),
        }
    }

    fn attribute(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        let (xml, prefixes) = self.source();
        attribute(&xml[self.data().span.tag()], prefixes, ns, local)
    }

    /// The position among [`PartsView::list_styles`].
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }

    /// Where the style is declared.
    #[must_use]
    pub fn part(self) -> StylePart {
        self.data().part
    }

    /// `style:name`.
    #[must_use]
    pub fn name(self) -> Option<Cow<'a, str>> {
        self.attribute(Ns::Style, "name")
    }

    /// `style:display-name`.
    #[must_use]
    pub fn display_name(self) -> Option<Cow<'a, str>> {
        self.attribute(Ns::Style, "display-name")
    }

    /// `text:consecutive-numbering`.
    #[must_use]
    pub fn consecutive_numbering(self) -> Option<bool> {
        self.attribute(Ns::Text, "consecutive-numbering")
            .map(|value| value == "true")
    }

    /// The level declarations, in declaration order.
    pub fn levels(self) -> impl ExactSizeIterator<Item = ListLevel<'a>> + 'a {
        (0..self.data().levels.len()).map(move |level| ListLevel { style: self, level })
    }

    /// The declaration of level `level` (1-based, as `text:level`).
    #[must_use]
    pub fn level(self, level: u16) -> Option<ListLevel<'a>> {
        self.levels().find(|entry| entry.level() == level)
    }
}

/// One `text:list-level-style-*` declaration.
#[derive(Clone, Copy)]
pub struct ListLevel<'a> {
    style: ListStyle<'a>,
    level: usize,
}

impl std::fmt::Debug for ListLevel<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ListLevel")
            .field("level", &self.level())
            .finish_non_exhaustive()
    }
}

impl<'a> ListLevel<'a> {
    fn data(self) -> &'a LevelData {
        &self.style.data().levels[self.level]
    }

    fn tag_attribute(
        self,
        range: Option<&Range<u32>>,
        ns: Ns,
        local: &str,
    ) -> Option<Cow<'a, str>> {
        let (xml, prefixes) = self.style.source();
        let range = range?;
        attribute(
            &xml[range.start as usize..range.end as usize],
            prefixes,
            ns,
            local,
        )
    }

    fn own(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        let tag = self.data().span.start..self.data().span.open_end;
        self.tag_attribute(Some(&tag), ns, local)
    }

    /// `text:level`, 1-based.
    #[must_use]
    pub fn level(self) -> u16 {
        self.data().level
    }

    /// Whether the level numbers, shows a bullet or shows an image.
    #[must_use]
    pub fn kind(self) -> ListLevelKind {
        self.data().kind
    }

    /// `text:start-value` of a numbered level.
    #[must_use]
    pub fn start_value(self) -> Option<u32> {
        self.own(Ns::Text, "start-value")?.parse().ok()
    }

    /// `style:num-format` (`1`, `a`, `A`, `i`, `I`, or empty for none).
    #[must_use]
    pub fn num_format(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Style, "num-format")
    }

    /// `style:num-prefix`.
    #[must_use]
    pub fn prefix(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Style, "num-prefix")
    }

    /// `style:num-suffix`.
    #[must_use]
    pub fn suffix(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Style, "num-suffix")
    }

    /// `text:display-levels`: how many levels the label shows.
    #[must_use]
    pub fn display_levels(self) -> Option<u16> {
        self.own(Ns::Text, "display-levels")?.parse().ok()
    }

    /// `text:bullet-char` of a bullet level.
    #[must_use]
    pub fn bullet_char(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Text, "bullet-char")
    }

    /// `text:style-name`: the character style of the label.
    #[must_use]
    pub fn style_name(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Text, "style-name")
    }

    /// `text:label-followed-by` of the label alignment (`listtab`,
    /// `space`, `nothing`).
    #[must_use]
    pub fn label_followed_by(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(
            self.data().alignment.as_ref(),
            Ns::Text,
            "label-followed-by",
        )
    }

    /// `fo:text-align` of the level properties.
    #[must_use]
    pub fn text_align(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(self.data().properties.as_ref(), Ns::Fo, "text-align")
    }

    /// `fo:margin-left` of the label alignment.
    #[must_use]
    pub fn margin_left(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(self.data().alignment.as_ref(), Ns::Fo, "margin-left")
    }

    /// `fo:text-indent` of the label alignment.
    #[must_use]
    pub fn text_indent(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(self.data().alignment.as_ref(), Ns::Fo, "text-indent")
    }

    /// `text:space-before` of the level properties (the ODF 1.1 position
    /// mode).
    #[must_use]
    pub fn space_before(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(self.data().properties.as_ref(), Ns::Text, "space-before")
    }

    /// `text:min-label-width` of the level properties (the ODF 1.1
    /// position mode).
    #[must_use]
    pub fn min_label_width(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(self.data().properties.as_ref(), Ns::Text, "min-label-width")
    }

    /// The label's font: `style:font-name`, else `fo:font-family`.
    #[must_use]
    pub fn font(self) -> Option<Cow<'a, str>> {
        let text = self.data().text.as_ref();
        self.tag_attribute(text, Ns::Style, "font-name")
            .or_else(|| self.tag_attribute(text, Ns::Fo, "font-family"))
            .or_else(|| self.tag_attribute(self.data().properties.as_ref(), Ns::Style, "font-name"))
    }

    /// `fo:font-size` of the label.
    #[must_use]
    pub fn font_size(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(self.data().text.as_ref(), Ns::Fo, "font-size")
    }

    /// `fo:color` of the label.
    #[must_use]
    pub fn color(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(self.data().text.as_ref(), Ns::Fo, "color")
    }

    /// Whether the label is bold (`fo:font-weight`).
    #[must_use]
    pub fn bold(self) -> Option<bool> {
        self.tag_attribute(self.data().text.as_ref(), Ns::Fo, "font-weight")
            .map(|weight| {
                weight == "bold" || weight.parse::<u16>().is_ok_and(|weight| weight >= 600)
            })
    }

    /// Whether the label is italic (`fo:font-style`).
    #[must_use]
    pub fn italic(self) -> Option<bool> {
        self.tag_attribute(self.data().text.as_ref(), Ns::Fo, "font-style")
            .map(|style| matches!(&*style, "italic" | "oblique"))
    }

    /// `xlink:href` of an image level.
    #[must_use]
    pub fn image_href(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Xlink, "href")
    }
}

/// The kind of one master-page region.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RegionKind {
    /// `style:header`: every page, or odd pages when a left header exists.
    Header,
    /// `style:header-left`: even pages.
    HeaderLeft,
    /// `style:header-first`: the first page.
    HeaderFirst,
    /// `style:footer`.
    Footer,
    /// `style:footer-left`.
    FooterLeft,
    /// `style:footer-first`.
    FooterFirst,
}

impl RegionKind {
    fn of(local: &[u8]) -> Option<Self> {
        Some(match local {
            b"header" => Self::Header,
            b"header-left" => Self::HeaderLeft,
            b"header-first" => Self::HeaderFirst,
            b"footer" => Self::Footer,
            b"footer-left" => Self::FooterLeft,
            b"footer-first" => Self::FooterFirst,
            _ => return None,
        })
    }

    /// Whether the region is a header.
    #[must_use]
    pub const fn is_header(self) -> bool {
        matches!(self, Self::Header | Self::HeaderLeft | Self::HeaderFirst)
    }

    /// The region's element local name.
    #[must_use]
    pub const fn local_name(self) -> &'static str {
        match self {
            Self::Header => "header",
            Self::HeaderLeft => "header-left",
            Self::HeaderFirst => "header-first",
            Self::Footer => "footer",
            Self::FooterLeft => "footer-left",
            Self::FooterFirst => "footer-first",
        }
    }
}

/// One header or footer region of a master page (`styles.xml`).
#[derive(Clone, Copy)]
pub struct Region<'a> {
    view: PartsView<'a>,
    index: usize,
}

impl std::fmt::Debug for Region<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Region")
            .field("kind", &self.kind())
            .finish_non_exhaustive()
    }
}

impl<'a> Region<'a> {
    fn data(self) -> &'a RegionData {
        &self.view.parts.regions[self.index]
    }

    /// The position among [`PartsView::regions`].
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }

    /// Header or footer, and which pages.
    #[must_use]
    pub fn kind(self) -> RegionKind {
        self.data().kind
    }

    /// The position of the master page among the master pages.
    #[must_use]
    pub fn master_index(self) -> usize {
        self.data().master as usize
    }

    /// `style:name` of the master page.
    #[must_use]
    pub fn master_name(self) -> Option<Cow<'a, str>> {
        let master = &self.view.parts.masters[self.data().master as usize];
        attribute(
            &self.view.styles_xml()[master.span.tag()],
            &self.view.parts.style_prefixes,
            Ns::Style,
            "name",
        )
    }

    /// Whether the region is shown (`style:display`, default `true`).
    #[must_use]
    pub fn displayed(self) -> bool {
        attribute(
            &self.view.styles_xml()[self.data().span.tag()],
            &self.view.parts.style_prefixes,
            Ns::Style,
            "display",
        )
        .is_none_or(|value| value != "false")
    }

    /// The region's visible text: paragraphs separated by line feeds.
    ///
    /// # Errors
    ///
    /// Returns an error when the region's markup cannot be read.
    pub fn text(self) -> Result<String> {
        let content = &self.view.styles_xml()[self.data().span.content()];
        super::text::text_of(&self.view.parts.style_prefixes, Kind::Body, content)
    }

    /// The byte range of the region element in `styles.xml`.
    #[must_use]
    pub fn range(self) -> Range<usize> {
        self.data().span.range()
    }

    /// The byte range of the region's content in `styles.xml`.
    #[must_use]
    pub fn content_range(self) -> Range<usize> {
        self.data().span.content()
    }

    /// The region's source markup.
    #[must_use]
    pub fn xml(self) -> &'a str {
        &self.view.styles_xml()[self.range()]
    }
}

/// One `text:table-of-content` of the body.
#[derive(Clone, Copy)]
pub struct TableOfContents<'a> {
    view: PartsView<'a>,
    index: usize,
}

impl std::fmt::Debug for TableOfContents<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("TableOfContents")
            .field("index", &self.index)
            .finish_non_exhaustive()
    }
}

impl<'a> TableOfContents<'a> {
    fn data(self) -> &'a TocData {
        &self.view.parts.tocs[self.index]
    }

    fn tag_attribute(
        self,
        range: Option<Range<usize>>,
        ns: Ns,
        local: &str,
    ) -> Option<Cow<'a, str>> {
        attribute(
            &self.view.content_xml()[range?],
            self.view.body.prefixes(),
            ns,
            local,
        )
    }

    /// `text:name`.
    #[must_use]
    pub fn name(self) -> Option<Cow<'a, str>> {
        self.tag_attribute(Some(self.data().span.tag()), Ns::Text, "name")
    }

    /// `text:protected`.
    #[must_use]
    pub fn protected(self) -> Option<bool> {
        self.tag_attribute(Some(self.data().span.tag()), Ns::Text, "protected")
            .map(|value| value == "true")
    }

    /// `text:outline-level` of the source: the deepest heading level the
    /// index lists.
    #[must_use]
    pub fn outline_level(self) -> Option<u16> {
        let source = self
            .data()
            .source
            .as_ref()
            .map(|range| range.start as usize..range.end as usize);
        self.tag_attribute(source, Ns::Text, "outline-level")?
            .parse()
            .ok()
    }

    /// Whether the entry templates link each entry to its heading
    /// (`text:index-entry-link-start`).
    #[must_use]
    pub fn links(self) -> bool {
        self.data().links
    }

    /// Whether the entry templates show page numbers
    /// (`text:index-entry-page-number`).
    #[must_use]
    pub fn page_numbers(self) -> bool {
        self.data().page_numbers
    }

    /// The title template's text.
    #[must_use]
    pub fn title(self) -> Option<Cow<'a, str>> {
        let title = self.data().title?;
        let content = &self.view.content_xml()[title.content()];
        Some(xml::unescape(content.trim()))
    }

    /// The generated index body's visible text.
    ///
    /// # Errors
    ///
    /// Returns an error when the index body's markup cannot be read.
    pub fn text(self) -> Result<String> {
        let Some(body) = self.data().index_body else {
            return Ok(String::new());
        };
        let content = &self.view.content_xml()[body.content()];
        super::text::text_of(self.view.body.prefixes(), Kind::Body, content)
    }

    /// The byte range of the element in `content.xml`.
    #[must_use]
    pub fn range(self) -> Range<usize> {
        self.data().span.range()
    }
}

/// One drawing shape of the body (`draw:rect`, `draw:custom-shape`, ...).
#[derive(Clone, Copy)]
pub struct Shape<'a> {
    view: PartsView<'a>,
    index: usize,
}

impl std::fmt::Debug for Shape<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Shape")
            .field("element", &self.element())
            .finish_non_exhaustive()
    }
}

impl<'a> Shape<'a> {
    fn data(self) -> &'a ShapeData {
        &self.view.parts.shapes[self.index]
    }

    fn tag(self) -> &'a str {
        &self.view.content_xml()[self.data().span.tag()]
    }

    fn own(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        attribute(self.tag(), self.view.body.prefixes(), ns, local)
    }

    /// The element's local name: `rect`, `ellipse`, `custom-shape`, `g`, ...
    #[must_use]
    pub fn element(self) -> &'a str {
        let name = xml::tag_name(self.tag());
        name.rsplit(':').next().unwrap_or(name)
    }

    /// The geometry: a custom shape's `draw:enhanced-geometry/@draw:type`,
    /// else the element's local name.
    #[must_use]
    pub fn geometry(self) -> Cow<'a, str> {
        self.data()
            .geometry
            .as_ref()
            .and_then(|range| {
                attribute(
                    &self.view.content_xml()[range.start as usize..range.end as usize],
                    self.view.body.prefixes(),
                    Ns::Draw,
                    "type",
                )
            })
            .unwrap_or(Cow::Borrowed(self.element()))
    }

    /// `draw:name`.
    #[must_use]
    pub fn name(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Draw, "name")
    }

    /// `svg:width`.
    #[must_use]
    pub fn width(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Svg, "width")
    }

    /// `svg:height`.
    #[must_use]
    pub fn height(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Svg, "height")
    }

    /// `svg:x`.
    #[must_use]
    pub fn x(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Svg, "x")
    }

    /// `svg:y`.
    #[must_use]
    pub fn y(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Svg, "y")
    }

    /// `text:anchor-type`.
    #[must_use]
    pub fn anchor_type(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Text, "anchor-type")
    }

    /// `draw:style-name`.
    #[must_use]
    pub fn style_name(self) -> Option<Cow<'a, str>> {
        self.own(Ns::Draw, "style-name")
    }

    /// One attribute of the `style:graphic-properties` of the shape's
    /// automatic graphic style.
    fn graphic(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        let style = self.view.body.automatic_style(&self.style_name()?)?;
        let source = style.xml();
        let prefixes = self.view.body.prefixes();
        let mut cursor = 0;
        while let Some(found) = source[cursor..].find('<') {
            let start = cursor + found;
            let end = xml::tag_end(source.as_bytes(), start)?;
            let tag = &source[start..end];
            if prefixes.is(
                xml::tag_name(tag).as_bytes(),
                Ns::Style,
                "graphic-properties",
            ) {
                return attribute(tag, prefixes, ns, local);
            }
            cursor = end;
        }
        None
    }

    /// `draw:fill` of the graphic style (`none`, `solid`, ...).
    #[must_use]
    pub fn fill(self) -> Option<Cow<'a, str>> {
        self.graphic(Ns::Draw, "fill")
    }

    /// `draw:fill-color` of the graphic style.
    #[must_use]
    pub fn fill_color(self) -> Option<Cow<'a, str>> {
        self.graphic(Ns::Draw, "fill-color")
    }

    /// `draw:stroke` of the graphic style (`none`, `solid`, `dash`).
    #[must_use]
    pub fn stroke(self) -> Option<Cow<'a, str>> {
        self.graphic(Ns::Draw, "stroke")
    }

    /// `svg:stroke-color` of the graphic style.
    #[must_use]
    pub fn stroke_color(self) -> Option<Cow<'a, str>> {
        self.graphic(Ns::Svg, "stroke-color")
    }

    /// `svg:stroke-width` of the graphic style.
    #[must_use]
    pub fn stroke_width(self) -> Option<Cow<'a, str>> {
        self.graphic(Ns::Svg, "stroke-width")
    }

    /// `style:wrap` of the graphic style.
    #[must_use]
    pub fn wrap(self) -> Option<Cow<'a, str>> {
        self.graphic(Ns::Style, "wrap")
    }

    /// The text the shape holds: its paragraphs, separated by line feeds.
    ///
    /// # Errors
    ///
    /// Returns an error when the shape's markup cannot be read.
    pub fn text(self) -> Result<String> {
        let content = &self.view.content_xml()[self.data().span.content()];
        super::text::text_of(self.view.body.prefixes(), Kind::Body, content)
    }

    /// The byte range of the element in `content.xml`.
    #[must_use]
    pub fn range(self) -> Range<usize> {
        self.data().span.range()
    }
}

/// One embedded object directly in a frame of the body.
#[derive(Clone, Copy)]
pub struct EmbeddedObject<'a> {
    view: PartsView<'a>,
    index: usize,
}

impl std::fmt::Debug for EmbeddedObject<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("EmbeddedObject")
            .field("index", &self.index)
            .finish_non_exhaustive()
    }
}

impl<'a> EmbeddedObject<'a> {
    fn data(self) -> &'a ObjectData {
        &self.view.parts.objects[self.index]
    }

    fn frame(self, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
        attribute(
            &self.view.content_xml()[self.data().frame.tag()],
            self.view.body.prefixes(),
            ns,
            local,
        )
    }

    /// Whether the object is `draw:object-ole`.
    #[must_use]
    pub fn is_ole(self) -> bool {
        self.data().ole
    }

    /// `xlink:href` of the object: the package subdocument or file.
    #[must_use]
    pub fn href(self) -> Option<Cow<'a, str>> {
        attribute(
            &self.view.content_xml()[self.data().object.tag()],
            self.view.body.prefixes(),
            Ns::Xlink,
            "href",
        )
    }

    /// The inline MathML payload (`math:math`), when the object carries one.
    #[must_use]
    pub fn inline_math(self) -> Option<&'a str> {
        self.data()
            .math
            .map(|math| &self.view.content_xml()[math.range()])
    }

    /// `draw:name` of the frame.
    #[must_use]
    pub fn name(self) -> Option<Cow<'a, str>> {
        self.frame(Ns::Draw, "name")
    }

    /// `svg:width` of the frame.
    #[must_use]
    pub fn width(self) -> Option<Cow<'a, str>> {
        self.frame(Ns::Svg, "width")
    }

    /// `svg:height` of the frame.
    #[must_use]
    pub fn height(self) -> Option<Cow<'a, str>> {
        self.frame(Ns::Svg, "height")
    }

    /// `text:anchor-type` of the frame.
    #[must_use]
    pub fn anchor_type(self) -> Option<Cow<'a, str>> {
        self.frame(Ns::Text, "anchor-type")
    }

    /// The byte range of the frame in `content.xml`.
    #[must_use]
    pub fn range(self) -> Range<usize> {
        self.data().frame.range()
    }
}

/// One column of a body table (`table:table-column`).
#[derive(Clone, Copy)]
pub struct Column<'a> {
    view: PartsView<'a>,
    index: usize,
}

impl std::fmt::Debug for Column<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Column")
            .field("index", &self.index)
            .finish_non_exhaustive()
    }
}

impl<'a> Column<'a> {
    fn data(self) -> &'a ColumnData {
        &self.view.parts.columns[self.index]
    }

    /// `table:style-name`.
    #[must_use]
    pub fn style_name(self) -> Option<Cow<'a, str>> {
        attribute(
            &self.view.content_xml()[self.data().span.tag()],
            self.view.body.prefixes(),
            Ns::Table,
            "style-name",
        )
    }

    /// How many columns the declaration stands for
    /// (`table:number-columns-repeated`).
    #[must_use]
    pub fn repeat(self) -> u32 {
        self.data().repeat
    }

    /// `style:column-width` of the column's style
    /// (`style:table-column-properties`).
    #[must_use]
    pub fn width(self) -> Option<Cow<'a, str>> {
        let name = self.style_name()?;
        let body = self.view.body;
        let style = body
            .automatic_style(&name)
            .or_else(|| body.common_style(&name))?;
        let (source, prefixes) = style.source();
        let declaration = &source[style.data().start as usize..style.data().end as usize];
        let mut cursor = 0;
        while let Some(found) = declaration[cursor..].find('<') {
            let start = cursor + found;
            let end = xml::tag_end(declaration.as_bytes(), start)?;
            let tag = &declaration[start..end];
            if prefixes.is(
                xml::tag_name(tag).as_bytes(),
                Ns::Style,
                "table-column-properties",
            ) {
                return attribute(tag, prefixes, Ns::Style, "column-width");
            }
            cursor = end;
        }
        None
    }

    /// The byte range of the declaration in `content.xml`.
    #[must_use]
    pub fn range(self) -> Range<usize> {
        self.data().span.range()
    }
}

/// One `style:tab-stop` of a paragraph's automatic style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TabStop<'a> {
    position: Cow<'a, str>,
    kind: Option<Cow<'a, str>>,
    leader_style: Option<Cow<'a, str>>,
    leader_text: Option<Cow<'a, str>>,
    character: Option<Cow<'a, str>>,
}

impl<'a> TabStop<'a> {
    /// `style:position`, as written (`2.5cm`).
    #[must_use]
    pub fn position(&self) -> &str {
        &self.position
    }

    /// `style:type` (`left` when absent, `center`, `right`, `char`).
    #[must_use]
    pub fn kind(&self) -> &str {
        self.kind.as_deref().unwrap_or("left")
    }

    /// `style:type` as written, `None` when absent.
    #[must_use]
    pub fn kind_attribute(&self) -> Option<&str> {
        self.kind.as_deref()
    }

    /// `style:leader-style` (`none`, `dotted`, `dash`, `solid`, ...).
    #[must_use]
    pub fn leader_style(&self) -> Option<&str> {
        self.leader_style.as_deref()
    }

    /// `style:leader-text`: the leader character.
    #[must_use]
    pub fn leader_text(&self) -> Option<&str> {
        self.leader_text.as_deref()
    }

    /// `style:char`: the alignment character of a `char` stop.
    #[must_use]
    pub fn character(&self) -> Option<&str> {
        self.character.as_deref()
    }
}

/// The tab stops a paragraph or heading's automatic style declares
/// (`style:paragraph-properties/style:tab-stops`), in declaration order;
/// empty for any other node.
#[must_use]
pub fn tab_stops(node: Node<'_>) -> Vec<TabStop<'_>> {
    if !matches!(node.kind(), Kind::Paragraph | Kind::Heading) {
        return Vec::new();
    }
    let Some(style) = node.automatic_style() else {
        return Vec::new();
    };
    let source = style.xml();
    let prefixes = node.body().prefixes();
    let mut stops = Vec::new();
    let mut cursor = 0;
    while let Some(found) = source[cursor..].find('<') {
        let start = cursor + found;
        let Some(end) = xml::tag_end(source.as_bytes(), start) else {
            break;
        };
        let tag = &source[start..end];
        if prefixes.is(xml::tag_name(tag).as_bytes(), Ns::Style, "tab-stop") {
            let read = |local: &str| attribute(tag, prefixes, Ns::Style, local);
            if let Some(position) = read("position") {
                stops.push(TabStop {
                    position,
                    kind: read("type"),
                    leader_style: read("leader-style"),
                    leader_text: read("leader-text"),
                    character: read("char"),
                });
            }
        }
        cursor = end;
    }
    stops
}

/// The formula text of a MathML payload: its StarMath annotation when it
/// carries one, else the text of its token elements.
#[must_use]
pub fn formula_text(mathml: &str) -> Option<String> {
    let mut reader = Reader::from_str(mathml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut annotation: Option<String> = None;
    let mut in_annotation = false;
    let mut skip = 0usize;
    let mut tokens = String::new();
    let mut saw_math = false;
    loop {
        match reader.read_event().ok()? {
            Event::Start(tag) => {
                let name = tag.name();
                let local = xml::tag_name_local(name.as_ref());
                if local == b"math" {
                    saw_math = true;
                }
                if skip > 0 {
                    skip += 1;
                } else if local == b"annotation" {
                    let raw = std::str::from_utf8(tag.as_ref()).unwrap_or_default();
                    let star = raw.contains("StarMath");
                    if star && annotation.is_none() {
                        in_annotation = true;
                        annotation = Some(String::new());
                    } else {
                        skip = 1;
                    }
                } else if local == b"annotation-xml" {
                    skip = 1;
                }
            },
            Event::End(_) => {
                if skip > 0 {
                    skip -= 1;
                } else {
                    in_annotation = false;
                }
            },
            Event::Text(text) => {
                let raw = std::str::from_utf8(text.as_ref()).ok()?;
                if in_annotation {
                    if let Some(out) = annotation.as_mut() {
                        out.push_str(raw);
                    }
                } else if skip == 0 {
                    tokens.push_str(raw.trim());
                }
            },
            Event::GeneralRef(reference) => {
                let entity = format!("&{};", std::str::from_utf8(reference.as_ref()).ok()?);
                let decoded = xml::unescape(&entity);
                if in_annotation {
                    if let Some(out) = annotation.as_mut() {
                        out.push_str(&decoded);
                    }
                } else if skip == 0 {
                    tokens.push_str(&decoded);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !saw_math {
        return None;
    }
    match annotation {
        Some(text) if !text.trim().is_empty() => Some(text.trim().to_owned()),
        _ => Some(tokens),
    }
}

impl Region<'_> {
    /// The splice that replaces the region's content with one paragraph
    /// holding `text` (record 0328): the first paragraph's start tag is
    /// kept; spaces, tabs and line feeds are encoded as `text:s`,
    /// `text:tab` and `text:line-break`. `None` when the region already
    /// shows `text` (an exact no-op).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the region holds content a text
    /// set cannot keep (a field, a frame, a table, a bookmark, ...), so a
    /// set never drops it silently.
    pub(crate) fn text_replacement(self, text: &str) -> Result<Option<(Range<usize>, String)>> {
        let prefixes = &self.view.parts.style_prefixes;
        let styles = self.view.styles_xml();
        let span = self.data().span;
        let content = &styles[span.content()];
        let mut reader = Reader::from_str(content);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = false;
        let mut first: Option<String> = None;
        loop {
            let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
            let event = reader
                .read_event()
                .map_err(|error| Error::InvalidFormat(format!("ODT styles.xml: {error}")))?;
            let end = usize::try_from(reader.buffer_position()).unwrap_or(content.len());
            let tag = match &event {
                Event::Start(tag) | Event::Empty(tag) => tag,
                Event::Eof => break,
                _ => continue,
            };
            let name = tag.name();
            let allowed = matches!(
                prefixes.resolve(name.as_ref()),
                Some((
                    Ns::Text,
                    b"p" | b"h" | b"span" | b"s" | b"tab" | b"line-break" | b"soft-page-break"
                ))
            );
            if !allowed {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odt.body.parts",
                    "an ODT header or footer region holding this content cannot have its text replaced",
                )));
            }
            if first.is_none() {
                let raw = &content[start..end];
                let paragraph = matches!(
                    prefixes.resolve(name.as_ref()),
                    Some((Ns::Text, b"p" | b"h"))
                );
                if paragraph {
                    first = Some(raw.to_owned());
                }
            }
        }
        if self.text()? == text {
            return Ok(None);
        }
        let open = first.unwrap_or_else(|| {
            let mut tag = String::from("<");
            prefixes.push_name(&mut tag, Ns::Text, "p");
            tag.push('>');
            tag
        });
        let name = xml::tag_name(&open).to_owned();
        let mut paragraph = open_tag(&open);
        xml::encode_text(&mut paragraph, prefixes, text);
        paragraph.push_str("</");
        paragraph.push_str(&name);
        paragraph.push('>');
        let tag = &styles[span.tag()];
        if span.close_start == span.end {
            // An empty region element grows to hold the paragraph.
            let mut whole = open_tag(tag);
            whole.push_str(&paragraph);
            whole.push_str("</");
            whole.push_str(xml::tag_name(tag));
            whole.push('>');
            return Ok(Some((span.range(), whole)));
        }
        Ok(Some((span.content(), paragraph)))
    }
}

/// The open form of a start tag, an empty element's included.
fn open_tag(tag: &str) -> String {
    match tag.strip_suffix("/>") {
        Some(open) => format!("{}>", open.trim_end()),
        None => tag.to_owned(),
    }
}
