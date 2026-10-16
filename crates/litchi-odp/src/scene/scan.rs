//! The one-pass scanner behind [`Scene::parse`](super::Scene::parse).

use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::styles::StylesInsertion;
use super::{
    Cell, FrameContent, Notes, Object, ObjectKind, Page, Paragraph, Prefixes, Row, Run, RunKind,
    Scene, Story, Table, WrappedField, Wrapper,
};

pub(super) const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(super) const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(super) const PRESENTATION_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const FO_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DR3D_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const XLINK_NS: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NS: &[u8] = b"http://www.w3.org/XML/1998/namespace";

/// Largest text the scanner indexes.
const MAX_XML_BYTES: usize = 128 * 1024 * 1024;
/// Deepest element nesting the scanner follows.
const MAX_DEPTH: usize = 1024;
/// Most pages, and most objects per container, the scanner records.
const MAX_ITEMS: usize = 1 << 20;
/// Largest `text:s` count the projection expands, as the slide parser.
const MAX_SPACES: usize = 1_000_000;

/// What the element at one depth is to the scanner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Frame {
    /// Any element the scene does not track.
    Other,
    /// `office:presentation`: pages are its children.
    Body,
    /// `draw:page`.
    Page,
    /// `presentation:notes`.
    Notes,
    /// A page's `office:forms`, which precedes its drawing objects.
    Forms,
    /// `draw:a` around drawing objects in a container.
    Link,
    /// A drawing object (the top of the object stack).
    Object,
    /// A frame's `draw:text-box`.
    TextBox,
    /// A frame's `table:table`.
    Table,
    /// `table:table-rows` or `table:table-header-rows`.
    RowGroup,
    /// `table:table-row`.
    Row,
    /// A table cell (the top of the story stack).
    Cell,
    /// `text:list`.
    List,
    /// `text:list-item` or `text:list-header`.
    ListItem,
    /// `text:p` or `text:h`.
    Paragraph,
    /// A tracked run element of a paragraph (span, link, field).
    Run,
    /// An element inside a tracked run or a text control.
    InRun,
    /// A subtree whose text the projection excludes.
    Opaque,
    /// An object's `svg:title`.
    Title,
    /// An object's `svg:desc`.
    Description,
    /// `office:automatic-styles`.
    AutomaticStyles,
}

/// The slide parser's paragraph text rules.
#[derive(Default)]
pub(super) struct Projection {
    value: String,
    trailing_collapsible_space: bool,
}

impl Projection {
    pub(super) fn push_text(&mut self, text: &str) {
        for character in text.chars() {
            if character.is_whitespace() {
                if !self.value.is_empty()
                    && !self
                        .value
                        .chars()
                        .next_back()
                        .is_some_and(char::is_whitespace)
                {
                    self.value.push(' ');
                    self.trailing_collapsible_space = true;
                }
            } else {
                self.value.push(character);
                self.trailing_collapsible_space = false;
            }
        }
    }

    pub(super) fn push_explicit(&mut self, character: char, count: usize) {
        self.value.extend(std::iter::repeat_n(character, count));
        self.trailing_collapsible_space = false;
    }

    pub(super) fn finish(mut self) -> String {
        if self.trailing_collapsible_space {
            self.value.pop();
        }
        self.value
    }
}

struct PageBuilder {
    start: usize,
    start_tag: Range<usize>,
    name: Option<String>,
    style_name: Option<String>,
    master_page: Option<String>,
    layout: Option<String>,
    objects: Vec<Object>,
    notes: Option<NotesBuilder>,
    insertion: usize,
}

struct NotesBuilder {
    start: usize,
    end: Option<usize>,
    objects: Vec<Object>,
}

/// Where a finished object goes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Owner {
    Page,
    Notes,
    Group,
}

struct ObjectBuilder {
    object: Object,
    owner: Owner,
    /// Whether a paragraph-level child was seen directly in the object.
    direct_story: bool,
    /// Whether another child followed a direct paragraph-level child.
    direct_interrupted: bool,
}

struct StoryBuilder {
    /// Byte offset where the story's content starts (text box content, or
    /// the first direct paragraph-level child).
    start: Option<usize>,
    /// The element of a self-closing container (`<draw:text-box/>`).
    empty: Option<Range<usize>>,
    end: usize,
    insertion: usize,
    paragraphs: Vec<Paragraph>,
    replaceable: bool,
    wrappers: Vec<WrapperFrame>,
}

struct WrapperFrame {
    start: usize,
    start_tag: Range<usize>,
    qname: String,
    item: bool,
    children: usize,
    sole: Option<usize>,
    /// Paragraphs whose innermost list item is this wrapper.
    items: Vec<usize>,
}

struct ParagraphBuilder {
    start: usize,
    start_tag: Range<usize>,
    content_start: usize,
    heading: bool,
    style_name: Option<String>,
    level: u8,
    text: Projection,
    runs: Vec<Run>,
    plain: Option<RunBuilder>,
    run: Option<RunBuilder>,
    wrappers: Vec<Wrapper>,
    /// The index of the innermost list item among the story's wrappers.
    item: Option<usize>,
    /// Inline content that is not text (record 0304).
    objects: Vec<(usize, String)>,
}

struct RunBuilder {
    kind: RunKind,
    start: usize,
    content_start: usize,
    content_end: usize,
    style_name: Option<String>,
    href: Option<String>,
    text: Projection,
    /// The direct children of the run element (controls included), its
    /// first one, and whether character data sits directly in it: a span
    /// that holds one field, or a link that holds one span, is recorded as
    /// such (record 0316).
    children: u8,
    child: Option<Child>,
    direct_text: bool,
}

/// The first direct child element of a run.
struct Child {
    /// A field-like inline element (not a span, link or control).
    field: bool,
    span: bool,
    start: usize,
    tag_end: usize,
    content_end: Option<usize>,
    end: Option<usize>,
    style_name: Option<String>,
}

impl RunBuilder {
    fn new(kind: RunKind, start: usize, content_start: usize, content_end: usize) -> Self {
        Self {
            kind,
            start,
            content_start,
            content_end,
            style_name: None,
            href: None,
            text: Projection::default(),
            children: 0,
            child: None,
            direct_text: false,
        }
    }

    /// The sole child element, when it is the run's whole content.
    fn sole_child(&self) -> Option<&Child> {
        let child = self.child.as_ref()?;
        (self.children == 1 && !self.direct_text && child.end.is_some()).then_some(child)
    }
}

struct TableBuilder {
    start: usize,
    name: Option<String>,
    style_name: Option<String>,
    columns: usize,
    rows: Vec<Row>,
    row: Option<(usize, Vec<Cell>)>,
}

struct CellBuilder {
    start: usize,
    covered: bool,
    column_span: u32,
    row_span: u32,
}

struct Scanner<'x> {
    /// Bytes skipped before the scanned text (a byte-order mark).
    base: usize,
    reader: NsReader<&'x [u8]>,
    frames: Vec<Frame>,
    pages: Vec<Page>,
    page: Option<PageBuilder>,
    objects: Vec<ObjectBuilder>,
    stories: Vec<StoryBuilder>,
    paragraph: Option<ParagraphBuilder>,
    table: Option<TableBuilder>,
    cells: Vec<CellBuilder>,
    alt: Option<String>,
    prefixes: Prefixes,
    root_seen: bool,
    body_end: Option<usize>,
    /// The content of `office:automatic-styles`, indexed on demand.
    styles_range: Option<Range<usize>>,
    styles_insertion: StylesInsertion,
}

pub(super) fn scan(xml: &str) -> Result<Scene> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("ODP content exceeds the 128 MiB scene limit"));
    }
    let (base, body) = match xml.strip_prefix('\u{feff}') {
        Some(body) => ('\u{feff}'.len_utf8(), body),
        None => (0, xml),
    };
    let mut reader = NsReader::from_str(body);
    reader.config_mut().check_end_names = true;
    let mut scanner = Scanner {
        base,
        reader,
        frames: Vec::new(),
        pages: Vec::new(),
        page: None,
        objects: Vec::new(),
        stories: Vec::new(),
        paragraph: None,
        table: None,
        cells: Vec::new(),
        alt: None,
        prefixes: Prefixes::default(),
        root_seen: false,
        body_end: None,
        styles_range: None,
        styles_insertion: StylesInsertion::None,
    };
    scanner.run()?;
    Ok(Scene {
        pages: scanner.pages,
        prefixes: scanner.prefixes,
        body_end: scanner.body_end,
        len: xml.len(),
        styles_range: scanner.styles_range,
        styles_insertion: scanner.styles_insertion,
    })
}

/// The prefixes a document root binds to the ODF namespaces markup is
/// authored in (the first binding of each).
pub(super) fn root_prefixes(element: &BytesStart<'_>) -> Result<Prefixes> {
    let mut prefixes = Prefixes::default();
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid ODP attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        let value = attribute.value.as_ref();
        let prefix = std::str::from_utf8(prefix)
            .map_err(|error| invalid(format!("invalid ODP namespace prefix: {error}")))?
            .to_owned();
        let slot = if value == DRAW_NS {
            &mut prefixes.draw
        } else if value == TEXT_NS {
            &mut prefixes.text
        } else if value == SVG_NS {
            &mut prefixes.svg
        } else if value == PRESENTATION_NS {
            &mut prefixes.presentation
        } else if value == STYLE_NS {
            &mut prefixes.style
        } else if value == FO_NS {
            &mut prefixes.fo
        } else if value == OFFICE_NS {
            &mut prefixes.office
        } else {
            continue;
        };
        if slot.is_none() {
            *slot = Some(prefix);
        }
    }
    Ok(prefixes)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn position(value: u64) -> Result<usize> {
    usize::try_from(value).map_err(|_| invalid("ODP content offset does not fit in memory"))
}

fn namespace<'a>(resolved: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match resolved {
        ResolveResult::Bound(Namespace(namespace)) => Some(namespace),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

impl Scanner<'_> {
    fn run(&mut self) -> Result<()> {
        loop {
            let start = self.base + position(self.reader.buffer_position())?;
            let (resolved, event) = self
                .reader
                .read_resolved_event()
                .map_err(|error| invalid(format!("invalid ODP content XML: {error}")))?;
            let element_namespace = namespace(&resolved).map(<[u8]>::to_vec);
            let end = self.base + position(self.reader.buffer_position())?;
            match event {
                Event::Start(element) => {
                    if self.frames.len() >= MAX_DEPTH {
                        return Err(invalid("ODP content exceeds the scene depth limit"));
                    }
                    let frame =
                        self.open(element_namespace.as_deref(), &element, start, end, false)?;
                    self.frames.push(frame);
                },
                Event::Empty(element) => {
                    let frame =
                        self.open(element_namespace.as_deref(), &element, start, end, true)?;
                    self.close(frame, start, end, end, true)?;
                },
                Event::End(_) => {
                    let frame = self
                        .frames
                        .pop()
                        .ok_or_else(|| invalid("ODP content has an unmatched end tag"))?;
                    self.close(frame, start, start, end, false)?;
                },
                Event::Text(text) => self.text(&decode_text(&text)?, start, end),
                Event::CData(data) => {
                    let text = std::str::from_utf8(&data)
                        .map_err(|error| invalid(format!("invalid ODP CDATA: {error}")))?
                        .to_owned();
                    self.text(&text, start, end);
                },
                Event::GeneralRef(reference) => {
                    self.text(&decode_reference(&reference)?, start, end);
                },
                Event::Eof => break,
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
            }
        }
        if !self.frames.is_empty() {
            return Err(invalid("ODP content ends inside an element"));
        }
        Ok(())
    }

    fn parent(&self) -> Option<Frame> {
        self.frames.last().copied()
    }

    /// Whether the innermost frame that can own drawing objects is a page,
    /// notes page or group, seen through `draw:a` links.
    fn object_owner(&self) -> Option<Owner> {
        for frame in self.frames.iter().rev() {
            match frame {
                Frame::Link => {},
                Frame::Page => return Some(Owner::Page),
                Frame::Notes => return Some(Owner::Notes),
                Frame::Object => {
                    return self
                        .objects
                        .last()
                        .filter(|object| object.object.kind == ObjectKind::Group)
                        .map(|_| Owner::Group);
                },
                _ => return None,
            }
        }
        None
    }

    fn open(
        &mut self,
        ns: Option<&[u8]>,
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<Frame> {
        let local = element.local_name();
        let local = local.as_ref();
        if !self.root_seen {
            self.root_seen = true;
            self.read_prefixes(element)?;
        }
        let parent = self.parent();
        // Inside a paragraph: runs, controls and opaque inline content.
        match parent {
            Some(Frame::Paragraph) => {
                return self.open_inline(ns, local, element, start, end, empty);
            },
            Some(Frame::Run | Frame::InRun) => {
                if parent == Some(Frame::Run) {
                    self.note_run_child(ns, local, element, start, end)?;
                }
                if self.is_control(ns, local) {
                    self.push_control(element, start, end, false)?;
                    return Ok(Frame::InRun);
                }
                if !(ns == Some(TEXT_NS) && local == b"span") {
                    self.note_object(element, start);
                }
                if is_opaque(ns, local) {
                    return Ok(Frame::Opaque);
                }
                return Ok(Frame::InRun);
            },
            Some(Frame::Opaque) => return Ok(Frame::Opaque),
            Some(Frame::Title | Frame::Description) => return Ok(Frame::Opaque),
            _ => {},
        }
        // `office:automatic-styles` and its children sit at depth 1 and 2.
        if self.frames.len() <= 2
            && let Some(frame) = self.open_style_element(ns, local, start, end, empty)
        {
            return Ok(frame);
        }
        if ns == Some(OFFICE_NS) && local == b"presentation" {
            return Ok(Frame::Body);
        }
        if ns == Some(DRAW_NS) && local == b"page" && parent == Some(Frame::Body) {
            self.open_page(element, start, end, empty)?;
            return Ok(Frame::Page);
        }
        if parent == Some(Frame::Page) {
            if ns == Some(PRESENTATION_NS) && local == b"notes" {
                if let Some(page) = &mut self.page {
                    page.notes = Some(NotesBuilder {
                        start,
                        end: if empty { Some(end) } else { None },
                        objects: Vec::new(),
                    });
                }
                return Ok(Frame::Notes);
            }
            if ns == Some(OFFICE_NS) && local == b"forms" {
                return Ok(Frame::Forms);
            }
        }
        if let Some(owner) = self.object_owner() {
            if ns == Some(DRAW_NS) && local == b"a" {
                return Ok(Frame::Link);
            }
            if let Some(kind) = object_kind(ns, local) {
                self.open_object(kind, owner, element, start, end, empty)?;
                return Ok(Frame::Object);
            }
            return Ok(Frame::Other);
        }
        if parent == Some(Frame::Object) {
            return self.open_object_child(ns, local, element, start, end, empty);
        }
        if matches!(
            parent,
            Some(Frame::TextBox | Frame::Cell | Frame::List | Frame::ListItem)
        ) {
            return self.open_story_child(ns, local, element, start, end, empty);
        }
        if parent == Some(Frame::Table) || parent == Some(Frame::RowGroup) {
            return self.open_table_child(ns, local, element, start, empty);
        }
        if parent == Some(Frame::Row) {
            return self.open_cell(ns, local, element, start, end, empty);
        }
        Ok(Frame::Other)
    }

    /// Note where `office:automatic-styles` holds its styles (indexed on
    /// demand, `Scene::automatic_styles`) and where `office:body` starts.
    fn open_style_element(
        &mut self,
        ns: Option<&[u8]>,
        local: &[u8],
        start: usize,
        end: usize,
        empty: bool,
    ) -> Option<Frame> {
        if self.frames.len() == 1 && ns == Some(OFFICE_NS) {
            if local == b"automatic-styles" {
                if empty {
                    self.styles_insertion = StylesInsertion::Expand(start..end);
                } else {
                    // `Before` and the range end are set at the end tag.
                    self.styles_insertion = StylesInsertion::None;
                    self.styles_range = Some(end..end);
                }
                return Some(Frame::AutomaticStyles);
            }
            if local == b"body" && self.styles_insertion == StylesInsertion::None {
                self.styles_insertion = StylesInsertion::Create(start);
            }
            return None;
        }
        // The styles' own subtree is not part of the scene.
        (self.parent() == Some(Frame::AutomaticStyles)).then_some(Frame::Opaque)
    }

    fn read_prefixes(&mut self, element: &BytesStart<'_>) -> Result<()> {
        self.prefixes = root_prefixes(element)?;
        Ok(())
    }

    fn attribute(
        &self,
        element: &BytesStart<'_>,
        ns: &[u8],
        local: &[u8],
    ) -> Result<Option<String>> {
        for attribute in element.attributes() {
            let attribute =
                attribute.map_err(|error| invalid(format!("invalid ODP attribute: {error}")))?;
            let (resolved, name) = self.reader.resolver().resolve_attribute(attribute.key);
            if namespace(&resolved) == Some(ns) && name.as_ref() == local {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, self.reader.decoder())
                    .map_err(|error| invalid(format!("invalid ODP attribute value: {error}")))?;
                return Ok(Some(value.into_owned()));
            }
        }
        Ok(None)
    }

    fn open_page(
        &mut self,
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<()> {
        if self.pages.len() >= MAX_ITEMS {
            return Err(invalid("ODP content exceeds the scene page limit"));
        }
        self.page = Some(PageBuilder {
            start,
            start_tag: start..end,
            name: self.attribute(element, DRAW_NS, b"name")?,
            style_name: self.attribute(element, DRAW_NS, b"style-name")?,
            master_page: self.attribute(element, DRAW_NS, b"master-page-name")?,
            layout: self.attribute(element, PRESENTATION_NS, b"presentation-page-layout-name")?,
            objects: Vec::new(),
            notes: None,
            insertion: if empty { usize::MAX } else { end },
        });
        Ok(())
    }

    fn open_object(
        &mut self,
        kind: ObjectKind,
        owner: Owner,
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<()> {
        let id = match self.attribute(element, DRAW_NS, b"id")? {
            Some(id) => Some(id),
            None => self.attribute(element, XML_NS, b"id")?,
        };
        let style_name = match self.attribute(element, DRAW_NS, b"style-name")? {
            Some(style) => Some(style),
            None => self.attribute(element, PRESENTATION_NS, b"style-name")?,
        };
        let placeholder = self
            .attribute(element, PRESENTATION_NS, b"placeholder")?
            .map(|value| value == "true");
        let object = Object {
            kind,
            content: FrameContent::None,
            range: start..end,
            start_tag: start..end,
            empty,
            name: self.attribute(element, DRAW_NS, b"name")?,
            class: self.attribute(element, PRESENTATION_NS, b"class")?,
            placeholder,
            id,
            style_name,
            text_style_name: self.attribute(element, DRAW_NS, b"text-style-name")?,
            layer: self.attribute(element, DRAW_NS, b"layer")?,
            z_index: self.attribute(element, DRAW_NS, b"z-index")?,
            x: self.attribute(element, SVG_NS, b"x")?,
            y: self.attribute(element, SVG_NS, b"y")?,
            width: self.attribute(element, SVG_NS, b"width")?,
            height: self.attribute(element, SVG_NS, b"height")?,
            transform: self.attribute(element, DRAW_NS, b"transform")?,
            href: None,
            title: None,
            description: None,
            story: None,
            table: None,
            children: Vec::new(),
        };
        self.objects.push(ObjectBuilder {
            object,
            owner,
            direct_story: false,
            direct_interrupted: false,
        });
        self.stories.push(StoryBuilder {
            start: None,
            empty: None,
            end,
            insertion: end,
            paragraphs: Vec::new(),
            replaceable: true,
            wrappers: Vec::new(),
        });
        Ok(())
    }

    fn open_object_child(
        &mut self,
        ns: Option<&[u8]>,
        local: &[u8],
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<Frame> {
        let is_frame = self
            .objects
            .last()
            .is_some_and(|object| object.object.kind == ObjectKind::Frame);
        if is_paragraph_level(ns, local) {
            if let Some(object) = self.objects.last_mut() {
                if object.direct_story && object.direct_interrupted {
                    if let Some(story) = self.stories.last_mut() {
                        story.replaceable = false;
                    }
                }
                object.direct_story = true;
            }
            if let Some(story) = self.stories.last_mut()
                && story.start.is_none()
            {
                story.start = Some(start);
                story.insertion = start;
            }
            return self.open_story_child(ns, local, element, start, end, empty);
        }
        if let Some(object) = self.objects.last_mut()
            && object.direct_story
        {
            object.direct_interrupted = true;
        }
        if ns == Some(SVG_NS) && local == b"title" {
            self.alt = Some(String::new());
            return Ok(Frame::Title);
        }
        if ns == Some(SVG_NS) && local == b"desc" {
            self.alt = Some(String::new());
            return Ok(Frame::Description);
        }
        if !is_frame {
            return Ok(Frame::Other);
        }
        let content = match (ns, local) {
            (Some(DRAW_NS), b"text-box") => FrameContent::TextBox,
            (Some(DRAW_NS), b"image") => FrameContent::Image,
            (Some(TABLE_NS), b"table") => FrameContent::Table,
            (Some(DRAW_NS), b"object" | b"object-ole") => FrameContent::Object,
            (Some(DRAW_NS), b"plugin") => FrameContent::Plugin,
            (Some(DRAW_NS), b"applet" | b"floating-frame") => FrameContent::Other,
            _ => return Ok(Frame::Other),
        };
        let href = self.attribute(element, XLINK_NS, b"href")?;
        let Some(object) = self.objects.last_mut() else {
            return Ok(Frame::Other);
        };
        if content == FrameContent::Image && object.object.href.is_none() {
            object.object.href.clone_from(&href);
        }
        if object.object.content != FrameContent::None {
            // A frame's first recognised child decides its content; later
            // children are replacement images and alternatives.
            if content == FrameContent::Image {
                object.object.href = href;
            }
            return Ok(Frame::Other);
        }
        object.object.content = content;
        if object.object.href.is_none() {
            object.object.href = href;
        }
        match content {
            FrameContent::TextBox => {
                if let Some(story) = self.stories.last_mut() {
                    story.start = Some(end);
                    story.end = end;
                    story.insertion = if empty { usize::MAX } else { end };
                    story.empty = empty.then_some(start..end);
                }
                Ok(Frame::TextBox)
            },
            FrameContent::Table => {
                self.table = Some(TableBuilder {
                    start,
                    name: self.attribute(element, TABLE_NS, b"name")?,
                    style_name: self.attribute(element, TABLE_NS, b"style-name")?,
                    columns: 0,
                    rows: Vec::new(),
                    row: None,
                });
                Ok(Frame::Table)
            },
            _ => Ok(Frame::Other),
        }
    }

    fn open_story_child(
        &mut self,
        ns: Option<&[u8]>,
        local: &[u8],
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<Frame> {
        let parent = self.parent();
        let direct = matches!(parent, Some(Frame::TextBox | Frame::Cell | Frame::Object));
        if direct && !is_paragraph_level(ns, local) {
            if let Some(story) = self.stories.last_mut() {
                story.replaceable = false;
            }
        }
        if ns == Some(TEXT_NS) && (local == b"p" || local == b"h") {
            self.open_paragraph(local == b"h", element, start, end)?;
            return Ok(Frame::Paragraph);
        }
        if ns == Some(TEXT_NS) && matches!(local, b"list" | b"list-item" | b"list-header") {
            let item = local != b"list";
            let qname = std::str::from_utf8(element.name().as_ref())
                .map_err(|error| invalid(format!("invalid ODP element name: {error}")))?
                .to_owned();
            if let Some(story) = self.stories.last_mut() {
                if empty {
                    register_child(story, None);
                } else {
                    story.wrappers.push(WrapperFrame {
                        start,
                        start_tag: start..end,
                        qname,
                        item,
                        children: 0,
                        sole: None,
                        items: Vec::new(),
                    });
                }
            }
            return Ok(if item { Frame::ListItem } else { Frame::List });
        }
        if let Some(story) = self.stories.last_mut()
            && !direct
        {
            // Another child inside a list wrapper (`text:number`, a nested
            // object): the wrapper does not hold its paragraph alone.
            register_child(story, None);
        }
        Ok(Frame::Other)
    }

    fn open_paragraph(
        &mut self,
        heading: bool,
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let style_name = self.attribute(element, TEXT_NS, b"style-name")?;
        let Some(story) = self.stories.last() else {
            return Ok(());
        };
        let wrappers: Vec<Wrapper> = story
            .wrappers
            .iter()
            .map(|wrapper| Wrapper {
                start_tag: wrapper.start_tag.clone(),
                qname: wrapper.qname.clone(),
                item: wrapper.item,
            })
            .collect();
        let level = u8::try_from(
            story
                .wrappers
                .iter()
                .filter(|wrapper| !wrapper.item)
                .count(),
        )
        .unwrap_or(u8::MAX);
        let item = story.wrappers.iter().rposition(|wrapper| wrapper.item);
        self.paragraph = Some(ParagraphBuilder {
            start,
            start_tag: start..end,
            content_start: end,
            heading,
            style_name,
            level,
            text: Projection::default(),
            runs: Vec::new(),
            plain: None,
            run: None,
            wrappers,
            item,
            objects: Vec::new(),
        });
        Ok(())
    }

    /// Count a direct child element of the open run, remembering the first.
    fn note_run_child(
        &mut self,
        ns: Option<&[u8]>,
        local: &[u8],
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let span = ns == Some(TEXT_NS) && local == b"span";
        let link = ns == Some(TEXT_NS) && local == b"a";
        let field = !(span || link || self.is_control(ns, local) || is_opaque(ns, local));
        let style_name = if span {
            self.attribute(element, TEXT_NS, b"style-name")?
        } else {
            None
        };
        if let Some(run) = self
            .paragraph
            .as_mut()
            .and_then(|paragraph| paragraph.run.as_mut())
        {
            if run.children == 0 {
                run.child = Some(Child {
                    field,
                    span,
                    start,
                    tag_end: end,
                    content_end: None,
                    end: None,
                    style_name,
                });
            }
            run.children = run.children.saturating_add(1);
        }
        Ok(())
    }

    /// Record inline content of the open paragraph that is not text: a
    /// frame, a note, an annotation, a field or a hyperlink (record 0304).
    fn note_object(&mut self, element: &BytesStart<'_>, start: usize) {
        if let Some(paragraph) = &mut self.paragraph
            && paragraph.objects.len() < 8
        {
            let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
            paragraph.objects.push((start, name));
        }
    }

    fn is_control(&self, ns: Option<&[u8]>, local: &[u8]) -> bool {
        ns == Some(TEXT_NS) && matches!(local, b"s" | b"tab" | b"line-break")
    }

    fn push_control(
        &mut self,
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        direct: bool,
    ) -> Result<()> {
        let local = element.local_name();
        let (character, count) = match local.as_ref() {
            b"tab" => ('\t', 1),
            b"line-break" => ('\n', 1),
            _ => {
                let count = self
                    .attribute(element, TEXT_NS, b"c")?
                    .map(|value| {
                        value
                            .trim()
                            .parse::<usize>()
                            .map_err(|_| invalid(format!("invalid text:s count '{value}'")))
                    })
                    .transpose()?
                    .unwrap_or(1);
                if count > MAX_SPACES {
                    return Err(invalid("text:s count exceeds the supported safety limit"));
                }
                (' ', count)
            },
        };
        let Some(paragraph) = &mut self.paragraph else {
            return Ok(());
        };
        paragraph.text.push_explicit(character, count);
        if direct {
            let plain = paragraph
                .plain
                .get_or_insert_with(|| RunBuilder::new(RunKind::Text, start, start, end));
            plain.content_end = end;
            plain.text.push_explicit(character, count);
        } else if let Some(run) = &mut paragraph.run {
            run.text.push_explicit(character, count);
        }
        Ok(())
    }

    fn open_inline(
        &mut self,
        ns: Option<&[u8]>,
        local: &[u8],
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<Frame> {
        if self.is_control(ns, local) {
            self.push_control(element, start, end, true)?;
            return Ok(Frame::InRun);
        }
        self.finish_plain();
        if !(ns == Some(TEXT_NS) && local == b"span") {
            self.note_object(element, start);
        }
        if is_opaque(ns, local) {
            return Ok(Frame::Opaque);
        }
        let (kind, style_name, href) = if ns == Some(TEXT_NS) && local == b"span" {
            (
                RunKind::Span,
                self.attribute(element, TEXT_NS, b"style-name")?,
                None,
            )
        } else if ns == Some(TEXT_NS) && local == b"a" {
            (
                RunKind::Link,
                self.attribute(element, TEXT_NS, b"style-name")?,
                self.attribute(element, XLINK_NS, b"href")?,
            )
        } else {
            (RunKind::Field, None, None)
        };
        if let Some(paragraph) = &mut self.paragraph {
            let mut run = RunBuilder::new(kind, start, end, end);
            run.style_name = style_name;
            run.href = href;
            paragraph.run = Some(run);
        }
        let _ = empty;
        Ok(Frame::Run)
    }

    fn finish_plain(&mut self) {
        if let Some(paragraph) = &mut self.paragraph
            && let Some(plain) = paragraph.plain.take()
        {
            paragraph.runs.push(Run {
                kind: RunKind::Text,
                range: plain.start..plain.content_end,
                content: plain.content_start..plain.content_end,
                start_tag: plain.start..plain.start,
                style_name: None,
                href: None,
                text: plain.text.finish(),
                field: None,
                inner_span: None,
            });
        }
    }

    fn open_table_child(
        &mut self,
        ns: Option<&[u8]>,
        local: &[u8],
        element: &BytesStart<'_>,
        start: usize,
        empty: bool,
    ) -> Result<Frame> {
        if ns != Some(TABLE_NS) {
            return Ok(Frame::Other);
        }
        match local {
            b"table-column" => {
                let repeated = self
                    .attribute(element, TABLE_NS, b"number-columns-repeated")?
                    .and_then(|value| value.trim().parse::<usize>().ok())
                    .unwrap_or(1)
                    .max(1);
                if let Some(table) = &mut self.table {
                    table.columns = table.columns.saturating_add(repeated);
                }
                Ok(Frame::Other)
            },
            b"table-rows" | b"table-header-rows" => Ok(Frame::RowGroup),
            b"table-row" => {
                if let Some(table) = &mut self.table {
                    if table.rows.len() >= MAX_ITEMS {
                        return Err(invalid("ODP table exceeds the scene row limit"));
                    }
                    table.row = Some((start, Vec::new()));
                }
                if empty {
                    return Ok(Frame::Row);
                }
                Ok(Frame::Row)
            },
            _ => Ok(Frame::Other),
        }
    }

    fn open_cell(
        &mut self,
        ns: Option<&[u8]>,
        local: &[u8],
        element: &BytesStart<'_>,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<Frame> {
        if ns != Some(TABLE_NS) || !matches!(local, b"table-cell" | b"covered-table-cell") {
            return Ok(Frame::Other);
        }
        let span = |value: Option<String>| {
            value
                .and_then(|value| value.trim().parse::<u32>().ok())
                .unwrap_or(1)
                .max(1)
        };
        self.cells.push(CellBuilder {
            start,
            covered: local == b"covered-table-cell",
            column_span: span(self.attribute(element, TABLE_NS, b"number-columns-spanned")?),
            row_span: span(self.attribute(element, TABLE_NS, b"number-rows-spanned")?),
        });
        self.stories.push(StoryBuilder {
            start: Some(end),
            empty: empty.then_some(start..end),
            end,
            insertion: if empty { usize::MAX } else { end },
            paragraphs: Vec::new(),
            replaceable: true,
            wrappers: Vec::new(),
        });
        Ok(Frame::Cell)
    }

    fn text(&mut self, text: &str, start: usize, end: usize) {
        match self.parent() {
            Some(Frame::Paragraph) => {
                if let Some(paragraph) = &mut self.paragraph {
                    paragraph.text.push_text(text);
                    let plain = paragraph
                        .plain
                        .get_or_insert_with(|| RunBuilder::new(RunKind::Text, start, start, end));
                    plain.content_end = end;
                    plain.text.push_text(text);
                }
            },
            Some(frame @ (Frame::Run | Frame::InRun)) => {
                if let Some(paragraph) = &mut self.paragraph {
                    paragraph.text.push_text(text);
                    if let Some(run) = &mut paragraph.run {
                        run.text.push_text(text);
                        run.direct_text |= frame == Frame::Run;
                    }
                }
            },
            Some(Frame::Title | Frame::Description) => {
                if let Some(alt) = &mut self.alt {
                    alt.push_str(text);
                }
            },
            _ => {},
        }
    }

    fn close(
        &mut self,
        frame: Frame,
        _start: usize,
        content_end: usize,
        end: usize,
        empty: bool,
    ) -> Result<()> {
        match frame {
            Frame::Other => {},
            Frame::AutomaticStyles => {
                if !empty {
                    self.styles_insertion = StylesInsertion::Before(content_end);
                    if let Some(range) = &mut self.styles_range {
                        range.end = content_end;
                    }
                }
            },
            Frame::Forms => {
                if let Some(page) = &mut self.page
                    && page.objects.is_empty()
                {
                    page.insertion = end;
                }
            },
            Frame::Body => self.body_end = Some(content_end),
            Frame::Page => self.close_page(content_end, end, empty)?,
            Frame::Notes => {
                if let Some(page) = &mut self.page
                    && let Some(notes) = &mut page.notes
                {
                    notes.end = Some(end);
                }
            },
            Frame::Object => self.close_object(content_end, end)?,
            Frame::TextBox => {
                if let Some(story) = self.stories.last_mut() {
                    story.end = content_end;
                }
            },
            Frame::Table => {
                if let Some(table) = self.table.take()
                    && let Some(object) = self.objects.last_mut()
                {
                    object.object.table = Some(Table {
                        range: table.start..end,
                        name: table.name,
                        style_name: table.style_name,
                        columns: table.columns,
                        rows: table.rows,
                    });
                }
            },
            Frame::Row => {
                if let Some(table) = &mut self.table
                    && let Some((row_start, cells)) = table.row.take()
                {
                    table.rows.push(Row {
                        range: row_start..end,
                        cells,
                    });
                }
            },
            Frame::Cell => {
                let story = self.stories.pop();
                if let (Some(cell), Some(story)) = (self.cells.pop(), story) {
                    let story = finish_story(story, content_end.max(cell.start));
                    if let Some(table) = &mut self.table
                        && let Some((_, cells)) = &mut table.row
                    {
                        cells.push(Cell {
                            range: cell.start..end,
                            covered: cell.covered,
                            column_span: cell.column_span,
                            row_span: cell.row_span,
                            story,
                        });
                    }
                }
            },
            Frame::List | Frame::ListItem => {
                if let Some(story) = self.stories.last_mut() {
                    if empty {
                        // Registered at open.
                    } else if let Some(wrapper) = story.wrappers.pop() {
                        let range = wrapper.start..end;
                        for &index in &wrapper.items {
                            if let Some(paragraph) = story.paragraphs.get_mut(index) {
                                paragraph.item = Some(range.clone());
                            }
                        }
                        let sole = if wrapper.children == 1 {
                            wrapper.sole
                        } else {
                            None
                        };
                        if let Some(index) = sole
                            && let Some(paragraph) = story.paragraphs.get_mut(index)
                        {
                            paragraph.removal = range;
                        }
                        register_child(story, sole);
                    }
                    if self.frames.last().is_some_and(|frame| {
                        matches!(frame, Frame::TextBox | Frame::Cell | Frame::Object)
                    }) {
                        story.end = end;
                    }
                }
            },
            Frame::Paragraph => self.close_paragraph(content_end, end, empty),
            Frame::Run => {
                if let Some(paragraph) = &mut self.paragraph
                    && let Some(mut run) = paragraph.run.take()
                {
                    run.content_end = content_end;
                    let sole = run.sole_child().map(|child| {
                        let child_end = child.end.unwrap_or(child.tag_end);
                        (
                            child.field,
                            child.span,
                            child.start..child.tag_end,
                            child.tag_end..child.content_end.unwrap_or(child_end),
                            child.start..child_end,
                            child.style_name.clone(),
                        )
                    });
                    let text = run.text.finish();
                    let mut kind = run.kind;
                    let mut field = None;
                    let mut inner_span = None;
                    match sole {
                        // A span holding one field is that field, formatted
                        // (record 0316).
                        Some((true, _, tag, content, range, _))
                            if kind == RunKind::Span && !text.is_empty() =>
                        {
                            kind = RunKind::Field;
                            field = Some(WrappedField {
                                range,
                                start_tag: tag,
                                content,
                            });
                        },
                        // A link holding one span carries that span's
                        // formatting.
                        Some((_, true, tag, _, _, style)) if kind == RunKind::Link => {
                            inner_span = Some((tag, style));
                        },
                        _ => {},
                    }
                    let tag_end = if run.content_start > run.start {
                        run.content_start
                    } else {
                        end
                    };
                    if kind != RunKind::Field || !text.is_empty() {
                        paragraph.runs.push(Run {
                            kind,
                            range: run.start..end,
                            content: run.content_start..run.content_end.max(run.content_start),
                            start_tag: run.start..tag_end,
                            style_name: run.style_name,
                            href: run.href,
                            text,
                            field,
                            inner_span,
                        });
                    }
                }
            },
            Frame::Title | Frame::Description => {
                let text = self.alt.take().map(|text| text.trim().to_owned());
                if let Some(object) = self.objects.last_mut() {
                    if frame == Frame::Title {
                        object.object.title = text;
                    } else {
                        object.object.description = text;
                    }
                }
            },
            Frame::InRun | Frame::Opaque => {
                // The first direct child of a run closes.
                if self.parent() == Some(Frame::Run)
                    && let Some(child) = self
                        .paragraph
                        .as_mut()
                        .and_then(|paragraph| paragraph.run.as_mut())
                        .and_then(|run| run.child.as_mut())
                    && child.end.is_none()
                {
                    child.content_end = Some(content_end);
                    child.end = Some(end);
                }
            },
            Frame::Link | Frame::RowGroup => {},
        }
        Ok(())
    }

    fn close_page(&mut self, content_end: usize, end: usize, empty: bool) -> Result<()> {
        let Some(page) = self.page.take() else {
            return Ok(());
        };
        let notes = page.notes.map(|notes| Notes {
            range: notes.start..notes.end.unwrap_or(content_end),
            objects: notes.objects,
        });
        let insertion = if empty {
            None
        } else if page.insertion == usize::MAX {
            None
        } else {
            Some(page.insertion)
        };
        self.pages.push(Page {
            range: page.start..end,
            start_tag: page.start_tag,
            empty,
            name: page.name,
            style_name: page.style_name,
            master_page: page.master_page,
            layout: page.layout,
            objects: page.objects,
            notes,
            insertion,
        });
        Ok(())
    }

    fn close_object(&mut self, content_end: usize, end: usize) -> Result<()> {
        let (Some(mut builder), Some(story)) = (self.objects.pop(), self.stories.pop()) else {
            return Err(invalid("ODP scene object stack is unbalanced"));
        };
        builder.object.range = builder.object.range.start..end;
        let text_box = builder.object.content == FrameContent::TextBox;
        if text_box || builder.direct_story {
            let end = story.end;
            builder.object.story = Some(finish_story(story, end));
        }
        let _ = content_end;
        let object = builder.object;
        match builder.owner {
            Owner::Group => {
                if let Some(parent) = self.objects.last_mut() {
                    if parent.object.children.len() >= MAX_ITEMS {
                        return Err(invalid("ODP group exceeds the scene object limit"));
                    }
                    parent.object.children.push(object);
                }
            },
            Owner::Notes => {
                if let Some(page) = &mut self.page
                    && let Some(notes) = &mut page.notes
                {
                    notes.objects.push(object);
                }
            },
            Owner::Page => {
                if let Some(page) = &mut self.page {
                    if page.objects.len() >= MAX_ITEMS {
                        return Err(invalid("ODP page exceeds the scene object limit"));
                    }
                    page.insertion = end;
                    page.objects.push(object);
                }
            },
        }
        Ok(())
    }

    fn close_paragraph(&mut self, content_end: usize, end: usize, empty: bool) {
        self.finish_plain();
        let Some(paragraph) = self.paragraph.take() else {
            return;
        };
        let Some(story) = self.stories.last_mut() else {
            return;
        };
        let index = story.paragraphs.len();
        let range = paragraph.start..end;
        story.paragraphs.push(Paragraph {
            range: range.clone(),
            start_tag: paragraph.start_tag,
            content: paragraph.content_start..content_end.max(paragraph.content_start),
            empty,
            heading: paragraph.heading,
            style_name: paragraph.style_name,
            level: paragraph.level,
            text: paragraph.text.finish(),
            runs: paragraph.runs,
            wrappers: paragraph.wrappers,
            removal: range,
            item: None,
            objects: paragraph.objects,
        });
        if let Some(item) = paragraph.item
            && let Some(wrapper) = story.wrappers.get_mut(item)
        {
            wrapper.items.push(index);
        }
        register_child(story, Some(index));
        if story.wrappers.is_empty() {
            story.end = end;
        }
    }
}

/// Count one child of the innermost list wrapper, remembering a paragraph
/// that might be its only child.
fn register_child(story: &mut StoryBuilder, paragraph: Option<usize>) {
    if let Some(wrapper) = story.wrappers.last_mut() {
        wrapper.children = wrapper.children.saturating_add(1);
        wrapper.sole = if wrapper.children == 1 {
            paragraph
        } else {
            None
        };
    }
}

/// Close a story whose content ends at `end`.
fn finish_story(story: StoryBuilder, end: usize) -> Story {
    let start = story.start.unwrap_or(end);
    Story {
        range: start..end.max(start),
        paragraphs: story.paragraphs,
        replaceable: story.replaceable,
        insertion: story.insertion,
        empty: story.empty,
    }
}

fn object_kind(ns: Option<&[u8]>, local: &[u8]) -> Option<ObjectKind> {
    match ns? {
        DRAW_NS => Some(match local {
            b"frame" => ObjectKind::Frame,
            b"rect" => ObjectKind::Rectangle,
            b"ellipse" | b"circle" => ObjectKind::Ellipse,
            b"line" => ObjectKind::Line,
            b"connector" => ObjectKind::Connector,
            b"custom-shape" => ObjectKind::CustomShape,
            b"path" | b"polygon" | b"polyline" | b"regular-polygon" => ObjectKind::Polygon,
            b"caption" | b"measure" => ObjectKind::Annotation,
            b"control" => ObjectKind::Control,
            b"g" => ObjectKind::Group,
            b"page-thumbnail" => ObjectKind::PageThumbnail,
            _ => return None,
        }),
        DR3D_NS if local == b"scene" => Some(ObjectKind::Scene3d),
        _ => None,
    }
}

fn is_paragraph_level(ns: Option<&[u8]>, local: &[u8]) -> bool {
    ns == Some(TEXT_NS) && matches!(local, b"p" | b"h" | b"list")
}

/// Inline content whose text the projection excludes: annotations, notes
/// and drawing objects anchored in a paragraph.
fn is_opaque(ns: Option<&[u8]>, local: &[u8]) -> bool {
    match ns {
        Some(OFFICE_NS) => matches!(local, b"annotation" | b"annotation-end"),
        Some(TEXT_NS) => matches!(local, b"note" | b"ruby-text"),
        Some(DRAW_NS | DR3D_NS | TABLE_NS | PRESENTATION_NS) => true,
        _ => false,
    }
}

fn decode_text(text: &BytesText<'_>) -> Result<String> {
    text.xml_content(XmlVersion::Explicit1_0)
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| invalid(format!("invalid ODP text: {error}")))
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| invalid(format!("invalid ODP character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| invalid(format!("invalid ODP entity reference: {error}")))?;
    Ok(match name.as_ref() {
        "amp" => "&",
        "lt" => "<",
        "gt" => ">",
        "quot" => "\"",
        "apos" => "'",
        _ => {
            return Err(invalid(format!(
                "unsupported ODP entity reference '&{name};'"
            )));
        },
    }
    .to_owned())
}
