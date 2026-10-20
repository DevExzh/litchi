//! A byte-range index of the presentation pages in `content.xml`.
//!
//! [`Scene::parse`] reads one presentation body once (a packaged
//! `content.xml`, or a flat `office:document`) and records, for every
//! `draw:page`, its drawing objects in document order, their text stories
//! (paragraphs and runs), tables and speaker notes. Every entry keeps the
//! exact byte range it occupies in the scanned text, so callers can address
//! the producer's own markup without regenerating it.
//!
//! Text is projected with the slide parser's rules: runs of XML whitespace
//! collapse to one space, `text:s`, `text:tab` and `text:line-break` are
//! explicit characters, and a paragraph's trailing collapsible space is
//! dropped. A story's text is its paragraphs joined by `\n`.
//!
//! The index is read-only. [`Change`] values describe edits against it;
//! [`crate::edit::Edit::splice_scene`] plans them as checked byte splices of
//! `content.xml`, so every byte outside the changed ranges keeps its source
//! provenance.
//!
//! ```
//! use litchi_odp::scene::Scene;
//!
//! let xml = concat!(
//!     r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
//!     r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
//!     r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
//!     r#"<office:body><office:presentation><draw:page draw:name="One">"#,
//!     r#"<draw:frame draw:name="Box"><draw:text-box><text:p>Hello <text:span>world</text:span></text:p>"#,
//!     r#"</draw:text-box></draw:frame></draw:page></office:presentation></office:body>"#,
//!     r#"</office:document-content>"#,
//! );
//! let scene = Scene::parse(xml)?;
//! let page = &scene.pages()[0];
//! assert_eq!(page.name(), Some("One"));
//! let story = page.objects()[0].story().expect("a text box");
//! assert_eq!(story.paragraphs()[0].text(), "Hello world");
//! assert_eq!(story.paragraphs()[0].runs().len(), 2);
//! # Ok::<(), litchi_core::Error>(())
//! ```

mod change;
mod markup;
pub mod parts;
mod scan;
mod styles;
#[cfg(test)]
mod tests;

use std::ops::Range;

use litchi_core::Result;

use change::Fragment;
pub use change::{Attribute, Change, NewPage, NewTextBox, ObjectRef, Scope, ShapeAction, StoryRef};
pub use styles::{
    AutomaticStyle, AutomaticStyles, CommonStyles, TextProperties, TextPropertyUpdate,
};

/// Apply `changes` to the `content.xml` of `source` as checked splices.
///
/// Returns `None` when the changes are exact no-ops.
pub(crate) fn apply_changes(
    source: &crate::core::OwnedPackage,
    changes: &[Change],
) -> Result<Option<Vec<u8>>> {
    use litchi_odf_common::core::{AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication};

    let part = XmlSourcePart::load(source, "content.xml")?;
    let content = std::str::from_utf8(part.bytes()).map_err(|error| {
        litchi_core::Error::InvalidFormat(format!("ODP content.xml is not UTF-8: {error}"))
    })?;
    let scene = Scene::parse(content)?;
    let plan = change::plan(content, &scene, changes)?;
    if plan.is_empty() {
        return Ok(None);
    }
    let updated = plan.assemble(content);
    let mut publication = XmlSplicePublication::new(part.clone());
    for (range, fragment) in &plan.edits {
        let expected = content.as_bytes().get(range.clone()).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("invalid ODP scene splice range".to_string())
        })?;
        let proof = part.checked_range(range.clone(), expected)?;
        let fragment = match fragment {
            Fragment::Text(text) => AuthoredXmlFragment::text(text.as_bytes().to_vec())?,
            Fragment::Markup(markup) => AuthoredXmlFragment::markup(markup.as_bytes().to_vec())?,
            Fragment::StartTag(tag) => AuthoredXmlFragment::start_tag(tag.as_bytes().to_vec())?,
            Fragment::Relocated(source) => AuthoredXmlFragment::relocated(&part, source.clone())?,
            Fragment::Delete => AuthoredXmlFragment::deletion(),
        };
        publication.replace(proof, fragment)?;
    }
    litchi_odf_common::package::replace_content_xml_spliced(source, &updated, publication).map(Some)
}

/// The pages of one presentation body and the namespace prefixes its root
/// element binds.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Scene {
    pub(crate) pages: Vec<Page>,
    pub(crate) prefixes: Prefixes,
    /// Where an inserted page goes when the body has no page yet: just
    /// before the end tag of `office:presentation`.
    pub(crate) body_end: Option<usize>,
    pub(crate) len: usize,
    /// The content of `office:automatic-styles`, indexed on demand by
    /// [`Scene::automatic_styles`].
    pub(crate) styles_range: Option<Range<usize>>,
    /// Where new automatic styles go.
    pub(crate) styles_insertion: styles::StylesInsertion,
}

/// The prefixes the document root binds to the ODF namespaces markup is
/// authored in; `None` when the root does not bind the namespace.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct Prefixes {
    pub(crate) draw: Option<String>,
    pub(crate) text: Option<String>,
    pub(crate) svg: Option<String>,
    pub(crate) presentation: Option<String>,
    pub(crate) style: Option<String>,
    pub(crate) fo: Option<String>,
    pub(crate) office: Option<String>,
}

impl Scene {
    /// Index the presentation pages of `xml`, a `content.xml` part or a flat
    /// `office:document`.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_core::Error::InvalidFormat`] for malformed XML and
    /// [`litchi_core::Error::ResourceLimit`]-class failures (as
    /// `InvalidFormat`) when the text exceeds the scan's size, depth or
    /// event bounds.
    pub fn parse(xml: &str) -> Result<Self> {
        scan::scan(xml)
    }

    /// The presentation pages, in document order.
    #[must_use]
    pub fn pages(&self) -> &[Page] {
        &self.pages
    }

    /// The page at zero-based `index`.
    #[must_use]
    pub fn page(&self, index: usize) -> Option<&Page> {
        self.pages.get(index)
    }

    /// The length in bytes of the scanned text.
    #[must_use]
    pub fn len(&self) -> usize {
        self.len
    }

    /// Whether the scanned text was empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Index the automatic styles of `xml`, the text this scene indexes
    /// (`office:automatic-styles`: each `style:style`, its family, parent
    /// and `style:text-properties`). The scan only records where they are,
    /// so reading a presentation does not pay for them.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_core::Error::InvalidFormat`] when `xml` is not the
    /// text this scene indexes or its styles are malformed.
    pub fn automatic_styles(&self, xml: &str) -> Result<AutomaticStyles> {
        if xml.len() != self.len {
            return Err(litchi_core::Error::InvalidFormat(
                "the ODP scene does not index this content".to_string(),
            ));
        }
        styles::parse(xml, self.styles_range.clone(), &self.prefixes)
    }

    /// Resolve an object reference against this scene.
    #[must_use]
    pub fn object(&self, object: &ObjectRef) -> Option<&Object> {
        let page = self.pages.get(object.page)?;
        let roots = match object.scope {
            Scope::Page => page.objects.as_slice(),
            Scope::Notes => page.notes.as_ref()?.objects.as_slice(),
        };
        let (&first, rest) = object.path.split_first()?;
        let mut current = roots.get(first)?;
        for &index in rest {
            current = current.children.get(index)?;
        }
        Some(current)
    }

    /// Resolve a story reference: an object's text story or one of its
    /// table cells' stories.
    #[must_use]
    pub fn story(&self, story: &StoryRef) -> Option<&Story> {
        let object = self.object(&story.object)?;
        match story.cell {
            None => object.story.as_ref(),
            Some((row, cell)) => {
                Some(&object.table.as_ref()?.rows.get(row)?.cells.get(cell)?.story)
            },
        }
    }
}

/// One `draw:page`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Page {
    pub(crate) range: Range<usize>,
    pub(crate) start_tag: Range<usize>,
    pub(crate) empty: bool,
    pub(crate) name: Option<String>,
    pub(crate) style_name: Option<String>,
    pub(crate) master_page: Option<String>,
    pub(crate) layout: Option<String>,
    pub(crate) objects: Vec<Object>,
    pub(crate) notes: Option<Notes>,
    /// Where an inserted drawing object goes: after the last object, else
    /// after `office:forms`, else after the start tag; `None` for an empty
    /// (self-closing) page.
    pub(crate) insertion: Option<usize>,
}

impl Page {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// `draw:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// `draw:style-name`, the drawing-page style.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// `draw:master-page-name`.
    #[must_use]
    pub fn master_page(&self) -> Option<&str> {
        self.master_page.as_deref()
    }

    /// `presentation:presentation-page-layout-name`.
    #[must_use]
    pub fn layout(&self) -> Option<&str> {
        self.layout.as_deref()
    }

    /// The page's top-level drawing objects, in document order (z-order).
    #[must_use]
    pub fn objects(&self) -> &[Object] {
        &self.objects
    }

    /// The speaker-notes page, when the slide has one.
    #[must_use]
    pub fn notes(&self) -> Option<&Notes> {
        self.notes.as_ref()
    }

    /// The text of the first title placeholder (`presentation:class="title"`).
    #[must_use]
    pub fn title(&self) -> Option<String> {
        self.objects
            .iter()
            .find(|object| object.class.as_deref() == Some("title"))
            .and_then(Object::story)
            .map(Story::text)
    }
}

/// A slide's `presentation:notes` page.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Notes {
    pub(crate) range: Range<usize>,
    pub(crate) objects: Vec<Object>,
}

impl Notes {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// The notes page's drawing objects (usually a page thumbnail and the
    /// notes text frame).
    #[must_use]
    pub fn objects(&self) -> &[Object] {
        &self.objects
    }

    /// The position of the notes text frame among [`Self::objects`]: the
    /// first `presentation:class="notes"` object with a story, else the
    /// first object with a story.
    #[must_use]
    pub fn text_object(&self) -> Option<usize> {
        self.objects
            .iter()
            .position(|object| object.class.as_deref() == Some("notes") && object.story.is_some())
            .or_else(|| {
                self.objects
                    .iter()
                    .position(|object| object.story.is_some())
            })
    }

    /// The notes text: the text frame's story text.
    #[must_use]
    pub fn text(&self) -> Option<String> {
        self.text_object()
            .and_then(|index| self.objects[index].story.as_ref())
            .map(Story::text)
    }
}

/// The drawing element of an [`Object`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ObjectKind {
    /// `draw:frame`: a text box, picture, table, embedded object or media.
    Frame,
    /// `draw:rect`.
    Rectangle,
    /// `draw:ellipse` or `draw:circle`.
    Ellipse,
    /// `draw:line`.
    Line,
    /// `draw:connector`.
    Connector,
    /// `draw:custom-shape`.
    CustomShape,
    /// `draw:path`, `draw:polygon`, `draw:polyline` or `draw:regular-polygon`.
    Polygon,
    /// `draw:caption` or `draw:measure`.
    Annotation,
    /// `draw:control`.
    Control,
    /// `draw:g`.
    Group,
    /// `draw:page-thumbnail`.
    PageThumbnail,
    /// A `dr3d:scene`.
    Scene3d,
}

/// What a `draw:frame` holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FrameContent {
    /// Not a frame, or a frame without a recognised child.
    None,
    /// `draw:text-box`.
    TextBox,
    /// `draw:image`.
    Image,
    /// `table:table`.
    Table,
    /// `draw:object` or `draw:object-ole` (charts and OLE objects).
    Object,
    /// `draw:plugin` (audio and video).
    Plugin,
    /// Another frame child (`draw:applet`, `draw:floating-frame`).
    Other,
}

/// One drawing object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Object {
    pub(crate) kind: ObjectKind,
    pub(crate) content: FrameContent,
    pub(crate) range: Range<usize>,
    pub(crate) start_tag: Range<usize>,
    pub(crate) empty: bool,
    pub(crate) name: Option<String>,
    pub(crate) class: Option<String>,
    pub(crate) placeholder: Option<bool>,
    pub(crate) id: Option<String>,
    pub(crate) style_name: Option<String>,
    pub(crate) text_style_name: Option<String>,
    pub(crate) layer: Option<String>,
    pub(crate) z_index: Option<String>,
    pub(crate) x: Option<String>,
    pub(crate) y: Option<String>,
    pub(crate) width: Option<String>,
    pub(crate) height: Option<String>,
    pub(crate) transform: Option<String>,
    pub(crate) href: Option<String>,
    pub(crate) title: Option<String>,
    pub(crate) description: Option<String>,
    pub(crate) story: Option<Story>,
    pub(crate) table: Option<Table>,
    pub(crate) children: Vec<Object>,
}

impl Object {
    /// The drawing element.
    #[must_use]
    pub fn kind(&self) -> ObjectKind {
        self.kind
    }

    /// What a frame holds; [`FrameContent::None`] for other elements.
    #[must_use]
    pub fn content(&self) -> FrameContent {
        self.content
    }

    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// `draw:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// `presentation:class` (`title`, `outline`, `subtitle`, `notes`, ...).
    #[must_use]
    pub fn class(&self) -> Option<&str> {
        self.class.as_deref()
    }

    /// `presentation:placeholder`.
    #[must_use]
    pub fn placeholder(&self) -> Option<bool> {
        self.placeholder
    }

    /// `draw:id`, else `xml:id`.
    #[must_use]
    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
    }

    /// `draw:style-name`, else `presentation:style-name`.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// `draw:text-style-name`.
    #[must_use]
    pub fn text_style_name(&self) -> Option<&str> {
        self.text_style_name.as_deref()
    }

    /// `draw:layer`.
    #[must_use]
    pub fn layer(&self) -> Option<&str> {
        self.layer.as_deref()
    }

    /// The lexical `draw:z-index`.
    #[must_use]
    pub fn z_index(&self) -> Option<&str> {
        self.z_index.as_deref()
    }

    /// The lexical `svg:x`.
    #[must_use]
    pub fn x(&self) -> Option<&str> {
        self.x.as_deref()
    }

    /// The lexical `svg:y`.
    #[must_use]
    pub fn y(&self) -> Option<&str> {
        self.y.as_deref()
    }

    /// The lexical `svg:width`.
    #[must_use]
    pub fn width(&self) -> Option<&str> {
        self.width.as_deref()
    }

    /// The lexical `svg:height`.
    #[must_use]
    pub fn height(&self) -> Option<&str> {
        self.height.as_deref()
    }

    /// The lexical `draw:transform`.
    #[must_use]
    pub fn transform(&self) -> Option<&str> {
        self.transform.as_deref()
    }

    /// The `xlink:href` of a frame's image, object or plugin child (the
    /// image's when a frame holds several).
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// The `svg:title` alternative text.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// The `svg:desc` description.
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    /// The object's text story: a frame's `draw:text-box`, or the
    /// paragraphs a shape holds directly.
    #[must_use]
    pub fn story(&self) -> Option<&Story> {
        self.story.as_ref()
    }

    /// The table a frame holds.
    #[must_use]
    pub fn table(&self) -> Option<&Table> {
        self.table.as_ref()
    }

    /// A group's member objects, in document order.
    #[must_use]
    pub fn children(&self) -> &[Object] {
        &self.children
    }

    /// The object's text: its story's text, else empty.
    #[must_use]
    pub fn text(&self) -> String {
        self.story.as_ref().map(Story::text).unwrap_or_default()
    }
}

/// A sequence of paragraphs: a text box, a shape's text or a table cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Story {
    /// The content range of the container (`draw:text-box` or table cell),
    /// or, for a shape's direct paragraphs, from the first paragraph-level
    /// child to the end of the last.
    pub(crate) range: Range<usize>,
    pub(crate) paragraphs: Vec<Paragraph>,
    /// Whether every child of the container is a paragraph, heading or
    /// list, so the container's whole content may be replaced.
    pub(crate) replaceable: bool,
    /// Where a paragraph goes in an empty story: the container's content
    /// start, or `usize::MAX` when the container is a self-closing element.
    pub(crate) insertion: usize,
    /// The self-closing container element (`<draw:text-box/>`,
    /// `<table:table-cell/>`), which a write expands.
    pub(crate) empty: Option<Range<usize>>,
}

impl Story {
    /// The bytes the story's paragraphs occupy.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// The paragraphs, in document order, including those inside lists.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// The paragraphs' text joined by `\n`.
    #[must_use]
    pub fn text(&self) -> String {
        let mut text = String::new();
        for (index, paragraph) in self.paragraphs.iter().enumerate() {
            if index > 0 {
                text.push('\n');
            }
            text.push_str(&paragraph.text);
        }
        text
    }
}

/// A `text:p` or `text:h`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Paragraph {
    pub(crate) range: Range<usize>,
    pub(crate) start_tag: Range<usize>,
    pub(crate) content: Range<usize>,
    pub(crate) empty: bool,
    pub(crate) heading: bool,
    pub(crate) style_name: Option<String>,
    pub(crate) level: u8,
    pub(crate) text: String,
    pub(crate) runs: Vec<Run>,
    /// The start tags of the enclosing list wrappers inside the story,
    /// outermost first, with their qualified names.
    pub(crate) wrappers: Vec<Wrapper>,
    /// The largest enclosing range that holds this paragraph alone: the
    /// paragraph itself or a list item or list with no other child.
    pub(crate) removal: Range<usize>,
    /// The innermost enclosing `text:list-item`, when there is one.
    pub(crate) item: Option<Range<usize>>,
    /// The start offsets and names of the inline content that is not text
    /// (frames, notes, annotations, fields, hyperlinks), at most eight:
    /// a text set that would replace one is refused (record 0304).
    pub(crate) objects: Vec<(usize, String)>,
}

/// One enclosing list wrapper of a paragraph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Wrapper {
    pub(crate) start_tag: Range<usize>,
    pub(crate) qname: String,
    pub(crate) item: bool,
}

impl Paragraph {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// Whether this is a `text:h` heading.
    #[must_use]
    pub fn is_heading(&self) -> bool {
        self.heading
    }

    /// `text:style-name`.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// The list nesting depth: 0 outside lists, 1 in a top-level list.
    #[must_use]
    pub fn level(&self) -> u8 {
        self.level
    }

    /// The projected text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The runs, in document order.
    #[must_use]
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }
}

/// What a [`Run`] is in its paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RunKind {
    /// Character data directly in the paragraph, with any `text:s`,
    /// `text:tab` and `text:line-break` between it.
    Text,
    /// A `text:span`.
    Span,
    /// A `text:a` hyperlink.
    Link,
    /// Another inline element: a field, bookmark or reference mark. A
    /// `text:span` whose whole content is one field with text is that
    /// field, formatted by the span's style (record 0316).
    Field,
}

/// One run of a paragraph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Run {
    pub(crate) kind: RunKind,
    pub(crate) range: Range<usize>,
    pub(crate) content: Range<usize>,
    /// The start tag of the run's element; empty for a plain-text run.
    pub(crate) start_tag: Range<usize>,
    pub(crate) style_name: Option<String>,
    pub(crate) href: Option<String>,
    pub(crate) text: String,
    /// The field element of a field run its formatting span wraps.
    pub(crate) field: Option<WrappedField>,
    /// The start tag and style of the one span a link's content is.
    pub(crate) inner_span: Option<(Range<usize>, Option<String>)>,
}

/// The field element inside the span of a formatted field run.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WrappedField {
    pub(crate) range: Range<usize>,
    pub(crate) start_tag: Range<usize>,
    pub(crate) content: Range<usize>,
}

impl Run {
    /// What the run is.
    #[must_use]
    pub fn kind(&self) -> RunKind {
        self.kind
    }

    /// The bytes of the run.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// The `text:style-name` of a span, of a link (its link character
    /// style) or of the span that formats a field run.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// The `text:style-name` of the one `text:span` a link's whole content
    /// is: the formatting of the link text (record 0316).
    #[must_use]
    pub fn link_text_style_name(&self) -> Option<&str> {
        self.inner_span.as_ref()?.1.as_deref()
    }

    /// A link's `xlink:href`.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// The projected text, collapsed within the run.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A `table:table` held by a frame.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    pub(crate) range: Range<usize>,
    pub(crate) name: Option<String>,
    pub(crate) style_name: Option<String>,
    pub(crate) columns: usize,
    pub(crate) rows: Vec<Row>,
}

impl Table {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// `table:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// `table:style-name`.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// The declared column count (`table:table-column` repeats included).
    #[must_use]
    pub fn columns(&self) -> usize {
        self.columns
    }

    /// The rows, header rows included, in document order.
    #[must_use]
    pub fn rows(&self) -> &[Row] {
        &self.rows
    }
}

/// A `table:table-row`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Row {
    pub(crate) range: Range<usize>,
    pub(crate) cells: Vec<Cell>,
}

impl Row {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// The cells, covered cells included.
    #[must_use]
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }
}

/// A `table:table-cell` or `table:covered-table-cell`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cell {
    pub(crate) range: Range<usize>,
    pub(crate) covered: bool,
    pub(crate) column_span: u32,
    pub(crate) row_span: u32,
    pub(crate) story: Story,
}

impl Cell {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// Whether this is a `table:covered-table-cell`.
    #[must_use]
    pub fn is_covered(&self) -> bool {
        self.covered
    }

    /// `table:number-columns-spanned`, 1 by default.
    #[must_use]
    pub fn column_span(&self) -> u32 {
        self.column_span
    }

    /// `table:number-rows-spanned`, 1 by default.
    #[must_use]
    pub fn row_span(&self) -> u32 {
        self.row_span
    }

    /// The cell's paragraphs.
    #[must_use]
    pub fn story(&self) -> &Story {
        &self.story
    }

    /// The cell text.
    #[must_use]
    pub fn text(&self) -> String {
        self.story.text()
    }
}
