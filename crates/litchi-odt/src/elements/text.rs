//! Text-related ODF elements.
//!
//! This module provides classes for text elements like paragraphs, spans,
//! headings, and other text content elements.

mod codec;
mod model;
mod validation;

use super::element::{Element, ElementBase};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub use codec::Elements;
pub use model::{Block, Kind, LinkActuate, LinkShow};

/// Compatibility facade for the historic text-element collection name.
pub type TextElements = Elements;

/// Internal name retained for the decoder's block slots.
pub(crate) type TextBlock = Block;

const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_TEXT_BLOCKS: usize = 1_000_000;
const MAX_TEXT_DEPTH: usize = 4_096;
const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_SPACE_COUNT: usize = 1_000_000;
/// Qualified name of the unnumbered leading block of an ODF list.
const LIST_HEADER_TAG: &str = "text:list-header";

/// A text paragraph element
#[derive(Debug, Clone)]
pub struct Paragraph {
    element: Element,
}

impl Default for Paragraph {
    fn default() -> Self {
        Self::new()
    }
}

impl Paragraph {
    /// Create a new paragraph
    pub fn new() -> Self {
        Self {
            element: Element::new("text:p"),
        }
    }

    /// Append a validated, inert dynamic text field to this paragraph.
    pub fn add_dynamic_text_field(
        &mut self,
        field: &crate::elements::field::DynamicTextField,
    ) -> Result<()> {
        self.element.add_child(field.to_element()?);
        Ok(())
    }

    /// Create paragraph from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:p" {
            return Err(Error::InvalidFormat(
                "Element is not a paragraph".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Wrap an element as a paragraph without a tag check; used when
    /// converting a `text:numbered-paragraph`.
    pub(crate) fn from_element_unchecked(element: Element) -> Self {
        Self { element }
    }

    /// Get the text content of the paragraph
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the paragraph
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get all text spans within this paragraph
    pub fn spans(&self) -> Result<Vec<Span>> {
        let mut spans = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:span"
                && let Ok(span) = Span::from_element(child.clone())
            {
                spans.push(span);
            }
        }
        Ok(spans)
    }

    /// Get all runs (text spans) within this paragraph.
    ///
    /// This is an alias for `spans()` to match the unified document API.
    pub fn runs(&self) -> Result<Vec<Span>> {
        self.spans()
    }

    /// Add a text span to this paragraph
    pub fn add_span(&mut self, span: Span) {
        self.element.add_child(span.element);
    }

    /// Append a validated `text:a` hyperlink to this paragraph.
    ///
    /// Hyperlinks are inline content, so this method can be combined with
    /// spans and other supported paragraph children to form rich text.
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) -> Result<()> {
        hyperlink.validate()?;
        self.element.add_child(hyperlink.element);
        Ok(())
    }

    /// Return direct `text:a` hyperlink children in document order.
    pub fn hyperlinks(&self) -> Result<Vec<Hyperlink>> {
        self.element
            .children
            .iter()
            .filter(|child| child.tag_name() == "text:a")
            .cloned()
            .map(Hyperlink::from_element)
            .collect()
    }

    /// Check if this paragraph is a heading
    pub fn is_heading(&self) -> bool {
        false // Paragraphs are not headings
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }
}

/// A `text:numbered-paragraph`: a paragraph with explicit list numbering.
///
/// The numbering attributes are inert metadata: numbers are never computed,
/// list styles are never resolved, and no list structure is altered.
#[derive(Clone, Debug)]
pub struct NumberedParagraph {
    element: Element,
}

impl Default for NumberedParagraph {
    fn default() -> Self {
        Self::new()
    }
}

impl NumberedParagraph {
    /// Create a new numbered paragraph.
    pub fn new() -> Self {
        Self {
            element: Element::new("text:numbered-paragraph"),
        }
    }

    /// Create from an existing element, validating the tag.
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:numbered-paragraph" {
            return Err(Error::InvalidFormat(
                "Element is not a numbered paragraph".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the text content of the paragraph.
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the paragraph.
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the underlying element.
    pub fn element(&self) -> &Element {
        &self.element
    }

    /// The `text:style-name` of the paragraph style.
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// The `text:level` nesting level of the paragraph.
    pub fn level(&self) -> Option<Result<u32>> {
        self.element.get_attribute("text:level").map(|value| {
            value.parse::<u32>().map_err(|_error| {
                Error::InvalidFormat("text:level is not a non-negative integer".to_string())
            })
        })
    }

    /// The `text:list-id` identifying the list the paragraph belongs to.
    pub fn list_id(&self) -> Option<&str> {
        self.element.get_attribute("text:list-id")
    }

    /// The `text:start-value` restarting numbering at this paragraph.
    pub fn start_value(&self) -> Option<Result<i32>> {
        self.element.get_attribute("text:start-value").map(|value| {
            value.parse::<i32>().map_err(|_error| {
                Error::InvalidFormat("text:start-value is not an integer".to_string())
            })
        })
    }

    /// Convert into a plain paragraph view of the same content.
    pub fn into_paragraph(self) -> Paragraph {
        Paragraph::from_element_unchecked(self.element)
    }
}

impl From<NumberedParagraph> for Element {
    fn from(para: NumberedParagraph) -> Element {
        para.element
    }
}

impl From<Paragraph> for Element {
    fn from(para: Paragraph) -> Element {
        para.element
    }
}

/// A text span element (formatted text within a paragraph)
#[derive(Debug, Clone)]
pub struct Span {
    element: Element,
}

impl Default for Span {
    fn default() -> Self {
        Self::new()
    }
}

impl Span {
    /// Create a new span
    pub fn new() -> Self {
        Self {
            element: Element::new("text:span"),
        }
    }

    /// Create span from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:span" {
            return Err(Error::InvalidFormat("Element is not a span".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the text content of the span
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the span
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }

    /// Check if the text is bold.
    ///
    /// Returns `None` if the style doesn't specify bold formatting.
    pub fn bold(&self) -> Option<bool> {
        // In ODF, formatting is typically in styles, not directly on elements
        // For now, return None to indicate formatting should be resolved via styles
        None
    }

    /// Check if the text is italic.
    ///
    /// Returns `None` if the style doesn't specify italic formatting.
    pub fn italic(&self) -> Option<bool> {
        // In ODF, formatting is typically in styles, not directly on elements
        None
    }

    /// Check if the text has strikethrough.
    ///
    /// Returns `None` if the style doesn't specify strikethrough formatting.
    pub fn strikethrough(&self) -> Option<bool> {
        // In ODF, formatting is typically in styles, not directly on elements
        None
    }

    /// Get the vertical position (superscript/subscript).
    ///
    /// Returns `None` if the text is in normal position.
    pub fn vertical_position(&self) -> Option<litchi_core::style::text::pos::VerticalPosition> {
        // In ODF, vertical position is typically in styles
        None
    }
}

impl From<Span> for Element {
    fn from(span: Span) -> Element {
        span.element
    }
}

/// A hyperlink element (text:a)
#[derive(Debug, Clone)]
pub struct Hyperlink {
    element: Element,
}

/// Historic prefixed alias retained for callers that already use the ODT
/// facade.  New code should use [`LinkShow`] inside this text context.
pub type TextHyperlinkShow = LinkShow;

/// Historic prefixed alias retained for callers that already use the ODT
/// facade.  New code should use [`LinkActuate`] inside this text context.
pub type TextHyperlinkActuate = LinkActuate;

impl Default for Hyperlink {
    fn default() -> Self {
        Self::new()
    }
}

impl Hyperlink {
    /// Create a new `text:a` hyperlink with its `XLink` type set to `simple`.
    ///
    /// Set a target with [`Self::set_href`] before inserting the hyperlink,
    /// or use [`Self::with_href`] to create a fully valid link in one step.
    pub fn new() -> Self {
        let mut element = Element::new("text:a");
        element.set_attribute("xlink:type", "simple");
        Self { element }
    }

    /// Create a validated simple hyperlink with visible text.
    pub fn with_href(href: impl AsRef<str>, text: impl AsRef<str>) -> Result<Self> {
        let href = href.as_ref();
        validation::href(href)?;
        let mut hyperlink = Self::new();
        hyperlink.set_href(href);
        hyperlink.set_text(text.as_ref());
        Ok(hyperlink)
    }

    /// Create a hyperlink wrapper from an existing `text:a` element.
    ///
    /// This preserves parsed attributes without validating them. Call
    /// [`Self::validate`] before reusing the element for authoring.
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:a" {
            return Err(Error::InvalidFormat(
                "Element is not a hyperlink".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the hyperlink URL
    pub fn href(&self) -> Option<&str> {
        self.element.get_attribute("xlink:href")
    }

    /// Set the hyperlink URL
    pub fn set_href(&mut self, href: &str) {
        self.element.set_attribute("xlink:href", href);
    }

    /// Get the optional hyperlink name (`office:name`).
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("office:name")
    }

    /// Set the optional hyperlink name (`office:name`).
    pub fn set_name(&mut self, name: &str) {
        self.element.set_attribute("office:name", name);
    }

    /// Get the optional hyperlink title (`office:title`).
    pub fn title(&self) -> Option<&str> {
        self.element.get_attribute("office:title")
    }

    /// Set the optional hyperlink title (`office:title`).
    pub fn set_title(&mut self, title: &str) {
        self.element.set_attribute("office:title", title);
    }

    /// Get the optional target frame (`office:target-frame-name`).
    pub fn target_frame_name(&self) -> Option<&str> {
        self.element.get_attribute("office:target-frame-name")
    }

    /// Set the optional target frame (`office:target-frame-name`).
    pub fn set_target_frame_name(&mut self, target_frame_name: &str) {
        self.element
            .set_attribute("office:target-frame-name", target_frame_name);
    }

    /// Get the link text content
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the link text content
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the link type (simple, locator, etc.)
    pub fn link_type(&self) -> Option<&str> {
        self.element.get_attribute("xlink:type")
    }

    /// Get the optional `XLink` display behavior (`xlink:show`).
    ///
    /// Returns `None` when the attribute is absent or malformed; use
    /// [`Self::validate`] to distinguish those cases.
    pub fn show(&self) -> Option<LinkShow> {
        self.element
            .get_attribute("xlink:show")
            .and_then(LinkShow::parse)
    }

    /// Set or omit the `XLink` display behavior (`xlink:show`).
    pub fn set_show(&mut self, show: Option<LinkShow>) {
        match show {
            Some(show) => self.element.set_attribute("xlink:show", show.as_str()),
            None => self.element.remove_attribute("xlink:show"),
        }
    }

    /// Get the optional explicit `XLink` activation behavior (`xlink:actuate`).
    ///
    /// Returns `None` when the attribute is absent or malformed; use
    /// [`Self::validate`] to distinguish those cases.
    pub fn actuate(&self) -> Option<LinkActuate> {
        self.element
            .get_attribute("xlink:actuate")
            .and_then(LinkActuate::parse)
    }

    /// Set or omit the explicit `XLink` activation behavior (`xlink:actuate`).
    pub fn set_actuate(&mut self, actuate: Option<LinkActuate>) {
        match actuate {
            Some(actuate) => self
                .element
                .set_attribute("xlink:actuate", actuate.as_str()),
            None => self.element.remove_attribute("xlink:actuate"),
        }
    }

    /// Get the unvisited link style name (`text:style-name`).
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the unvisited link style name (`text:style-name`).
    pub fn set_style_name(&mut self, style_name: &str) {
        self.element.set_attribute("text:style-name", style_name);
    }

    /// Get the visited style name
    pub fn visited_style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:visited-style-name")
    }

    /// Set the visited link style name (`text:visited-style-name`).
    pub fn set_visited_style_name(&mut self, style_name: &str) {
        self.element
            .set_attribute("text:visited-style-name", style_name);
    }

    /// Validate attributes required for safe authoring of an ODF `text:a`.
    pub fn validate(&self) -> Result<()> {
        validation::hyperlink(&self.element)
    }
}

impl From<Hyperlink> for Element {
    fn from(link: Hyperlink) -> Element {
        link.element
    }
}

/// A bookmark element (text:bookmark)
#[derive(Debug, Clone)]
pub struct Bookmark {
    element: Element,
}

impl Bookmark {
    /// Create a new bookmark
    pub fn new(name: &str) -> Self {
        let mut element = Element::new("text:bookmark");
        element.set_attribute("text:name", name);
        Self { element }
    }

    /// Create bookmark from element
    pub fn from_element(element: Element) -> Result<Self> {
        let tag = element.tag_name();
        if tag != "text:bookmark" && tag != "text:bookmark-start" && tag != "text:bookmark-end" {
            return Err(Error::InvalidFormat(
                "Element is not a bookmark".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the bookmark name
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("text:name")
    }

    /// Set the bookmark name
    pub fn set_name(&mut self, name: &str) {
        self.element.set_attribute("text:name", name);
    }

    /// Check if this is a bookmark-start element
    pub fn is_start(&self) -> bool {
        self.element.tag_name() == "text:bookmark-start"
    }

    /// Check if this is a bookmark-end element
    pub fn is_end(&self) -> bool {
        self.element.tag_name() == "text:bookmark-end"
    }
}

impl From<Bookmark> for Element {
    fn from(bookmark: Bookmark) -> Element {
        bookmark.element
    }
}

/// A heading element
#[derive(Debug, Clone)]
pub struct Heading {
    element: Element,
}

impl Heading {
    /// Create a new heading
    pub fn new(level: u8) -> Self {
        let mut element = Element::new("text:h");
        element.set_attribute("text:outline-level", &level.to_string());
        Self { element }
    }

    /// Create heading from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:h" {
            return Err(Error::InvalidFormat("Element is not a heading".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the text content of the heading
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the heading
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the outline level
    pub fn level(&self) -> Option<u8> {
        self.element
            .get_int_attribute("text:outline-level")
            .and_then(|value| u8::try_from(value).ok())
    }

    /// Set the outline level
    pub fn set_level(&mut self, level: u8) {
        self.element
            .set_attribute("text:outline-level", &level.to_string());
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }

    /// Check if this is a heading
    pub fn is_heading(&self) -> bool {
        true
    }
}

impl From<Heading> for Element {
    fn from(heading: Heading) -> Element {
        heading.element
    }
}

/// A text list element
#[derive(Debug, Clone)]
pub struct List {
    element: Element,
}

impl Default for List {
    fn default() -> Self {
        Self::new()
    }
}

impl List {
    /// Create a new list
    pub fn new() -> Self {
        Self {
            element: Element::new("text:list"),
        }
    }

    /// Create list from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:list" {
            return Err(Error::InvalidFormat("Element is not a list".to_string()));
        }
        Ok(Self { element })
    }

    /// Get list items
    pub fn items(&self) -> Result<Vec<ListItem>> {
        let mut items = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:list-item"
                && let Ok(item) = ListItem::from_element(child.clone())
            {
                items.push(item);
            }
        }
        Ok(items)
    }

    /// Get the optional unnumbered header block of the list.
    ///
    /// ODF allows a single `text:list-header` before the first `text:list-item`.
    /// It holds ordinary paragraphs that receive no list label, and is content
    /// the author typed just like any item, so it is reported separately from
    /// [`items`](Self::items) rather than dropped.
    pub fn header(&self) -> Result<Option<ListHeader>> {
        for child in &self.element.children {
            if child.tag_name() == LIST_HEADER_TAG {
                return ListHeader::from_element(child.clone()).map(Some);
            }
        }
        Ok(None)
    }

    /// Add a list item
    pub fn add_item(&mut self, item: ListItem) {
        self.element.add_child(item.element);
    }

    /// Set the unnumbered header block of the list, replacing any existing one.
    pub fn set_header(&mut self, header: ListHeader) {
        self.element
            .children
            .retain(|child| child.tag_name() != LIST_HEADER_TAG);
        self.element.children.insert(0, header.element);
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }
}

impl From<List> for Element {
    fn from(list: List) -> Element {
        list.element
    }
}

/// The unnumbered leading block of a list (`text:list-header`).
///
/// A list header holds the same paragraph content a list item does, but the
/// list-level style gives it no number or bullet. ODF permits at most one, and
/// it must precede every `text:list-item`.
#[derive(Debug, Clone)]
pub struct ListHeader {
    element: Element,
}

impl Default for ListHeader {
    fn default() -> Self {
        Self::new()
    }
}

impl ListHeader {
    /// Create an empty list header.
    pub fn new() -> Self {
        Self {
            element: Element::new(LIST_HEADER_TAG),
        }
    }

    /// Wrap a parsed `text:list-header` element.
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != LIST_HEADER_TAG {
            return Err(Error::InvalidFormat(
                "Element is not a list header".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the flattened text content of the header.
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the header.
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the paragraphs the header contains.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut paragraphs = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:p"
                && let Ok(paragraph) = Paragraph::from_element(child.clone())
            {
                paragraphs.push(paragraph);
            }
        }
        Ok(paragraphs)
    }

    /// Add a paragraph to the header.
    pub fn add_paragraph(&mut self, paragraph: Paragraph) {
        self.element.add_child(paragraph.into());
    }
}

impl From<ListHeader> for Element {
    fn from(header: ListHeader) -> Element {
        header.element
    }
}

/// A list item element
#[derive(Debug, Clone)]
pub struct ListItem {
    element: Element,
}

impl Default for ListItem {
    fn default() -> Self {
        Self::new()
    }
}

impl ListItem {
    /// Create a new list item
    pub fn new() -> Self {
        Self {
            element: Element::new("text:list-item"),
        }
    }

    /// Create list item from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:list-item" {
            return Err(Error::InvalidFormat(
                "Element is not a list item".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the text content of the list item
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the list item
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get nested paragraphs
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut paragraphs = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:p"
                && let Ok(para) = Paragraph::from_element(child.clone())
            {
                paragraphs.push(para);
            }
        }
        Ok(paragraphs)
    }

    /// Add a paragraph to this list item
    pub fn add_paragraph(&mut self, paragraph: Paragraph) {
        self.element.add_child(paragraph.element);
    }
}

impl From<ListItem> for Element {
    fn from(item: ListItem) -> Element {
        item.element
    }
}

/// A page break element
#[derive(Debug, Clone)]
pub struct PageBreak {
    element: Element,
}

impl Default for PageBreak {
    fn default() -> Self {
        Self::new()
    }
}

impl PageBreak {
    /// Create a new page break
    pub fn new() -> Self {
        let mut element = Element::new("text:p");
        element.set_attribute("text:style-name", "PageBreak");
        Self { element }
    }

    /// Create page break from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:p" {
            return Err(Error::InvalidFormat(
                "Element is not a page break".to_string(),
            ));
        }
        if element.get_attribute("text:style-name") != Some("PageBreak") {
            return Err(Error::InvalidFormat(
                "Element is not a page break".to_string(),
            ));
        }
        Ok(Self { element })
    }
}

impl From<PageBreak> for Element {
    fn from(pb: PageBreak) -> Element {
        pb.element
    }
}

struct ActiveTextBlock {
    element: Element,
    depth: usize,
    text: String,
    /// Index of the reserved output slot, so completed blocks keep the order in
    /// which they *started* rather than the order in which they closed.
    slot: usize,
}

/// Parse every `text:p` and `text:h` in `xml_content` into flat text blocks.
///
/// ODF allows a paragraph to contain further paragraphs through frames
/// (`draw:frame`/`draw:text-box`, `draw:custom-shape`), inline annotations
/// (`office:annotation`), sections, and tables nested inside those frames.
/// Each nested paragraph becomes its own block, emitted in document (start)
/// order so an enclosing paragraph precedes the frame content it carries.
///
/// Stored tracked-change definitions, note bodies, and ruby pronunciation runs
/// are excluded from the visible flow; they are exposed through their own
/// dedicated readers instead.
pub(crate) fn parse_text_blocks(xml_content: &str) -> Result<Vec<TextBlock>> {
    let mut reader = NsReader::from_str(xml_content);
    let mut buffer = Vec::new();
    let mut blocks: Vec<Option<TextBlock>> = Vec::new();
    let mut active: Vec<ActiveTextBlock> = Vec::new();
    let mut document_depth = 0usize;
    let mut tracked_changes_depth = 0usize;
    // Depth of the note-body/ruby-text subtree whose content is suppressed.
    let mut skipped_depth = 0usize;
    let mut total_text_bytes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF text XML: {error}")))?;
        let text_namespace =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = document_depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF text nesting depth overflow".to_string())
                })?;
                if document_depth > MAX_TEXT_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODF text nesting exceeds {MAX_TEXT_DEPTH} levels"
                    )));
                }

                if tracked_changes_depth > 0 {
                    tracked_changes_depth += 1;
                    buffer.clear();
                    continue;
                }
                if is_text_element(text_namespace, element, b"tracked-changes") {
                    tracked_changes_depth = 1;
                    buffer.clear();
                    continue;
                }

                // A nested block owns its whole subtree, so the element that
                // opens it is counted once — against the new block, never also
                // against the block that encloses it. Every other element counts
                // against the innermost open block so that block still closes on
                // the right end tag.
                let starts_block = skipped_depth == 0 && is_text_block(text_namespace, element);
                if !starts_block && let Some(current) = active.last_mut() {
                    current.depth += 1;
                }

                if starts_block {
                    active.push(ActiveTextBlock {
                        element: make_text_block_element(&reader, element)?,
                        depth: 1,
                        text: String::new(),
                        slot: reserve_text_block(&mut blocks)?,
                    });
                } else if skipped_depth > 0 {
                    skipped_depth += 1;
                } else if let Some(current) = active.last_mut() {
                    if is_text_element(text_namespace, element, b"note-body")
                        || is_text_element(text_namespace, element, b"ruby-text")
                    {
                        skipped_depth = 1;
                    } else {
                        append_text_control(&reader, text_namespace, element, &mut current.text)?;
                    }
                }
            },
            Event::Empty(ref element) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if is_text_element(text_namespace, element, b"note-body")
                    || is_text_element(text_namespace, element, b"ruby-text")
                {
                    // An empty suppressed run contributes nothing either way.
                } else if is_text_block(text_namespace, element) {
                    let slot = reserve_text_block(&mut blocks)?;
                    store_text_block(
                        make_text_block_element(&reader, element)?,
                        String::new(),
                        slot,
                        &mut blocks,
                        &mut total_text_bytes,
                    )?;
                } else if let Some(current) = active.last_mut() {
                    append_text_control(&reader, text_namespace, element, &mut current.text)?;
                }
            },
            Event::Text(ref value) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if let Some(current) = active.last_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODF text content: {error}"))
                        })?;
                    append_checked(&mut current.text, &decoded)?;
                }
            },
            Event::CData(ref value) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if let Some(current) = active.last_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODF text CDATA: {error}"))
                        })?;
                    append_checked(&mut current.text, &decoded)?;
                }
            },
            Event::GeneralRef(ref reference)
                if tracked_changes_depth == 0 && skipped_depth == 0 =>
            {
                if let Some(current) = active.last_mut() {
                    let decoded = decode_reference(reference)?;
                    append_checked(&mut current.text, &decoded)?;
                }
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF text element stack underflow".to_string())
                })?;
                if tracked_changes_depth > 0 {
                    tracked_changes_depth -= 1;
                    buffer.clear();
                    continue;
                }
                skipped_depth = skipped_depth.saturating_sub(1);
                if let Some(current) = active.last_mut() {
                    current.depth = current.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("ODF text block stack underflow".to_string())
                    })?;
                    if current.depth == 0 {
                        let current = active
                            .pop()
                            .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))?;
                        store_text_block(
                            current.element,
                            current.text,
                            current.slot,
                            &mut blocks,
                            &mut total_text_bytes,
                        )?;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if !active.is_empty() || tracked_changes_depth != 0 || skipped_depth != 0 || document_depth != 0
    {
        return Err(Error::InvalidFormat(
            "incomplete ODF text XML structure".to_string(),
        ));
    }
    Ok(blocks.into_iter().flatten().collect())
}

fn is_text_element(text_namespace: bool, element: &BytesStart<'_>, local_name: &[u8]) -> bool {
    text_namespace && element.local_name().as_ref() == local_name
}

fn is_text_block(text_namespace: bool, element: &BytesStart<'_>) -> bool {
    matches!(element.local_name().as_ref(), b"p" | b"h") && text_namespace
}

fn make_text_block_element(reader: &NsReader<&[u8]>, source: &BytesStart<'_>) -> Result<Element> {
    let tag_name = match source.local_name().as_ref() {
        b"p" => "text:p",
        b"h" => "text:h",
        _ => {
            return Err(Error::InvalidFormat(
                "element is not an ODF paragraph or heading".to_string(),
            ));
        },
    };
    let mut element = Element::new(tag_name);
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODF text attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let local_name = std::str::from_utf8(local_name.as_ref()).map_err(|_error| {
            Error::InvalidFormat("non-UTF-8 ODF text attribute name".to_string())
        })?;
        let name = match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE => {
                format!("text:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => {
                format!("xlink:{local_name}")
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XML_NAMESPACE => {
                format!("xml:{local_name}")
            },
            ResolveResult::Bound(_) | ResolveResult::Unbound => {
                std::str::from_utf8(attribute.key.as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat("non-UTF-8 ODF text attribute name".to_string())
                    })?
                    .to_string()
            },
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unknown ODF text attribute namespace prefix '{}'",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODF text attribute value: {error}"))
            })?;
        if element.has_attribute(&name) {
            return Err(Error::InvalidFormat(format!(
                "duplicate ODF text attribute '{name}'"
            )));
        }
        element.set_attribute(&name, &value);
    }
    Ok(element)
}

fn append_text_control(
    reader: &NsReader<&[u8]>,
    text_namespace: bool,
    element: &BytesStart<'_>,
    output: &mut String,
) -> Result<()> {
    if !text_namespace {
        return Ok(());
    }
    match element.local_name().as_ref() {
        b"s" => {
            let count = text_space_count(reader, element)?.unwrap_or(1);
            if count > MAX_SPACE_COUNT {
                return Err(Error::InvalidFormat(format!(
                    "text:s count exceeds {MAX_SPACE_COUNT}"
                )));
            }
            let new_len = output
                .len()
                .checked_add(count)
                .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
            if new_len > MAX_TEXT_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "ODF text exceeds {MAX_TEXT_BYTES} bytes"
                )));
            }
            output.extend(std::iter::repeat_n(' ', count));
        },
        b"tab" => append_checked(output, "\t")?,
        b"line-break" => append_checked(output, "\n")?,
        _ => {},
    }
    Ok(())
}

fn text_space_count(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<usize>> {
    let mut count = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid text:s attribute: {error}")))?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE)
            && local_name.as_ref() == b"c"
        {
            if count.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate expanded text:c attribute".to_string(),
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid text:c attribute: {error}"))
                })?;
            let value = value.parse().map_err(|_error| {
                Error::InvalidFormat("text:c must be a non-negative integer".to_string())
            })?;
            count = Some(value);
        }
    }
    Ok(count)
}

fn append_checked(output: &mut String, value: &str) -> Result<()> {
    let new_len = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
    if new_len > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BYTES} bytes"
        )));
    }
    output.push_str(value);
    Ok(())
}

/// Message used when the open-block stack and the element stack disagree.
const BLOCK_STACK_ERROR: &str = "inconsistent ODF text block stack";

/// Reserve an output slot for a block that has just started, returning its
/// index. The slot keeps document order while the block is still being read.
fn reserve_text_block(blocks: &mut Vec<Option<TextBlock>>) -> Result<usize> {
    if blocks.len() >= MAX_TEXT_BLOCKS {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BLOCKS} paragraphs and headings"
        )));
    }
    let slot = blocks.len();
    blocks.push(None);
    Ok(slot)
}

/// Fill a previously reserved slot with the finished block, accounting the
/// collected text against the overall size budget.
fn store_text_block(
    mut element: Element,
    text: String,
    slot: usize,
    blocks: &mut [Option<TextBlock>],
    total_text_bytes: &mut usize,
) -> Result<()> {
    *total_text_bytes = total_text_bytes
        .checked_add(text.len())
        .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
    if *total_text_bytes > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BYTES} bytes"
        )));
    }
    element.set_text(&text);
    let block = match element.tag_name() {
        "text:p" => TextBlock::Paragraph(Paragraph::from_element(element)?),
        "text:h" => TextBlock::Heading(Heading::from_element(element)?),
        _ => {
            return Err(Error::InvalidFormat(
                "element is not an ODF paragraph or heading".to_string(),
            ));
        },
    };
    let target = blocks
        .get_mut(slot)
        .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))?;
    *target = Some(block);
    Ok(())
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODF text character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODF text entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported ODF text entity '&{name};'"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ========== Paragraph Tests ==========
    #[test]
    fn test_paragraph_new() {
        let para = Paragraph::new();
        assert_eq!(para.text().unwrap(), "");
        assert!(!para.is_heading());
    }

    #[test]
    fn test_paragraph_set_text() {
        let mut para = Paragraph::new();
        para.set_text("Hello World");
        assert_eq!(para.text().unwrap(), "Hello World");
    }

    #[test]
    fn test_paragraph_from_element() {
        let element = Element::new("text:p");
        let para = Paragraph::from_element(element).unwrap();
        assert_eq!(para.text().unwrap(), "");
    }

    #[test]
    fn test_paragraph_from_element_wrong_tag() {
        let element = Element::new("text:span");
        assert!(Paragraph::from_element(element).is_err());
    }

    #[test]
    fn test_paragraph_style_name() {
        let mut para = Paragraph::new();
        assert!(para.style_name().is_none());

        para.set_style_name("BodyText");
        assert_eq!(para.style_name(), Some("BodyText"));
    }

    #[test]
    fn test_paragraph_spans() {
        let mut para = Paragraph::new();
        let span1 = Span::new();
        let span2 = Span::new();

        para.add_span(span1);
        para.add_span(span2);

        let spans = para.spans().unwrap();
        assert_eq!(spans.len(), 2);
    }

    // ========== Span Tests ==========
    #[test]
    fn test_span_new() {
        let span = Span::new();
        assert_eq!(span.text().unwrap(), "");
    }

    #[test]
    fn test_span_set_text() {
        let mut span = Span::new();
        span.set_text("Hello");
        assert_eq!(span.text().unwrap(), "Hello");
    }

    #[test]
    fn test_span_style_name() {
        let mut span = Span::new();
        span.set_style_name("Bold");
        assert_eq!(span.style_name(), Some("Bold"));
    }

    #[test]
    fn test_span_formatting_none() {
        let span = Span::new();
        assert_eq!(span.bold(), None);
        assert_eq!(span.italic(), None);
        assert_eq!(span.strikethrough(), None);
        assert_eq!(span.vertical_position(), None);
    }

    // ========== Hyperlink Tests ==========
    #[test]
    fn test_hyperlink_new() {
        let link = Hyperlink::new();
        assert_eq!(link.text().unwrap(), "");
        assert_eq!(link.link_type(), Some("simple"));
    }

    #[test]
    fn test_hyperlink_href() {
        let mut link = Hyperlink::new();
        assert!(link.href().is_none());

        link.set_href("https://example.com");
        assert_eq!(link.href(), Some("https://example.com"));
    }

    #[test]
    fn test_hyperlink_text() {
        let mut link = Hyperlink::new();
        link.set_text("Click here");
        assert_eq!(link.text().unwrap(), "Click here");
    }

    #[test]
    fn hyperlink_metadata_and_paragraph_insertion_round_trip() {
        let mut link =
            Hyperlink::with_href("https://example.test/a?x=1&y=2", "Click here").unwrap();
        link.set_name("example-link");
        link.set_title("Example & more");
        link.set_target_frame_name("_blank");
        link.set_show(Some(TextHyperlinkShow::New));
        link.set_actuate(Some(TextHyperlinkActuate::OnRequest));
        link.set_style_name("Internet_20_link");
        link.set_visited_style_name("Visited_20_Internet_20_link");
        link.validate().unwrap();

        let mut paragraph = Paragraph::new();
        paragraph.add_hyperlink(link.clone()).unwrap();
        assert_eq!(paragraph.text().unwrap(), "Click here");
        let links = paragraph.hyperlinks().unwrap();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].href(), link.href());
        assert_eq!(links[0].name(), Some("example-link"));
        assert_eq!(links[0].title(), Some("Example & more"));
        assert_eq!(links[0].target_frame_name(), Some("_blank"));
        assert_eq!(links[0].show(), Some(TextHyperlinkShow::New));
        assert_eq!(links[0].actuate(), Some(TextHyperlinkActuate::OnRequest));
        assert_eq!(links[0].style_name(), Some("Internet_20_link"));
        assert_eq!(
            links[0].visited_style_name(),
            Some("Visited_20_Internet_20_link")
        );
    }

    #[test]
    fn hyperlink_insertion_rejects_missing_or_unsafe_targets() {
        assert!(Hyperlink::with_href("", "empty").is_err());
        assert!(Hyperlink::with_href("https://example.test/\nnext", "unsafe").is_err());

        let mut paragraph = Paragraph::new();
        assert!(paragraph.add_hyperlink(Hyperlink::new()).is_err());

        let mut missing_type = Element::new("text:a");
        missing_type.set_attribute("xlink:href", "https://example.test/");
        assert!(
            paragraph
                .add_hyperlink(Hyperlink::from_element(missing_type).unwrap())
                .is_err()
        );

        let mut invalid_type = Element::new("text:a");
        invalid_type.set_attribute("xlink:type", "extended");
        invalid_type.set_attribute("xlink:href", "https://example.test/");
        assert!(
            paragraph
                .add_hyperlink(Hyperlink::from_element(invalid_type).unwrap())
                .is_err()
        );

        let mut invalid_show = Element::new("text:a");
        invalid_show.set_attribute("xlink:type", "simple");
        invalid_show.set_attribute("xlink:href", "https://example.test/");
        invalid_show.set_attribute("xlink:show", "embed");
        assert!(
            paragraph
                .add_hyperlink(Hyperlink::from_element(invalid_show).unwrap())
                .is_err()
        );

        let mut invalid_metadata = Hyperlink::with_href("#bookmark", "Bookmark").unwrap();
        invalid_metadata.set_title("not\0valid");
        assert!(paragraph.add_hyperlink(invalid_metadata).is_err());
    }

    // ========== Bookmark Tests ==========
    #[test]
    fn test_bookmark_new() {
        let bookmark = Bookmark::new("section1");
        assert_eq!(bookmark.name(), Some("section1"));
        assert!(!bookmark.is_start());
        assert!(!bookmark.is_end());
    }

    #[test]
    fn test_bookmark_set_name() {
        let mut bookmark = Bookmark::new("old");
        bookmark.set_name("new");
        assert_eq!(bookmark.name(), Some("new"));
    }

    #[test]
    fn test_bookmark_from_element() {
        let mut element = Element::new("text:bookmark-start");
        element.set_attribute("text:name", "start");

        let bookmark = Bookmark::from_element(element).unwrap();
        assert!(bookmark.is_start());
    }

    // ========== Heading Tests ==========
    #[test]
    fn test_heading_new() {
        let heading = Heading::new(1);
        assert_eq!(heading.level(), Some(1));
        assert!(heading.is_heading());
    }

    #[test]
    fn test_heading_set_level() {
        let mut heading = Heading::new(1);
        heading.set_level(2);
        assert_eq!(heading.level(), Some(2));
    }

    #[test]
    fn test_heading_text() {
        let mut heading = Heading::new(1);
        heading.set_text("Title");
        assert_eq!(heading.text().unwrap(), "Title");
    }

    #[test]
    fn test_heading_style_name() {
        let mut heading = Heading::new(1);
        heading.set_style_name("Heading1");
        assert_eq!(heading.style_name(), Some("Heading1"));
    }

    // ========== List Tests ==========
    #[test]
    fn test_list_new() {
        let list = List::new();
        assert!(list.items().unwrap().is_empty());
    }

    #[test]
    fn test_list_add_item() {
        let mut list = List::new();
        let item = ListItem::new();
        list.add_item(item);

        assert_eq!(list.items().unwrap().len(), 1);
    }

    #[test]
    fn test_list_style_name() {
        let mut list = List::new();
        list.set_style_name("BulletList");
        assert_eq!(list.style_name(), Some("BulletList"));
    }

    // ========== ListItem Tests ==========
    #[test]
    fn test_list_item_new() {
        let item = ListItem::new();
        assert_eq!(item.text().unwrap(), "");
    }

    #[test]
    fn test_list_item_paragraphs() {
        let mut item = ListItem::new();
        let para = Paragraph::new();
        item.add_paragraph(para);

        assert_eq!(item.paragraphs().unwrap().len(), 1);
    }

    // ========== PageBreak Tests ==========
    #[test]
    fn test_page_break_new() {
        let pb = PageBreak::new();
        assert_eq!(
            pb.element.get_attribute("text:style-name"),
            Some("PageBreak")
        );
    }

    // ========== TextElements Tests ==========
    #[test]
    fn test_text_elements_parse_paragraphs() {
        let xml = r#"<office:text xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <text:p>First paragraph</text:p>
            <text:p>Second paragraph</text:p>
        </office:text>"#;

        let paragraphs = TextElements::parse_paragraphs(xml).unwrap();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].text().unwrap(), "First paragraph");
        assert_eq!(paragraphs[1].text().unwrap(), "Second paragraph");
    }

    #[test]
    fn test_text_elements_parse_headings() {
        let xml = r#"<office:text xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <text:h text:outline-level="1">Heading 1</text:h>
            <text:p>Paragraph</text:p>
            <text:h text:outline-level="2">Heading 2</text:h>
        </office:text>"#;

        let headings = TextElements::parse_headings(xml).unwrap();
        assert_eq!(headings.len(), 2);
        assert_eq!(headings[0].level(), Some(1));
        assert_eq!(headings[0].text().unwrap(), "Heading 1");
        assert_eq!(headings[1].level(), Some(2));
    }

    #[test]
    fn typed_codec_blocks_expose_contextual_semantics() {
        let xml = r#"<office:text xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <text:h text:outline-level="1">Heading</text:h>
            <text:p>Body</text:p>
        </office:text>"#;

        let blocks = Elements::parse(xml).unwrap();
        assert_eq!(blocks.len(), 2);
        assert_eq!(blocks[0].kind(), Kind::Heading);
        assert_eq!(blocks[0].as_heading().unwrap().text().unwrap(), "Heading");
        assert_eq!(blocks[1].kind(), Kind::Paragraph);
        assert_eq!(blocks[1].as_paragraph().unwrap().text().unwrap(), "Body");
    }

    #[test]
    fn test_text_elements_extract_text() {
        let xml = r#"<office:text xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <text:p>First paragraph</text:p>
            <text:p>Second paragraph</text:p>
        </office:text>"#;

        let text = TextElements::extract_text(xml).unwrap();
        assert!(text.contains("First paragraph"));
        assert!(text.contains("Second paragraph"));
    }

    #[test]
    fn test_text_elements_extract_text_with_lists() {
        let xml = r#"<office:text xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <text:list>
                <text:list-item><text:p>Item 1</text:p></text:list-item>
                <text:list-item><text:p>Item 2</text:p></text:list-item>
            </text:list>
        </office:text>"#;

        let text = TextElements::extract_text(xml).unwrap();
        assert!(text.contains("Item 1"));
        assert!(text.contains("Item 2"));
    }

    #[test]
    fn parses_arbitrary_prefixes_entities_cdata_and_odf_whitespace() {
        let xml = r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:h t:outline-level="3">A &amp; <![CDATA[B]]></t:h><t:p t:style-name="Body">C<t:s t:c="2"/>D<t:tab/>E<t:line-break/>F&#x21;</t:p><t:p/></o:text>"#;

        assert_eq!(
            TextElements::extract_text(xml).unwrap(),
            "A & B\nC  D\tE\nF!\n"
        );
        let paragraphs = TextElements::parse_paragraphs(xml).unwrap();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].style_name(), Some("Body"));
        assert_eq!(paragraphs[0].text().unwrap(), "C  D\tE\nF!");
        let headings = TextElements::parse_headings(xml).unwrap();
        assert_eq!(headings.len(), 1);
        assert_eq!(headings[0].level(), Some(3));
    }

    #[test]
    fn skips_tracked_change_definitions_and_rejects_malformed_or_excessive_text() {
        let tracked = r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:tracked-changes><t:changed-region><t:deletion><t:p>Deleted</t:p></t:deletion></t:changed-region></t:tracked-changes><t:p>Visible</t:p></o:text>"#;
        assert_eq!(TextElements::extract_text(tracked).unwrap(), "Visible");

        assert!(TextElements::extract_text("<t:p>").is_err());
        let excessive = format!(
            r#"<t:p xmlns:t="{}">A<t:s t:c="{}"/></t:p>"#,
            std::str::from_utf8(TEXT_NAMESPACE).unwrap(),
            MAX_SPACE_COUNT + 1
        );
        assert!(TextElements::extract_text(&excessive).is_err());
        let zero = format!(
            r#"<t:p xmlns:t="{}">A<t:s t:c="0"/>B</t:p>"#,
            std::str::from_utf8(TEXT_NAMESPACE).unwrap()
        );
        assert_eq!(TextElements::extract_text(&zero).unwrap(), "AB");
    }

    #[test]
    fn keeps_note_citations_but_excludes_note_bodies_from_outer_paragraphs() {
        let xml = r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p>Before<t:note t:note-class="footnote" t:id="n1"><t:note-citation>1</t:note-citation><t:note-body><t:p>Hidden body</t:p><t:list><t:list-item><t:p>Hidden item</t:p></t:list-item></t:list></t:note-body></t:note>After</t:p></o:text>"#;
        assert_eq!(TextElements::extract_text(xml).unwrap(), "Before1After");
        let paragraphs = TextElements::parse_paragraphs(xml).unwrap();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].text().unwrap(), "Before1After");
    }

    #[test]
    fn keeps_ruby_base_but_excludes_pronunciation_from_visible_text() {
        let xml = r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p>Before<t:ruby><t:ruby-base>漢字</t:ruby-base><t:ruby-text>かんじ</t:ruby-text></t:ruby>After</t:p></o:text>"#;
        assert_eq!(TextElements::extract_text(xml).unwrap(), "Before漢字After");
    }
}
