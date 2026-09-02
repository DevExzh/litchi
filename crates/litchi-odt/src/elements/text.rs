//! Text-related ODF elements.
//!
//! This module provides classes for text elements like paragraphs, spans,
//! headings, and other text content elements.

mod codec;
mod model;
mod validation;

use super::element::{Element, ElementBase, try_owned_string, try_prefixed_name};
use crate::binding_tracker::BindingTracker;
use litchi_core::{Error, Result, SequentialTextWriter, TextObjectKind, TextOutputError};
use memchr::memmem;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{LocalName, Namespace, QName, ResolveResult};
use quick_xml::reader::{NsReader, Reader};
use std::{collections::VecDeque, io::Write};

pub use codec::Elements;
pub use model::{Block, Kind, LinkActuate, LinkShow};

/// Compatibility facade for the historic text-element collection name.
pub type TextElements = Elements;

/// Internal name retained for the decoder's block slots.
pub(crate) type TextBlock = Block;

pub(crate) const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_TEXT_BLOCKS: usize = 1_000_000;
const MAX_TEXT_DEPTH: usize = 4_096;
const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_SPACE_COUNT: usize = 1_000_000;
/// Qualified name of the unnumbered leading block of an ODF list.
const LIST_HEADER_TAG: &str = "text:list-header";

fn try_push<T>(items: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    items
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })?;
    items.push(value);
    Ok(())
}

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
        self.element.try_get_text_recursive()
    }

    #[cfg(test)]
    pub(crate) fn into_text(self) -> String {
        self.element.into_text_recursive()
    }

    /// Set the text content of the paragraph
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get all text spans within this paragraph
    pub fn spans(&self) -> Result<Vec<Span>> {
        let mut spans = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:span" {
                let span = Span::from_element(child.try_clone()?)?;
                try_push(&mut spans, span, "ODT span projection")?;
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
        let mut hyperlinks = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:a" {
                try_push(
                    &mut hyperlinks,
                    Hyperlink::from_element(child.try_clone()?)?,
                    "ODT hyperlink projection",
                )?;
            }
        }
        Ok(hyperlinks)
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
        self.element.try_get_text_recursive()
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
        self.element.try_get_text_recursive()
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
        self.element.try_get_text_recursive()
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
        self.element.try_get_text_recursive()
    }

    #[cfg(test)]
    pub(crate) fn into_text(self) -> String {
        self.element.into_text_recursive()
    }

    /// Convert this heading to a paragraph with fallible text/style projection.
    pub fn try_into_paragraph(self) -> Result<Paragraph> {
        let text = self.element.try_get_text_recursive()?;
        let mut element = Element::try_new("text:p")?;
        element.try_set_text(&text, "ODT heading paragraph text")?;
        if let Some(style) = self.element.get_attribute("text:style-name") {
            element.try_set_attribute("text:style-name", style, "ODT heading paragraph style")?;
        }
        Ok(Paragraph { element })
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
            if child.tag_name() == "text:list-item" {
                let item = ListItem::from_element(child.try_clone()?)?;
                try_push(&mut items, item, "ODT list-item projection")?;
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
                return ListHeader::from_element(child.try_clone()?).map(Some);
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
        self.element.try_get_text_recursive()
    }

    /// Set the text content of the header.
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the paragraphs the header contains.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut paragraphs = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:p" {
                let paragraph = Paragraph::from_element(child.try_clone()?)?;
                try_push(&mut paragraphs, paragraph, "ODT paragraph projection")?;
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
        self.element.try_get_text_recursive()
    }

    /// Set the text content of the list item
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get nested paragraphs
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut paragraphs = Vec::new();
        for child in &self.element.children {
            if child.tag_name() == "text:p" {
                let para = Paragraph::from_element(child.try_clone()?)?;
                try_push(&mut paragraphs, para, "ODT paragraph projection")?;
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

/// Text-only counterpart of [`ActiveTextBlock`] for the discard-but-validate
/// extraction path: identical depth/slot bookkeeping, no retained `Element`.
struct ActiveTextBlockText {
    depth: usize,
    text: String,
    /// Index of the reserved output slot, so completed blocks keep the order in
    /// which they *started* rather than the order in which they closed.
    slot: usize,
}

struct ActiveSelectedTextBlock {
    element: Option<Element>,
    depth: usize,
    text: RetainedText,
}

struct RetainedText {
    value: Option<String>,
    len: usize,
}

impl RetainedText {
    fn new(retain: bool) -> Self {
        Self {
            value: retain.then(String::new),
            len: 0,
        }
    }

    fn append(&mut self, value: &str) -> Result<()> {
        self.len = self
            .len
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
        if self.len > MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODF text exceeds {MAX_TEXT_BYTES} bytes"
            )));
        }
        if let Some(output) = &mut self.value {
            output
                .try_reserve(value.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT selected text projection",
                    source,
                })?;
            output.push_str(value);
        }
        Ok(())
    }

    fn append_spaces(&mut self, count: usize) -> Result<()> {
        self.len = self
            .len
            .checked_add(count)
            .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
        if self.len > MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODF text exceeds {MAX_TEXT_BYTES} bytes"
            )));
        }
        if let Some(output) = &mut self.value {
            output
                .try_reserve(count)
                .map_err(|source| Error::Allocation {
                    resource: "ODT selected text projection",
                    source,
                })?;
            output.extend(std::iter::repeat_n(' ', count));
        }
        Ok(())
    }
}

struct ParagraphOutput {
    target: usize,
    next_paragraph: usize,
    paragraph: Option<Paragraph>,
}

impl ParagraphOutput {
    fn begin(&mut self, source: &BytesStart<'_>) -> bool {
        if source.local_name().as_ref() != b"p" {
            return false;
        }
        let retain = self.next_paragraph == self.target;
        self.next_paragraph += 1;
        retain
    }

    fn store(&mut self, mut element: Element, text: String) -> Result<()> {
        element.set_text_owned(text);
        self.paragraph = Some(Paragraph::from_element(element)?);
        Ok(())
    }
}

struct BlockOutput {
    target: usize,
    next_block: usize,
    block: Option<TextBlock>,
}

impl BlockOutput {
    fn begin(&mut self, source: &BytesStart<'_>) -> bool {
        if !is_text_block_name(source.local_name().as_ref()) {
            return false;
        }
        let retain = self.next_block == self.target;
        self.next_block += 1;
        retain
    }

    fn store(&mut self, mut element: Element, text: String) -> Result<()> {
        element.set_text_owned(text);
        self.block = Some(match element.tag_name() {
            "text:p" => TextBlock::Paragraph(Paragraph::from_element(element)?),
            "text:h" => TextBlock::Heading(Heading::from_element(element)?),
            _ => {
                return Err(Error::InvalidFormat(
                    "element is not an ODF paragraph or heading".to_string(),
                ));
            },
        });
        Ok(())
    }
}

trait SelectedTextBlockOutput {
    fn begin(&mut self, source: &BytesStart<'_>) -> bool;
    fn store(&mut self, element: Element, text: String) -> Result<()>;
}

impl SelectedTextBlockOutput for ParagraphOutput {
    fn begin(&mut self, source: &BytesStart<'_>) -> bool {
        ParagraphOutput::begin(self, source)
    }

    fn store(&mut self, element: Element, text: String) -> Result<()> {
        ParagraphOutput::store(self, element, text)
    }
}

impl SelectedTextBlockOutput for BlockOutput {
    fn begin(&mut self, source: &BytesStart<'_>) -> bool {
        BlockOutput::begin(self, source)
    }

    fn store(&mut self, element: Element, text: String) -> Result<()> {
        BlockOutput::store(self, element, text)
    }
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
    parse_text_blocks_with_ownership(xml_content, false)
}

/// Retained-element owned variant, now used only as the test oracle for the
/// discard-but-validate extraction path ([`parse_text_block_texts`]).
#[cfg(test)]
pub(crate) fn parse_text_blocks_owned(xml_content: &str) -> Result<Vec<TextBlock>> {
    parse_text_blocks_with_ownership(xml_content, true)
}

pub(crate) fn parse_paragraph_at(xml_content: &str, index: usize) -> Result<Option<Paragraph>> {
    let mut output = ParagraphOutput {
        target: index,
        next_paragraph: 0,
        paragraph: None,
    };
    parse_selected_paragraph(xml_content, &mut output)?;
    Ok(output.paragraph)
}

pub(crate) fn parse_block_at(xml_content: &str, index: usize) -> Result<Option<TextBlock>> {
    let mut output = BlockOutput {
        target: index,
        next_block: 0,
        block: None,
    };
    parse_selected_paragraph(xml_content, &mut output)?;
    Ok(output.block)
}

/// Scan visible paragraph and heading starts without retaining XML elements
/// or text. The caller is responsible for ODT package-root validation; this
/// routine mirrors the text parser's namespace and suppression rules.
#[cfg(test)]
pub(crate) fn scan_text_block_kinds(xml_content: &str) -> Result<Vec<Kind>> {
    let mut reader = NsReader::from_str(xml_content);
    let mut buffer = Vec::new();
    let mut handler = TextBlockKindHandler::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF text XML: {error}")))?;
        let text_namespace =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE);
        handler.on_event(text_namespace, &event)?;
        if matches!(&event, Event::Eof) {
            break;
        }
        buffer.clear();
    }

    handler.finish()
}

/// Borrowing event handler for the visible paragraph and heading catalog.
///
/// The state machine is shared by the standalone scanner and the catalog's
/// fused validation pass so their block ordering, suppression rules, limits,
/// and allocation resources remain identical.
pub(crate) struct TextBlockKindHandler {
    kinds: Vec<Kind>,
    active_depths: Vec<usize>,
    document_depth: usize,
    tracked_changes_depth: usize,
    skipped_depth: usize,
}

impl TextBlockKindHandler {
    pub(crate) fn new() -> Self {
        Self {
            kinds: Vec::new(),
            active_depths: Vec::new(),
            document_depth: 0,
            tracked_changes_depth: 0,
            skipped_depth: 0,
        }
    }

    pub(crate) fn on_event(&mut self, text_namespace: bool, event: &Event<'_>) -> Result<()> {
        match event {
            Event::Start(element) => {
                self.document_depth = self.document_depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF text nesting depth overflow".to_string())
                })?;
                if self.document_depth > MAX_TEXT_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODF text nesting exceeds {MAX_TEXT_DEPTH} levels"
                    )));
                }

                if self.tracked_changes_depth > 0 {
                    self.tracked_changes_depth =
                        self.tracked_changes_depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODF tracked-change nesting depth overflow".to_string(),
                            )
                        })?;
                } else if is_text_element(text_namespace, element, b"tracked-changes") {
                    self.tracked_changes_depth = 1;
                } else {
                    let block_kind = (self.skipped_depth == 0)
                        .then(|| text_block_kind(text_namespace, element))
                        .flatten();
                    if block_kind.is_none()
                        && let Some(current) = self.active_depths.last_mut()
                    {
                        *current = current.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODF text block nesting depth overflow".to_string(),
                            )
                        })?;
                    }

                    if let Some(kind) = block_kind {
                        push_text_block_kind(&mut self.kinds, kind)?;
                        self.active_depths
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "ODT source text-block scan stack",
                                source,
                            })?;
                        self.active_depths.push(1);
                    } else if self.skipped_depth > 0 {
                        self.skipped_depth =
                            self.skipped_depth.checked_add(1).ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODF suppressed text depth overflow".to_string(),
                                )
                            })?;
                    } else if self.active_depths.last().is_some()
                        && (is_text_element(text_namespace, element, b"note-body")
                            || is_text_element(text_namespace, element, b"ruby-text"))
                    {
                        self.skipped_depth = 1;
                    }
                }
            },
            Event::Empty(element) if self.tracked_changes_depth == 0 && self.skipped_depth == 0 => {
                if let Some(kind) = text_block_kind(text_namespace, element) {
                    push_text_block_kind(&mut self.kinds, kind)?;
                }
            },
            Event::End(_) => {
                self.document_depth = self.document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF text element stack underflow".to_string())
                })?;
                if self.tracked_changes_depth > 0 {
                    self.tracked_changes_depth -= 1;
                } else {
                    self.skipped_depth = self.skipped_depth.saturating_sub(1);
                    if let Some(current) = self.active_depths.last_mut() {
                        *current = current.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat("ODT text block stack underflow".to_string())
                        })?;
                        if *current == 0 {
                            self.active_depths.pop().ok_or_else(|| {
                                Error::InvalidFormat(BLOCK_STACK_ERROR.to_string())
                            })?;
                        }
                    }
                }
            },
            Event::Eof => {},
            _ => {},
        }

        Ok(())
    }

    pub(crate) fn finish(self) -> Result<Vec<Kind>> {
        if !self.active_depths.is_empty()
            || self.document_depth != 0
            || self.tracked_changes_depth != 0
            || self.skipped_depth != 0
        {
            return Err(Error::InvalidFormat(
                "incomplete ODF text XML structure".to_string(),
            ));
        }
        Ok(self.kinds)
    }
}

fn push_text_block_kind(kinds: &mut Vec<Kind>, kind: Kind) -> Result<()> {
    if kinds.len() >= MAX_TEXT_BLOCKS {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BLOCKS} paragraphs and headings"
        )));
    }
    kinds.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODT source text-block catalog",
        source,
    })?;
    kinds.push(kind);
    Ok(())
}

fn parse_text_blocks_with_ownership(xml_content: &str, own_text: bool) -> Result<Vec<TextBlock>> {
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
                    active.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT active text-block stack",
                        source,
                    })?;
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
                        own_text,
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
                            own_text,
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
    let completed = blocks.iter().filter(|block| block.is_some()).count();
    let mut output = Vec::new();
    output
        .try_reserve(completed)
        .map_err(|source| Error::Allocation {
            resource: "ODT completed text-block projection",
            source,
        })?;
    output.extend(blocks.into_iter().flatten());
    Ok(output)
}

/// Whether the raw attributes of an element could declare a namespace.
///
/// `NamespaceResolver::push` adds a binding only for attribute keys that are
/// exactly `xmlns` or start with `xmlns:` (`as_namespace_binding`), and both
/// shapes contain the literal substring `xmlns` in the raw attribute slice, so
/// a slice without that substring can never change the binding content. The
/// converse does not hold — an attribute *value* may contain the substring —
/// so `true` is only a conservative "maybe declares".
fn may_declare_namespace(element: &BytesStart<'_>) -> bool {
    let raw = element.attributes_raw();
    raw.len() >= b"xmlns".len() && memmem::find(raw, b"xmlns").is_some()
}

/// Attribute-name resolution and value decoding for the text-attribute
/// helpers, abstracting over the two namespace machineries driving the text
/// loops: `NsReader` (retained and selected paths) and the change-0227
/// plain-[`Reader`] + [`BindingTracker`] pair (discard path). Both sides
/// replicate `NamespaceResolver::resolve_attribute` (`use_default = false`)
/// and expose the driving reader's decoder, so
/// [`validate_text_block_attributes`] and [`text_space_count`] behave
/// byte-identically under either.
trait TextAttributeResolver {
    /// Resolve an attribute qualified name exactly like
    /// `NamespaceResolver::resolve_attribute`.
    fn resolve_attribute<'n>(&self, name: QName<'n>) -> (ResolveResult<'_>, LocalName<'n>);
    /// The driving reader's decoder, for attribute value normalization.
    fn decoder(&self) -> Decoder;
}

impl TextAttributeResolver for NsReader<&[u8]> {
    fn resolve_attribute<'n>(&self, name: QName<'n>) -> (ResolveResult<'_>, LocalName<'n>) {
        self.resolver().resolve_attribute(name)
    }

    fn decoder(&self) -> Decoder {
        // `NsReader` has no inherent `decoder`; it derefs to the inner
        // `Reader`, whose decoder this exposes.
        (**self).decoder()
    }
}

/// The [`BindingTracker`] + decoder pair of the discard path (change 0227).
/// The decoder is the plain reader's, the same UTF-8 pass-through
/// `NsReader::from_str` carries.
struct TrackedAttributes<'a> {
    tracker: &'a BindingTracker,
    decoder: Decoder,
}

impl TextAttributeResolver for TrackedAttributes<'_> {
    fn resolve_attribute<'n>(&self, name: QName<'n>) -> (ResolveResult<'_>, LocalName<'n>) {
        self.tracker.resolve_attribute(name)
    }

    fn decoder(&self) -> Decoder {
        self.decoder
    }
}

/// Build the discard path's attribute resolver for one call site (the
/// tracker is mutably borrowed between resolutions, so the adapter cannot
/// outlive a single call).
fn tracked_attributes(tracker: &BindingTracker, decoder: Decoder) -> TrackedAttributes<'_> {
    TrackedAttributes { tracker, decoder }
}

/// Last-resolution memo for the text-namespace classification of
/// [`parse_text_block_texts`].
///
/// quick-xml resolves every element event by reverse-scanning all live
/// bindings (`NamespaceResolver::resolve_prefix`), which dominates the
/// discard path when a document repeats the same prefix on thousands of
/// sibling blocks. The resolution is a pure function of the binding-stack
/// content and the queried prefix, and the content changes only when a
/// binding is pushed or popped:
///
/// - A `push` adds bindings only for `xmlns`/`xmlns:*` attribute keys, so a
///   `Start`/`Empty` element failing [`may_declare_namespace`] cannot change
///   the content; one passing it conservatively invalidates the memo *before*
///   the element's own resolution (its declarations are already in scope).
/// - A `pop` removes exactly the bindings of the scope being closed, so only
///   a scope that declared can lose bindings. `NsReader` defers the pop into
///   the next read, so the caller tracks declaring scopes on a stack and
///   bumps the content version before the first resolution after the pop.
///
/// A memo hit therefore means byte-identical binding content for the same
/// prefix, and reusing the verdict is exact — never a stale approximation.
struct TextNamespaceMemo {
    /// Content version the cached verdict was resolved against
    /// (`u64::MAX` before the first resolution, so no real version matches).
    version: u64,
    /// Whether the cached element name carried a prefix.
    has_prefix: bool,
    /// Cached prefix bytes; meaningful only when `has_prefix` is set.
    prefix: Vec<u8>,
    /// Cached verdict: whether the name resolves to `TEXT_NAMESPACE`.
    is_text: bool,
}

impl TextNamespaceMemo {
    fn new() -> Self {
        Self {
            version: u64::MAX,
            has_prefix: false,
            prefix: Vec::new(),
            is_text: false,
        }
    }

    /// Classify an element name as belonging to the text namespace, reusing
    /// the last verdict when the binding content and prefix are unchanged.
    fn is_text_namespace(
        &mut self,
        tracker: &BindingTracker,
        content_version: u64,
        element: &BytesStart<'_>,
    ) -> Result<bool> {
        let prefix = element.name().prefix().map(|prefix| prefix.into_inner());
        if self.version == content_version
            && self.has_prefix == prefix.is_some()
            && (!self.has_prefix || self.prefix.as_slice() == prefix.unwrap_or(b""))
        {
            return Ok(self.is_text);
        }
        // Miss: resolve directly, exactly as `resolve_event` does for element
        // events (`use_default = true`, baked into `resolve_prefix`).
        let is_text = matches!(
            tracker.resolve_prefix(element.name().prefix()),
            ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE
        );
        self.has_prefix = prefix.is_some();
        if let Some(prefix) = prefix {
            self.prefix
                .try_reserve(prefix.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT namespace memo prefix",
                    source,
                })?;
            self.prefix.clear();
            self.prefix.extend_from_slice(prefix);
        }
        self.is_text = is_text;
        // Stamp the version last so a failed refresh leaves the memo
        // conservatively stale rather than falsely current.
        self.version = content_version;
        Ok(is_text)
    }
}

/// Parse every `text:p` and `text:h`, retaining only the collected text.
///
/// Behaves exactly like [`parse_text_blocks_owned`] followed by
/// `Block::into_text` on every block — same event handling, suppression
/// rules, limits, and start-ordered output — but validates each block's
/// attributes without building the retained `Element`, using the established
/// discard pattern of `parse_selected_text_block_element` with
/// `retain = false` ([`validate_text_block_attributes`]): no tag-name copy,
/// no `QualifiedName` triple-allocation, no attribute map, and no owned
/// attribute values. Block elements never have children in this parser, so
/// `into_text_recursive` equals the accumulated text and dropping the tree
/// loses nothing. Only OOM-only `Error::Allocation` sites vanish
/// (`Element::try_new`, `QualifiedName::try_from_string`, `try_set_attribute`,
/// `try_set_text`); the `from_element` tag check that disappears is
/// unreachable, since the tag derives from the same `b"p" | b"h"` match that
/// started the block.
pub(crate) fn parse_text_block_texts(xml_content: &str) -> Result<Vec<String>> {
    // Plain reader + hand-rolled binding maintenance (change 0227): the
    // tracker replicates the push/pop `NsReader` performs inside its read
    // (the `BindingTracker` byte-exactness contract), and the borrowing read
    // drops the per-event buffer copy of `read_event_into`.
    let mut reader = Reader::from_str(xml_content);
    let decoder = reader.decoder();
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    let mut blocks: Vec<Option<String>> = Vec::new();
    let mut active: Vec<ActiveTextBlockText> = Vec::new();
    let mut document_depth = 0usize;
    let mut tracked_changes_depth = 0usize;
    // Depth of the note-body/ruby-text subtree whose content is suppressed.
    let mut skipped_depth = 0usize;
    let mut total_text_bytes = 0usize;
    // Last-resolution memo state (see `TextNamespaceMemo`): a monotone
    // content version of the binding stack, one flag per open `Start` scope
    // recording whether that scope could declare namespaces, and a marker
    // that the previous event's deferred pop removed bindings.
    let mut memo = TextNamespaceMemo::new();
    let mut content_version = 0u64;
    let mut declared_stack: Vec<bool> = Vec::new();
    let mut pop_removed_bindings = false;

    loop {
        // The deferred pop of the previous `End`/`Empty` scope runs before
        // the read, exactly where `NsReader::read_event_impl` applies it.
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        // Borrowing read: events borrow `xml_content` directly. The plain
        // `Reader` is the same tokenizer `NsReader` wraps, with the same
        // default configuration, so the tokenization error stream is
        // unchanged.
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF text XML: {error}")))?;
        // The deferred pop above removed bindings when this flag is set;
        // invalidate the memo before any resolution below.
        if pop_removed_bindings {
            content_version += 1;
            pop_removed_bindings = false;
        }
        // This bookkeeping precedes every `continue` below because the
        // tracker maintains bindings inside skipped subtrees too.
        //
        // The push for a `Start`/`Empty` runs before the classification, so
        // a namespace error preempts the event exactly where `NsReader`'s
        // read returned `Err`. A push error is a real `NamespaceError`,
        // whose `Display` is what `quick_xml::Error::Namespace` forwards to,
        // so the message is byte-identical to the historical failure.
        //
        // `resolve_event` maps `Start`/`Empty` to
        // `resolve_prefix(name().prefix(), use_default = true)` and every
        // other event to `Unbound` (the `End` verdict was computed but never
        // consumed here), so only `Start`/`Empty` go through the memo.
        let text_namespace = match event {
            Event::Start(ref element) => {
                tracker.push(element).map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODF text XML: {error}"))
                })?;
                let declares = may_declare_namespace(element);
                try_push(
                    &mut declared_stack,
                    declares,
                    "ODT namespace-declaration stack",
                )?;
                // The element's own declarations are in scope for its
                // resolution (the push ran above).
                if declares {
                    content_version += 1;
                }
                memo.is_text_namespace(&tracker, content_version, element)?
            },
            Event::Empty(ref element) => {
                tracker.push(element).map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODF text XML: {error}"))
                })?;
                // The scope an `Empty` element opens closes immediately:
                // defer its pop to the top of the next iteration.
                pending_pop = true;
                if may_declare_namespace(element) {
                    // Push for this event, then a deferred pop before the
                    // next one: both change the binding content.
                    content_version += 1;
                    pop_removed_bindings = true;
                }
                memo.is_text_namespace(&tracker, content_version, element)?
            },
            _ => false,
        };
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
                    continue;
                }
                if is_text_element(text_namespace, element, b"tracked-changes") {
                    tracked_changes_depth = 1;
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
                    active.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT active text-block stack",
                        source,
                    })?;
                    // Attribute validation precedes slot reservation, matching
                    // the retained path's `make_text_block_element`-before-
                    // `reserve_text_block` evaluation order.
                    validate_text_block_attributes(
                        &tracked_attributes(&tracker, decoder),
                        element,
                    )?;
                    let slot = reserve_text_block_text(&mut blocks)?;
                    active.push(ActiveTextBlockText {
                        depth: 1,
                        text: String::new(),
                        slot,
                    });
                } else if skipped_depth > 0 {
                    skipped_depth += 1;
                } else if let Some(current) = active.last_mut() {
                    if is_text_element(text_namespace, element, b"note-body")
                        || is_text_element(text_namespace, element, b"ruby-text")
                    {
                        skipped_depth = 1;
                    } else {
                        append_text_control(
                            &tracked_attributes(&tracker, decoder),
                            text_namespace,
                            element,
                            &mut current.text,
                        )?;
                    }
                }
            },
            Event::Empty(ref element) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if is_text_element(text_namespace, element, b"note-body")
                    || is_text_element(text_namespace, element, b"ruby-text")
                {
                    // An empty suppressed run contributes nothing either way.
                } else if is_text_block(text_namespace, element) {
                    // The retained path reserves the slot before validating
                    // (`reserve_text_block` before `make_text_block_element`);
                    // keep that order.
                    let slot = reserve_text_block_text(&mut blocks)?;
                    validate_text_block_attributes(
                        &tracked_attributes(&tracker, decoder),
                        element,
                    )?;
                    store_text_block_text(String::new(), slot, &mut blocks, &mut total_text_bytes)?;
                } else if let Some(current) = active.last_mut() {
                    append_text_control(
                        &tracked_attributes(&tracker, decoder),
                        text_namespace,
                        element,
                        &mut current.text,
                    )?;
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
                // The deferred pop of this scope executes at the top of the
                // next iteration; record now whether it will remove bindings.
                // An unbalanced stack is malformed input the reader rejects
                // anyway, so conservatively treat it as declaring.
                pending_pop = true;
                if declared_stack.pop().unwrap_or(true) {
                    pop_removed_bindings = true;
                }
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF text element stack underflow".to_string())
                })?;
                if tracked_changes_depth > 0 {
                    tracked_changes_depth -= 1;
                    continue;
                }
                skipped_depth = skipped_depth.saturating_sub(1);
                if let Some(current) = active.last_mut() {
                    current.depth = current.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("ODT text block stack underflow".to_string())
                    })?;
                    if current.depth == 0 {
                        let current = active
                            .pop()
                            .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))?;
                        store_text_block_text(
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
    }

    if !active.is_empty() || tracked_changes_depth != 0 || skipped_depth != 0 || document_depth != 0
    {
        return Err(Error::InvalidFormat(
            "incomplete ODF text XML structure".to_string(),
        ));
    }
    let completed = blocks.iter().filter(|block| block.is_some()).count();
    let mut output = Vec::new();
    output
        .try_reserve(completed)
        .map_err(|source| Error::Allocation {
            resource: "ODT completed text-block projection",
            source,
        })?;
    output.extend(blocks.into_iter().flatten());
    Ok(output)
}

/// Bounded decoded-text accounting for the one-pass sink parser.
///
/// Unlike the retained extraction path, this budget is charged while text is
/// still distributed among nested active blocks. That prevents a deeply
/// nested document from obtaining one full text ceiling per active string.
struct SinkTextBudget {
    decoded_bytes: usize,
}

impl SinkTextBudget {
    fn charge(&mut self, additional: usize) -> Result<()> {
        self.decoded_bytes = self
            .decoded_bytes
            .checked_add(additional)
            .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
        if self.decoded_bytes > MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODF text exceeds {MAX_TEXT_BYTES} bytes"
            )));
        }
        Ok(())
    }
}

/// Return the exact UTF-8 byte length produced by XML 1.0 EOL
/// normalization, without allocating the normalized string.
///
/// The sink parser uses `Reader::from_str`, so event bytes are UTF-8. Keeping
/// the validation here makes the precharge fallible and preserves the decoder
/// error boundary before `xml_content` is asked to materialize a normalized
/// value. XML 1.0 changes `CRLF` to one `LF`; a lone `CR` and every other
/// UTF-8 byte keep their length.
fn normalized_xml10_decoded_len(raw: &[u8], context: &str) -> Result<usize> {
    std::str::from_utf8(raw)
        .map_err(|error| Error::InvalidFormat(format!("invalid ODF {context}: {error}")))?;
    let mut length = raw.len();
    let mut index = 0;
    while index < raw.len() {
        if raw[index] == b'\r' {
            if raw.get(index + 1) == Some(&b'\n') {
                length = length
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
                index += 2;
            } else {
                index += 1;
            }
        } else {
            index += 1;
        }
    }
    Ok(length)
}

/// Start-order slots for blocks that have completed out of order.
///
/// A nested block can close before its containing paragraph. Completed slots
/// remain bounded and are emitted as soon as the contiguous start-order
/// frontier advances. The queue therefore never requires a retained document
/// projection, while preserving the existing nested-block ordering contract.
struct PendingSinkBlocks {
    pending: VecDeque<Option<String>>,
    next_slot: usize,
    next_emit: usize,
    block_count: usize,
}

impl PendingSinkBlocks {
    fn new() -> Self {
        Self {
            pending: VecDeque::new(),
            next_slot: 0,
            next_emit: 0,
            block_count: 0,
        }
    }

    fn reserve(&mut self) -> Result<usize> {
        if self.block_count >= MAX_TEXT_BLOCKS {
            return Err(Error::InvalidFormat(format!(
                "ODF text exceeds {MAX_TEXT_BLOCKS} paragraphs and headings"
            )));
        }
        self.pending
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT pending text-block frontier",
                source,
            })?;
        let slot = self.next_slot;
        self.next_slot = self
            .next_slot
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODF text block slot overflow".to_string()))?;
        self.block_count += 1;
        self.pending.push_back(None);
        Ok(slot)
    }

    fn complete<'options, 'output, W: Write + ?Sized>(
        &mut self,
        slot: usize,
        text: String,
        writer: &mut SequentialTextWriter<'options, 'output, W>,
    ) -> std::result::Result<(), TextOutputError<Error>> {
        let index = slot
            .checked_sub(self.next_emit)
            .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))
            .map_err(|error| writer.document_error(error))?;
        let target = self
            .pending
            .get_mut(index)
            .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))
            .map_err(|error| writer.document_error(error))?;
        if target.is_some() {
            return Err(writer.document_error(Error::InvalidFormat(BLOCK_STACK_ERROR.to_string())));
        }
        *target = Some(text);

        while matches!(self.pending.front(), Some(Some(_))) {
            let value = self
                .pending
                .pop_front()
                .and_then(|value| value)
                .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))
                .map_err(|error| writer.document_error(error))?;
            writer.write_object(TextObjectKind::Paragraph, &value)?;
            self.next_emit = self
                .next_emit
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("ODF text block slot overflow".to_string()))
                .map_err(|error| writer.document_error(error))?;
        }
        Ok(())
    }
}

fn append_sink_checked(
    budget: &mut SinkTextBudget,
    output: &mut String,
    value: &str,
) -> Result<()> {
    budget.charge(value.len())?;
    append_sink_precharged(output, value, value.len())
}

/// Append a decoded fragment after its exact byte length has been charged.
fn append_sink_precharged(output: &mut String, value: &str, charged_len: usize) -> Result<()> {
    if value.len() != charged_len {
        return Err(Error::InvalidFormat(
            "ODF sink text precharge mismatch".to_string(),
        ));
    }
    output
        .try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT sink text-block content",
            source,
        })?;
    output.push_str(value);
    Ok(())
}

fn append_sink_spaces(
    budget: &mut SinkTextBudget,
    output: &mut String,
    count: usize,
) -> Result<()> {
    budget.charge(count)?;
    output
        .try_reserve(count)
        .map_err(|source| Error::Allocation {
            resource: "ODT sink text-block content",
            source,
        })?;
    output.extend(std::iter::repeat_n(' ', count));
    Ok(())
}

fn append_sink_text_control<R: TextAttributeResolver + ?Sized>(
    resolver: &R,
    text_namespace: bool,
    element: &BytesStart<'_>,
    output: &mut String,
    budget: &mut SinkTextBudget,
) -> Result<()> {
    if !text_namespace {
        return Ok(());
    }
    match element.local_name().as_ref() {
        b"s" => {
            let count = text_space_count(resolver, element)?.unwrap_or(1);
            if count > MAX_SPACE_COUNT {
                return Err(Error::InvalidFormat(format!(
                    "text:s count exceeds {MAX_SPACE_COUNT}"
                )));
            }
            append_sink_spaces(budget, output, count)?;
        },
        b"tab" => append_sink_checked(budget, output, "\t")?,
        b"line-break" => append_sink_checked(budget, output, "\n")?,
        _ => {},
    }
    Ok(())
}

/// Decode and append one general reference without allocating an owned
/// intermediate string on the sink path.
fn append_sink_reference(
    reference: &BytesRef<'_>,
    budget: &mut SinkTextBudget,
    output: &mut String,
) -> Result<()> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODF text character reference: {error}"))
    })? {
        let mut encoded = [0_u8; 4];
        let value = character.encode_utf8(&mut encoded);
        budget.charge(value.len())?;
        return append_sink_precharged(output, value, value.len());
    }

    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODF text entity: {error}")))?;
    let value = match name.as_ref() {
        "amp" => "&",
        "lt" => "<",
        "gt" => ">",
        "quot" => "\"",
        "apos" => "'",
        _ => {
            return Err(Error::InvalidFormat(format!(
                "unsupported ODF text entity '&{name};'"
            )));
        },
    };
    budget.charge(value.len())?;
    append_sink_precharged(output, value, value.len())
}

/// Parse visible ODT text once and feed completed blocks to a shared sink.
///
/// This intentionally remains a dedicated parser rather than adding a mode to
/// the retained hot paths. It keeps only active block strings and the bounded
/// start-order frontier, and applies the same namespace, suppression,
/// attribute, control, and structural-limit rules as the established parser.
pub(crate) fn write_text_blocks_to_writer<'options, 'output, W: Write + ?Sized>(
    xml_content: &str,
    writer: &mut SequentialTextWriter<'options, 'output, W>,
) -> std::result::Result<(), TextOutputError<Error>> {
    macro_rules! parse {
        ($expression:expr) => {
            $expression.map_err(|error| writer.document_error(error))?
        };
    }

    let mut reader = Reader::from_str(xml_content);
    let decoder = reader.decoder();
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    let mut active: Vec<ActiveTextBlockText> = Vec::new();
    let mut pending = PendingSinkBlocks::new();
    let mut document_depth = 0usize;
    let mut tracked_changes_depth = 0usize;
    let mut skipped_depth = 0usize;
    let mut budget = SinkTextBudget { decoded_bytes: 0 };
    let mut memo = TextNamespaceMemo::new();
    let mut content_version = 0u64;
    let mut declared_stack: Vec<bool> = Vec::new();
    let mut pop_removed_bindings = false;

    loop {
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event = parse!(
            reader
                .read_event()
                .map_err(|error| Error::InvalidFormat(format!("invalid ODF text XML: {error}")))
        );
        if pop_removed_bindings {
            content_version += 1;
            pop_removed_bindings = false;
        }
        let text_namespace = match event {
            Event::Start(ref element) => {
                parse!(tracker.push(element).map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODF text XML: {error}"))
                }));
                let declares = may_declare_namespace(element);
                parse!(try_push(
                    &mut declared_stack,
                    declares,
                    "ODT namespace-declaration stack",
                ));
                if declares {
                    content_version += 1;
                }
                parse!(memo.is_text_namespace(&tracker, content_version, element))
            },
            Event::Empty(ref element) => {
                parse!(tracker.push(element).map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODF text XML: {error}"))
                }));
                pending_pop = true;
                if may_declare_namespace(element) {
                    content_version += 1;
                    pop_removed_bindings = true;
                }
                parse!(memo.is_text_namespace(&tracker, content_version, element))
            },
            _ => false,
        };

        match event {
            Event::Start(ref element) => {
                document_depth = parse!(document_depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF text nesting depth overflow".to_string())
                }));
                if document_depth > MAX_TEXT_DEPTH {
                    return Err(writer.document_error(Error::InvalidFormat(format!(
                        "ODF text nesting exceeds {MAX_TEXT_DEPTH} levels"
                    ))));
                }
                if tracked_changes_depth > 0 {
                    tracked_changes_depth += 1;
                    continue;
                }
                if is_text_element(text_namespace, element, b"tracked-changes") {
                    tracked_changes_depth = 1;
                    continue;
                }

                let starts_block = skipped_depth == 0 && is_text_block(text_namespace, element);
                if !starts_block && let Some(current) = active.last_mut() {
                    current.depth += 1;
                }
                if starts_block {
                    parse!(active.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT active text-block stack",
                        source,
                    }));
                    parse!(validate_text_block_attributes(
                        &tracked_attributes(&tracker, decoder),
                        element,
                    ));
                    let slot = parse!(pending.reserve());
                    active.push(ActiveTextBlockText {
                        depth: 1,
                        text: String::new(),
                        slot,
                    });
                } else if skipped_depth > 0 {
                    skipped_depth += 1;
                } else if let Some(current) = active.last_mut() {
                    if is_text_element(text_namespace, element, b"note-body")
                        || is_text_element(text_namespace, element, b"ruby-text")
                    {
                        skipped_depth = 1;
                    } else {
                        parse!(append_sink_text_control(
                            &tracked_attributes(&tracker, decoder),
                            text_namespace,
                            element,
                            &mut current.text,
                            &mut budget,
                        ));
                    }
                }
            },
            Event::Empty(ref element) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if is_text_element(text_namespace, element, b"note-body")
                    || is_text_element(text_namespace, element, b"ruby-text")
                {
                    // An empty suppressed run contributes nothing.
                } else if is_text_block(text_namespace, element) {
                    let slot = parse!(pending.reserve());
                    parse!(validate_text_block_attributes(
                        &tracked_attributes(&tracker, decoder),
                        element,
                    ));
                    pending.complete(slot, String::new(), writer)?;
                } else if let Some(current) = active.last_mut() {
                    parse!(append_sink_text_control(
                        &tracked_attributes(&tracker, decoder),
                        text_namespace,
                        element,
                        &mut current.text,
                        &mut budget,
                    ));
                }
            },
            Event::Text(ref value) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if let Some(current) = active.last_mut() {
                    let decoded_len =
                        parse!(normalized_xml10_decoded_len(value.as_ref(), "text content",));
                    parse!(budget.charge(decoded_len));
                    let decoded =
                        parse!(value.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODF text content: {error}"))
                        }));
                    parse!(append_sink_precharged(
                        &mut current.text,
                        &decoded,
                        decoded_len,
                    ));
                }
            },
            Event::CData(ref value) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if let Some(current) = active.last_mut() {
                    let decoded_len =
                        parse!(normalized_xml10_decoded_len(value.as_ref(), "text CDATA",));
                    parse!(budget.charge(decoded_len));
                    let decoded =
                        parse!(value.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODF text CDATA: {error}"))
                        }));
                    parse!(append_sink_precharged(
                        &mut current.text,
                        &decoded,
                        decoded_len,
                    ));
                }
            },
            Event::GeneralRef(ref reference)
                if tracked_changes_depth == 0 && skipped_depth == 0 =>
            {
                if let Some(current) = active.last_mut() {
                    parse!(append_sink_reference(
                        reference,
                        &mut budget,
                        &mut current.text,
                    ));
                }
            },
            Event::End(_) => {
                pending_pop = true;
                if declared_stack.pop().unwrap_or(true) {
                    pop_removed_bindings = true;
                }
                document_depth = parse!(document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODF text element stack underflow".to_string())
                }));
                if tracked_changes_depth > 0 {
                    tracked_changes_depth -= 1;
                    continue;
                }
                skipped_depth = skipped_depth.saturating_sub(1);
                if let Some(current) = active.last_mut() {
                    current.depth = parse!(current.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("ODF text block stack underflow".to_string())
                    }));
                    if current.depth == 0 {
                        let current = active
                            .pop()
                            .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))
                            .map_err(|error| writer.document_error(error))?;
                        pending.complete(current.slot, current.text, writer)?;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !active.is_empty() || tracked_changes_depth != 0 || skipped_depth != 0 || document_depth != 0
    {
        return Err(writer.document_error(Error::InvalidFormat(
            "incomplete ODF text XML structure".to_string(),
        )));
    }
    if !pending.pending.is_empty() {
        return Err(writer.document_error(Error::InvalidFormat(BLOCK_STACK_ERROR.to_string())));
    }
    Ok(())
}

fn parse_selected_paragraph<O: SelectedTextBlockOutput + ?Sized>(
    xml_content: &str,
    output: &mut O,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml_content);
    let mut buffer = Vec::new();
    let mut active: Vec<ActiveSelectedTextBlock> = Vec::new();
    let mut block_count = 0usize;
    let mut document_depth = 0usize;
    let mut tracked_changes_depth = 0usize;
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

                let starts_block = skipped_depth == 0 && is_text_block(text_namespace, element);
                if !starts_block && let Some(current) = active.last_mut() {
                    current.depth += 1;
                }

                if starts_block {
                    count_text_block(&mut block_count)?;
                    let retain = output.begin(element);
                    active.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT selected text-block stack",
                        source,
                    })?;
                    active.push(ActiveSelectedTextBlock {
                        element: parse_selected_text_block_element(&reader, element, retain)?,
                        depth: 1,
                        text: RetainedText::new(retain),
                    });
                } else if skipped_depth > 0 {
                    skipped_depth += 1;
                } else if let Some(current) = active.last_mut() {
                    if is_text_element(text_namespace, element, b"note-body")
                        || is_text_element(text_namespace, element, b"ruby-text")
                    {
                        skipped_depth = 1;
                    } else {
                        append_selected_text_control(
                            &reader,
                            text_namespace,
                            element,
                            &mut current.text,
                        )?;
                    }
                }
            },
            Event::Empty(ref element) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if is_text_element(text_namespace, element, b"note-body")
                    || is_text_element(text_namespace, element, b"ruby-text")
                {
                    // An empty suppressed run contributes nothing either way.
                } else if is_text_block(text_namespace, element) {
                    count_text_block(&mut block_count)?;
                    let retain = output.begin(element);
                    finish_selected_text_block(
                        parse_selected_text_block_element(&reader, element, retain)?,
                        RetainedText::new(retain),
                        output,
                        &mut total_text_bytes,
                    )?;
                } else if let Some(current) = active.last_mut() {
                    append_selected_text_control(
                        &reader,
                        text_namespace,
                        element,
                        &mut current.text,
                    )?;
                }
            },
            Event::Text(ref value) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if let Some(current) = active.last_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODF text content: {error}"))
                        })?;
                    current.text.append(&decoded)?;
                }
            },
            Event::CData(ref value) if tracked_changes_depth == 0 && skipped_depth == 0 => {
                if let Some(current) = active.last_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODF text CDATA: {error}"))
                        })?;
                    current.text.append(&decoded)?;
                }
            },
            Event::GeneralRef(ref reference)
                if tracked_changes_depth == 0 && skipped_depth == 0 =>
            {
                if let Some(current) = active.last_mut() {
                    let decoded = decode_reference(reference)?;
                    current.text.append(&decoded)?;
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
                        finish_selected_text_block(
                            current.element,
                            current.text,
                            output,
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
    Ok(())
}

fn is_text_element(text_namespace: bool, element: &BytesStart<'_>, local_name: &[u8]) -> bool {
    text_namespace && element.local_name().as_ref() == local_name
}

fn is_text_block(text_namespace: bool, element: &BytesStart<'_>) -> bool {
    text_block_kind(text_namespace, element).is_some()
}

fn is_text_block_name(local_name: &[u8]) -> bool {
    matches!(local_name, b"p" | b"h")
}

fn text_block_kind(text_namespace: bool, element: &BytesStart<'_>) -> Option<Kind> {
    if !text_namespace {
        return None;
    }
    match element.local_name().as_ref() {
        b"p" => Some(Kind::Paragraph),
        b"h" => Some(Kind::Heading),
        _ => None,
    }
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
    let mut element = Element::try_new(tag_name)?;
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
                try_prefixed_name("text", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => {
                try_prefixed_name("xlink", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XML_NAMESPACE => {
                try_prefixed_name("xml", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(_) | ResolveResult::Unbound => {
                std::str::from_utf8(attribute.key.as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat("non-UTF-8 ODF text attribute name".to_string())
                    })
                    .and_then(|name| try_owned_string(name, "ODT text attribute name"))?
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
        element.try_set_attribute(&name, &value, "ODT text-block attribute")?;
    }
    Ok(element)
}

fn parse_selected_text_block_element(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    retain: bool,
) -> Result<Option<Element>> {
    let _tag_name = match source.local_name().as_ref() {
        b"p" => "text:p",
        b"h" => "text:h",
        _ => {
            return Err(Error::InvalidFormat(
                "element is not an ODF paragraph or heading".to_string(),
            ));
        },
    };
    if retain {
        return make_text_block_element(reader, source).map(Some);
    }
    let mut discarded_names = Vec::new();
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
                try_prefixed_name("text", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => {
                try_prefixed_name("xlink", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XML_NAMESPACE => {
                try_prefixed_name("xml", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(_) | ResolveResult::Unbound => {
                std::str::from_utf8(attribute.key.as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat("non-UTF-8 ODF text attribute name".to_string())
                    })
                    .and_then(|name| try_owned_string(name, "ODT text attribute name"))?
            },
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unknown ODF text attribute namespace prefix '{}'",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        };
        let _value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODF text attribute value: {error}"))
            })?;
        if discarded_names.iter().any(|existing| existing == &name) {
            return Err(Error::InvalidFormat(format!(
                "duplicate ODF text attribute '{name}'"
            )));
        }
        try_push(
            &mut discarded_names,
            name,
            "ODT discarded attribute-name projection",
        )?;
    }
    Ok(None)
}

/// Validate one text block's attributes exactly like
/// [`make_text_block_element`] — same checks in the same order producing the
/// same errors, including decode-before-duplicate — but retain nothing:
/// names live only in a scratch `Vec` for duplicate detection. This is the
/// discard branch of [`parse_selected_text_block_element`] lifted into a
/// shared shape; both copies must stay in lockstep with
/// [`make_text_block_element`].
fn validate_text_block_attributes<R: TextAttributeResolver + ?Sized>(
    resolver: &R,
    source: &BytesStart<'_>,
) -> Result<()> {
    let mut discarded_names = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODF text attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local_name) = resolver.resolve_attribute(attribute.key);
        let local_name = std::str::from_utf8(local_name.as_ref()).map_err(|_error| {
            Error::InvalidFormat("non-UTF-8 ODF text attribute name".to_string())
        })?;
        let name = match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE => {
                try_prefixed_name("text", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => {
                try_prefixed_name("xlink", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XML_NAMESPACE => {
                try_prefixed_name("xml", local_name, "ODT text attribute name")?
            },
            ResolveResult::Bound(_) | ResolveResult::Unbound => {
                std::str::from_utf8(attribute.key.as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat("non-UTF-8 ODF text attribute name".to_string())
                    })
                    .and_then(|name| try_owned_string(name, "ODT text attribute name"))?
            },
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unknown ODF text attribute namespace prefix '{}'",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        };
        let _value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, resolver.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODF text attribute value: {error}"))
            })?;
        if discarded_names.iter().any(|existing| existing == &name) {
            return Err(Error::InvalidFormat(format!(
                "duplicate ODF text attribute '{name}'"
            )));
        }
        try_push(
            &mut discarded_names,
            name,
            "ODT discarded attribute-name projection",
        )?;
    }
    Ok(())
}

fn append_selected_text_control(
    reader: &NsReader<&[u8]>,
    text_namespace: bool,
    element: &BytesStart<'_>,
    output: &mut RetainedText,
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
            output.append_spaces(count)?;
        },
        b"tab" => output.append("\t")?,
        b"line-break" => output.append("\n")?,
        _ => {},
    }
    Ok(())
}

fn append_text_control<R: TextAttributeResolver + ?Sized>(
    resolver: &R,
    text_namespace: bool,
    element: &BytesStart<'_>,
    output: &mut String,
) -> Result<()> {
    if !text_namespace {
        return Ok(());
    }
    match element.local_name().as_ref() {
        b"s" => {
            let count = text_space_count(resolver, element)?.unwrap_or(1);
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
            output
                .try_reserve(count)
                .map_err(|source| Error::Allocation {
                    resource: "ODT text projection",
                    source,
                })?;
            output.extend(std::iter::repeat_n(' ', count));
        },
        b"tab" => append_checked(output, "\t")?,
        b"line-break" => append_checked(output, "\n")?,
        _ => {},
    }
    Ok(())
}

fn text_space_count<R: TextAttributeResolver + ?Sized>(
    resolver: &R,
    element: &BytesStart<'_>,
) -> Result<Option<usize>> {
    let mut count = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid text:s attribute: {error}")))?;
        let (namespace, local_name) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE)
            && local_name.as_ref() == b"c"
        {
            if count.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate expanded text:c attribute".to_string(),
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, resolver.decoder())
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
    output
        .try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT text projection",
            source,
        })?;
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
    blocks.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODT text-block projection",
        source,
    })?;
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
    own_text: bool,
) -> Result<()> {
    *total_text_bytes = total_text_bytes
        .checked_add(text.len())
        .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
    if *total_text_bytes > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BYTES} bytes"
        )));
    }
    if own_text {
        element.set_text_owned(text);
    } else {
        element.try_set_text(&text, "ODT text-block content")?;
    }
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

/// Text-only counterpart of [`reserve_text_block`] with the identical limit
/// and message, for the discard-but-validate extraction path.
fn reserve_text_block_text(blocks: &mut Vec<Option<String>>) -> Result<usize> {
    if blocks.len() >= MAX_TEXT_BLOCKS {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BLOCKS} paragraphs and headings"
        )));
    }
    let slot = blocks.len();
    blocks.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODT text-block projection",
        source,
    })?;
    blocks.push(None);
    Ok(slot)
}

/// Text-only counterpart of [`store_text_block`]: fill a previously reserved
/// slot with the finished block's text, accounting it against the overall
/// size budget with the identical checks and messages.
fn store_text_block_text(
    text: String,
    slot: usize,
    blocks: &mut [Option<String>],
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
    let target = blocks
        .get_mut(slot)
        .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))?;
    *target = Some(text);
    Ok(())
}

fn count_text_block(block_count: &mut usize) -> Result<()> {
    if *block_count >= MAX_TEXT_BLOCKS {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BLOCKS} paragraphs and headings"
        )));
    }
    *block_count += 1;
    Ok(())
}

fn finish_selected_text_block<O: SelectedTextBlockOutput + ?Sized>(
    element: Option<Element>,
    text: RetainedText,
    output: &mut O,
    total_text_bytes: &mut usize,
) -> Result<()> {
    *total_text_bytes = total_text_bytes
        .checked_add(text.len)
        .ok_or_else(|| Error::InvalidFormat("ODF text size overflow".to_string()))?;
    if *total_text_bytes > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF text exceeds {MAX_TEXT_BYTES} bytes"
        )));
    }
    if let Some(element) = element {
        output.store(
            element,
            text.value
                .ok_or_else(|| Error::InvalidFormat(BLOCK_STACK_ERROR.to_string()))?,
        )?;
    }
    Ok(())
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODF text character reference: {error}"))
    })? {
        let mut encoded = [0_u8; 4];
        return try_owned_string(
            character.encode_utf8(&mut encoded),
            "ODT text character reference",
        );
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODF text entity: {error}")))?;
    match name.as_ref() {
        "amp" => try_owned_string("&", "ODT text entity reference"),
        "lt" => try_owned_string("<", "ODT text entity reference"),
        "gt" => try_owned_string(">", "ODT text entity reference"),
        "quot" => try_owned_string("\"", "ODT text entity reference"),
        "apos" => try_owned_string("'", "ODT text entity reference"),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported ODF text entity '&{name};'"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::{Path, PathBuf};

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
        assert_eq!(
            TextElements::parse_paragraph_at(xml, 1)
                .unwrap()
                .unwrap()
                .text()
                .unwrap(),
            "Second paragraph"
        );
        assert!(TextElements::parse_paragraph_at(xml, 2).unwrap().is_none());
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

        let malformed_tail = format!(
            r#"<t:p xmlns:t="{}">Selected</t:p><t:p>unfinished"#,
            std::str::from_utf8(TEXT_NAMESPACE).unwrap()
        );
        assert!(TextElements::parse_paragraph_at(&malformed_tail, 0).is_err());

        let excessive_tail = format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{}"><t:p>Selected</t:p><t:p><t:s t:c="{}"/></t:p></o:text>"#,
            std::str::from_utf8(TEXT_NAMESPACE).unwrap(),
            MAX_SPACE_COUNT + 1
        );
        assert!(TextElements::parse_paragraph_at(&excessive_tail, 0).is_err());

        let duplicate_tail = format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:a="{0}" xmlns:b="{0}"><a:p>Selected</a:p><a:p a:style-name="one" b:style-name="two">Invalid</a:p></o:text>"#,
            std::str::from_utf8(TEXT_NAMESPACE).unwrap()
        );
        assert!(TextElements::parse_paragraph_at(&duplicate_tail, 0).is_err());
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

    // ========== 0216 discard-but-validate extraction pins ==========

    /// The pre-0216 `extract_text` implementation, kept as the parity
    /// oracle: retained owned blocks, `into_text` per block, '\n' join.
    fn oracle_extract_text(xml_content: &str) -> String {
        let mut blocks = parse_text_blocks_owned(xml_content)
            .expect("oracle parse must succeed for parity inputs")
            .into_iter();
        let Some(first) = blocks.next() else {
            return String::new();
        };
        let mut output = first.into_text();
        for block in blocks {
            output.push('\n');
            output.push_str(&block.into_text());
        }
        output
    }

    fn assert_extract_text_parity(xml: &str) {
        assert_eq!(
            TextElements::extract_text(xml).unwrap(),
            oracle_extract_text(xml),
            "discard-mode extract_text diverges from the retained path"
        );
    }

    #[test]
    fn discard_mode_matches_retained_path_across_block_shapes() {
        let text_ns = std::str::from_utf8(TEXT_NAMESPACE).unwrap();
        let fixtures = [
            // Plain paragraphs.
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p>First</t:p><t:p>Second</t:p></o:text>"#,
            // Heading, styled paragraph, entities, CDATA, whitespace controls,
            // general reference, empty block.
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:h t:outline-level="3">A &amp; <![CDATA[B]]></t:h><t:p t:style-name="Body">C<t:s t:c="2"/>D<t:tab/>E<t:line-break/>F&#x21;</t:p><t:p/></o:text>"#,
            // Lists.
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:list><t:list-item><t:p>Item 1</t:p></t:list-item><t:list-item><t:p>Item 2</t:p></t:list-item></t:list></o:text>"#,
            // Nested frame text box: the inner block completes before the
            // outer one, but start order must be preserved.
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><t:p>Outer<d:frame><d:text-box><t:p>Inner</t:p></d:text-box></d:frame></t:p></o:text>"#,
            // Attribute-carrying blocks: text/xlink/xml prefixes, unbound and
            // foreign-bound attribute names.
            r##"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:f="urn:example:foreign"><t:p t:style-name="B" xlink:href="#x" xml:id="i1" plain="p" f:extra="e">Text</t:p></o:text>"##,
            // No blocks at all.
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"></o:text>"#,
        ];
        for xml in fixtures {
            assert_extract_text_parity(xml);
        }
        // Tracked-change definitions stay excluded.
        let tracked = format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{text_ns}"><t:tracked-changes><t:changed-region><t:deletion><t:p>Deleted</t:p></t:deletion></t:changed-region></t:tracked-changes><t:p>Visible</t:p></o:text>"#
        );
        assert_extract_text_parity(&tracked);
        // Note bodies stay suppressed while citations remain.
        let note = format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{text_ns}"><t:p>Before<t:note t:note-class="footnote" t:id="n1"><t:note-citation>1</t:note-citation><t:note-body><t:p>Hidden</t:p></t:note-body></t:note>After</t:p></o:text>"#
        );
        assert_extract_text_parity(&note);
        // Ruby pronunciation stays excluded.
        let ruby = format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{text_ns}"><t:p>Before<t:ruby><t:ruby-base>漢字</t:ruby-base><t:ruby-text>かんじ</t:ruby-text></t:ruby>After</t:p></o:text>"#
        );
        assert_extract_text_parity(&ruby);
    }

    fn extract_text_error(xml: &str) -> String {
        match TextElements::extract_text(xml) {
            Err(error) => error.to_string(),
            Ok(text) => panic!("expected invalid-format error, got {text:?}"),
        }
    }

    fn oracle_extract_text_error(xml: &str) -> String {
        match parse_text_blocks_owned(xml) {
            Err(error) => error.to_string(),
            Ok(blocks) => panic!("expected invalid-format error, got {} blocks", blocks.len()),
        }
    }

    #[test]
    fn discard_mode_preserves_attribute_error_messages_and_precedence() {
        let text_ns = std::str::from_utf8(TEXT_NAMESPACE).unwrap();
        // Malformed attribute syntax.
        let malformed = format!(r#"<t:p xmlns:t="{text_ns}" broken>x</t:p>"#);
        let message = extract_text_error(&malformed);
        assert_eq!(message, oracle_extract_text_error(&malformed));
        assert!(message.starts_with("Invalid format: invalid ODF text attribute:"));
        // Malformed attribute in the second block, after retained text.
        let second_block = format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{text_ns}"><t:p>First</t:p><t:p broken>Second</t:p></o:text>"#
        );
        assert_eq!(
            extract_text_error(&second_block),
            oracle_extract_text_error(&second_block)
        );
        // Raw duplicate qualified name (iterator-level detection).
        let raw_duplicate =
            format!(r#"<t:p xmlns:t="{text_ns}" t:style-name="a" t:style-name="b">x</t:p>"#);
        let message = extract_text_error(&raw_duplicate);
        assert_eq!(message, oracle_extract_text_error(&raw_duplicate));
        assert!(message.starts_with("Invalid format: invalid ODF text attribute:"));
        // Duplicate resolved name through two prefixes (decode precedes the
        // duplicate error in both paths).
        let resolved_duplicate = format!(
            r#"<t:p xmlns:t="{text_ns}" xmlns:a="{text_ns}" xmlns:b="{text_ns}" a:style-name="one" b:style-name="two">x</t:p>"#
        );
        let message = extract_text_error(&resolved_duplicate);
        assert_eq!(message, oracle_extract_text_error(&resolved_duplicate));
        assert!(message.contains("duplicate ODF text attribute 'text:style-name'"));
        // Unknown namespace prefix, on both a paired and an empty block.
        let unknown = format!(r#"<t:p xmlns:t="{text_ns}" u:foo="1">x</t:p>"#);
        let message = extract_text_error(&unknown);
        assert_eq!(message, oracle_extract_text_error(&unknown));
        assert!(message.contains("unknown ODF text attribute namespace prefix 'u'"));
        let unknown_empty = format!(r#"<t:p xmlns:t="{text_ns}" u:foo="1"/>"#);
        assert_eq!(
            extract_text_error(&unknown_empty),
            oracle_extract_text_error(&unknown_empty)
        );
        // Undecodable attribute value.
        let bad_value = format!(r#"<t:p xmlns:t="{text_ns}" t:style-name="&#xD800;">x</t:p>"#);
        let message = extract_text_error(&bad_value);
        assert_eq!(message, oracle_extract_text_error(&bad_value));
        assert!(message.starts_with("Invalid format: invalid ODF text attribute value:"));
    }

    #[test]
    fn discard_mode_preserves_depth_and_structure_limits() {
        let text_ns = std::str::from_utf8(TEXT_NAMESPACE).unwrap();
        // Nesting beyond MAX_TEXT_DEPTH fails identically in both paths.
        let deep = format!(
            r#"<t:p xmlns:t="{text_ns}">{}x{}</t:p>"#,
            "<t:a>".repeat(MAX_TEXT_DEPTH),
            "</t:a>".repeat(MAX_TEXT_DEPTH),
        );
        assert_eq!(extract_text_error(&deep), oracle_extract_text_error(&deep));
        // Incomplete structure fails identically.
        let incomplete = format!(r#"<t:p xmlns:t="{text_ns}">unfinished"#);
        assert_eq!(
            extract_text_error(&incomplete),
            oracle_extract_text_error(&incomplete)
        );
    }

    // ========== 0225 namespace-memo differential pins ==========

    /// Drive both the memoized discard path and the direct-resolution
    /// retained oracle over `xml`, asserting identical per-block text or an
    /// identical error message.
    fn assert_memo_matches_direct_resolution(xml: &str) {
        let memoized = parse_text_block_texts(xml);
        let oracle = parse_text_blocks_owned(xml)
            .map(|blocks| blocks.into_iter().map(Block::into_text).collect::<Vec<_>>());
        match (memoized, oracle) {
            (Ok(memoized), Ok(oracle)) => assert_eq!(
                memoized, oracle,
                "memoized classification diverges from direct resolution"
            ),
            (Err(memoized), Err(oracle)) => assert_eq!(
                memoized.to_string(),
                oracle.to_string(),
                "memoized error diverges from direct resolution"
            ),
            (memoized, oracle) => panic!(
                "memoized/direct outcome mismatch: {memoized:?} vs {:?}",
                oracle.map(|blocks| blocks.len())
            ),
        }
    }

    /// Replay the bookkeeping of `parse_text_block_texts` event by event,
    /// asserting every memo verdict equals a fresh direct resolution against
    /// the live resolver. The tracker is maintained in parallel with the
    /// `NsReader` oracle — deferred pop at the top of the iteration, push
    /// before the classification — replicating the production loop, so a
    /// verdict match also pins tracker-vs-`NsReader` resolution parity.
    /// Panic-free here is not a concern: test-only.
    fn assert_memo_classification_matches(xml: &str) {
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        let mut tracker = BindingTracker::new();
        let mut pending_pop = false;
        let mut memo = TextNamespaceMemo::new();
        let mut content_version = 0u64;
        let mut declared_stack: Vec<bool> = Vec::new();
        let mut pop_removed_bindings = false;
        loop {
            if pending_pop {
                tracker.pop();
                pending_pop = false;
            }
            let event = reader
                .read_event_into(&mut buffer)
                .expect("test XML must parse");
            if pop_removed_bindings {
                content_version += 1;
                pop_removed_bindings = false;
            }
            match event {
                Event::Start(ref element) => {
                    tracker
                        .push(element)
                        .expect("test XML namespaces must be valid");
                    let declares = may_declare_namespace(element);
                    declared_stack.push(declares);
                    if declares {
                        content_version += 1;
                    }
                    assert_memo_hit_matches(&reader, &tracker, &mut memo, content_version, element);
                },
                Event::Empty(ref element) => {
                    tracker
                        .push(element)
                        .expect("test XML namespaces must be valid");
                    pending_pop = true;
                    if may_declare_namespace(element) {
                        content_version += 1;
                        pop_removed_bindings = true;
                    }
                    assert_memo_hit_matches(&reader, &tracker, &mut memo, content_version, element);
                },
                Event::End(_) => {
                    pending_pop = true;
                    if declared_stack.pop().unwrap_or(true) {
                        pop_removed_bindings = true;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
    }

    fn assert_memo_hit_matches(
        reader: &NsReader<&[u8]>,
        tracker: &BindingTracker,
        memo: &mut TextNamespaceMemo,
        content_version: u64,
        element: &BytesStart<'_>,
    ) {
        let direct = matches!(
            reader.resolver().resolve_prefix(element.name().prefix(), true),
            ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE
        );
        let memoized = memo
            .is_text_namespace(tracker, content_version, element)
            .expect("test XML fits the memo prefix buffer");
        assert_eq!(
            memoized,
            direct,
            "memo verdict diverges from direct resolution for {:?}",
            element.name()
        );
    }

    #[test]
    fn memoized_resolution_matches_direct_resolution_under_rebinding() {
        let text_ns = std::str::from_utf8(TEXT_NAMESPACE).unwrap();
        let office_ns = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        let fixtures: Vec<String> = vec![
            // Same prefix rebound to a foreign namespace at depth: the block
            // inside the scope is not text, the sibling after the scope
            // closes is (catches a memo stale across the deferred pop).
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}"><t:p>Before</t:p><o:wrapper xmlns:t="urn:example:foreign"><t:p>Shadowed</t:p></o:wrapper><t:p>After</t:p></o:text>"#
            ),
            // Nested rebinding: rebound to foreign, rebound back to the text
            // namespace at a deeper level, then restored outward.
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}"><o:a xmlns:t="urn:example:one"><t:p>S1</t:p><o:b xmlns:t="{text_ns}"><t:p>Inner</t:p></o:b><t:p>S2</t:p></o:a><t:p>Outer</t:p></o:text>"#
            ),
            // Unbinding via `xmlns:t=""` (resolution becomes `Unknown`), then
            // the outer declaration takes over again after the scope closes.
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}"><t:p>Before</t:p><o:wrapper xmlns:t=""><t:p>Shadowed</t:p></o:wrapper><t:p>After</t:p></o:text>"#
            ),
            // Default namespace makes prefix-less `<p>` a text block,
            // interleaved with prefixed blocks; a scope unsetting the default
            // suppresses prefix-less blocks inside it.
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}" xmlns="{text_ns}"><p>Plain</p><t:p>Prefixed</t:p><o:wrapper xmlns=""><p>NotText</p></o:wrapper><p>PlainAgain</p></o:text>"#
            ),
            // An `Empty` element carrying a declaration: its push and
            // deferred pop bracket the event itself, and the following
            // sibling must resolve against the restored outer binding.
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}"><o:marker xmlns:t="urn:example:foreign"/><t:p>After</t:p></o:text>"#
            ),
            // An attribute *value* containing the substring `xmlns` forces a
            // conservative invalidation without declaring anything.
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}"><t:p t:style-name="xmlns:shadow">A</t:p><t:p>B</t:p></o:text>"#
            ),
            // Foreign-prefix blocks before and after a rebinding scope.
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}" xmlns:f="urn:example:foreign"><f:p>Foreign</f:p><t:p>Text1</t:p><o:w xmlns:t="urn:example:foreign"><t:p>Shadowed</t:p></o:w><f:p>Foreign2</f:p><t:p>Text2</t:p></o:text>"#
            ),
            // Error parity: an unknown attribute prefix on a block after a
            // rebinding scope closes fails identically in both paths.
            format!(
                r#"<o:text xmlns:o="{office_ns}" xmlns:t="{text_ns}"><o:w xmlns:t="urn:example:foreign"><t:p>S</t:p></o:w><t:p u:foo="1">x</t:p></o:text>"#
            ),
        ];
        let expected: [&[&str]; 7] = [
            &["Before", "After"],
            &["Inner", "Outer"],
            &["Before", "After"],
            &["Plain", "Prefixed", "PlainAgain"],
            &["After"],
            &["A", "B"],
            &["Text1", "Text2"],
        ];
        for (index, xml) in fixtures.iter().enumerate() {
            assert_memo_matches_direct_resolution(xml);
            assert_memo_classification_matches(xml);
            if let Some(expected) = expected.get(index) {
                assert_eq!(
                    parse_text_block_texts(xml).unwrap(),
                    *expected,
                    "unexpected extracted text for fixture {index}"
                );
            }
        }
    }

    // ========== 0227 binding-tracker differential pins ==========

    /// Namespace-error and attribute-resolution parity between the
    /// tracker-driven discard path (change 0227) and the `NsReader`-driven
    /// retained oracle: identical outcomes and byte-identical error strings.
    #[test]
    fn tracker_driven_path_matches_ns_reader_oracle_on_adversarial_namespaces() {
        let text_ns = std::str::from_utf8(TEXT_NAMESPACE).unwrap();
        let xml_ns = "http://www.w3.org/XML/1998/namespace";
        let xmlns_ns = "http://www.w3.org/2000/xmlns/";

        // --- Reserved-prefix and reserved-URI errors (push failures) ---
        let error_fixtures: Vec<String> = vec![
            // Declaring the `xmlns` prefix itself.
            format!(r#"<t:p xmlns:t="{text_ns}" xmlns:xmlns="urn:example:x">x</t:p>"#),
            // Binding `xml` to a foreign URI.
            format!(r#"<t:p xmlns:t="{text_ns}" xmlns:xml="urn:example:x">x</t:p>"#),
            // Binding another prefix to the reserved xml URI.
            format!(r#"<t:p xmlns:t="{text_ns}" xmlns:q="{xml_ns}">x</t:p>"#),
            // Binding a prefix to the reserved xmlns URI.
            format!(r#"<t:p xmlns:t="{text_ns}" xmlns:q="{xmlns_ns}">x</t:p>"#),
            // The same failures on a nested element mid-stream, after a
            // successful block: the push error preempts the event exactly
            // where `NsReader`'s read error did.
            format!(
                r#"<t:p xmlns:t="{text_ns}">A</t:p><t:p xmlns:t="{text_ns}"><t:span xmlns:xmlns="urn:example:x">B</t:span></t:p>"#
            ),
            // A namespace error on an `Empty` element.
            format!(r#"<t:p xmlns:t="{text_ns}">A</t:p><t:s xmlns:xml="urn:example:x"/>"#),
            // Malformed declaration mid-scan (`xmlns:t` with no value): the
            // tokenizer and both readers reject it identically.
            format!(r#"<t:p xmlns:t="{text_ns}"><t:s xmlns:t=">#</t:p>"#),
        ];
        for xml in &error_fixtures {
            assert_memo_matches_direct_resolution(xml);
            let message = extract_text_error(xml);
            assert!(
                message.starts_with("Invalid format: invalid ODF text XML:"),
                "unexpected error for {xml}: {message}"
            );
        }

        // --- Declaration-limit parity: 256 declarations pass, 257 fail ---
        // (`xmlns:t` below accounts for one declaration on the tag).
        let declarations = |count: usize| {
            (0..count)
                .map(|index| format!(r#"xmlns:d{index}="urn:example:{index}""#))
                .collect::<Vec<_>>()
                .join(" ")
        };
        let within_limit = format!(r#"<t:p xmlns:t="{text_ns}" {}>x</t:p>"#, declarations(255));
        assert_memo_matches_direct_resolution(&within_limit);
        assert_eq!(parse_text_block_texts(&within_limit).unwrap(), ["x"]);
        let over_limit = format!(r#"<t:p xmlns:t="{text_ns}" {}>x</t:p>"#, declarations(256));
        assert_memo_matches_direct_resolution(&over_limit);
        assert!(
            extract_text_error(&over_limit).starts_with("Invalid format: invalid ODF text XML:")
        );

        // --- Benign reserved bindings and attribute resolution ---
        let ok_fixtures: Vec<(String, Vec<&str>)> = vec![
            // Rebinding `xml` to its reserved URI is a no-op.
            (
                format!(r#"<t:p xmlns:t="{text_ns}" xmlns:xml="{xml_ns}">x</t:p>"#),
                vec!["x"],
            ),
            // `text:c` on `text:s` resolves through the tracker's
            // `resolve_attribute` under a second prefix bound to the text
            // namespace.
            (
                format!(r#"<t:p xmlns:t="{text_ns}" xmlns:a="{text_ns}">A<t:s a:c="3"/>B</t:p>"#),
                vec!["A   B"],
            ),
            // An unprefixed `c` attribute does NOT fall back to the default
            // namespace (`use_default = false` for attributes): `text:s`
            // contributes the default single space.
            (
                format!(r#"<t:p xmlns:t="{text_ns}" xmlns="{text_ns}">A<t:s c="5"/>B</t:p>"#),
                vec!["A B"],
            ),
            // An emptied binding (`xmlns:t=""`) shadows the outer prefix
            // inside the scope: the inner `t:s` element name itself resolves
            // to `Unknown`, so the whole control is skipped rather than
            // counted.
            (
                format!(
                    r#"<t:p xmlns:t="{text_ns}">A<t:s><t:inner xmlns:t=""><t:s t:c="9"/></t:inner></t:s>B</t:p>"#
                ),
                vec!["A B"],
            ),
        ];
        for (xml, expected) in &ok_fixtures {
            assert_memo_matches_direct_resolution(xml);
            assert_eq!(
                parse_text_block_texts(xml).unwrap(),
                *expected,
                "unexpected extracted text for {xml}"
            );
        }
    }

    #[test]
    fn memoized_classification_matches_direct_resolution_on_odt_corpus() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let mut files = Vec::new();
        collect_odt_corpus(&root, &mut files);
        files.sort();
        assert!(!files.is_empty(), "no .odt corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(xml) = odt_content_xml(path) else {
                continue;
            };
            assert_memo_matches_direct_resolution(&xml);
            assert_memo_classification_matches(&xml);
            compared += 1;
        }
        assert!(compared > 0, "no .odt corpus fixtures yielded content.xml");
    }

    fn collect_odt_corpus(directory: &Path, files: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_odt_corpus(&path, files);
            } else if path
                .extension()
                .is_some_and(|extension| extension == "odt" || extension == "fodt")
            {
                files.push(path);
            }
        }
    }

    fn odt_content_xml(path: &Path) -> Option<String> {
        let bytes = std::fs::read(path).ok()?;
        if path
            .extension()
            .is_some_and(|extension| extension == "fodt")
        {
            return String::from_utf8(bytes).ok();
        }
        let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
        let entry = reader.read("content.xml").ok()?;
        String::from_utf8(entry).ok()
    }
}
