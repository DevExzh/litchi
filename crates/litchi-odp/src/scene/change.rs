//! Scene changes and their checked byte-splice plans.

use std::collections::BTreeMap;
use std::ops::Range;

use litchi_core::{Error, Result};

use super::markup::{Encoded, StartTag, encode, escape_attribute, open_tag, template};
use super::{Object, ObjectKind, Page, Paragraph, RunKind, Scene, Story};

/// Which object list of a page an [`ObjectRef`] indexes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Scope {
    /// The slide's drawing objects.
    Page,
    /// The objects of the slide's notes page.
    Notes,
}

/// An object addressed by its page, list and position path (a top-level
/// position, then positions inside nested groups).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ObjectRef {
    /// Zero-based page position.
    pub page: usize,
    /// The page's object list.
    pub scope: Scope,
    /// Zero-based positions: the top-level object, then group members.
    pub path: Vec<usize>,
}

impl ObjectRef {
    /// A top-level object of a slide.
    #[must_use]
    pub fn new(page: usize, object: usize) -> Self {
        Self {
            page,
            scope: Scope::Page,
            path: vec![object],
        }
    }
}

/// A story: an object's text, or one table cell's text.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct StoryRef {
    /// The object holding the story or the table.
    pub object: ObjectRef,
    /// A table cell as zero-based `(row, cell)` positions.
    pub cell: Option<(usize, usize)>,
}

impl StoryRef {
    /// An object's own story.
    #[must_use]
    pub fn object(object: ObjectRef) -> Self {
        Self { object, cell: None }
    }
}

/// An object attribute [`Change::SetAttribute`] writes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Attribute {
    /// `draw:name`.
    Name,
    /// `svg:x`, an ODF length.
    X,
    /// `svg:y`, an ODF length.
    Y,
    /// `svg:width`, an ODF length.
    Width,
    /// `svg:height`, an ODF length.
    Height,
}

/// A new text box for [`Change::InsertTextBox`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewTextBox {
    /// `draw:name`.
    pub name: Option<String>,
    /// `svg:x`, `svg:y`, `svg:width` and `svg:height` as ODF lengths.
    pub geometry: [String; 4],
    /// The text; `\n` separates paragraphs.
    pub text: String,
}

/// A new blank page for [`Change::InsertPage`].
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct NewPage {
    /// `draw:name`; a unique `page<N>` name when absent.
    pub name: Option<String>,
    /// The page whose drawing-page style, master page and layout the new
    /// page copies; the preceding page (or the first) when absent.
    pub like: Option<usize>,
    /// The master page for a presentation without pages to copy from.
    pub master: Option<String>,
}

/// One edit of a scene, resolved against the scene it is planned on.
///
/// Positions in one [`crate::edit::Edit::splice_scene`] call all refer to
/// the scene before the call; two changes in one call must not touch the
/// same bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    /// Replace a story with plain paragraphs (`\n` separates them), keeping
    /// the first paragraph's style and list wrappers.
    SetStoryText {
        /// The story.
        story: StoryRef,
        /// The new text.
        text: String,
    },
    /// Replace one paragraph's content with plain text.
    SetParagraphText {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// The new text; `\n` is a line break, `\t` a tab.
        text: String,
    },
    /// Replace one run's content with plain text.
    SetRunText {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// Zero-based run position.
        run: usize,
        /// The new text.
        text: String,
    },
    /// Insert a paragraph at a final position, styled like its neighbour.
    InsertParagraph {
        /// The story.
        story: StoryRef,
        /// Zero-based final position.
        index: usize,
        /// The text.
        text: String,
    },
    /// Remove one paragraph (with a list item or list it alone fills).
    RemoveParagraph {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
    },
    /// Remove one run.
    RemoveRun {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// Zero-based run position.
        run: usize,
    },
    /// Set or remove an object attribute.
    SetAttribute {
        /// The object.
        object: ObjectRef,
        /// The attribute.
        attribute: Attribute,
        /// The new lexical value; `None` removes it (names only).
        value: Option<String>,
    },
    /// Rename a page.
    SetPageName {
        /// Zero-based page position.
        page: usize,
        /// The new unique name.
        name: String,
    },
    /// Append a text box to a slide's objects.
    InsertTextBox {
        /// Zero-based page position.
        page: usize,
        /// The text box.
        text_box: NewTextBox,
    },
    /// Remove an object.
    RemoveObject {
        /// The object.
        object: ObjectRef,
    },
    /// Move a top-level object to a final z-order position.
    MoveObject {
        /// The object.
        object: ObjectRef,
        /// Zero-based final position among the page's objects.
        to: usize,
    },
    /// Exchange two top-level objects of one page.
    SwapObjects {
        /// One object.
        first: ObjectRef,
        /// The other object, on the same page.
        second: ObjectRef,
    },
    /// Insert a blank page at a final position.
    InsertPage {
        /// Zero-based final position.
        index: usize,
        /// The page.
        page: NewPage,
    },
    /// Remove a page; the last page cannot be removed.
    RemovePage {
        /// Zero-based page position.
        page: usize,
    },
    /// Move a page to a final position.
    MovePage {
        /// Zero-based page position.
        page: usize,
        /// Zero-based final position.
        to: usize,
    },
    /// Exchange two pages.
    SwapPages {
        /// One page.
        first: usize,
        /// The other page.
        second: usize,
    },
}

/// A checked replacement fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Fragment {
    /// Escaped character data.
    Text(String),
    /// Balanced markup.
    Markup(String),
    /// One rewritten start tag.
    StartTag(String),
    /// Source markup relocated from another range of the same part.
    Relocated(Range<usize>),
    /// Removal.
    Delete,
}

impl Fragment {
    pub(crate) fn bytes<'a>(&'a self, xml: &'a str) -> &'a str {
        match self {
            Self::Text(text) | Self::Markup(text) | Self::StartTag(text) => text,
            Self::Relocated(range) => &xml[range.clone()],
            Self::Delete => "",
        }
    }
}

/// The splices of one change list, sorted and disjoint.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct Plan {
    pub(crate) edits: Vec<(Range<usize>, Fragment)>,
}

impl Plan {
    /// Whether the plan changes nothing.
    pub(crate) fn is_empty(&self) -> bool {
        self.edits.is_empty()
    }

    /// Assemble the changed text.
    pub(crate) fn assemble(&self, xml: &str) -> String {
        let added: usize = self
            .edits
            .iter()
            .map(|(_, fragment)| fragment.bytes(xml).len())
            .sum();
        let mut out = String::with_capacity(xml.len() + added);
        let mut cursor = 0;
        for (range, fragment) in &self.edits {
            out.push_str(&xml[cursor..range.start]);
            out.push_str(fragment.bytes(xml));
            cursor = range.end;
        }
        out.push_str(&xml[cursor..]);
        out
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unsupported(message: impl Into<String>) -> Error {
    Error::Unsupported(message.into())
}

fn missing(what: &str) -> Error {
    Error::InvalidFormat(format!("the ODP scene has no such {what}"))
}

struct Planner<'a> {
    xml: &'a str,
    scene: &'a Scene,
    edits: Vec<(Range<usize>, Fragment)>,
    /// Attribute changes per start tag: range, self-closing, changes.
    tags: BTreeMap<usize, (Range<usize>, Vec<(String, Option<String>)>)>,
    /// New page names chosen in this plan.
    names: Vec<String>,
}

/// Plan `changes` against `scene`, the index of `xml`.
pub(crate) fn plan(xml: &str, scene: &Scene, changes: &[Change]) -> Result<Plan> {
    if scene.len != xml.len() {
        return Err(invalid("the ODP scene does not index this content"));
    }
    let mut planner = Planner {
        xml,
        scene,
        edits: Vec::new(),
        tags: BTreeMap::new(),
        names: Vec::new(),
    };
    for change in changes {
        planner.change(change)?;
    }
    let mut edits = planner.edits;
    for (_, (range, attributes)) in planner.tags {
        let raw = &xml[range.clone()];
        let tag = StartTag::parse(raw)?;
        let rewritten = tag.write(&attributes, raw.ends_with("/>"), false);
        if rewritten != raw {
            edits.push((range, Fragment::StartTag(rewritten)));
        }
    }
    // Insertions at one position keep their call order; a stable sort on
    // the start keeps it.
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    let mut merged: Vec<(Range<usize>, Fragment)> = Vec::with_capacity(edits.len());
    for (range, fragment) in edits {
        if let Some((last, last_fragment)) = merged.last_mut() {
            let both_insertions =
                last.start == last.end && range.start == range.end && last.start == range.start;
            if both_insertions {
                if matches!(last_fragment, Fragment::Relocated(_))
                    || matches!(fragment, Fragment::Relocated(_))
                {
                    return Err(invalid("two ODP scene changes insert at the same position"));
                }
                let joined = format!("{}{}", last_fragment.bytes(xml), fragment.bytes(xml));
                *last_fragment = Fragment::Markup(joined);
                continue;
            }
            if range.start < last.end {
                return Err(invalid("two ODP scene changes touch the same markup"));
            }
        }
        merged.push((range, fragment));
    }
    Ok(Plan { edits: merged })
}

impl<'a> Planner<'a> {
    fn prefix(&self, prefix: Option<&'a String>, namespace: &str) -> Result<&'a str> {
        prefix.map(String::as_str).ok_or_else(|| {
            unsupported(format!(
                "the ODP content root does not bind the {namespace} namespace"
            ))
        })
    }

    fn text_prefix(&self) -> Result<&'a str> {
        self.prefix(self.scene.prefixes.text.as_ref(), "text")
    }

    fn page(&self, page: usize) -> Result<&'a Page> {
        self.scene.pages.get(page).ok_or_else(|| missing("page"))
    }

    fn object(&self, object: &ObjectRef) -> Result<&'a Object> {
        self.scene.object(object).ok_or_else(|| missing("object"))
    }

    fn story(&self, story: &StoryRef) -> Result<&'a Story> {
        self.scene.story(story).ok_or_else(|| missing("text story"))
    }

    fn paragraph(&self, story: &StoryRef, paragraph: usize) -> Result<&'a Paragraph> {
        self.story(story)?
            .paragraphs
            .get(paragraph)
            .ok_or_else(|| missing("paragraph"))
    }

    fn push(&mut self, range: Range<usize>, fragment: Fragment) {
        self.edits.push((range, fragment));
    }

    fn change(&mut self, change: &Change) -> Result<()> {
        match change {
            Change::SetStoryText { story, text } => self.set_story_text(story, text),
            Change::SetParagraphText {
                story,
                paragraph,
                text,
            } => self.set_paragraph_text(story, *paragraph, text),
            Change::SetRunText {
                story,
                paragraph,
                run,
                text,
            } => self.set_run_text(story, *paragraph, *run, text),
            Change::InsertParagraph { story, index, text } => {
                self.insert_paragraph(story, *index, text)
            },
            Change::RemoveParagraph { story, paragraph } => {
                let removal = self.paragraph(story, *paragraph)?.removal.clone();
                self.push(removal, Fragment::Delete);
                Ok(())
            },
            Change::RemoveRun {
                story,
                paragraph,
                run,
            } => {
                let paragraph = self.paragraph(story, *paragraph)?;
                let run = paragraph.runs.get(*run).ok_or_else(|| missing("run"))?;
                self.push(run.range.clone(), Fragment::Delete);
                Ok(())
            },
            Change::SetAttribute {
                object,
                attribute,
                value,
            } => self.set_attribute(object, *attribute, value.as_deref()),
            Change::SetPageName { page, name } => self.set_page_name(*page, name),
            Change::InsertTextBox { page, text_box } => self.insert_text_box(*page, text_box),
            Change::RemoveObject { object } => self.remove_object(object),
            Change::MoveObject { object, to } => self.move_object(object, *to),
            Change::SwapObjects { first, second } => self.swap_objects(first, second),
            Change::InsertPage { index, page } => self.insert_page(*index, page),
            Change::RemovePage { page } => self.remove_page(*page),
            Change::MovePage { page, to } => self.move_page(*page, *to),
            Change::SwapPages { first, second } => self.swap_pages(*first, *second),
        }
    }

    /// The paragraphs of `text` (one per line) in the style of `story`'s
    /// first paragraph, its list wrappers included.
    fn story_markup(&self, story: &Story, text: &str) -> Result<String> {
        let prefix = self.text_prefix()?;
        let (wrappers, start, end) = match story.paragraphs.first() {
            Some(first) => (
                first.wrappers.as_slice(),
                template(&self.xml[first.start_tag.clone()])?,
                paragraph_qname(&self.xml[first.start_tag.clone()])?,
            ),
            None => (&[][..], format!("<{prefix}:p>"), format!("{prefix}:p")),
        };
        let mut out = String::with_capacity(text.len() + 64);
        let (outer, inner) = match wrappers.split_first() {
            Some((outer, inner)) => (Some(outer), inner),
            None => (None, &[][..]),
        };
        if let Some(outer) = outer {
            out.push_str(&template(&self.xml[outer.start_tag.clone()])?);
        }
        for line in text.split('\n') {
            for wrapper in inner {
                out.push_str(&template(&self.xml[wrapper.start_tag.clone()])?);
            }
            out.push_str(&start);
            out.push_str(&encode(prefix, line)?.markup);
            out.push_str("</");
            out.push_str(&end);
            out.push('>');
            for wrapper in inner.iter().rev() {
                out.push_str("</");
                out.push_str(&wrapper.qname);
                out.push('>');
            }
        }
        if let Some(outer) = outer {
            out.push_str("</");
            out.push_str(&outer.qname);
            out.push('>');
        }
        Ok(out)
    }

    fn set_story_text(&mut self, reference: &StoryRef, text: &str) -> Result<()> {
        let story = self.story(reference)?;
        if story.text() == text {
            return Ok(());
        }
        let markup = self.story_markup(story, text)?;
        if let Some(empty) = &story.empty {
            return self.expand(empty.clone(), &markup);
        }
        if !story.replaceable {
            return Err(unsupported(
                "this ODP text story holds content other than paragraphs and lists",
            ));
        }
        self.push(story.range.clone(), Fragment::Markup(markup));
        Ok(())
    }

    fn set_paragraph_text(&mut self, story: &StoryRef, index: usize, text: &str) -> Result<()> {
        let paragraph = self.paragraph(story, index)?;
        if paragraph.text == text {
            return Ok(());
        }
        let prefix = self.text_prefix()?;
        let encoded = encode(prefix, text)?;
        self.replace_content(
            paragraph.range.clone(),
            paragraph.start_tag.clone(),
            paragraph.content.clone(),
            paragraph.empty,
            &encoded,
        )
    }

    /// Replace the content of an element (a paragraph or a span) with
    /// encoded text: a character-data splice when possible, else the whole
    /// element.
    fn replace_content(
        &mut self,
        range: Range<usize>,
        start_tag: Range<usize>,
        content: Range<usize>,
        empty: bool,
        encoded: &Encoded,
    ) -> Result<()> {
        if !empty && encoded.plain {
            let fragment = if encoded.markup.is_empty() {
                Fragment::Delete
            } else if encoded.markup.trim().is_empty() {
                return self.replace_element(range, start_tag, encoded);
            } else {
                Fragment::Text(encoded.markup.clone())
            };
            if content.is_empty() && fragment == Fragment::Delete {
                return Ok(());
            }
            self.push(content, fragment);
            return Ok(());
        }
        self.replace_element(range, start_tag, encoded)
    }

    fn replace_element(
        &mut self,
        range: Range<usize>,
        start_tag: Range<usize>,
        encoded: &Encoded,
    ) -> Result<()> {
        let raw = &self.xml[start_tag];
        let name = StartTag::parse(raw)?.name.to_owned();
        let mut markup = open_tag(raw)?;
        markup.push_str(&encoded.markup);
        markup.push_str("</");
        markup.push_str(&name);
        markup.push('>');
        self.push(range, Fragment::Markup(markup));
        Ok(())
    }

    fn set_run_text(
        &mut self,
        story: &StoryRef,
        index: usize,
        run: usize,
        text: &str,
    ) -> Result<()> {
        let paragraph = self.paragraph(story, index)?;
        let selected = paragraph.runs.get(run).ok_or_else(|| missing("run"))?;
        if selected.text == text {
            return Ok(());
        }
        let prefix = self.text_prefix()?;
        let encoded = encode(prefix, text)?;
        match selected.kind {
            RunKind::Span | RunKind::Link => {
                let tag_end = if selected.content.start > selected.range.start {
                    selected.content.start
                } else {
                    selected.range.end
                };
                let empty = selected.content.start == selected.range.end;
                self.replace_content(
                    selected.range.clone(),
                    selected.range.start..tag_end,
                    selected.content.clone(),
                    empty,
                    &encoded,
                )
            },
            RunKind::Text => {
                if encoded.plain && !encoded.markup.trim().is_empty() {
                    self.push(selected.range.clone(), Fragment::Text(encoded.markup));
                    return Ok(());
                }
                if encoded.markup.is_empty() {
                    self.push(selected.range.clone(), Fragment::Delete);
                    return Ok(());
                }
                // Rebuild the paragraph around the run.
                let raw = &self.xml[paragraph.start_tag.clone()];
                let name = StartTag::parse(raw)?.name.to_owned();
                let mut markup = open_tag(raw)?;
                markup.push_str(&self.xml[paragraph.content.start..selected.range.start]);
                markup.push_str(&encoded.markup);
                markup.push_str(&self.xml[selected.range.end..paragraph.content.end]);
                markup.push_str("</");
                markup.push_str(&name);
                markup.push('>');
                self.push(paragraph.range.clone(), Fragment::Markup(markup));
                Ok(())
            },
            RunKind::Field => Err(unsupported("ODP field runs are computed and cannot be set")),
        }
    }

    fn insert_paragraph(&mut self, reference: &StoryRef, index: usize, text: &str) -> Result<()> {
        let story = self.story(reference)?;
        let prefix = self.text_prefix()?;
        let content = encode(prefix, text)?.markup;
        if index > story.paragraphs.len() {
            return Err(invalid("the ODP paragraph position is outside the story"));
        }
        let Some(neighbour) = story
            .paragraphs
            .get(index.saturating_sub(1))
            .or_else(|| story.paragraphs.first())
        else {
            let markup = format!("<{prefix}:p>{content}</{prefix}:p>");
            if let Some(empty) = &story.empty {
                return self.expand(empty.clone(), &markup);
            }
            if story.insertion == usize::MAX {
                return Err(unsupported("this ODP text story has no insertion point"));
            }
            self.push(story.insertion..story.insertion, Fragment::Markup(markup));
            return Ok(());
        };
        let after = index > 0;
        let p_start = template(&self.xml[neighbour.start_tag.clone()])?;
        let p_name = paragraph_qname(&self.xml[neighbour.start_tag.clone()])?;
        let mut markup = String::new();
        let item = neighbour
            .wrappers
            .iter()
            .rev()
            .find(|wrapper| wrapper.item)
            .zip(neighbour.item.clone());
        let at = match item {
            Some((wrapper, range)) => {
                markup.push_str(&template(&self.xml[wrapper.start_tag.clone()])?);
                markup.push_str(&p_start);
                markup.push_str(&content);
                markup.push_str(&format!("</{p_name}></{}>", wrapper.qname));
                if after { range.end } else { range.start }
            },
            None => {
                markup.push_str(&p_start);
                markup.push_str(&content);
                markup.push_str(&format!("</{p_name}>"));
                if after {
                    neighbour.range.end
                } else {
                    neighbour.range.start
                }
            },
        };
        self.push(at..at, Fragment::Markup(markup));
        Ok(())
    }

    /// Replace a self-closing element with its open form holding `content`.
    fn expand(&mut self, element: Range<usize>, content: &str) -> Result<()> {
        let raw = &self.xml[element.clone()];
        let name = StartTag::parse(raw)?.name.to_owned();
        let mut markup = open_tag(raw)?;
        markup.push_str(content);
        markup.push_str("</");
        markup.push_str(&name);
        markup.push('>');
        self.push(element, Fragment::Markup(markup));
        Ok(())
    }

    fn tag(&mut self, range: &Range<usize>, key: String, value: Option<String>) {
        let entry = self
            .tags
            .entry(range.start)
            .or_insert_with(|| (range.clone(), Vec::new()));
        if let Some(existing) = entry.1.iter_mut().find(|(name, _)| *name == key) {
            existing.1 = value;
        } else {
            entry.1.push((key, value));
        }
    }

    fn set_attribute(
        &mut self,
        reference: &ObjectRef,
        attribute: Attribute,
        value: Option<&str>,
    ) -> Result<()> {
        let object = self.object(reference)?;
        let (prefix, local, current) = match attribute {
            Attribute::Name => (
                self.prefix(self.scene.prefixes.draw.as_ref(), "drawing")?,
                "name",
                object.name.as_deref(),
            ),
            geometry => {
                if matches!(object.kind, ObjectKind::Line | ObjectKind::Connector) {
                    return Err(unsupported(
                        "ODP lines and connectors are placed by their end points",
                    ));
                }
                if value.is_none() {
                    return Err(unsupported("ODP object geometry cannot be removed"));
                }
                let (local, current) = match geometry {
                    Attribute::X => ("x", object.x.as_deref()),
                    Attribute::Y => ("y", object.y.as_deref()),
                    Attribute::Width => ("width", object.width.as_deref()),
                    _ => ("height", object.height.as_deref()),
                };
                (
                    self.prefix(self.scene.prefixes.svg.as_ref(), "SVG")?,
                    local,
                    current,
                )
            },
        };
        if current == value {
            return Ok(());
        }
        let range = object.start_tag.clone();
        self.tag(
            &range,
            format!("{prefix}:{local}"),
            value.map(str::to_owned),
        );
        Ok(())
    }

    fn set_page_name(&mut self, index: usize, name: &str) -> Result<()> {
        let page = self.page(index)?;
        if page.name.as_deref() == Some(name) {
            return Ok(());
        }
        if name.is_empty() {
            return Err(invalid("an ODP page name cannot be empty"));
        }
        if self
            .scene
            .pages
            .iter()
            .enumerate()
            .any(|(other, page)| other != index && page.name.as_deref() == Some(name))
            || self.names.iter().any(|taken| taken == name)
        {
            return Err(invalid("another ODP page already has this name"));
        }
        if let Some(old) = &page.name
            && referenced(self.xml, old, &page.start_tag)
        {
            return Err(unsupported(
                "the ODP page name is referenced by links or custom shows",
            ));
        }
        let prefix = self.prefix(self.scene.prefixes.draw.as_ref(), "drawing")?;
        let range = page.start_tag.clone();
        self.names.push(name.to_owned());
        self.tag(&range, format!("{prefix}:name"), Some(name.to_owned()));
        Ok(())
    }

    fn insert_text_box(&mut self, index: usize, text_box: &NewTextBox) -> Result<()> {
        let page = self.page(index)?;
        let draw = self.prefix(self.scene.prefixes.draw.as_ref(), "drawing")?;
        let svg = self.prefix(self.scene.prefixes.svg.as_ref(), "SVG")?;
        let story = Story {
            range: 0..0,
            paragraphs: Vec::new(),
            replaceable: true,
            insertion: 0,
            empty: None,
        };
        let paragraphs = self.story_markup(&story, &text_box.text)?;
        let mut markup = format!("<{draw}:frame");
        if let Some(name) = &text_box.name {
            markup.push_str(&format!(" {draw}:name=\"{}\"", escape_attribute(name)));
        }
        for (local, value) in ["x", "y", "width", "height"].iter().zip(&text_box.geometry) {
            markup.push_str(&format!(" {svg}:{local}=\"{}\"", escape_attribute(value)));
        }
        markup.push_str(&format!(
            "><{draw}:text-box>{paragraphs}</{draw}:text-box></{draw}:frame>"
        ));
        if page.empty {
            return self.expand(page.range.clone(), &markup);
        }
        let at = page
            .insertion
            .ok_or_else(|| unsupported("this ODP page has no insertion point"))?;
        self.push(at..at, Fragment::Markup(markup));
        Ok(())
    }

    fn objects(&self, reference: &ObjectRef) -> Result<&'a [Object]> {
        let page = self.page(reference.page)?;
        let mut objects = match reference.scope {
            Scope::Page => page.objects.as_slice(),
            Scope::Notes => {
                &page
                    .notes
                    .as_ref()
                    .ok_or_else(|| missing("notes page"))?
                    .objects
            },
        };
        let Some((_, parents)) = reference.path.split_last() else {
            return Err(missing("object"));
        };
        for &index in parents {
            objects = &objects.get(index).ok_or_else(|| missing("group"))?.children;
        }
        Ok(objects)
    }

    fn remove_object(&mut self, reference: &ObjectRef) -> Result<()> {
        if reference.scope == Scope::Notes {
            return Err(unsupported("ODP notes-page objects cannot be removed"));
        }
        let object = self.object(reference)?;
        if let Some(id) = &object.id
            && referenced(self.xml, id, &object.start_tag)
        {
            return Err(unsupported(
                "the ODP object is referenced by an animation or another object",
            ));
        }
        self.push(object.range.clone(), Fragment::Delete);
        Ok(())
    }

    fn move_object(&mut self, reference: &ObjectRef, to: usize) -> Result<()> {
        let objects = self.objects(reference)?;
        let from = *reference.path.last().ok_or_else(|| missing("object"))?;
        let object = objects.get(from).ok_or_else(|| missing("object"))?;
        if to >= objects.len() {
            return Err(invalid("the ODP object position is outside the page"));
        }
        if to == from {
            return Ok(());
        }
        let at = if to > from {
            objects[to].range.end
        } else {
            objects[to].range.start
        };
        self.push(object.range.clone(), Fragment::Delete);
        self.push(at..at, Fragment::Relocated(object.range.clone()));
        Ok(())
    }

    fn swap_objects(&mut self, first: &ObjectRef, second: &ObjectRef) -> Result<()> {
        if first.page != second.page
            || first.scope != second.scope
            || first.path.len() != second.path.len()
            || first.path[..first.path.len() - 1] != second.path[..second.path.len() - 1]
        {
            return Err(unsupported(
                "ODP objects can be exchanged within one list only",
            ));
        }
        if first == second {
            return Ok(());
        }
        let left = self.object(first)?.range.clone();
        let right = self.object(second)?.range.clone();
        self.push(left.clone(), Fragment::Relocated(right.clone()));
        self.push(right, Fragment::Relocated(left));
        Ok(())
    }

    fn unique_page_name(&self) -> String {
        let mut number = self.scene.pages.len() + 1;
        loop {
            let name = format!("page{number}");
            if !self
                .scene
                .pages
                .iter()
                .any(|page| page.name.as_deref() == Some(name.as_str()))
                && !self.names.contains(&name)
            {
                return name;
            }
            number += 1;
        }
    }

    fn insert_page(&mut self, index: usize, page: &NewPage) -> Result<()> {
        let count = self.scene.pages.len();
        if index > count {
            return Err(invalid("the ODP page position is outside the presentation"));
        }
        let like = page
            .like
            .unwrap_or_else(|| index.saturating_sub(1).min(count.saturating_sub(1)));
        let model = self.scene.pages.get(like);
        let master = match (page.master.as_deref(), model) {
            (Some(master), _) => master,
            (None, Some(model)) => model
                .master_page
                .as_deref()
                .ok_or_else(|| unsupported("the model ODP page names no master page"))?,
            (None, None) => {
                return Err(unsupported(
                    "a new ODP page needs a master page or an existing page to copy from",
                ));
            },
        };
        let draw = self.prefix(self.scene.prefixes.draw.as_ref(), "drawing")?;
        let name = match &page.name {
            Some(name) => {
                if self
                    .scene
                    .pages
                    .iter()
                    .any(|page| page.name.as_deref() == Some(name))
                    || self.names.iter().any(|taken| taken == name)
                {
                    return Err(invalid("another ODP page already has this name"));
                }
                name.clone()
            },
            None => self.unique_page_name(),
        };
        let mut markup = format!("<{draw}:page {draw}:name=\"{}\"", escape_attribute(&name));
        if let Some(style) = model.and_then(|model| model.style_name.as_ref()) {
            markup.push_str(&format!(
                " {draw}:style-name=\"{}\"",
                escape_attribute(style)
            ));
        }
        markup.push_str(&format!(
            " {draw}:master-page-name=\"{}\"",
            escape_attribute(master)
        ));
        if let (Some(layout), Some(presentation)) = (
            model.and_then(|model| model.layout.as_ref()),
            self.scene.prefixes.presentation.as_ref(),
        ) {
            markup.push_str(&format!(
                " {presentation}:presentation-page-layout-name=\"{}\"",
                escape_attribute(layout)
            ));
        }
        markup.push_str("/>");
        self.names.push(name);
        let at = match self.scene.pages.get(index) {
            Some(next) => next.range.start,
            None => match self.scene.pages.last() {
                Some(last) => last.range.end,
                None => self
                    .scene
                    .body_end
                    .ok_or_else(|| unsupported("the ODP content has no presentation body"))?,
            },
        };
        self.push(at..at, Fragment::Markup(markup));
        Ok(())
    }

    fn check_page_removable(&self, page: &Page) -> Result<()> {
        if let Some(name) = &page.name
            && referenced(self.xml, name, &page.start_tag)
        {
            return Err(unsupported(
                "the ODP page is referenced by links or custom shows",
            ));
        }
        Ok(())
    }

    fn remove_page(&mut self, index: usize) -> Result<()> {
        let page = self.page(index)?;
        if self.scene.pages.len() == 1 {
            return Err(invalid("the last ODP page cannot be removed"));
        }
        self.check_page_removable(page)?;
        self.push(page.range.clone(), Fragment::Delete);
        Ok(())
    }

    fn move_page(&mut self, index: usize, to: usize) -> Result<()> {
        let page = self.page(index)?;
        let count = self.scene.pages.len();
        if to >= count {
            return Err(invalid("the ODP page position is outside the presentation"));
        }
        if to == index {
            return Ok(());
        }
        let at = if to > index {
            self.scene.pages[to].range.end
        } else {
            self.scene.pages[to].range.start
        };
        self.push(page.range.clone(), Fragment::Delete);
        self.push(at..at, Fragment::Relocated(page.range.clone()));
        Ok(())
    }

    fn swap_pages(&mut self, first: usize, second: usize) -> Result<()> {
        let left = self.page(first)?.range.clone();
        let right = self.page(second)?.range.clone();
        if first == second {
            return Ok(());
        }
        self.push(left.clone(), Fragment::Relocated(right.clone()));
        self.push(right, Fragment::Relocated(left));
        Ok(())
    }
}

/// The qualified name of a paragraph's start tag.
fn paragraph_qname(raw: &str) -> Result<String> {
    Ok(StartTag::parse(raw)?.name.to_owned())
}

/// Whether `token` occurs in `xml` as a whole token (bounded by quotes,
/// `#`, commas or spaces) outside `own`, the start tag that defines it.
fn referenced(xml: &str, token: &str, own: &Range<usize>) -> bool {
    if token.is_empty() {
        return false;
    }
    let escaped = escape_attribute(token);
    let boundary =
        |byte: Option<u8>| matches!(byte, None | Some(b'"' | b'\'' | b'#' | b',' | b' '));
    let bytes = xml.as_bytes();
    let mut from = 0;
    while let Some(found) = xml[from..].find(escaped.as_str()) {
        let start = from + found;
        let end = start + escaped.len();
        from = start + 1;
        if own.contains(&start) {
            continue;
        }
        let before = start.checked_sub(1).map(|index| bytes[index]);
        let after = bytes.get(end).copied();
        if boundary(before) && boundary(after) && matches!(before, Some(b'"' | b'#' | b',' | b' '))
        {
            return true;
        }
    }
    false
}
