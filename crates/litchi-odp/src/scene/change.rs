//! Scene changes and their checked byte-splice plans.

use std::collections::BTreeMap;
use std::ops::Range;

use litchi_core::{Error, Result};

use super::markup::{Encoded, StartTag, encode, escape_attribute, open_tag, template};
use super::styles::{StylesInsertion, TextPropertyUpdate};
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

/// The field elements [`Change::SetFieldText`] rewrites.
const TEXT_FIELDS: [&str; 16] = [
    "page-number",
    "page-count",
    "date",
    "time",
    "title",
    "subject",
    "description",
    "keywords",
    "initial-creator",
    "creator",
    "author-name",
    "creation-date",
    "creation-time",
    "modification-date",
    "modification-time",
    "file-name",
];

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
    /// Replace the displayed text of one field run (`text:page-number`,
    /// `text:date`, `text:title`, ...), keeping its element and attributes:
    /// a field refresh (record 0295). Bookmarks and other markers are
    /// refused.
    SetFieldText {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// Zero-based run position.
        run: usize,
        /// The new text.
        text: String,
    },
    /// Set character formatting on a paragraph (`run: None`, through its
    /// paragraph style) or on one run: a `text:span`; plain text or a field,
    /// which a new span wraps; a hyperlink, whose content span carries it
    /// (record 0316). A new automatic style derived from the node's own
    /// carries the properties; no existing style changes (record 0308).
    SetTextProperties {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// Zero-based run position, or `None` for the paragraph.
        run: Option<usize>,
        /// The properties to set.
        properties: TextPropertyUpdate,
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
    /// Set (`href: Some`) a link run's target, and optionally its
    /// `office:title` (`title: Some(None)` removes it), or remove the link
    /// (`href: None`), keeping its content (record 0328).
    SetRunLink {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// Zero-based run position: a [`super::RunKind::Link`] run.
        run: usize,
        /// The new `xlink:href`, or `None` to unwrap the link.
        href: Option<String>,
        /// The new `office:title`; `None` keeps the current one.
        title: Option<Option<String>>,
    },
    /// Make a plain-text or span run a hyperlink: the run is wrapped in a
    /// `text:a` to `href`, with an optional `office:title` (record 0328).
    LinkRun {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// Zero-based run position: a text or span run.
        run: usize,
        /// The link target.
        href: String,
        /// The link's `office:title`.
        title: Option<String>,
    },
    /// Set (`Some`) or remove (`None`) what clicking an object does: its
    /// click listener (`office:event-listeners`), or the `draw:a` it sits
    /// in (record 0328).
    SetShapeClick {
        /// The object.
        object: ObjectRef,
        /// The new action.
        click: Option<ShapeAction>,
    },
    /// Append a `text:line-break` to a paragraph (record 0328).
    InsertLineBreak {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
    },
    /// Remove one `text:line-break` of a paragraph
    /// ([`super::parts::line_breaks`] order, record 0328).
    RemoveLineBreak {
        /// The story.
        story: StoryRef,
        /// Zero-based paragraph position.
        paragraph: usize,
        /// Zero-based line-break position.
        index: usize,
    },
}

/// What clicking a shape does ([`Change::SetShapeClick`]).
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ShapeAction {
    /// Open a URL (`presentation:action="show"`).
    Url(String),
    /// Go to the page with this `draw:name` (`show` with `#name`).
    Page(String),
    /// Go to the first page.
    FirstPage,
    /// Go to the last page.
    LastPage,
    /// Go to the next page.
    NextPage,
    /// Go to the previous page.
    PreviousPage,
    /// End the show.
    Stop,
}

impl ShapeAction {
    /// `presentation:action` and `xlink:href` of the action.
    fn attributes(&self) -> (&'static str, Option<String>) {
        match self {
            Self::Url(url) => ("show", Some(url.clone())),
            Self::Page(name) => ("show", Some(format!("#{name}"))),
            Self::FirstPage => ("first-page", None),
            Self::LastPage => ("last-page", None),
            Self::NextPage => ("next-page", None),
            Self::PreviousPage => ("previous-page", None),
            Self::Stop => ("stop", None),
        }
    }
}

/// Where a new `office:event-listeners` goes in `object`: in a frame,
/// after its content and before its glue points, image map, title,
/// description and contour; in another shape, after its title and
/// description (the ODF content order).
fn listener_insertion(xml: &str, object: &Object) -> Result<usize> {
    let open_end = object.start_tag.end;
    let close_start = xml[object.range.clone()]
        .rfind("</")
        .map(|close| object.range.start + close)
        .ok_or_else(|| invalid("an ODP drawing element has no end tag"))?;
    let children = super::parts::direct_children(xml, object);
    let frame = object.kind == ObjectKind::Frame;
    if frame {
        let late = [
            "glue-point",
            "image-map",
            "title",
            "desc",
            "contour-polygon",
            "contour-path",
        ];
        Ok(children
            .iter()
            .find(|(local, _)| late.contains(&local.as_str()))
            .map_or(close_start, |(_, range)| range.start))
    } else {
        Ok(children
            .iter()
            .filter(|(local, _)| local == "title" || local == "desc")
            .map(|(_, range)| range.end)
            .max()
            .unwrap_or(open_end))
    }
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

/// The refusal of a text set that would replace `element`, content that
/// is not text (record 0304).
fn dropped(_element: &str) -> Error {
    unsupported(
        "an ODP text set would replace content that is not text: set the text of the runs \
         around it, or remove it first",
    )
}

fn unsupported(reason: &'static str) -> Error {
    Error::Refused(litchi_core::edit::Refusal::new(
        litchi_core::edit::RefusalKind::UnsupportedOperation,
        "odp.scene",
        reason,
    ))
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
    /// Character formatting staged per formatted node.
    formats: Vec<Format>,
}

/// How a formatted node points at its new automatic style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Pointer {
    /// Rewrite `text:style-name` on this start tag.
    StartTag,
    /// Wrap this plain-text range in a new `text:span`.
    Wrap,
}

/// The character formatting one node receives in a plan.
struct Format {
    /// The start tag (or the wrapped range) of the node.
    range: Range<usize>,
    pointer: Pointer,
    family: &'static str,
    /// The node's current style.
    base: Option<String>,
    properties: TextPropertyUpdate,
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
        formats: Vec::new(),
    };
    for change in changes {
        planner.change(change)?;
    }
    planner.finish_formats()?;
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
    fn prefix(&self, prefix: Option<&'a String>, _namespace: &str) -> Result<&'a str> {
        prefix
            .map(String::as_str)
            .ok_or_else(|| unsupported("the ODP content root does not bind the namespace"))
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
            Change::SetFieldText {
                story,
                paragraph,
                run,
                text,
            } => self.set_field_text(story, *paragraph, *run, text),
            Change::SetTextProperties {
                story,
                paragraph,
                run,
                properties,
            } => self.set_text_properties(story, *paragraph, *run, properties),
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
            Change::SetRunLink {
                story,
                paragraph,
                run,
                href,
                title,
            } => self.set_run_link(story, *paragraph, *run, href.as_deref(), title.as_ref()),
            Change::SetShapeClick { object, click } => self.set_shape_click(object, click.as_ref()),
            Change::LinkRun {
                story,
                paragraph,
                run,
                href,
                title,
            } => self.link_run(story, *paragraph, *run, href, title.as_deref()),
            Change::InsertLineBreak { story, paragraph } => {
                self.insert_line_break(story, *paragraph)
            },
            Change::RemoveLineBreak {
                story,
                paragraph,
                index,
            } => {
                let paragraph = self.paragraph(story, *paragraph)?;
                let breaks = super::parts::line_breaks(self.xml, paragraph);
                let range = breaks
                    .get(*index)
                    .cloned()
                    .ok_or_else(|| missing("line break"))?;
                self.push(range, Fragment::Delete);
                Ok(())
            },
        }
    }

    /// The prefix the content root binds to `uri`.
    fn root_prefix(&self, uri: &str, _namespace: &str) -> Result<String> {
        let start = self
            .xml
            .find('<')
            .ok_or_else(|| invalid("ODP content has no root element"))?;
        let mut cursor = start;
        // Skip the declaration and comments to the root start tag.
        while self.xml[cursor..].starts_with("<?") || self.xml[cursor..].starts_with("<!") {
            let end = self.xml[cursor..]
                .find('>')
                .ok_or_else(|| invalid("ODP content prologue is malformed"))?;
            cursor += end + 1;
            cursor += self.xml[cursor..]
                .find('<')
                .ok_or_else(|| invalid("ODP content has no root element"))?;
        }
        let end = self.xml[cursor..]
            .find('>')
            .ok_or_else(|| invalid("ODP content root is malformed"))?;
        let root = StartTag::parse(&self.xml[cursor..=cursor + end])?;
        root.prefix_of(uri)
            .map(str::to_owned)
            .ok_or_else(|| unsupported("the ODP content root does not bind the namespace"))
    }

    fn set_run_link(
        &mut self,
        story: &StoryRef,
        paragraph: usize,
        run: usize,
        href: Option<&str>,
        title: Option<&Option<String>>,
    ) -> Result<()> {
        let paragraph = self.paragraph(story, paragraph)?;
        let run = paragraph.runs.get(run).ok_or_else(|| missing("run"))?;
        if run.kind != RunKind::Link {
            return Err(unsupported("the ODP run is not a hyperlink"));
        }
        let Some(href) = href else {
            // Unwrap: delete the start and end tags, keep the content.
            if run.content.is_empty() && run.start_tag.end == run.range.end {
                self.push(run.range.clone(), Fragment::Delete);
                return Ok(());
            }
            self.push(run.start_tag.clone(), Fragment::Delete);
            self.push(run.content.end..run.range.end, Fragment::Delete);
            return Ok(());
        };
        let raw = &self.xml[run.start_tag.clone()];
        let tag = StartTag::parse(raw)?;
        let href_key = match tag.key_of_local("href") {
            Some(key) => key.to_owned(),
            None => format!(
                "{}:href",
                self.root_prefix("http://www.w3.org/1999/xlink", "XLink")?
            ),
        };
        if run.href.as_deref() != Some(href) {
            self.tag(&run.start_tag.clone(), href_key, Some(href.to_owned()));
        }
        if let Some(title) = title {
            let current = super::parts::link_title(self.xml, run);
            if current.as_ref() != title.as_ref() {
                let key = match tag.key_of_local("title") {
                    Some(key) => key.to_owned(),
                    None => format!(
                        "{}:title",
                        self.prefix(self.scene.prefixes.office.as_ref(), "office")?
                    ),
                };
                self.tag(&run.start_tag.clone(), key, title.clone());
            }
        }
        Ok(())
    }

    fn link_run(
        &mut self,
        story: &StoryRef,
        paragraph: usize,
        run: usize,
        href: &str,
        title: Option<&str>,
    ) -> Result<()> {
        let paragraph = self.paragraph(story, paragraph)?;
        let run = paragraph.runs.get(run).ok_or_else(|| missing("run"))?;
        if !matches!(run.kind, RunKind::Text | RunKind::Span) {
            return Err(unsupported(
                "only a plain-text or span ODP run can become a hyperlink",
            ));
        }
        let text = self.text_prefix()?;
        let xlink = self.root_prefix("http://www.w3.org/1999/xlink", "XLink")?;
        let mut open = format!(
            "<{text}:a {xlink}:type=\"simple\" {xlink}:href=\"{}\"",
            escape_attribute(href)
        );
        if let Some(title) = title {
            let office = self.prefix(self.scene.prefixes.office.as_ref(), "office")?;
            open.push_str(&format!(" {office}:title=\"{}\"", escape_attribute(title)));
        }
        open.push('>');
        let range = run.range.clone();
        let markup = format!("{open}{}</{text}:a>", &self.xml[range.clone()]);
        self.push(range, Fragment::Markup(markup));
        Ok(())
    }

    fn insert_line_break(&mut self, story: &StoryRef, paragraph: usize) -> Result<()> {
        let prefix = self.text_prefix()?;
        let paragraph = self.paragraph(story, paragraph)?;
        let markup = format!("<{prefix}:line-break/>");
        match super::parts::line_break_insertion(paragraph) {
            Some(offset) => self.push(offset..offset, Fragment::Markup(markup)),
            None => {
                let raw = &self.xml[paragraph.start_tag.clone()];
                let open = open_tag(raw)?;
                let qname = paragraph_qname(raw)?;
                self.push(
                    paragraph.range.clone(),
                    Fragment::Markup(format!("{open}{markup}</{qname}>")),
                );
            },
        }
        Ok(())
    }

    fn set_shape_click(
        &mut self,
        reference: &ObjectRef,
        click: Option<&ShapeAction>,
    ) -> Result<()> {
        let object = self.object(reference)?;
        let current = super::parts::shape_click(self.xml, object);
        if let Some(current) = &current {
            if current.is_wrapper() {
                return self.set_wrapper_click(object, current, click);
            }
        }
        if let (Some(current), Some(click)) = (&current, click) {
            let (action, href) = click.attributes();
            if current.action() == Some(action) && current.href() == href.as_deref() {
                return Ok(());
            }
        }
        let listener = |planner: &Self, click: &ShapeAction| -> Result<String> {
            let presentation =
                planner.prefix(planner.scene.prefixes.presentation.as_ref(), "presentation")?;
            let script = planner
                .root_prefix("urn:oasis:names:tc:opendocument:xmlns:script:1.0", "script")?;
            let (action, href) = click.attributes();
            let mut out = format!(
                "<{presentation}:event-listener {script}:event-name=\"dom:click\" {presentation}:action=\"{action}\""
            );
            if let Some(href) = href {
                let xlink = planner.root_prefix("http://www.w3.org/1999/xlink", "XLink")?;
                out.push_str(&format!(
                    " {xlink}:href=\"{}\" {xlink}:type=\"simple\" {xlink}:show=\"embed\" {xlink}:actuate=\"onRequest\"",
                    escape_attribute(&href)
                ));
            }
            out.push_str("/>");
            Ok(out)
        };
        match (current, click) {
            (None, None) => Ok(()),
            (Some(current), None) => {
                let range = match current.container() {
                    Some((container, 1)) => container,
                    _ => current
                        .listener()
                        .ok_or_else(|| missing("click listener"))?,
                };
                self.push(range, Fragment::Delete);
                Ok(())
            },
            (Some(current), Some(click)) => {
                let range = current
                    .listener()
                    .ok_or_else(|| missing("click listener"))?;
                let markup = listener(self, click)?;
                self.push(range, Fragment::Markup(markup));
                Ok(())
            },
            (None, Some(click)) => {
                if object.empty {
                    return Err(unsupported(
                        "an empty ODP drawing element cannot gain a click action",
                    ));
                }
                let office = self.prefix(self.scene.prefixes.office.as_ref(), "office")?;
                let markup = format!(
                    "<{office}:event-listeners>{}</{office}:event-listeners>",
                    listener(self, click)?
                );
                let offset = listener_insertion(self.xml, object)?;
                self.push(offset..offset, Fragment::Markup(markup));
                Ok(())
            },
        }
    }

    /// Set or remove the click of an object a `draw:a` holds.
    fn set_wrapper_click(
        &mut self,
        object: &Object,
        current: &super::parts::Click,
        click: Option<&ShapeAction>,
    ) -> Result<()> {
        let (start_tag, whole) =
            super::parts::wrapper_ranges(self.xml, object).ok_or_else(|| missing("draw:a link"))?;
        match click {
            None => {
                self.push(whole, Fragment::Relocated(object.range.clone()));
                Ok(())
            },
            Some(ShapeAction::Url(url)) => {
                if current.href() == Some(url.as_str()) {
                    return Ok(());
                }
                let tag = StartTag::parse(&self.xml[start_tag.clone()])?;
                let key = match tag.key_of_local("href") {
                    Some(key) => key.to_owned(),
                    None => format!(
                        "{}:href",
                        self.root_prefix("http://www.w3.org/1999/xlink", "XLink")?
                    ),
                };
                self.tag(&start_tag, key, Some(url.clone()));
                Ok(())
            },
            Some(_) => Err(unsupported(
                "a shape a draw:a link holds takes a URL; remove the link to give it another action",
            )),
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
        if let Some(paragraph) = story.paragraphs.iter().find(|p| !p.objects.is_empty()) {
            return Err(dropped(&paragraph.objects[0].1));
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
        if let Some((_, name)) = paragraph.objects.first() {
            return Err(dropped(name));
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
        if let Some((_, name)) = paragraph
            .objects
            .iter()
            .find(|(start, _)| selected.content.contains(start))
        {
            return Err(dropped(name));
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

    fn set_field_text(
        &mut self,
        story: &StoryRef,
        index: usize,
        run: usize,
        text: &str,
    ) -> Result<()> {
        let paragraph = self.paragraph(story, index)?;
        let selected = paragraph.runs.get(run).ok_or_else(|| missing("run"))?;
        if selected.kind != RunKind::Field {
            return Err(invalid("the ODP run is not a field"));
        }
        // A formatted field is the field element inside its span.
        let (element, content) = match &selected.field {
            Some(field) => (field.range.clone(), field.content.clone()),
            None => (selected.range.clone(), selected.content.clone()),
        };
        let tag_end = if content.start > element.start {
            content.start
        } else {
            element.end
        };
        let name = StartTag::parse(&self.xml[element.start..tag_end])?.name;
        let local = name.rsplit(':').next().unwrap_or(name);
        if !TEXT_FIELDS.contains(&local) {
            return Err(unsupported(
                "this ODP field run holds no displayed text to refresh",
            ));
        }
        if selected.text == text {
            return Ok(());
        }
        let range = element;
        let empty = content.start == range.end;
        let prefix = self.text_prefix()?;
        let encoded = encode(prefix, text)?;
        self.replace_content(
            range.clone(),
            range.start..tag_end,
            content,
            empty,
            &encoded,
        )
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

    fn set_text_properties(
        &mut self,
        story: &StoryRef,
        index: usize,
        run: Option<usize>,
        properties: &TextPropertyUpdate,
    ) -> Result<()> {
        properties.validate()?;
        let paragraph = self.paragraph(story, index)?;
        let (range, pointer, family, base) = match run {
            None => (
                paragraph.start_tag.clone(),
                Pointer::StartTag,
                "paragraph",
                paragraph.style_name.clone(),
            ),
            Some(run) => {
                let selected = paragraph.runs.get(run).ok_or_else(|| missing("run"))?;
                match (selected.kind, &selected.inner_span) {
                    // A span, or the span that formats a field (record
                    // 0316): derive its style.
                    (RunKind::Span, _) => (
                        selected.start_tag.clone(),
                        Pointer::StartTag,
                        "text",
                        selected.style_name.clone(),
                    ),
                    (RunKind::Field, _) if selected.field.is_some() => (
                        selected.start_tag.clone(),
                        Pointer::StartTag,
                        "text",
                        selected.style_name.clone(),
                    ),
                    // Plain text and a bare field: a new span wraps them.
                    (RunKind::Text | RunKind::Field, _) => {
                        (selected.range.clone(), Pointer::Wrap, "text", None)
                    },
                    // A link keeps its own (link character) style: the span
                    // its content is, or a new one inside it, carries the
                    // text's formatting (record 0316).
                    (RunKind::Link, Some((tag, style))) => {
                        (tag.clone(), Pointer::StartTag, "text", style.clone())
                    },
                    (RunKind::Link, None) => {
                        if selected.content.is_empty() {
                            return Err(unsupported(
                                "an empty ODP hyperlink run has no text to format",
                            ));
                        }
                        (selected.content.clone(), Pointer::Wrap, "text", None)
                    },
                }
            },
        };
        if properties.is_empty() {
            return Ok(());
        }
        if let Some(format) = self.formats.iter_mut().find(|format| format.range == range) {
            format.properties.merge(properties);
            return Ok(());
        }
        self.formats.push(Format {
            range,
            pointer,
            family,
            base,
            properties: properties.clone(),
        });
        Ok(())
    }

    /// Derive one automatic style per formatted node, point the node at
    /// it, and insert the new styles into `office:automatic-styles`.
    fn finish_formats(&mut self) -> Result<()> {
        let formats = std::mem::take(&mut self.formats);
        if formats.is_empty() {
            return Ok(());
        }
        let styles = self.scene.automatic_styles(self.xml)?;
        let mut definitions = String::new();
        let mut taken: Vec<String> = Vec::new();
        for format in formats {
            let deleted = format.pointer == Pointer::Wrap
                && self.edits.iter().any(|(range, fragment)| {
                    *range == format.range && *fragment == Fragment::Delete
                });
            if deleted {
                continue;
            }
            let base = format
                .base
                .as_deref()
                .and_then(|name| styles.get(format.family, name));
            if base.is_some_and(|style| style.text.satisfies(&format.properties)) {
                continue;
            }
            let style = self.prefix(self.scene.prefixes.style.as_ref(), "style")?;
            let fo = self.prefix(self.scene.prefixes.fo.as_ref(), "XSL-FO")?;
            let text = self.text_prefix()?;
            let name = fresh_style_name(&styles, format.family, &taken);
            let changes = format.properties.attributes(style, fo);
            let definition = match base {
                Some(base) => self.derive_style(base, &name, &changes, style)?,
                None => {
                    let mut out = format!(
                        "<{style}:style {style}:name=\"{}\" {style}:family=\"{}\"",
                        escape_attribute(&name),
                        format.family
                    );
                    if let Some(parent) = &format.base {
                        out.push_str(&format!(
                            " {style}:parent-style-name=\"{}\"",
                            escape_attribute(parent)
                        ));
                    }
                    out.push('>');
                    out.push_str(&text_properties_tag(style, &changes));
                    out.push_str(&format!("</{style}:style>"));
                    out
                },
            };
            definitions.push_str(&definition);
            match format.pointer {
                Pointer::StartTag => {
                    self.tag(
                        &format.range,
                        format!("{text}:style-name"),
                        Some(name.clone()),
                    );
                },
                Pointer::Wrap => {
                    // A text set of the same run in this plan is folded
                    // into the wrapper; a deleted run needs none.
                    let staged = self
                        .edits
                        .iter()
                        .position(|(range, _)| *range == format.range);
                    let inner = match staged {
                        Some(position) => {
                            let (_, fragment) = self.edits.remove(position);
                            fragment.bytes(self.xml).to_owned()
                        },
                        None => self.xml[format.range.clone()].to_owned(),
                    };
                    let markup = format!(
                        "<{text}:span {text}:style-name=\"{}\">{inner}</{text}:span>",
                        escape_attribute(&name)
                    );
                    self.push(format.range, Fragment::Markup(markup));
                },
            }
            taken.push(name);
        }
        if definitions.is_empty() {
            return Ok(());
        }
        match &self.scene.styles_insertion {
            StylesInsertion::Before(at) => {
                self.push(*at..*at, Fragment::Markup(definitions));
            },
            StylesInsertion::Expand(range) => {
                self.expand(range.clone(), &definitions)?;
            },
            StylesInsertion::Create(at) => {
                let office = self.prefix(self.scene.prefixes.office.as_ref(), "office")?;
                self.push(
                    *at..*at,
                    Fragment::Markup(format!(
                        "<{office}:automatic-styles>{definitions}</{office}:automatic-styles>"
                    )),
                );
            },
            StylesInsertion::None => {
                return Err(unsupported(
                    "this ODP content has no place for automatic styles",
                ));
            },
        }
        Ok(())
    }

    /// A copy of the automatic style `base` named `name`, with `changes`
    /// applied to its `style:text-properties` (added when absent).
    fn derive_style(
        &self,
        base: &super::AutomaticStyle,
        name: &str,
        changes: &[(String, Option<String>)],
        style: &str,
    ) -> Result<String> {
        let rename = [(format!("{style}:name"), Some(name.to_owned()))];
        let raw_start = &self.xml[base.start_tag.clone()];
        let start = StartTag::parse(raw_start)?;
        if base.empty {
            let mut out = start.write(&rename, false, false);
            out.push_str(&text_properties_tag(style, changes));
            out.push_str("</");
            out.push_str(start.name);
            out.push('>');
            return Ok(out);
        }
        let mut out = start.write(&rename, false, false);
        let body = base.start_tag.end;
        match &base.text_tag {
            Some(tag) => {
                let raw = &self.xml[tag.clone()];
                let rewritten = StartTag::parse(raw)?.write(changes, raw.ends_with("/>"), false);
                out.push_str(&self.xml[body..tag.start]);
                out.push_str(&rewritten);
                out.push_str(&self.xml[tag.end..base.range.end]);
            },
            None => {
                let at = base.text_insertion.clamp(body, base.range.end);
                out.push_str(&self.xml[body..at]);
                out.push_str(&text_properties_tag(style, changes));
                out.push_str(&self.xml[at..base.range.end]);
            },
        }
        Ok(out)
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
/// An automatic style name unused in the content and in this plan: `T<n>`
/// for text styles, `P<n>` for paragraph styles.
fn fresh_style_name(styles: &super::AutomaticStyles, family: &str, taken: &[String]) -> String {
    let letter = if family == "paragraph" { 'P' } else { 'T' };
    let mut number = styles.as_slice().len() + taken.len() + 1;
    loop {
        let candidate = format!("{letter}{number}");
        if !taken.contains(&candidate) && !styles.has_name(&candidate) {
            return candidate;
        }
        number += 1;
    }
}

/// A new `style:text-properties` element carrying `changes`.
fn text_properties_tag(style: &str, changes: &[(String, Option<String>)]) -> String {
    let mut out = format!("<{style}:text-properties");
    for (key, value) in changes {
        if let Some(value) = value {
            out.push(' ');
            out.push_str(key);
            out.push_str("=\"");
            out.push_str(&escape_attribute(value));
            out.push('"');
        }
    }
    out.push_str("/>");
    out
}

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
