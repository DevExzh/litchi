//! Lazy traversal of the scene index: one cursor per parent axis.

use litchi_odp::scene::{Object, ObjectRef, Scope, StoryRef};

use super::node::{Item, Node, is_placeholder, notes_story, object_kind, page_comments};
use super::spec::Kind;
use crate::Result;
use crate::path::navigate::KindSet;

/// The children of one parent, filtered by kind, in document order.
pub(crate) struct Children<'f> {
    parent: Node<'f>,
    kinds: KindSet<Kind>,
    axis: Axis,
    next: usize,
    counters: [usize; Kind::COUNT],
    /// Annotation inventory indices of a slide, read once.
    comments: Vec<usize>,
    /// The parts of record 0328 listed after the structural children, when
    /// a query names them.
    tail: Option<std::vec::IntoIter<Node<'f>>>,
}

/// Whether `kinds` is a view's or a wildcard's set rather than a query
/// naming kinds: the parts of record 0328 are listed only when asked for
/// by name, so views and `*` list the structural children they listed
/// before.
fn wildcard(kinds: KindSet<Kind>) -> bool {
    kinds.contains(Kind::Connector) && kinds.contains(Kind::Comment)
}

enum Axis {
    Empty,
    /// Slides, then master pages, then layouts.
    Root(u8),
    /// A slide's or group's objects; a slide continues with its notes and
    /// comments.
    Objects {
        list: ObjectRef,
        slide: Option<usize>,
        phase: u8,
    },
    Paragraphs(StoryRef),
    Runs(StoryRef, usize),
    Rows(ObjectRef),
    Cells(ObjectRef, usize),
}

impl<'f> Children<'f> {
    pub(crate) fn of(parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let mut comments = Vec::new();
        let axis = match parent.item() {
            Item::Presentation => Axis::Root(0),
            Item::Slide(page) => {
                if kinds.contains(Kind::Comment) {
                    comments = page_comments(parent.file(), *page)?.collect();
                }
                Axis::Objects {
                    list: ObjectRef {
                        page: *page,
                        scope: Scope::Page,
                        path: Vec::new(),
                    },
                    slide: Some(*page),
                    phase: 0,
                }
            },
            Item::Object(object, Kind::Group) => Axis::Objects {
                list: object.clone(),
                slide: None,
                phase: 0,
            },
            Item::Object(object, Kind::Table) => Axis::Rows(object.clone()),
            Item::Object(..) | Item::Notes(_) | Item::Cell(..) => {
                parent.own_story().map_or(Axis::Empty, Axis::Paragraphs)
            },
            Item::Paragraph(story, index) => Axis::Runs(story.clone(), *index),
            Item::Row(object, row) => Axis::Cells(object.clone(), *row),
            Item::Run(..)
            | Item::Comment(_)
            | Item::Master(_)
            | Item::Layout(_)
            | Item::ShapeLink(_)
            | Item::RunLink(..)
            | Item::LineBreak(..)
            | Item::Column(..)
            | Item::Series(..)
            | Item::Axis(..) => Axis::Empty,
        };
        Ok(Self {
            parent: parent.clone(),
            kinds,
            axis,
            next: 0,
            counters: [0; Kind::COUNT],
            comments,
            tail: None,
        })
    }

    /// The part children of the parent the query asks for (record 0328).
    fn extras(&self) -> Result<Vec<Node<'f>>> {
        let kinds = self.kinds;
        if wildcard(kinds) {
            return Ok(Vec::new());
        }
        let parent = &self.parent;
        let scene = parent.scene();
        let file = parent.file();
        let mut out = Vec::new();
        match parent.item() {
            Item::Object(object, Kind::Shape) if kinds.contains(Kind::Hyperlink) => {
                if super::parts::shape_has_link(file, scene, object)? {
                    out.push(parent.child(Item::ShapeLink(object.clone()), "hyperlink", None));
                }
            },
            Item::Run(story, paragraph, run) if kinds.contains(Kind::Hyperlink) => {
                let link = scene
                    .story(story)
                    .and_then(|story| story.paragraphs().get(*paragraph))
                    .and_then(|paragraph| paragraph.runs().get(*run))
                    .is_some_and(|run| run.kind() == litchi_odp::scene::RunKind::Link);
                if link {
                    out.push(parent.child(
                        Item::RunLink(story.clone(), *paragraph, *run),
                        "hyperlink",
                        None,
                    ));
                }
            },
            Item::Paragraph(story, paragraph) if kinds.contains(Kind::LineBreak) => {
                for index in 0..super::parts::line_break_count(file, scene, story, *paragraph)? {
                    out.push(parent.child(
                        Item::LineBreak(story.clone(), *paragraph, index),
                        Kind::LineBreak.name(),
                        Some(index),
                    ));
                }
            },
            Item::Object(object, Kind::Table) if kinds.contains(Kind::Col) => {
                for index in 0..super::parts::column_count(file, scene, object)? {
                    out.push(parent.child(
                        Item::Column(object.clone(), index),
                        Kind::Col.name(),
                        Some(index),
                    ));
                }
            },
            Item::Object(object, Kind::Chart) => {
                let parts = file.parts()?;
                if let Some(chart) = parts.chart(object) {
                    if kinds.contains(Kind::Series) {
                        for index in 0..chart.series.len() {
                            out.push(parent.child(
                                Item::Series(object.clone(), index),
                                Kind::Series.name(),
                                Some(index),
                            ));
                        }
                    }
                    if kinds.contains(Kind::Axis) {
                        for index in 0..chart.axes.len() {
                            out.push(parent.child(
                                Item::Axis(object.clone(), index),
                                Kind::Axis.name(),
                                Some(index),
                            ));
                        }
                    }
                }
            },
            _ => {},
        }
        Ok(out)
    }

    /// The next part child, once the structural children are listed.
    fn next_extra(&mut self) -> Result<Option<Node<'f>>> {
        if self.tail.is_none() {
            self.tail = Some(self.extras()?.into_iter());
        }
        Ok(self.tail.as_mut().and_then(Iterator::next))
    }

    fn objects(&self, list: &ObjectRef) -> &'f [Object] {
        let scene = self.parent.scene();
        if list.path.is_empty() {
            return scene.page(list.page).map_or(&[][..], |page| page.objects());
        }
        scene.object(list).map_or(&[][..], Object::children)
    }

    fn next_root(&mut self, phase: u8) -> Result<Option<Node<'f>>> {
        let scene = self.parent.scene();
        let (kind, count, item): (Kind, usize, fn(usize) -> Item) = match phase {
            0 => (Kind::Slide, scene.pages().len(), Item::Slide),
            1 => (
                Kind::SlideMaster,
                self.parent.file().masters()?.len(),
                Item::Master,
            ),
            2 => (
                Kind::SlideLayout,
                self.parent.file().layouts()?.len(),
                Item::Layout,
            ),
            _ => return Ok(None),
        };
        if !self.kinds.contains(kind) || self.next >= count {
            self.next = 0;
            self.axis = Axis::Root(phase + 1);
            return self.step();
        }
        let index = self.next;
        self.next += 1;
        Ok(Some(self.parent.child(
            item(index),
            kind.name(),
            Some(index),
        )))
    }

    fn next_object(
        &mut self,
        list: ObjectRef,
        slide: Option<usize>,
        phase: u8,
    ) -> Result<Option<Node<'f>>> {
        match phase {
            0 => {
                let objects = self.objects(&list);
                while self.next < objects.len() {
                    let index = self.next;
                    self.next += 1;
                    let object = &objects[index];
                    let physical = object_kind(object);
                    let mut path = list.path.clone();
                    path.push(index);
                    let reference = ObjectRef {
                        page: list.page,
                        scope: list.scope,
                        path,
                    };
                    // A query naming video, audio or equation objects (and
                    // not the media or embedded objects they are) lists
                    // them as such (record 0328).
                    let mut kind = physical;
                    if !wildcard(self.kinds)
                        && ((physical == Kind::Media
                            && !self.kinds.contains(Kind::Media)
                            && (self.kinds.contains(Kind::Video)
                                || self.kinds.contains(Kind::Audio)))
                            || (physical == Kind::Ole
                                && !self.kinds.contains(Kind::Ole)
                                && self.kinds.contains(Kind::Equation)))
                    {
                        if let Some(semantic) =
                            super::parts::semantic_kind(self.parent.file(), &reference, object)?
                        {
                            kind = semantic;
                        }
                    }
                    let ordinal = self.counters[kind as usize];
                    self.counters[kind as usize] += 1;
                    let wanted = self.kinds.contains(kind)
                        || (is_placeholder(object) && self.kinds.contains(Kind::Placeholder));
                    if !wanted {
                        continue;
                    }
                    return Ok(Some(self.parent.child(
                        Item::Object(reference, kind),
                        kind.name(),
                        Some(ordinal),
                    )));
                }
                let Some(page) = slide else {
                    self.axis = Axis::Empty;
                    return Ok(None);
                };
                self.axis = Axis::Objects {
                    list,
                    slide,
                    phase: 1,
                };
                self.next = 0;
                if self.kinds.contains(Kind::Notes)
                    && notes_story(self.parent.scene(), page).is_some()
                {
                    return Ok(Some(self.parent.child(Item::Notes(page), "notes", None)));
                }
                self.step()
            },
            _ => {
                if self.next < self.comments.len() {
                    let ordinal = self.next;
                    self.next += 1;
                    let index = self.comments[ordinal];
                    return Ok(Some(self.parent.child(
                        Item::Comment(index),
                        Kind::Comment.name(),
                        Some(ordinal),
                    )));
                }
                self.axis = Axis::Empty;
                Ok(None)
            },
        }
    }

    fn step(&mut self) -> Result<Option<Node<'f>>> {
        let scene = self.parent.scene();
        match std::mem::replace(&mut self.axis, Axis::Empty) {
            Axis::Empty => self.next_extra(),
            Axis::Root(phase) => {
                self.axis = Axis::Root(phase);
                self.next_root(phase)
            },
            Axis::Objects { list, slide, phase } => {
                self.axis = Axis::Objects {
                    list: list.clone(),
                    slide,
                    phase,
                };
                self.next_object(list, slide, phase)
            },
            Axis::Paragraphs(story) => {
                let count = scene
                    .story(&story)
                    .map_or(0, |story| story.paragraphs().len());
                if !self.kinds.contains(Kind::Paragraph) || self.next >= count {
                    return self.next_extra();
                }
                let index = self.next;
                self.next += 1;
                let node = self.parent.child(
                    Item::Paragraph(story.clone(), index),
                    Kind::Paragraph.name(),
                    Some(index),
                );
                self.axis = Axis::Paragraphs(story);
                Ok(Some(node))
            },
            Axis::Runs(story, paragraph) => {
                let count = scene
                    .story(&story)
                    .and_then(|story| story.paragraphs().get(paragraph))
                    .map_or(0, |paragraph| paragraph.runs().len());
                if !self.kinds.contains(Kind::Run) || self.next >= count {
                    return self.next_extra();
                }
                let index = self.next;
                self.next += 1;
                let node = self.parent.child(
                    Item::Run(story.clone(), paragraph, index),
                    Kind::Run.name(),
                    Some(index),
                );
                self.axis = Axis::Runs(story, paragraph);
                Ok(Some(node))
            },
            Axis::Rows(object) => {
                let count = scene
                    .object(&object)
                    .and_then(Object::table)
                    .map_or(0, |table| table.rows().len());
                if !self.kinds.contains(Kind::Row) || self.next >= count {
                    return self.next_extra();
                }
                let index = self.next;
                self.next += 1;
                let node = self.parent.child(
                    Item::Row(object.clone(), index),
                    Kind::Row.name(),
                    Some(index),
                );
                self.axis = Axis::Rows(object);
                Ok(Some(node))
            },
            Axis::Cells(object, row) => {
                let count = scene
                    .object(&object)
                    .and_then(Object::table)
                    .and_then(|table| table.rows().get(row))
                    .map_or(0, |row| row.cells().len());
                if !self.kinds.contains(Kind::Cell) || self.next >= count {
                    return Ok(None);
                }
                let index = self.next;
                self.next += 1;
                let node = self.parent.child(
                    Item::Cell(object.clone(), row, index),
                    Kind::Cell.name(),
                    Some(index),
                );
                self.axis = Axis::Cells(object, row);
                Ok(Some(node))
            },
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.step() {
            Ok(Some(node)) => Some(Ok(node)),
            Ok(None) => None,
            Err(error) => {
                self.axis = Axis::Empty;
                Some(Err(error))
            },
        }
    }
}
