//! The PPTX adapter of path resolution and queries (record 0010 §4.1,
//! §4.6, §5.3).
//!
//! [`Nav`] is a copyable reference to one [`Pptx`] handle. Resolution asks
//! it for one child per path step and runs only that step's leaf read:
//!
//! | Step | Leaf read |
//! |---|---|
//! | `/slide[N]`, `[last()]` | the catalog; nothing is decoded |
//! | `/slide[N]/shape[K]`, `picture[K]`, `table[K]`, `chart[K]`, `group[K]`, `connector[K]`, `smartart[K]`, `ole[K]` | the slide's scene (the leaf memoizes it per slide payload); each kind counts the direct elements of its own kind (D5) |
//! | `…[@name=X]` (or an explicitly raw `…[@id=N]`) | the same scene; `@name` is caseless and ignores a `!!` morph prefix on either side, as OfficeCLI's `MatchesShapeName` |
//! | `/slide[N]/placeholder[K]`, `placeholder[type]`, `placeholder[@type=T][@idx=I]` | the slide's direct `p:sp` with a `p:ph`; an ambiguous identity is refused |
//! | `…/group[K]/…` | the group's direct elements |
//! | `…/shape[K]/paragraph[P]/run[R]`, `linebreak[B]`, `…/shape[K]/run[R]` | the shape's text body, parsed once per node |
//! | `…/table[K]/tr[R]/tc[C]` | the graphic frame's table |
//! | `/slide[N]/notes` | the slide's notes slide only |
//! | `/slide[N]/comment[M]` | the slide's comments part and the comment authors |
//!
//! Elements the leaf cannot answer yet answer their spec gap
//! (`unsupported_capability`); kinds that cannot occur under a parent are
//! `unsupported_path`.

use std::borrow::Cow;
use std::fmt;
use std::sync::Arc;

use litchi_core::FileFormat;
use litchi_core::selector::Position;
use litchi_pptx::shape::owned::{Handle as ShapeHandle, Handles};

use super::Pptx;
use super::children::{Children, Listing};
use super::node::{
    CellRef, ChartItem, CommentItem, Item, LayoutItem, MasterItem, ModernCommentItem, Node,
    NotesItem, ReplyItem, SlideItem, TableItem, TextItem, TreeContainer, TreeItem,
    placeholder_type,
};
use super::spec::{Kind, PLACEHOLDER_TYPE, SPEC};
use super::{steps, text};
use crate::Result;
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, PathStep, Value};
use crate::path::navigate::{Available, ElementKind, KindSet, Lookup, Navigate, Resolved, Select};
use crate::path::vocabulary::{Case, Identity};
use crate::query::inspect::Inspect;

/// The PPTX adapter: path resolution and queries over one handle.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    file: &'f Pptx,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(file: &'f Pptx) -> Self {
        Self { file }
    }

    /// The presentation root.
    pub(crate) fn root_node(self) -> Node<'f> {
        Node::new(self.file, Item::Presentation)
    }

    fn slide(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let count = self.file.slides_in_catalog()?;
        let position = match select {
            Select::None => (count > 0).then_some(0),
            Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
            Select::Last => count.checked_sub(1),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(
            match position.and_then(|position| u32::try_from(position).ok()) {
                Some(position) => Lookup::Found(Node::new(
                    self.file,
                    Item::Slide(SlideItem::new(position, false)),
                )),
                None => Lookup::Missing,
            },
        )
    }

    /// An element of a slide's or group's shape tree.
    fn tree(
        self,
        container: TreeContainer,
        handles: Handles,
        kind: Kind,
        select: Select<'_>,
    ) -> Lookup<Node<'f>, GapSpec> {
        let candidates = handles
            .filter(|handle| steps::tree_kind(handle) == Some(kind))
            .enumerate();
        let found = match select {
            Select::None => candidates.into_iter().next(),
            Select::Ordinal(position) => candidates.into_iter().nth(position.get()),
            Select::Last => candidates.into_iter().last(),
            Select::Identity(identities) => {
                let mut matching = candidates.into_iter().filter(|(_, handle)| {
                    identities
                        .iter()
                        .all(|(identity, value)| tree_identity(handle, identity, value))
                });
                let first = matching.next();
                let more = matching.count();
                if more > 0 {
                    return Lookup::Ambiguous(more + 1);
                }
                first
            },
            _ => return Lookup::NotApplicable,
        };
        match found {
            Some((ordinal, handle)) => Lookup::Found(Node::new(
                self.file,
                Item::Tree(tree_item(
                    container,
                    kind,
                    handle,
                    u32::try_from(ordinal).unwrap_or(u32::MAX),
                )),
            )),
            None => Lookup::Missing,
        }
    }

    /// A placeholder among a slide's or group's direct `p:sp` (OfficeCLI
    /// `GetPlaceholderNode`, `ResolvePlaceholderShape` and
    /// `FindPlaceholderByBinding`): resolves to the shape.
    fn placeholder(
        self,
        container: TreeContainer,
        handles: Handles,
        select: Select<'_>,
    ) -> Lookup<Node<'f>, GapSpec> {
        // `(position among shapes, handle)` of each placeholder shape.
        let mut shapes = 0_u32;
        let placeholders = handles.filter_map(move |handle| {
            if steps::tree_kind(&handle) != Some(Kind::Shape) {
                return None;
            }
            let ordinal = shapes;
            shapes += 1;
            handle
                .view()
                .placeholder()
                .is_some()
                .then_some((ordinal, handle))
        });
        let found = match select {
            Select::None => placeholders.into_iter().next(),
            Select::Ordinal(position) => placeholders.into_iter().nth(position.get()),
            Select::Last => placeholders.into_iter().last(),
            Select::Token(token) => {
                let Some(wanted) = PLACEHOLDER_TYPE.get(token) else {
                    return Lookup::Missing;
                };
                // A bare type token matches an explicit `type` only, as
                // OfficeCLI's `ResolvePlaceholderShape` compares the enum.
                placeholders.into_iter().find(|(_, handle)| {
                    handle
                        .view()
                        .placeholder()
                        .and_then(|placeholder| placeholder.kind())
                        .and_then(|kind| PLACEHOLDER_TYPE.get(kind))
                        == Some(wanted)
                })
            },
            Select::Identity(identities) => {
                let mut matching = placeholders.into_iter().filter(|(_, handle)| {
                    identities
                        .iter()
                        .all(|(identity, value)| placeholder_identity(handle, identity, value))
                });
                let first = matching.next();
                let more = matching.count();
                if more > 0 {
                    return Lookup::Ambiguous(more + 1);
                }
                first
            },
            _ => return Lookup::NotApplicable,
        };
        match found {
            Some((ordinal, handle)) => Lookup::Found(Node::new(
                self.file,
                Item::Tree(tree_item(container, Kind::Shape, handle, ordinal)),
            )),
            None => Lookup::Missing,
        }
    }

    fn notes(self, slide: u32, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        match select {
            Select::None | Select::Last => {},
            Select::Ordinal(position) if position.get() == 0 => {},
            Select::Ordinal(_) => return Ok(Lookup::Missing),
            _ => return Ok(Lookup::NotApplicable),
        }
        Ok(match self.file.slide_notes(slide as usize)? {
            Some(text) => Lookup::Found(Node::new(
                self.file,
                Item::Notes(NotesItem {
                    slide,
                    text: Arc::from(text),
                }),
            )),
            None => Lookup::Missing,
        })
    }

    fn comment(self, slide: u32, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        if !matches!(select, Select::None | Select::Ordinal(_) | Select::Last) {
            return Ok(Lookup::NotApplicable);
        }
        let Some(list) = self.file.slide_comments(slide as usize)? else {
            return Ok(Lookup::Missing);
        };
        let count = list.entries().len();
        let index = match select {
            Select::None => (count > 0).then_some(0),
            Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
            _ => count.checked_sub(1),
        };
        Ok(match index.and_then(|index| u32::try_from(index).ok()) {
            Some(index) => Lookup::Found(Node::new(
                self.file,
                Item::Comment(CommentItem { slide, list, index }),
            )),
            None => Lookup::Missing,
        })
    }

    fn modern_comment(self, slide: u32, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        if !matches!(select, Select::None | Select::Ordinal(_) | Select::Last) {
            return Ok(Lookup::NotApplicable);
        }
        let Some(list) = self.file.slide_modern_comments(slide as usize)? else {
            return Ok(Lookup::Missing);
        };
        let index = match pick_index(list.entries().len(), select) {
            Pick::Index(index) => index,
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        };
        Ok(Lookup::Found(Node::new(
            self.file,
            Item::ModernComment(ModernCommentItem { slide, list, index }),
        )))
    }

    fn reply(
        self,
        comment: &ModernCommentItem,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        if !matches!(select, Select::None | Select::Ordinal(_) | Select::Last) {
            return Ok(Lookup::NotApplicable);
        }
        let Some(record) = comment.comment() else {
            return Ok(Lookup::Missing);
        };
        let index = match pick_index(record.replies().len(), select) {
            Pick::Index(index) => index,
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        };
        Ok(Lookup::Found(Node::new(
            self.file,
            Item::Reply(ReplyItem {
                slide: comment.slide,
                list: comment.list.clone(),
                comment: comment.index,
                index,
            }),
        )))
    }

    fn master(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        if !matches!(select, Select::None | Select::Ordinal(_) | Select::Last) {
            return Ok(Lookup::NotApplicable);
        }
        let count = self.file.masters_in_catalog()?;
        let index = match pick_index(count, select) {
            Pick::Index(index) => index,
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        };
        Ok(Lookup::Found(Node::new(
            self.file,
            Item::Master(MasterItem::new(index, false)),
        )))
    }

    fn layout(self, master: u32, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        if !matches!(select, Select::None | Select::Ordinal(_) | Select::Last) {
            return Ok(Lookup::NotApplicable);
        }
        let count = self.file.layouts_in_catalog(master as usize)?;
        let index = match pick_index(count, select) {
            Pick::Index(index) => index,
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        };
        Ok(Lookup::Found(Node::new(
            self.file,
            Item::Layout(LayoutItem::new(master, index, false)),
        )))
    }

    /// Resolve the root `/slidelayout[N]` view over masters in presentation
    /// order. The nested form remains the authoritative master/layout path.
    fn root_layout(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        if !matches!(select, Select::None | Select::Ordinal(_) | Select::Last) {
            return Ok(Lookup::NotApplicable);
        }
        let wanted = match select {
            Select::None => 0,
            Select::Ordinal(position) => position.get(),
            Select::Last => usize::MAX,
            _ => unreachable!(),
        };
        let mut seen = 0_usize;
        let mut last = None;
        for master in 0..self.file.masters_in_catalog()? {
            let count = self.file.layouts_in_catalog(master)?;
            if count != 0 {
                last = Some((master as u32, (count - 1) as u32));
            }
            if wanted != usize::MAX && wanted < seen.saturating_add(count) {
                let position = wanted - seen;
                return Ok(Lookup::Found(Node::new(
                    self.file,
                    Item::Layout(LayoutItem::new(master as u32, position as u32, false)),
                )));
            }
            seen = seen.saturating_add(count);
        }
        if wanted == usize::MAX {
            if let Some((master, position)) = last {
                return Ok(Lookup::Found(Node::new(
                    self.file,
                    Item::Layout(LayoutItem::new(master, position, false)),
                )));
            }
        }
        Ok(Lookup::Missing)
    }

    fn paragraph(self, tree: &TreeItem, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let Some(body) = tree.body()? else {
            return Ok(Lookup::Missing);
        };
        let count = body.paragraphs.len();
        let index = match pick_index(count, select) {
            Pick::Index(index) => index,
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        };
        Ok(Lookup::Found(Node::new(
            self.file,
            Item::Paragraph(TextItem {
                shape: tree.clone(),
                body: Arc::clone(body),
                paragraph: index,
                run: 0,
                ordinal: index,
                cell: CellRef::NONE,
            }),
        )))
    }

    /// A paragraph of a table cell's text body, by path only: cell text
    /// is not listed in views or unscoped queries (record 0333).
    fn cell_paragraph(
        self,
        cell: &TableItem,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let Some(body) = text::cell_body(cell)? else {
            return Ok(Lookup::Missing);
        };
        let Some(packed) = CellRef::new(cell.row, cell.column) else {
            return Ok(Lookup::NotApplicable);
        };
        let index = match pick_index(body.paragraphs.len(), select) {
            Pick::Index(index) => index,
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        };
        Ok(Lookup::Found(Node::new(
            self.file,
            Item::Paragraph(TextItem {
                shape: cell.table.clone(),
                body: Arc::new(body),
                paragraph: index,
                run: 0,
                ordinal: index,
                cell: packed,
            }),
        )))
    }

    /// A paragraph of a slide's speaker notes (the notes page's body
    /// placeholder), by path only (record 0333).
    fn notes_paragraph(self, slide: u32, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let Some(shape) = super::text_structure::notes_body(self.file, slide)? else {
            return Ok(Lookup::Missing);
        };
        let Some(body) = shape.body()?.cloned() else {
            return Ok(Lookup::Missing);
        };
        let index = match pick_index(body.paragraphs.len(), select) {
            Pick::Index(index) => index,
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        };
        Ok(Lookup::Found(Node::new(
            self.file,
            Item::Paragraph(TextItem {
                shape,
                body,
                paragraph: index,
                run: 0,
                ordinal: index,
                cell: CellRef::NONE,
            }),
        )))
    }

    /// A run of a shape counted across its paragraphs (OfficeCLI's
    /// `/slide[N]/shape[M]/run[K]`).
    fn shape_run(self, tree: &TreeItem, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let Some(body) = tree.body()? else {
            return Ok(Lookup::Missing);
        };
        let runs = body
            .paragraphs
            .iter()
            .enumerate()
            .flat_map(|(paragraph, item)| {
                item.runs
                    .iter()
                    .enumerate()
                    .filter(|(_, run)| text::listed_run(run) == Some(Kind::Run))
                    .enumerate()
                    // `ordinal` is local to this paragraph and is retained
                    // for the node's canonical paragraph/run address.
                    .map(move |(ordinal, (run, _))| (paragraph, run, ordinal))
            })
            // The selector on `/shape/run[N]` is global to the shape.  Keep
            // that traversal position separate from the paragraph-local
            // ordinal above; using one value made the second paragraph's
            // first run resolve to a nonexistent local run.
            .scan(0_usize, |global, (paragraph, run, local)| {
                let selected = *global;
                *global = (*global).saturating_add(1);
                Some((paragraph, run, local, selected))
            });
        let found = match select {
            Select::None => runs.into_iter().next(),
            Select::Ordinal(position) => runs.into_iter().nth(position.get()),
            Select::Last => runs.into_iter().last(),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match found {
            Some((paragraph, run, ordinal, _selected)) => Lookup::Found(Node::new(
                self.file,
                Item::Run(TextItem {
                    shape: tree.clone(),
                    body: Arc::clone(body),
                    paragraph: to_u32(paragraph),
                    run: to_u32(run),
                    ordinal: to_u32(ordinal),
                    cell: CellRef::NONE,
                }),
            )),
            None => Lookup::Missing,
        })
    }

    /// A run or line break of a paragraph.
    fn run(self, text: &TextItem, kind: Kind, select: Select<'_>) -> Lookup<Node<'f>, GapSpec> {
        let Some(paragraph) = text.paragraph() else {
            return Lookup::Missing;
        };
        let runs = paragraph
            .runs
            .iter()
            .enumerate()
            .filter(|(_, run)| text::listed_run(run) == Some(kind))
            .enumerate();
        let found = match select {
            Select::None => runs.into_iter().next(),
            Select::Ordinal(position) => runs.into_iter().nth(position.get()),
            Select::Last => runs.into_iter().last(),
            _ => return Lookup::NotApplicable,
        };
        match found {
            Some((ordinal, (run, _))) => {
                let item = TextItem {
                    run: to_u32(run),
                    ordinal: to_u32(ordinal),
                    ..text.clone()
                };
                Lookup::Found(Node::new(
                    self.file,
                    if kind == Kind::Run {
                        Item::Run(item)
                    } else {
                        Item::LineBreak(item)
                    },
                ))
            },
            None => Lookup::Missing,
        }
    }

    /// A chart's series (by position) or axis (by position or `@role`).
    fn chart_child(
        self,
        chart: &TreeItem,
        kind: Kind,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let Some(summary) = chart.chart(self.file)? else {
            return Ok(Lookup::Missing);
        };
        let make = |index: usize| {
            let item = ChartItem {
                chart: chart.clone(),
                summary: Arc::clone(&summary),
                index: to_u32(index),
            };
            Node::new(
                self.file,
                if kind == Kind::Series {
                    Item::Series(item)
                } else {
                    Item::Axis(item)
                },
            )
        };
        if kind == Kind::Axis
            && let Select::Identity(identities) = select
        {
            let mut found = None;
            for (identity, value) in identities.iter() {
                if !identity.key.eq_ignore_ascii_case("role") {
                    return Ok(Lookup::NotApplicable);
                }
                let Some(role) = crate::format::chart::Role::parse(value) else {
                    return Ok(Lookup::Missing);
                };
                found = summary.axes.iter().position(|axis| axis.role == role);
            }
            return Ok(found.map_or(Lookup::Missing, |index| Lookup::Found(make(index))));
        }
        if kind == Kind::Series
            && let Select::Identity(identities) = select
        {
            // `@id` is the series' stable `c:idx`.
            let mut found = None;
            for (identity, value) in identities.iter() {
                if !identity.key.eq_ignore_ascii_case("id") {
                    return Ok(Lookup::NotApplicable);
                }
                let Ok(id) = value.trim().parse::<u32>() else {
                    return Ok(Lookup::Missing);
                };
                found = summary.series.iter().position(|series| series.index == id);
            }
            return Ok(found.map_or(Lookup::Missing, |index| Lookup::Found(make(index))));
        }
        let count = if kind == Kind::Series {
            summary.series.len()
        } else {
            summary.axes.len()
        };
        Ok(match pick_index(count, select) {
            Pick::Index(index) => Lookup::Found(make(index as usize)),
            Pick::Missing => Lookup::Missing,
            Pick::NotApplicable => Lookup::NotApplicable,
        })
    }

    /// A SmartArt's or node's child node, by position (record 0284).
    fn diagram_node(self, parent: &Item, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let children = super::smartart::children(self.file, parent)?;
        Ok(match pick_index(children.len(), select) {
            Pick::Index(index) => match children.into_iter().nth(index as usize) {
                Some(item) => Lookup::Found(Node::new(self.file, item)),
                None => Lookup::Missing,
            },
            Pick::Missing => Lookup::Missing,
            Pick::NotApplicable => Lookup::NotApplicable,
        })
    }

    /// The item an index selector picks from a list.
    fn pick_item(self, items: Vec<Item>, select: Select<'_>) -> Lookup<Node<'f>, GapSpec> {
        match pick_index(items.len(), select) {
            Pick::Index(index) => match items.into_iter().nth(index as usize) {
                Some(item) => Lookup::Found(Node::new(self.file, item)),
                None => Lookup::Missing,
            },
            Pick::Missing => Lookup::Missing,
            Pick::NotApplicable => Lookup::NotApplicable,
        }
    }

    /// The presentation theme, a singleton (record 0321).
    fn theme(self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        match pick_index(1, select) {
            Pick::Index(_) => {},
            Pick::Missing => return Ok(Lookup::Missing),
            Pick::NotApplicable => return Ok(Lookup::NotApplicable),
        }
        Ok(match self.file.presentation_theme()? {
            Some(theme) => Lookup::Found(Node::new(
                self.file,
                Item::Theme(super::theme::ThemeItem {
                    theme: Arc::new(theme),
                }),
            )),
            None => Lookup::Missing,
        })
    }

    fn row(self, table: &TreeItem, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let rows = table.table_grid()?.map_or(0, |grid| grid.len());
        Ok(match pick_index(rows, select) {
            Pick::Index(row) => Lookup::Found(Node::new(
                self.file,
                Item::Row(TableItem {
                    table: table.clone(),
                    row,
                    column: 0,
                }),
            )),
            Pick::Missing => Lookup::Missing,
            Pick::NotApplicable => Lookup::NotApplicable,
        })
    }

    fn cell(self, row: &TableItem, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let cells = match row.table.table_grid()? {
            Some(grid) => grid
                .get(row.row as usize)
                .map_or(0, |cells| *cells as usize),
            None => text::row_cells(&row.table.handle, row.row)?,
        };
        Ok(match pick_index(cells, select) {
            Pick::Index(column) => Lookup::Found(Node::new(
                self.file,
                Item::Cell(TableItem {
                    column,
                    ..row.clone()
                }),
            )),
            Pick::Missing => Lookup::Missing,
            Pick::NotApplicable => Lookup::NotApplicable,
        })
    }

    /// Resolve one path step after the outer source fence has been entered.
    fn child_unfenced(
        self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let nav = self;
        let kind = step.kind;
        if let Some(gap) = kind.spec().gap {
            return Ok(Lookup::Gap(gap));
        }
        let select = step.select;
        if kind.is_tree_kind() || kind == Kind::Placeholder {
            let Some((container, handles)) = tree_container(nav, parent)? else {
                return Ok(Lookup::NotApplicable);
            };
            return Ok(if kind == Kind::Placeholder {
                nav.placeholder(container, handles, select)
            } else {
                nav.tree(container, handles, kind, select)
            });
        }
        if matches!(kind, Kind::Video | Kind::Audio | Kind::Media) {
            let Some((container, handles)) = tree_container(nav, parent)? else {
                return Ok(Lookup::NotApplicable);
            };
            let found = super::media::select(handles, kind)?;
            if let Select::Identity(identities) = select {
                let mut matching = found.into_iter().filter(|(_, handle)| {
                    identities
                        .iter()
                        .all(|(identity, value)| tree_identity(handle, identity, value))
                });
                let first = matching.next();
                let more = matching.count();
                if more > 0 {
                    return Ok(Lookup::Ambiguous(more + 1));
                }
                return Ok(match first {
                    Some((ordinal, handle)) => Lookup::Found(Node::new(
                        nav.file,
                        Item::Tree(TreeItem::in_container(container, kind, handle, ordinal)),
                    )),
                    None => Lookup::Missing,
                });
            }
            return Ok(match pick_index(found.len(), select) {
                Pick::Index(index) => match found.into_iter().nth(index as usize) {
                    Some((ordinal, handle)) => Lookup::Found(Node::new(
                        nav.file,
                        Item::Tree(TreeItem::in_container(container, kind, handle, ordinal)),
                    )),
                    None => Lookup::Missing,
                },
                Pick::Missing => Lookup::Missing,
                Pick::NotApplicable => Lookup::NotApplicable,
            });
        }
        match (parent.item(), kind) {
            (Item::Presentation, Kind::Slide) => nav.slide(select),
            (Item::Slide(slide), Kind::Notes) => nav.notes(slide.position, select),
            (Item::Slide(slide), Kind::Comment) => nav.comment(slide.position, select),
            (Item::Slide(slide), Kind::ModernComment) => nav.modern_comment(slide.position, select),
            (Item::Master(master), Kind::SlideLayout) => nav.layout(master.position, select),
            (Item::Presentation, Kind::SlideMaster) => nav.master(select),
            (Item::Presentation, Kind::SlideLayout) => nav.root_layout(select),
            (Item::Presentation, Kind::Theme) => nav.theme(select),
            (Item::ModernComment(comment), Kind::Reply) => nav.reply(comment, select),
            (Item::Tree(tree), Kind::Paragraph) if tree.kind == Kind::Shape => {
                nav.paragraph(tree, select)
            },
            (Item::Tree(tree), Kind::Run) if tree.kind == Kind::Shape => {
                nav.shape_run(tree, select)
            },
            (Item::Tree(tree), Kind::Row) if tree.kind == Kind::Table => nav.row(tree, select),
            (Item::Tree(tree), Kind::Hyperlink) if tree.kind == Kind::Shape => {
                Ok(match pick_index(1, select) {
                    Pick::Index(_) => match super::links::of_shape(nav.file, tree)? {
                        Some(item) => {
                            Lookup::Found(Node::new(nav.file, Item::Hyperlink(Box::new(item))))
                        },
                        None => Lookup::Missing,
                    },
                    Pick::Missing => Lookup::Missing,
                    Pick::NotApplicable => Lookup::NotApplicable,
                })
            },
            (Item::Slide(slide), Kind::Equation) => {
                let found = super::equations::items(nav.file, slide.position, None)?;
                Ok(nav.pick_item(found, select))
            },
            (Item::Tree(tree), Kind::Equation) if tree.kind == Kind::Shape => {
                let Some(slide) = tree.container.slide() else {
                    return Ok(Lookup::NotApplicable);
                };
                let found = super::equations::items(nav.file, slide, Some(tree))?;
                Ok(nav.pick_item(found, select))
            },
            (Item::Run(text), Kind::Hyperlink) => Ok(match pick_index(1, select) {
                Pick::Index(_) => match super::links::of_run(nav.file, text)? {
                    Some(item) => {
                        Lookup::Found(Node::new(nav.file, Item::Hyperlink(Box::new(item))))
                    },
                    None => Lookup::Missing,
                },
                Pick::Missing => Lookup::Missing,
                Pick::NotApplicable => Lookup::NotApplicable,
            }),
            (Item::Tree(tree), Kind::Col) if tree.kind == Kind::Table => {
                let columns = super::columns::children(tree)?;
                Ok(match pick_index(columns.len(), select) {
                    Pick::Index(index) => match columns.into_iter().nth(index as usize) {
                        Some(item) => Lookup::Found(Node::new(nav.file, item)),
                        None => Lookup::Missing,
                    },
                    Pick::Missing => Lookup::Missing,
                    Pick::NotApplicable => Lookup::NotApplicable,
                })
            },
            (Item::Tree(tree), Kind::Series | Kind::Axis) if tree.kind == Kind::Chart => {
                nav.chart_child(tree, kind, select)
            },
            (Item::Tree(tree), Kind::SmartArtNode) if tree.kind == Kind::SmartArt => {
                nav.diagram_node(parent.item(), select)
            },
            (Item::DiagramNode(_), Kind::SmartArtNode) => nav.diagram_node(parent.item(), select),
            (Item::Row(row), Kind::Cell) => nav.cell(row, select),
            (Item::Paragraph(text), Kind::Run | Kind::LineBreak) => Ok(nav.run(text, kind, select)),
            (Item::Cell(cell), Kind::Paragraph) => nav.cell_paragraph(cell, select),
            (Item::Notes(notes), Kind::Paragraph) => nav.notes_paragraph(notes.slide, select),
            _ => Ok(Lookup::NotApplicable),
        }
    }
}

/// What an index selector picks among `count` children.
enum Pick {
    Index(u32),
    Missing,
    NotApplicable,
}

fn pick_index(count: usize, select: Select<'_>) -> Pick {
    let index = match select {
        Select::None => (count > 0).then_some(0),
        Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
        Select::Last => count.checked_sub(1),
        _ => return Pick::NotApplicable,
    };
    match index.and_then(|index| u32::try_from(index).ok()) {
        Some(index) => Pick::Index(index),
        None => Pick::Missing,
    }
}

fn to_u32(value: usize) -> u32 {
    u32::try_from(value).unwrap_or(u32::MAX)
}

/// Whether `id` is written as `value`: the decimal form, compared exactly
/// as OfficeCLI compares `Id.Value.ToString()`.
fn id_is(id: Option<u32>, value: &str) -> bool {
    id.is_some_and(|id| {
        let mut digits = [0_u8; 10];
        let mut start = digits.len();
        let mut rest = id;
        loop {
            start -= 1;
            // `rest % 10` is one digit.
            digits[start] = b'0' + (rest % 10) as u8;
            rest /= 10;
            if rest == 0 {
                break;
            }
        }
        &digits[start..] == value.as_bytes()
    })
}

/// OfficeCLI's `MatchesShapeName`: caseless, ignoring a `!!` morph prefix
/// on either side.
pub(crate) fn name_matches(held: Option<&str>, wanted: &str) -> bool {
    let Some(held) = held else {
        return false;
    };
    Case::Caseless.matches(held, wanted)
        || held
            .strip_prefix("!!")
            .is_some_and(|held| Case::Caseless.matches(held, wanted))
        || wanted
            .strip_prefix("!!")
            .is_some_and(|wanted| Case::Caseless.matches(held, wanted))
}

fn tree_identity(handle: &ShapeHandle, identity: Identity, value: &str) -> bool {
    match identity.key {
        "id" => id_is(handle.view().id(), value),
        "name" => name_matches(handle.view().name(), value),
        _ => false,
    }
}

/// OfficeCLI's `FindPlaceholderByBinding` keys: `@type` (the schema default
/// `obj` when absent), `@idx` (`0` when absent), `@id`, `@name`. Native
/// `@idx` and `@id` predicates arrive here only after raw-path admission.
fn placeholder_identity(handle: &ShapeHandle, identity: Identity, value: &str) -> bool {
    match identity.key {
        "type" => PLACEHOLDER_TYPE
            .get(value)
            .is_some_and(|wanted| placeholder_type(handle) == Some(wanted)),
        "idx" => id_is(
            handle
                .view()
                .placeholder()
                .map(|placeholder| placeholder.index()),
            value,
        ),
        "id" => id_is(handle.view().id(), value),
        "name" => name_matches(handle.view().name(), value),
        _ => false,
    }
}

/// Whether a container node holds a shape tree: a slide or a group.
fn tree_container(nav: Nav<'_>, parent: &Node<'_>) -> Result<Option<(TreeContainer, Handles)>> {
    Ok(match parent.item() {
        Item::Slide(slide) => Some((
            TreeContainer::Slide(slide.position),
            slide.scene(nav.file)?.roots(),
        )),
        Item::Master(master) => Some((
            TreeContainer::Master(master.position),
            master.scene(nav.file)?.roots(),
        )),
        Item::Layout(layout) => Some((
            TreeContainer::Layout {
                master: layout.master,
                position: layout.position,
            },
            layout.scene(nav.file)?.roots(),
        )),
        Item::Tree(tree) if tree.kind == Kind::Group => {
            Some((tree.container, tree.handle.children()))
        },
        _ => None,
    })
}

fn tree_item(container: TreeContainer, kind: Kind, handle: ShapeHandle, ordinal: u32) -> TreeItem {
    match container {
        TreeContainer::Slide(slide) => TreeItem::new(slide, kind, handle, ordinal),
        TreeContainer::Master(master) => TreeItem::new_master(master, kind, handle, ordinal),
        TreeContainer::Layout { master, position } => {
            TreeItem::new_layout(master, position, kind, handle, ordinal)
        },
        container @ TreeContainer::Notes(_) => {
            TreeItem::in_container(container, kind, handle, ordinal)
        },
    }
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Pptx;

    fn check_source(&self) -> Result<()> {
        self.file.check_source_now()
    }

    fn fence_key(&self) -> usize {
        self.file.fence_key()
    }

    fn root(&self) -> Result<Node<'f>> {
        self.file.check_owner()?;
        Ok(self.root_node())
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        self.file.finish(|| (*self).child_unfenced(parent, step))
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Children<'f>> {
        self.file
            .finish(|| Ok(Children::of(parent, kinds, Listing::Walk)))
    }

    fn available(
        &self,
        parent: &Node<'f>,
        _step: &Resolved<'_, Kind>,
        out: &mut Available,
    ) -> Result<()> {
        self.file.finish(|| available_unfenced(*self, parent, out))
    }
}

fn available_unfenced<'f>(nav: Nav<'f>, parent: &Node<'f>, out: &mut Available) -> Result<()> {
    let mut counts = [0_usize; Kind::COUNT];
    let mut order: Vec<Kind> = Vec::new();
    let mut note = |kind: Kind| {
        if counts[kind as usize] == 0 {
            order.push(kind);
        }
        counts[kind as usize] += 1;
    };
    if let Some((_, handles)) = tree_container(nav, parent)? {
        for handle in handles {
            if let Some(kind) = steps::tree_kind(&handle) {
                note(kind);
            }
        }
    }
    // A scene-bearing parent can also have catalog/related-part children.
    // Keep this independent of the shape-tree probe above: otherwise a
    // slide's notes/comments and a master's layouts disappear from
    // `available`, which makes a failed path report claim they do not
    // exist.
    match parent.item() {
        Item::Presentation => {
            for _ in 0..nav.file.slides_in_catalog()? {
                note(Kind::Slide);
            }
            for _ in 0..nav.file.masters_in_catalog()? {
                note(Kind::SlideMaster);
            }
        },
        Item::Master(master) => {
            for _ in 0..nav.file.layouts_in_catalog(master.position as usize)? {
                note(Kind::SlideLayout);
            }
        },
        Item::Slide(slide) => {
            if let Some(list) = nav.file.slide_modern_comments(slide.position as usize)? {
                for _ in list.entries() {
                    note(Kind::ModernComment);
                }
            }
            if let Some(list) = nav.file.slide_comments(slide.position as usize)? {
                for _ in list.entries() {
                    note(Kind::Comment);
                }
            }
            if nav.file.slide_notes(slide.position as usize)?.is_some() {
                note(Kind::Notes);
            }
        },
        Item::ModernComment(comment) => {
            if let Some(record) = comment.comment() {
                for _ in record.replies() {
                    note(Kind::Reply);
                }
            }
        },
        Item::Tree(tree) if tree.kind == Kind::Shape => {
            if let Some(body) = tree.body()? {
                for _ in &body.paragraphs {
                    note(Kind::Paragraph);
                }
            }
        },
        Item::Tree(tree) if tree.kind == Kind::Table => {
            for _ in 0..text::table_shape(&tree.handle)?.0 {
                note(Kind::Row);
            }
        },
        Item::Tree(tree) if tree.kind == Kind::SmartArt => {
            for _ in super::smartart::children(nav.file, parent.item())? {
                note(Kind::SmartArtNode);
            }
        },
        Item::DiagramNode(_) => {
            for _ in super::smartart::children(nav.file, parent.item())? {
                note(Kind::SmartArtNode);
            }
        },
        Item::Tree(tree) if tree.kind == Kind::Chart => {
            if let Some(summary) = tree.chart(nav.file)? {
                for _ in &summary.series {
                    note(Kind::Series);
                }
                for _ in &summary.axes {
                    note(Kind::Axis);
                }
            }
        },
        Item::Row(row) => {
            for _ in 0..text::row_cells(&row.table.handle, row.row)? {
                note(Kind::Cell);
            }
        },
        Item::Paragraph(text) => {
            if let Some(paragraph) = text.paragraph() {
                for run in &paragraph.runs {
                    if let Some(kind) = text::listed_run(run) {
                        note(kind);
                    }
                }
            }
        },
        _ => {},
    }
    for kind in order {
        out.kind(kind.name(), counts[kind as usize]);
    }
    Ok(())
}

impl<'f> Inspect<'f> for Nav<'f> {
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Slide, Kind::SlideMaster, Kind::SlideLayout];
    const PLACEHOLDER_SHAPES: bool = true;

    fn kind(&self, node: &Node<'f>) -> Kind {
        node.element()
    }

    fn is_placeholder(&self, node: &Node<'f>) -> bool {
        node.is_placeholder()
    }

    fn type_name(&self, node: &Node<'f>) -> &'static str {
        node.type_name()
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        self.file.finish(|| match node.item() {
            // A slide's text is all its text, for `slide:contains(...)`.
            Item::Slide(slide) => Ok(self
                .file
                .slide_text(slide.position as usize)?
                .map(Cow::Owned)),
            _ => node.text_member(),
        })
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        self.file.finish(|| node.format_member())
    }

    fn prop<'n>(&'n self, node: &'n Node<'f>, index: usize) -> Result<Option<Value<'n>>> {
        self.file.finish(|| node.prop_member(index))
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        kind: Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        match node.item() {
            Item::Tree(tree) => {
                let ordinal = u32::try_from(ordinal.get()).ok();
                fmt::Display::fmt(&steps::tree_step(&tree.handle, kind, ordinal, true), out)
            },
            Item::Notes(_) => fmt::Display::fmt(&PathStep::new("notes"), out),
            _ => fmt::Display::fmt(&PathStep::at(kind.name(), ordinal), out),
        }
    }
}

impl Kind {
    /// The number of kinds.
    pub(crate) const COUNT: usize = 37;
}
