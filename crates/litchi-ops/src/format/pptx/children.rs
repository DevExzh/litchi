//! The children of a PPTX node, in document order (record 0010 §2.5,
//! §4.6, §5.3).
//!
//! One lazy iterator serves the node view and the query walker:
//!
//! | Parent | Children |
//! |---|---|
//! | presentation | slides, in catalog order |
//! | slide | its shape tree's direct elements in z order, each kind with its own counter (D5); for queries, then its notes and comments |
//! | group | the group's direct elements, likewise |
//! | shape | the paragraphs of its text body |
//! | paragraph | its runs and line breaks (fields are not listed, as in OfficeCLI) |
//! | table | its rows; a row, its cells |
//! | chart | its series, then its axes (record 0270) |
//! | smartart, node | its child nodes (record 0284) |
//!
//! Each child is read when it is pulled, so a walk that stops early decodes
//! no later sibling; a slide's scene is read when the iterator is built for
//! a walk or view that wants its shapes. After the first error the iterator
//! is fused.

use std::sync::Arc;

use litchi_drawingml::text::body::Body;
use litchi_pptx::presentation::modern_comments::List as ModernList;
use litchi_pptx::presentation::related::Comments;
use litchi_pptx::shape::owned::Handles;

use super::Pptx;
use super::node::{
    CellRef, CommentItem, Item, LayoutItem, MasterItem, ModernCommentItem, Node, NotesItem,
    ReplyItem, SlideItem, TableItem, TextItem, TreeContainer, TreeItem,
};
use super::spec::Kind;
use super::steps;
use super::text;
use crate::error::{Error, Result};
use crate::path::navigate::KindSet;

/// Who asks for the children.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Listing {
    /// The node view. `leaves` is set when the children are the view's
    /// leaves: a slide listed there shows catalog data only.
    View { leaves: bool },
    /// The query walker, which may also ask for a slide's notes and comments.
    Walk,
}

/// The children of a PPTX node.
pub(crate) struct Children<'f> {
    file: &'f Pptx,
    kinds: KindSet<Kind>,
    state: State,
    /// Children a query walk lists after the main axis (record 0321): the
    /// theme, table columns, hyperlinks, media and equations. The node view
    /// never lists them, so existing views keep their shape.
    extra: std::vec::IntoIter<Item>,
}

enum State {
    Done,
    Failed(Option<Error>),
    /// Slides `next..end`.
    Slides {
        next: u32,
        end: u32,
        listed: bool,
        masters: Option<(u32, bool)>,
    },
    /// Slide masters in presentation order.
    Masters {
        next: u32,
        end: u32,
        listed: bool,
    },
    /// A slide's or group's direct elements (when wanted), then, for a
    /// query, the slide's notes and comments.
    Tree {
        container: TreeContainer,
        handles: Option<Handles>,
        counters: [u32; 8],
        layouts: Option<(u32, u32, bool)>,
        notes: bool,
        comments: bool,
        modern_comments: bool,
    },
    /// A slide's legacy comments.
    Comments {
        slide: u32,
        list: Comments,
        next: u32,
        modern: bool,
    },
    /// A slide's modern threaded comments.
    ModernComments {
        slide: u32,
        list: ModernList,
        next: u32,
    },
    /// Replies of one modern threaded comment.
    Replies {
        slide: u32,
        list: ModernList,
        comment: u32,
        next: u32,
    },
    Paragraphs {
        shape: TreeItem,
        body: Arc<Body>,
        next: u32,
    },
    Runs {
        text: TextItem,
        next: u32,
        runs: u32,
        breaks: u32,
    },
    Rows {
        table: TreeItem,
        next: u32,
        end: u32,
    },
    Cells {
        row: TableItem,
        next: u32,
        end: u32,
    },
    /// A chart's series, then its axes (record 0270); a SmartArt's or
    /// node's child nodes (record 0284).
    Chart(std::vec::IntoIter<Item>),
}

/// The shape-tree kinds, in counter-slot order.
const TREE_KINDS: [Kind; 8] = [
    Kind::Shape,
    Kind::Picture,
    Kind::Table,
    Kind::Chart,
    Kind::Group,
    Kind::Connector,
    Kind::SmartArt,
    Kind::Ole,
];

/// The counter slot of a shape-tree kind.
fn slot(kind: Kind) -> usize {
    TREE_KINDS
        .iter()
        .position(|tree| *tree == kind)
        .unwrap_or(TREE_KINDS.len() - 1)
}

impl<'f> Children<'f> {
    /// The children of `parent` whose kinds are in `kinds`.
    pub(crate) fn of(parent: &Node<'f>, kinds: KindSet<Kind>, listing: Listing) -> Self {
        let file = parent.file();
        let (state, extra) = match State::of(file, parent, kinds, listing) {
            Ok(state) if listing == Listing::Walk => {
                match super::elements::walk_children(file, parent.item(), kinds) {
                    Ok(extra) => (state, extra),
                    Err(error) => (State::Failed(Some(error)), Vec::new()),
                }
            },
            Ok(state) => (state, Vec::new()),
            Err(error) => (State::Failed(Some(error)), Vec::new()),
        };
        Self {
            file,
            kinds,
            state,
            extra: extra.into_iter(),
        }
    }

    fn step(&mut self) -> Result<Option<Node<'f>>> {
        self.file.check_owner()?;
        let result = self.step_unfenced();
        let result = match result {
            Ok(None) => Ok(self.extra.next().map(|item| Node::new(self.file, item))),
            other => other,
        };
        self.file.check_owner().and(result)
    }

    /// Pull one child while the caller's source fence is active.
    fn step_unfenced(&mut self) -> Result<Option<Node<'f>>> {
        let (file, kinds) = (self.file, self.kinds);
        loop {
            match &mut self.state {
                State::Done => return Ok(None),
                State::Failed(error) => {
                    return match error.take() {
                        Some(error) => Err(error),
                        None => Ok(None),
                    };
                },
                State::Slides {
                    next,
                    end,
                    listed,
                    masters,
                } => {
                    if *next >= *end {
                        if let Some((master_end, master_listed)) = masters.take() {
                            self.state = State::Masters {
                                next: 0,
                                end: master_end,
                                listed: master_listed,
                            };
                            continue;
                        }
                        return Ok(None);
                    }
                    let position = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::Slide(SlideItem::new(position, *listed)),
                    )));
                },
                State::Masters { next, end, listed } => {
                    if *next >= *end {
                        return Ok(None);
                    }
                    let position = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::Master(MasterItem::new(position, *listed)),
                    )));
                },
                State::Tree {
                    container,
                    handles,
                    counters,
                    layouts,
                    notes,
                    comments,
                    modern_comments,
                } => {
                    if layouts.is_some() {
                        let master = match *container {
                            TreeContainer::Master(master) => master,
                            _ => u32::MAX,
                        };
                        let (next, end, listed) = layouts.as_mut().expect("checked above");
                        if *next < *end {
                            let position = *next;
                            *next += 1;
                            return Ok(Some(Node::new(
                                file,
                                Item::Layout(LayoutItem::new(master, position, *listed)),
                            )));
                        }
                        *layouts = None;
                    }
                    if let Some(handles) = handles {
                        if let Some(node) = next_tree(file, kinds, *container, handles, counters) {
                            return Ok(Some(node));
                        }
                    }
                    // The shape tree is done: a slide's notes, legacy
                    // comments, then modern comments, when requested.
                    *handles = None;
                    let Some(slide) = container.slide() else {
                        return Ok(None);
                    };
                    if *notes {
                        *notes = false;
                        if let Some(text) = file.slide_notes(slide as usize)? {
                            return Ok(Some(Node::new(
                                file,
                                Item::Notes(NotesItem {
                                    slide,
                                    text: Arc::from(text),
                                }),
                            )));
                        }
                        continue;
                    }
                    if *comments {
                        *comments = false;
                        self.state = match file.slide_comments(slide as usize)? {
                            Some(list) => State::Comments {
                                slide,
                                list,
                                next: 0,
                                modern: *modern_comments,
                            },
                            None => match file.slide_modern_comments(slide as usize)? {
                                Some(list) if *modern_comments => State::ModernComments {
                                    slide,
                                    list,
                                    next: 0,
                                },
                                _ => State::Done,
                            },
                        };
                        continue;
                    }
                    if *modern_comments {
                        *modern_comments = false;
                        self.state = match file.slide_modern_comments(slide as usize)? {
                            Some(list) => State::ModernComments {
                                slide,
                                list,
                                next: 0,
                            },
                            None => State::Done,
                        };
                        continue;
                    }
                    return Ok(None);
                },
                State::Comments {
                    slide,
                    list,
                    next,
                    modern,
                } => {
                    if *next as usize >= list.entries().len() {
                        if *modern {
                            *modern = false;
                            self.state = match file.slide_modern_comments(*slide as usize)? {
                                Some(list) => State::ModernComments {
                                    slide: *slide,
                                    list,
                                    next: 0,
                                },
                                None => State::Done,
                            };
                            continue;
                        }
                        return Ok(None);
                    }
                    let index = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::Comment(CommentItem {
                            slide: *slide,
                            list: list.clone(),
                            index,
                        }),
                    )));
                },
                State::ModernComments { slide, list, next } => {
                    if *next as usize >= list.entries().len() {
                        return Ok(None);
                    }
                    let index = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::ModernComment(ModernCommentItem {
                            slide: *slide,
                            list: list.clone(),
                            index,
                        }),
                    )));
                },
                State::Replies {
                    slide,
                    list,
                    comment,
                    next,
                } => {
                    let Some(record) = list.entries().get(*comment as usize) else {
                        return Ok(None);
                    };
                    if *next as usize >= record.replies().len() {
                        return Ok(None);
                    }
                    let index = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::Reply(ReplyItem {
                            slide: *slide,
                            list: list.clone(),
                            comment: *comment,
                            index,
                        }),
                    )));
                },
                State::Paragraphs { shape, body, next } => {
                    if *next as usize >= body.paragraphs.len() {
                        return Ok(None);
                    }
                    let paragraph = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::Paragraph(TextItem {
                            shape: shape.clone(),
                            body: Arc::clone(body),
                            paragraph,
                            run: 0,
                            ordinal: paragraph,
                            cell: CellRef::NONE,
                        }),
                    )));
                },
                State::Runs {
                    text,
                    next,
                    runs,
                    breaks,
                } => {
                    let Some(paragraph) = text.paragraph() else {
                        return Ok(None);
                    };
                    while let Some(run) = paragraph.runs.get(*next as usize) {
                        let index = *next;
                        *next += 1;
                        let (counter, kind) = match text::listed_run(run) {
                            Some(Kind::Run) => (&mut *runs, Kind::Run),
                            Some(_) => (&mut *breaks, Kind::LineBreak),
                            None => continue,
                        };
                        let ordinal = *counter;
                        *counter += 1;
                        if !kinds.contains(kind) {
                            continue;
                        }
                        let item = TextItem {
                            run: index,
                            ordinal,
                            ..text.clone()
                        };
                        let item = if kind == Kind::Run {
                            Item::Run(item)
                        } else {
                            Item::LineBreak(item)
                        };
                        return Ok(Some(Node::new(file, item)));
                    }
                    return Ok(None);
                },
                State::Rows { table, next, end } => {
                    if *next >= *end {
                        return Ok(None);
                    }
                    let row = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::Row(TableItem {
                            table: table.clone(),
                            row,
                            column: 0,
                        }),
                    )));
                },
                State::Chart(items) => {
                    return Ok(items.next().map(|item| Node::new(file, item)));
                },
                State::Cells { row, next, end } => {
                    if *next >= *end {
                        return Ok(None);
                    }
                    let column = *next;
                    *next += 1;
                    return Ok(Some(Node::new(
                        file,
                        Item::Cell(TableItem {
                            column,
                            ..row.clone()
                        }),
                    )));
                },
            }
        }
    }
}

/// The next element of a shape tree whose kind is wanted, counting every
/// element of each kind on the way (D5).
fn next_tree<'f>(
    file: &'f Pptx,
    kinds: KindSet<Kind>,
    container: TreeContainer,
    handles: &mut Handles,
    counters: &mut [u32; 8],
) -> Option<Node<'f>> {
    for handle in handles.by_ref() {
        let Some(kind) = steps::tree_kind(&handle) else {
            continue;
        };
        let counter = &mut counters[slot(kind)];
        let ordinal = *counter;
        *counter += 1;
        if kinds.contains(kind) {
            return Some(Node::new(
                file,
                Item::Tree(match container {
                    TreeContainer::Slide(slide) => TreeItem::new(slide, kind, handle, ordinal),
                    TreeContainer::Master(master) => {
                        TreeItem::new_master(master, kind, handle, ordinal)
                    },
                    TreeContainer::Layout { master, position } => {
                        TreeItem::new_layout(master, position, kind, handle, ordinal)
                    },
                    container @ TreeContainer::Notes(_) => {
                        TreeItem::in_container(container, kind, handle, ordinal)
                    },
                }),
            ));
        }
    }
    None
}

impl State {
    fn of(file: &Pptx, parent: &Node<'_>, kinds: KindSet<Kind>, listing: Listing) -> Result<Self> {
        let wants_tree = TREE_KINDS.iter().any(|kind| kinds.contains(*kind));
        Ok(match parent.item() {
            Item::Presentation => {
                let listed = matches!(listing, Listing::View { leaves: true });
                let master_end = if kinds.contains(Kind::SlideMaster) {
                    u32::try_from(file.masters_in_catalog()?).unwrap_or(u32::MAX)
                } else {
                    0
                };
                if kinds.contains(Kind::Slide) {
                    let end = u32::try_from(file.slides_in_catalog()?).unwrap_or(u32::MAX);
                    Self::Slides {
                        next: 0,
                        end,
                        listed,
                        masters: (master_end > 0).then_some((master_end, listed)),
                    }
                } else if master_end > 0 {
                    Self::Masters {
                        next: 0,
                        end: master_end,
                        listed,
                    }
                } else {
                    Self::Done
                }
            },
            Item::Slide(slide) => {
                let walk = listing == Listing::Walk;
                let notes = walk && kinds.contains(Kind::Notes);
                let comments = walk && kinds.contains(Kind::Comment);
                let modern_comments = walk && kinds.contains(Kind::ModernComment);
                // A listed slide shows no children; the view never lists it
                // at depth 0, but a caller of `children` may ask.
                let handles = if wants_tree {
                    Some(slide.scene(file)?.roots())
                } else {
                    None
                };
                if handles.is_none() && !notes && !comments && !modern_comments {
                    return Ok(Self::Done);
                }
                Self::Tree {
                    container: TreeContainer::Slide(slide.position),
                    handles,
                    counters: [0; 8],
                    layouts: None,
                    notes,
                    comments,
                    modern_comments,
                }
            },
            Item::Master(master) => {
                let wants_layouts = kinds.contains(Kind::SlideLayout);
                let layout_end = if wants_layouts {
                    u32::try_from(file.layouts_in_catalog(master.position as usize)?)
                        .unwrap_or(u32::MAX)
                } else {
                    0
                };
                let handles = if wants_tree {
                    Some(master.scene(file)?.roots())
                } else {
                    None
                };
                if handles.is_none() && layout_end == 0 {
                    return Ok(Self::Done);
                }
                Self::Tree {
                    container: TreeContainer::Master(master.position),
                    handles,
                    counters: [0; 8],
                    layouts: (layout_end > 0).then_some((
                        0,
                        layout_end,
                        matches!(listing, Listing::View { leaves: true }),
                    )),
                    notes: false,
                    comments: false,
                    modern_comments: false,
                }
            },
            Item::Layout(layout) => {
                if !wants_tree {
                    return Ok(Self::Done);
                }
                Self::Tree {
                    container: TreeContainer::Layout {
                        master: layout.master,
                        position: layout.position,
                    },
                    handles: Some(layout.scene(file)?.roots()),
                    counters: [0; 8],
                    layouts: None,
                    notes: false,
                    comments: false,
                    modern_comments: false,
                }
            },
            Item::Tree(tree) => match tree.kind {
                Kind::Group if wants_tree => Self::Tree {
                    container: tree.container,
                    handles: Some(tree.handle.children()),
                    counters: [0; 8],
                    layouts: None,
                    notes: false,
                    comments: false,
                    modern_comments: false,
                },
                Kind::Shape if kinds.contains(Kind::Paragraph) => match tree.body()? {
                    Some(body) => Self::Paragraphs {
                        shape: tree.clone(),
                        body: Arc::clone(body),
                        next: 0,
                    },
                    None => Self::Done,
                },
                Kind::Chart if kinds.contains(Kind::Series) || kinds.contains(Kind::Axis) => {
                    let items: Vec<Item> = super::chart::children(file, tree)?
                        .into_iter()
                        .filter(|item| match item {
                            Item::Series(_) => kinds.contains(Kind::Series),
                            Item::Axis(_) => kinds.contains(Kind::Axis),
                            _ => false,
                        })
                        .collect();
                    Self::Chart(items.into_iter())
                },
                Kind::SmartArt if kinds.contains(Kind::SmartArtNode) => {
                    Self::Chart(super::smartart::children(file, parent.item())?.into_iter())
                },
                Kind::Table if kinds.contains(Kind::Row) => {
                    let rows = tree.table_grid()?.map_or(0, |grid| grid.len());
                    Self::Rows {
                        table: tree.clone(),
                        next: 0,
                        end: u32::try_from(rows).unwrap_or(u32::MAX),
                    }
                },
                _ => Self::Done,
            },
            Item::Paragraph(text)
                if kinds.contains(Kind::Run) || kinds.contains(Kind::LineBreak) =>
            {
                Self::Runs {
                    text: text.clone(),
                    next: 0,
                    runs: 0,
                    breaks: 0,
                }
            },
            Item::DiagramNode(_) if kinds.contains(Kind::SmartArtNode) => {
                Self::Chart(super::smartart::children(file, parent.item())?.into_iter())
            },
            Item::Row(row) if kinds.contains(Kind::Cell) => {
                let cells = match row.table.table_grid()? {
                    Some(grid) => grid
                        .get(row.row as usize)
                        .map_or(0, |cells| *cells as usize),
                    None => text::row_cells(&row.table.handle, row.row)?,
                };
                Self::Cells {
                    row: row.clone(),
                    next: 0,
                    end: u32::try_from(cells).unwrap_or(u32::MAX),
                }
            },
            Item::ModernComment(comment) if kinds.contains(Kind::Reply) => {
                if comment.comment().is_none() {
                    return Ok(Self::Done);
                }
                Self::Replies {
                    slide: comment.slide,
                    list: comment.list.clone(),
                    comment: comment.index,
                    next: 0,
                }
            },
            _ => Self::Done,
        })
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        if matches!(self.state, State::Done) && self.extra.len() == 0 {
            return None;
        }
        match self.step() {
            Ok(Some(node)) => Some(Ok(node)),
            Ok(None) => {
                self.state = State::Done;
                None
            },
            Err(error) => {
                self.state = State::Done;
                self.extra = Vec::new().into_iter();
                Some(Err(error))
            },
        }
    }
}

impl std::iter::FusedIterator for Children<'_> {}
