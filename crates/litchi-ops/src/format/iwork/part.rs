//! Lazily built parts of an iWork object tree (record 0305).
//!
//! The first read builds the eager tree: the root and the containers every
//! read needs (Numbers sheets and tables; the Keynote root alone). The rest
//! is held by parts, built as reads reach them:
//!
//! - the slides of a Keynote presentation are one part whose items are
//!   materialized one by one (a slide's properties read the theme's layout
//!   catalog, decoded on the first slide item);
//! - the children of a Keynote slide are one part, built whole the first
//!   time a read reaches them, and again with the slide's character runs
//!   appended when a read reaches a paragraph's children;
//! - the cells and merged ranges of a Numbers table are one part whose items
//!   are materialized one by one: a cell step looks its cell up in the
//!   leaf's sorted cell list and builds only that item.
//!
//! A step's kind and ordinal are known without materializing the item.
//! Every item keeps its place ([`At`]) once built, so nodes taken before a
//! part grew stay valid.

#![allow(
    irrefutable_let_patterns,
    reason = "a build with one lazily built format has fewer kinds of part"
)]

#[cfg(feature = "keynote")]
use std::sync::OnceLock;

use super::Iwork;
use super::tree::{At, Item};
#[cfg(feature = "keynote")]
use super::tree::{ROOT, Tree};

/// One lazily built part.
pub(crate) enum Part {
    /// The slides of a Keynote presentation.
    #[cfg(feature = "keynote")]
    Slides(super::keynote::Slides),
    /// The children of a Keynote slide.
    #[cfg(feature = "keynote")]
    Slide(Whole),
    /// The cells and merged ranges of a Numbers table.
    #[cfg(feature = "numbers")]
    Table(super::numbers::Cells),
}

/// A part built whole, then again with its deferred items appended.
#[cfg(feature = "keynote")]
pub(crate) struct Whole {
    /// The leaf position of the element whose children the part holds.
    position: usize,
    /// That element's canonical path.
    path: Box<[crate::node::PathStep<'static>]>,
    built: OnceLock<Tree>,
    /// The part with its deferred items: every item of `built` keeps its
    /// index here.
    full: OnceLock<Tree>,
}

#[cfg(feature = "keynote")]
impl Whole {
    pub(crate) fn new(position: usize, path: Box<[crate::node::PathStep<'static>]>) -> Self {
        Self {
            position,
            path,
            built: OnceLock::new(),
            full: OnceLock::new(),
        }
    }
}

/// The children of one item: indices in one part.
#[derive(Clone, Copy)]
pub(crate) struct Kids<'f> {
    part: u32,
    list: List<'f>,
}

#[derive(Clone, Copy)]
enum List<'f> {
    /// The item's child list.
    Items(&'f [u32]),
    /// Every slot of a part materialized item by item, in order.
    #[cfg(any(feature = "keynote", feature = "numbers"))]
    Slots(u32),
}

impl<'f> Kids<'f> {
    pub(crate) const fn len(&self) -> usize {
        match self.list {
            List::Items(items) => items.len(),
            #[cfg(any(feature = "keynote", feature = "numbers"))]
            List::Slots(count) => count as usize,
        }
    }

    pub(crate) fn get(&self, position: usize) -> Option<At> {
        let index = match self.list {
            List::Items(items) => *items.get(position)?,
            #[cfg(any(feature = "keynote", feature = "numbers"))]
            List::Slots(count) => u32::try_from(position).ok().filter(|slot| *slot < count)?,
        };
        Some(At::new(self.part, index))
    }

    /// The children in order.
    pub(crate) fn iter(self) -> impl DoubleEndedIterator<Item = At> + Clone + 'f {
        (0..self.len()).filter_map(move |position| self.get(position))
    }
}

impl Iwork {
    fn part(&self, part: u32) -> &Part {
        &self.built().parts[part as usize - 1]
    }

    /// The tree of a Keynote slide's part, built on the first call.
    #[cfg(feature = "keynote")]
    fn whole<'f>(&'f self, whole: &'f Whole) -> &'f Tree {
        if let Some(full) = whole.full.get() {
            return full;
        }
        whole.built.get_or_init(|| self.build_whole(whole, false))
    }

    #[cfg(feature = "keynote")]
    fn build_whole(&self, whole: &Whole, deferred: bool) -> Tree {
        match &self.0.owner {
            super::Owner::Keynote(package) => {
                super::keynote::build_slide(package, whole.position, &whole.path, deferred)
            },
            #[allow(unreachable_patterns, reason = "only Keynote slides build whole parts")]
            _ => Tree::root_only(),
        }
    }

    /// The item at `at`.
    pub(crate) fn item(&self, at: At) -> &Item {
        if at.part == 0 {
            return self.tree().item(at.index);
        }
        match *self.part(at.part) {
            #[cfg(feature = "keynote")]
            Part::Slides(ref slides) => slides.item(self, at.index),
            #[cfg(feature = "keynote")]
            Part::Slide(ref whole) => self.whole(whole).item(at.index),
            #[cfg(feature = "numbers")]
            Part::Table(ref cells) => cells.item(self, at.index),
        }
    }

    /// The kind ordinal of the item at `at`, without materializing it.
    pub(crate) fn kind_at(&self, at: At) -> u8 {
        if at.part != 0 {
            match *self.part(at.part) {
                #[cfg(feature = "keynote")]
                Part::Slides(_) => return super::keynote::Slides::KIND,
                #[cfg(feature = "numbers")]
                Part::Table(ref cells) => return cells.kind(at.index),
                #[allow(unreachable_patterns, reason = "the part arms depend on features")]
                _ => {},
            }
        }
        self.item(at).kind()
    }

    /// The position of the item at `at` among its parent's children of its
    /// kind, without materializing it.
    pub(crate) fn ordinal_at(&self, at: At) -> u32 {
        if at.part != 0 {
            match *self.part(at.part) {
                #[cfg(feature = "keynote")]
                Part::Slides(_) => return at.index,
                #[cfg(feature = "numbers")]
                Part::Table(ref cells) => return cells.ordinal(at.index),
                #[allow(unreachable_patterns, reason = "the part arms depend on features")]
                _ => {},
            }
        }
        self.item(at).ordinal()
    }

    /// The part holding the children of the item at `at` (0: its own),
    /// without materializing a slot item.
    fn child_part(&self, at: At) -> u32 {
        #[cfg(feature = "keynote")]
        if at.part != 0
            && let Part::Slides(_) = self.part(at.part)
        {
            return super::keynote::Slides::child_part(at.index);
        }
        self.item(at).part()
    }

    /// The child of kind `kind` at cell `row`, `column` among `kids`, when
    /// their part looks cells up directly; `None` when it does not.
    #[allow(
        clippy::option_option,
        reason = "`None` asks the caller to scan; `Some(None)` is a miss"
    )]
    pub(crate) fn cell_child(
        &self,
        kids: Kids<'_>,
        kind: u8,
        row: u32,
        column: u32,
    ) -> Option<Option<At>> {
        #[cfg(feature = "numbers")]
        if kids.part != 0
            && let List::Slots(_) = kids.list
            && let Part::Table(cells) = self.part(kids.part)
        {
            return cells
                .locate(self, kind, row, column)
                .map(|slot| slot.map(|slot| At::new(kids.part, slot)));
        }
        let _ = (kids, kind, row, column);
        None
    }

    /// The children of the item at `at`, building the part that holds them
    /// (and, for a format that adds some children on demand, completing the
    /// tree first: Pages section headers and footers, Keynote character
    /// runs).
    pub(crate) fn kids(&self, at: At) -> Kids<'_> {
        self.expand(at);
        match self.child_part(at) {
            0 => Kids {
                part: at.part,
                list: List::Items(self.item(at).children()),
            },
            part => match *self.part(part) {
                #[cfg(feature = "keynote")]
                Part::Slides(ref slides) => Kids {
                    part,
                    list: List::Slots(slides.len()),
                },
                #[cfg(feature = "keynote")]
                Part::Slide(ref whole) => Kids {
                    part,
                    list: List::Items(self.whole(whole).item(ROOT).children()),
                },
                #[cfg(feature = "numbers")]
                Part::Table(ref cells) => Kids {
                    part,
                    list: List::Slots(cells.len()),
                },
            },
        }
    }

    /// Complete the tree before the children of `at` are read, when the
    /// format builds them on demand.
    fn expand(&self, at: At) {
        #[cfg(feature = "pages")]
        if at.part == 0
            && self.0.full.get().is_none()
            && matches!(self.0.owner, super::Owner::Pages(_))
            && self.tree().item(at.index).type_name() == "section"
        {
            self.expand_all();
        }
        // Keynote character runs join a slide's part under their paragraphs
        // (record 0298).
        #[cfg(feature = "keynote")]
        if at.part != 0
            && let Part::Slide(whole) = self.part(at.part)
            && whole.full.get().is_none()
            && self.whole(whole).item(at.index).type_name() == "paragraph"
        {
            whole.full.get_or_init(|| self.build_whole(whole, true));
        }
        #[cfg(not(any(feature = "pages", feature = "keynote")))]
        let _ = at;
    }

    /// Build every part completely (views and whole-tree reads).
    pub(crate) fn expand_parts(&self) {
        #[cfg(feature = "keynote")]
        for part in &self.built().parts {
            if let Part::Slide(whole) = part {
                whole.full.get_or_init(|| self.build_whole(whole, true));
            }
        }
    }

    /// Every item in document order, with its place: each item followed by
    /// the items of the part holding its children, then every part's
    /// deferred items (Keynote character runs), part by part.
    pub(crate) fn ordered(&self) -> Vec<(At, &Item)> {
        self.expand_all();
        let tree = self.tree();
        let mut items = Vec::with_capacity(tree.len());
        for (index, item) in (0_u32..).zip(tree.items()) {
            self.push_ordered(At::new(0, index), item, &mut items);
        }
        #[cfg(feature = "keynote")]
        for (part, built) in (1_u32..).zip(self.built().parts.iter()) {
            if let Part::Slide(whole) = built {
                let tree = self.whole(whole);
                for index in tree.deferred()..u32::try_from(tree.len()).unwrap_or(u32::MAX) {
                    items.push((At::new(part, index), tree.item(index)));
                }
            }
        }
        items
    }

    /// Push `item` and, after it, the items of the part holding its
    /// children (a part's deferred items excepted).
    fn push_ordered<'f>(&'f self, at: At, item: &'f Item, items: &mut Vec<(At, &'f Item)>) {
        items.push((at, item));
        let part = item.part();
        if part == 0 {
            return;
        }
        match *self.part(part) {
            #[cfg(feature = "keynote")]
            Part::Slides(ref slides) => {
                for slot in 0..slides.len() {
                    self.push_ordered(At::new(part, slot), slides.item(self, slot), items);
                }
            },
            #[cfg(feature = "keynote")]
            Part::Slide(ref whole) => {
                let tree = self.whole(whole);
                for index in 1..tree.deferred() {
                    items.push((At::new(part, index), tree.item(index)));
                }
            },
            #[cfg(feature = "numbers")]
            Part::Table(ref cells) => {
                for slot in 0..cells.len() {
                    items.push((At::new(part, slot), cells.item(self, slot)));
                }
            },
        }
    }
}

#[cfg(all(test, any(feature = "keynote", feature = "numbers")))]
impl Iwork {
    /// What each part has built: `slides m/n` (materialized slide items,
    /// `+layouts` once the catalog was read), `slide`, `slide+runs` or `-`,
    /// `cells m/n`.
    fn part_states(&self) -> Vec<String> {
        self.built()
            .parts
            .iter()
            .map(|part| match *part {
                #[cfg(feature = "keynote")]
                Part::Slides(ref slides) => slides.state(),
                #[cfg(feature = "keynote")]
                Part::Slide(ref whole) => match (whole.built.get(), whole.full.get()) {
                    (_, Some(_)) => "slide+runs".to_owned(),
                    (Some(_), None) => "slide".to_owned(),
                    (None, None) => "-".to_owned(),
                },
                #[cfg(feature = "numbers")]
                Part::Table(ref cells) => cells.state(),
            })
            .collect()
    }
}

#[cfg(all(test, any(feature = "keynote", feature = "numbers")))]
mod tests {
    use crate::{File, open_bytes};

    fn open(path: &str) -> File {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork")
            .join(path);
        open_bytes(std::fs::read(path).expect("fixture")).expect("open")
    }

    fn states(file: &File) -> Vec<String> {
        match file {
            #[cfg(feature = "keynote")]
            File::Keynote(handle) => handle.part_states(),
            #[cfg(feature = "numbers")]
            File::Numbers(handle) => handle.part_states(),
            _ => unreachable!("an iWork fixture"),
        }
    }

    fn text(file: &File, path: &str) -> Option<String> {
        file.get(path)
            .expect("path")
            .view()
            .text()
            .expect("text")
            .map(std::borrow::Cow::into_owned)
    }

    #[cfg(feature = "keynote")]
    #[test]
    fn keynote_first_get_builds_only_the_slide_it_reaches() {
        let file = open("keynote/multi-slide.key");
        let title = text(&file, "/slide[2]/placeholder[title]");
        // Slide 2's children only; no slide item, no layout catalog.
        assert_eq!(states(&file), ["slides 0/5", "-", "slide", "-", "-", "-"]);
        assert_eq!(title.as_deref(), Some("Agenda"));
        // A slide's own properties materialize that slide and read the
        // catalog once.
        let node = file.get("/slide[3]").expect("slide");
        let format = node.view().format().expect("format");
        assert!(format.into_iter().any(|prop| prop.key() == "hidden"));
        assert_eq!(
            states(&file),
            ["slides 1/5+layouts", "-", "slide", "-", "-", "-"]
        );
        // A paragraph's children add slide 2's runs; nodes taken before stay
        // valid.
        let paragraph = file
            .get("/slide[2]/placeholder[title]/paragraph[1]")
            .expect("paragraph");
        let _ = paragraph.children().count();
        assert_eq!(
            states(&file),
            ["slides 1/5+layouts", "-", "slide+runs", "-", "-", "-"]
        );
        assert_eq!(text(&file, "/slide[2]/placeholder[title]"), title);
        // A view completes every part.
        file.view(crate::view::Mode::Outline).expect("outline");
        assert_eq!(
            states(&file),
            [
                "slides 5/5+layouts",
                "slide+runs",
                "slide+runs",
                "slide+runs",
                "slide+runs",
                "slide+runs"
            ]
        );
    }

    #[cfg(feature = "numbers")]
    #[test]
    fn numbers_cell_get_materializes_one_cell() {
        let file = open("synthetic/numbers/bench-grid.numbers");
        assert_eq!(text(&file, "/sheet[1]/table[1]/A3").as_deref(), Some("R2"));
        assert_eq!(states(&file), ["cells 1/110"]);
        // An unstored in-grid cell is missing without materializing any.
        assert!(file.get("/sheet[1]/table[1]/C3").is_err());
        assert_eq!(states(&file), ["cells 1/110"]);
        // Positional and last steps read the slot's kind and ordinal only.
        assert_eq!(
            text(&file, "/sheet[1]/table[1]/cell[2]").as_deref(),
            Some("C2")
        );
        assert_eq!(states(&file), ["cells 2/110"]);
        // A query matches cells by kind without materializing them; reading
        // them does.
        let count = file.query("cell").expect("query").count();
        assert_eq!(count, 110);
        assert_eq!(states(&file), ["cells 2/110"]);
        let texts = file
            .query("cell")
            .expect("query")
            .filter_map(|node| {
                node.expect("node")
                    .view()
                    .text()
                    .expect("text")
                    .map(|text| text.into_owned())
            })
            .count();
        assert_eq!(texts, 110);
        assert_eq!(states(&file), ["cells 110/110"]);
    }
}
