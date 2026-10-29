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
//!   leaf's sorted cell list and builds only that item;
//! - the contents of a Keynote slide table or a Pages body table (the
//!   table's own properties, its rows and cells) are built whole the first
//!   time a read reaches the table's properties or children (record 0312).
//!   The tree holding the table lists a stand-in item with the table's kind,
//!   ordinal, path and leaf positions; [`Iwork::item`] serves the stand-in's
//!   reads from the built table. A Pages body table's contents are a part
//!   of their own; a Keynote slide table's live in its slide's part, at
//!   indices flagged with [`TABLE`].
//!
//! A step's kind and ordinal are known without materializing the item.
//! Every item keeps its place ([`At`]) once built, so nodes taken before a
//! part grew stay valid.

#![allow(
    irrefutable_let_patterns,
    reason = "a build with one lazily built format has fewer kinds of part"
)]

#[cfg(any(feature = "keynote", feature = "pages"))]
use std::sync::OnceLock;

use super::Iwork;
use super::tree::{At, Item};
#[cfg(any(feature = "keynote", feature = "pages"))]
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
    /// The contents of a Pages body table.
    #[cfg(feature = "pages")]
    Contents(Contents),
}

/// The contents of one table: the table item itself (the anchor, at
/// [`ROOT`]), its rows and cells, built whole on the first read that needs
/// them from the stand-in item the holding tree lists (record 0312).
#[cfg(any(feature = "keynote", feature = "pages"))]
pub(crate) struct Contents {
    /// The leaf's selector of the table: a Keynote table's position among
    /// its slide's tables in z-order, a Pages body table's catalog position.
    selector: usize,
    /// The index of the stand-in in its holding tree (the eager Pages
    /// tree, a Keynote slide's part).
    stand_in: u32,
    built: OnceLock<Tree>,
}

#[cfg(any(feature = "keynote", feature = "pages"))]
impl Contents {
    pub(crate) fn new(selector: usize, stand_in: u32) -> Self {
        Self {
            selector,
            stand_in,
            built: OnceLock::new(),
        }
    }
}

/// The flag of an index in a Keynote slide's part that addresses an item
/// of one of the slide's tables: bits 21–30 hold the table, bits 0–20 the
/// item's index in the table's contents.
#[cfg(feature = "keynote")]
pub(crate) const TABLE: u32 = 1 << 31;
/// The bits of a flagged index holding the item's index in its table.
#[cfg(feature = "keynote")]
const TABLE_ITEM: u32 = (1 << 21) - 1;
/// How many items one Keynote slide table's contents can hold; a larger
/// table keeps its frame without rows.
#[cfg(feature = "keynote")]
pub(crate) const TABLE_ITEMS: usize = TABLE_ITEM as usize;
/// How many tables one Keynote slide's part addresses; later tables keep
/// their frames without rows.
#[cfg(feature = "keynote")]
pub(crate) const SLIDE_TABLES: usize = 1 << 10;

/// The flagged base index of the items of table `table` of a slide.
#[cfg(feature = "keynote")]
const fn table_base(table: u32) -> u32 {
    TABLE | (table << 21)
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
    /// The contents of the slide's tables, the `k`-th held by the stand-in
    /// whose part is `k + 1` (record 0312); set by the first build.
    tables: OnceLock<Box<[Contents]>>,
}

#[cfg(feature = "keynote")]
impl Whole {
    pub(crate) fn new(position: usize, path: Box<[crate::node::PathStep<'static>]>) -> Self {
        Self {
            position,
            path,
            built: OnceLock::new(),
            full: OnceLock::new(),
            tables: OnceLock::new(),
        }
    }

    /// The contents held by the stand-in whose part is `part`.
    fn table(&self, part: u32) -> Option<&Contents> {
        self.tables.get()?.get(part.checked_sub(1)? as usize)
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
    /// The item's child list, each index with the flag bits of its part
    /// (a Keynote table's base, [`TABLE`]).
    Items(&'f [u32], u32),
    /// Every slot of a part materialized item by item, in order.
    #[cfg(any(feature = "keynote", feature = "numbers"))]
    Slots(u32),
}

impl<'f> Kids<'f> {
    pub(crate) const fn len(&self) -> usize {
        match self.list {
            List::Items(items, _) => items.len(),
            #[cfg(any(feature = "keynote", feature = "numbers"))]
            List::Slots(count) => count as usize,
        }
    }

    pub(crate) fn get(&self, position: usize) -> Option<At> {
        let index = match self.list {
            List::Items(items, flag) => *items.get(position)? | flag,
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
                let (tree, tables) =
                    super::keynote::build_slide(package, whole.position, &whole.path, deferred);
                // Both builds list the same tables; the first sets them.
                let _ = whole.tables.set(
                    tables
                        .into_iter()
                        .map(|(selector, stand_in)| Contents::new(selector, stand_in))
                        .collect(),
                );
                tree
            },
            #[allow(unreachable_patterns, reason = "only Keynote slides build whole parts")]
            _ => Tree::root_only(),
        }
    }

    /// The contents of a table whose stand-in is `stand_in`, built on the
    /// first call.
    #[cfg(any(feature = "keynote", feature = "pages"))]
    fn contents<'f>(&'f self, contents: &'f Contents, stand_in: &Item) -> &'f Tree {
        contents.built.get_or_init(|| match &self.0.owner {
            #[cfg(feature = "keynote")]
            super::Owner::Keynote(package) => {
                super::keynote::build_table(package, stand_in, contents.selector)
            },
            #[cfg(feature = "pages")]
            super::Owner::Pages(package) => {
                super::pages::build_table(package, stand_in, contents.selector)
            },
            #[allow(unreachable_patterns, reason = "only Keynote and Pages tables")]
            _ => Tree::root_only(),
        })
    }

    /// The table contents a stand-in item of part `at.part` holds, if it
    /// is one: its part is a table's.
    #[cfg(any(feature = "keynote", feature = "pages"))]
    fn held<'f>(&'f self, at: At, item: &'f Item) -> Option<(&'f Contents, u32)> {
        if at.is_chart() {
            return None;
        }
        let part = item.part();
        if part == 0 {
            return None;
        }
        match (at.part, self.part_or_slots(at.part)) {
            #[cfg(feature = "pages")]
            (0, _) => match *self.part(part) {
                Part::Contents(ref contents) => Some((contents, 0)),
                #[allow(unreachable_patterns, reason = "the part arms depend on features")]
                _ => None,
            },
            #[cfg(feature = "keynote")]
            (_, Some(Part::Slide(whole))) => whole
                .table(part)
                .map(|contents| (contents, table_base(part - 1))),
            _ => None,
        }
    }

    /// The part `part`, or `None` for the eager tree.
    #[cfg(any(feature = "keynote", feature = "pages"))]
    fn part_or_slots(&self, part: u32) -> Option<&Part> {
        (part != 0).then(|| self.part(part))
    }

    /// The item at `at` as its tree holds it: a table's stand-in, not the
    /// table (record 0312). Its kind, ordinal, path, leaf positions and
    /// part are the element's; its properties and children may not be.
    pub(crate) fn stand_in(&self, at: At) -> &Item {
        if at.is_chart() {
            return self.chart_item(at);
        }
        if at.part == 0 {
            return self.tree().item(at.index);
        }
        match *self.part(at.part) {
            #[cfg(feature = "keynote")]
            Part::Slides(ref slides) => slides.item(self, at.index),
            #[cfg(feature = "keynote")]
            Part::Slide(ref whole) => {
                if at.index & TABLE != 0 {
                    return self.table_item(whole, at.index);
                }
                self.whole(whole).item(at.index)
            },
            #[cfg(feature = "numbers")]
            Part::Table(ref cells) => cells.item(self, at.index),
            #[cfg(feature = "pages")]
            Part::Contents(ref contents) => self
                .contents(contents, self.tree().item(contents.stand_in))
                .item(at.index),
        }
    }

    /// The item at a flagged index of a Keynote slide's part.
    #[cfg(feature = "keynote")]
    fn table_item<'f>(&'f self, whole: &'f Whole, index: u32) -> &'f Item {
        let table = (index & !TABLE) >> 21;
        let tree = self.whole(whole);
        match whole.table(table + 1) {
            Some(contents) => self
                .contents(contents, tree.item(contents.stand_in))
                .item(index & TABLE_ITEM),
            // A flagged index is handed out only by its table's contents.
            None => tree.item(ROOT),
        }
    }

    /// The item at `at`.
    pub(crate) fn item(&self, at: At) -> &Item {
        let item = self.stand_in(at);
        #[cfg(any(feature = "keynote", feature = "pages"))]
        if let Some((contents, _)) = self.held(at, item) {
            return self.contents(contents, item).item(ROOT);
        }
        item
    }

    /// The kind ordinal of the item at `at`, without materializing it.
    pub(crate) fn kind_at(&self, at: At) -> u8 {
        if at.is_chart() {
            return self.chart_item(at).kind();
        }
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
        self.stand_in(at).kind()
    }

    /// The position of the item at `at` among its parent's children of its
    /// kind, without materializing it.
    pub(crate) fn ordinal_at(&self, at: At) -> u32 {
        if at.is_chart() {
            return self.chart_item(at).ordinal();
        }
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
        self.stand_in(at).ordinal()
    }

    /// The part holding the children of the item at `at` (0: its own),
    /// without materializing a slot item.
    fn child_part(&self, at: At) -> u32 {
        if at.is_chart() {
            return 0;
        }
        #[cfg(feature = "keynote")]
        if at.part != 0
            && let Part::Slides(_) = self.part(at.part)
        {
            return super::keynote::Slides::child_part(at.index);
        }
        self.stand_in(at).part()
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
        if at.is_chart() {
            return Kids {
                part: at.part,
                list: List::Items(&[], 0),
            };
        }
        self.expand(at);
        #[cfg(any(feature = "keynote", feature = "pages"))]
        if !self.is_slot(at) {
            let item = self.stand_in(at);
            if let Some((contents, base)) = self.held(at, item) {
                let tree = self.contents(contents, item);
                let part = if at.part == 0 { item.part() } else { at.part };
                return Kids {
                    part,
                    list: List::Items(tree.item(ROOT).children(), base),
                };
            }
        }
        match self.child_part(at) {
            0 => Kids {
                part: at.part,
                list: List::Items(self.stand_in(at).children(), Self::flag(at)),
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
                    list: List::Items(self.whole(whole).item(ROOT).children(), 0),
                },
                #[cfg(feature = "numbers")]
                Part::Table(ref cells) => Kids {
                    part,
                    list: List::Slots(cells.len()),
                },
                #[cfg(feature = "pages")]
                Part::Contents(_) => Kids {
                    part,
                    list: List::Items(&[], 0),
                },
            },
        }
    }

    /// Whether `at` is a slot of a part materialized item by item (a
    /// Keynote slide, a Numbers cell), which holds no table.
    #[cfg(any(feature = "keynote", feature = "pages"))]
    fn is_slot(&self, at: At) -> bool {
        if at.part == 0 {
            return false;
        }
        match *self.part(at.part) {
            #[cfg(feature = "keynote")]
            Part::Slides(_) => true,
            #[cfg(feature = "numbers")]
            Part::Table(_) => true,
            #[allow(unreachable_patterns, reason = "the part arms depend on features")]
            _ => false,
        }
    }

    /// The flag bits the children of the item at `at` share with it: a
    /// Keynote table's rows hold their cells at the table's flagged base.
    #[allow(
        clippy::missing_const_for_fn,
        reason = "const in builds without Keynote only"
    )]
    fn flag(at: At) -> u32 {
        #[cfg(feature = "keynote")]
        if at.index & TABLE != 0 {
            return at.index & !TABLE_ITEM;
        }
        let _ = at;
        0
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
            && at.index & TABLE == 0
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
    /// children (a part's deferred items excepted). A table's stand-in
    /// pushes the table, then its rows and cells.
    fn push_ordered<'f>(&'f self, at: At, item: &'f Item, items: &mut Vec<(At, &'f Item)>) {
        #[cfg(any(feature = "keynote", feature = "pages"))]
        if let Some((contents, base)) = self.held(at, item) {
            let tree = self.contents(contents, item);
            items.push((at, tree.item(ROOT)));
            let part = if at.part == 0 { item.part() } else { at.part };
            for index in 1..u32::try_from(tree.len()).unwrap_or(u32::MAX) {
                items.push((At::new(part, index | base), tree.item(index)));
            }
            return;
        }
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
                    self.push_ordered(At::new(part, index), tree.item(index), items);
                }
            },
            #[cfg(feature = "numbers")]
            Part::Table(ref cells) => {
                for slot in 0..cells.len() {
                    items.push((At::new(part, slot), cells.item(self, slot)));
                }
            },
            #[cfg(feature = "pages")]
            Part::Contents(_) => {},
        }
    }
}

#[cfg(all(test, any(feature = "keynote", feature = "numbers", feature = "pages")))]
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
                Part::Slide(ref whole) => {
                    let state = match (whole.built.get(), whole.full.get()) {
                        (_, Some(_)) => "slide+runs",
                        (Some(_), None) => "slide",
                        (None, None) => "-",
                    };
                    match whole.tables.get().filter(|tables| !tables.is_empty()) {
                        Some(tables) => format!(
                            "{state} tables {}/{}",
                            tables
                                .iter()
                                .filter(|table| table.built.get().is_some())
                                .count(),
                            tables.len()
                        ),
                        None => state.to_owned(),
                    }
                },
                #[cfg(feature = "numbers")]
                Part::Table(ref cells) => cells.state(),
                #[cfg(feature = "pages")]
                Part::Contents(ref contents) => if contents.built.get().is_some() {
                    "table"
                } else {
                    "-"
                }
                .to_owned(),
            })
            .collect()
    }
}

#[cfg(all(test, any(feature = "keynote", feature = "numbers", feature = "pages")))]
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
            #[cfg(feature = "pages")]
            File::Pages(handle) => handle.part_states(),
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

    #[cfg(feature = "keynote")]
    #[test]
    fn keynote_slide_tables_build_when_a_read_reaches_them() {
        let file = open("keynote/table-discovery.key");
        // The slide's other children leave the table unread.
        assert_eq!(
            text(&file, "/slide[1]/shape[1]").as_deref(),
            Some("2026-08-07")
        );
        assert_eq!(states(&file), ["slides 0/1", "slide tables 0/1"]);
        // A kind query and the table's kind and ordinal read the stand-in.
        assert_eq!(file.query("table").expect("query").count(), 1);
        assert_eq!(states(&file), ["slides 0/1", "slide tables 0/1"]);
        // A cell builds the table once.
        assert_eq!(
            text(&file, "/slide[1]/table[1]/tr[1]/tc[1]").as_deref(),
            Some("\"Buffa discovery\n\"")
        );
        assert_eq!(states(&file), ["slides 0/1", "slide tables 1/1"]);
        let node = file.get("/slide[1]/table[1]").expect("table");
        let format = node.view().format().expect("format");
        assert!(format.into_iter().any(|prop| prop.key() == "rows"));
        // A direct cell write commits without reading the table.
        let fresh = open("keynote/table-discovery.key");
        let edited = fresh
            .set(
                "/slide[1]/table[1]/tr[2]/tc[3]",
                crate::props! { "text" => "Q3" },
            )
            .expect("set");
        assert_eq!(states(&fresh), ["slides 0/1", "slide tables 0/1"]);
        assert_eq!(
            text(&edited, "/slide[1]/table[1]/tr[2]/tc[3]").as_deref(),
            Some("Q3")
        );
        // A view completes the table.
        let fresh = open("keynote/table-discovery.key");
        fresh.view(crate::view::Mode::Outline).expect("outline");
        assert_eq!(
            states(&fresh),
            ["slides 1/1+layouts", "slide+runs tables 1/1"]
        );
    }

    #[cfg(feature = "pages")]
    #[test]
    fn pages_body_tables_build_when_a_read_reaches_them() {
        let file = open("pages/body-table-catalog-native.pages");
        assert!(text(&file, "/body/p[1]").is_some());
        assert_eq!(states(&file), ["-", "-"]);
        assert_eq!(file.query("tbl").expect("query").count(), 2);
        assert_eq!(states(&file), ["-", "-"]);
        let node = file.get("/body/tbl[2]").expect("table");
        let format = node.view().format().expect("format");
        assert!(format.into_iter().any(|prop| prop.key() == "rows"));
        assert_eq!(states(&file), ["-", "table"]);
        assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[1]").as_deref(), Some(""));
        assert_eq!(states(&file), ["table", "table"]);
        let fresh = open("pages/body-table-catalog-native.pages");
        let edited = fresh
            .set("/body/tbl[1]/tr[1]/tc[1]", crate::props! { "text" => "Q3" })
            .expect("set");
        assert_eq!(states(&fresh), ["-", "-"]);
        assert_eq!(
            text(&edited, "/body/tbl[1]/tr[1]/tc[1]").as_deref(),
            Some("Q3")
        );
    }
}
