//! Queries: CSS-like selectors over a document (record 0010 §5, M1 step R3).
//!
//! A selector arrives as text, a parsed [`Query`], or a compiled [`Plan`]
//! through [`IntoQuery`]. It is compiled once per format into a [`Plan`]
//! that binds every element name to the format's kinds, every predicate key
//! to the kinds' property tables, and every literal to the key's value type
//! ([`crate::value`]); running the plan streams matches in document order
//! from one iterative depth-first walk, stopping as soon as the caller stops
//! pulling or `limit` is reached.
//!
//! # Semantics
//!
//! | Selector | Meaning |
//! |---|---|
//! | `paragraph`, `p`, `Table` | an element kind, through the format's alias table, ASCII case-insensitively |
//! | `title`, `textbox` | presentation shapes that are, or are not, title placeholders |
//! | `B`, `A1:C5` | workbook cells in a column or a range |
//! | `*` | the direct children of the scope, or of the body, slides or worksheets |
//! | `[key=v]`, `!=`, `~=`, `>`, `<`, `>=`, `<=`, `[key]` | OfficeCLI's `AttributeFilter` comparisons on the key's canonical text, plus typed equality (`[fill=red]` matches `#FF0000`); lengths in different units compare in EMU and are equal within 500 EMU; an absent key matches only `!=` |
//! | `[a and (b or not(c))]` | boolean filters |
//! | `[N]`, `[last()]` | on the selected element, the Nth or last match; on an ancestor, the Nth or last sibling of its kind; `row[3]` and `col[B]` keep their path meaning |
//! | `:contains(t)`, `:empty`, `:no-alt` | case-insensitive text test, no text, no alternative text |
//! | `:has(list)`, `:not(list)` | a descendant matches; the element itself does not match |
//! | `a > b`, `a b` | child and descendant combinators, chained freely |
//! | `a, b` | a union, each node yielded once, in document order |
//! | `/body/tbl[1]/…`, `Sheet1!…` | a scope resolved as a path |
//! | `row[Salary>5000]` | header-row predicates: columns named by a table's header (record 0010 §5.4) |
//!
//! Regular-expression literals (`~= r"…"`) are refused with
//! `unsupported_capability`: the workspace has no regex engine (record 0010
//! open question Q3). An unknown key is a [`Diagnostic`], not an error.
//!
//! # Example
//!
//! ```
//! use litchi_ops::query::{IntoQuery, Plan, Query};
//!
//! // Parse once; a format adapter compiles and runs it (M1 steps R4, R5).
//! let arg = "paragraph[style=Heading1]:contains(revenue)".into_query()?;
//! assert!(arg.query().is_some());
//!
//! let query = Query::parse("row[Region=West and Sales>1000]")?;
//! assert_eq!(query.alternatives().len(), 1);
//! # let _ = Plan::compile;
//! # Ok::<(), litchi_ops::Error>(())
//! ```
//!
//! # OfficeCLI deltas
//!
//! Record 0010 §11.4: matches stream as typed nodes in document order with
//! `limit`, not as collected node maps grouped by kind (M1-Δ19); `:contains`
//! is case-insensitive everywhere (M1-Δ10); `:has`, `:not` and combinators
//! take any selector, not only OfficeCLI's literal forms (M1-Δ21). This
//! implementation adds: `[last()]` on the selected element cannot be combined
//! with `,` alternatives (it would break document order); at most one
//! positional predicate applies to the selected element; positional
//! predicates inside `:has` and `:not` lists select among siblings; `:empty`
//! holds for no text or white space only in every format (OfficeCLI's Word
//! rule; its Excel handler requires empty text); a header-row query scans the
//! bound table's data rows by number, rows without a record included, only
//! when it is the query's single alternative, and otherwise matches rows that
//! have a record; and the diagnostics are computed from the property tables
//! before the walk rather than from the matched nodes after it.

// The format adapters of M1 step R4 and `File::query` (step R5) are the
// production callers of the query engine; until they land, only the
// test-only fake formats run it.
#![cfg_attr(
    not(test),
    allow(
        dead_code,
        reason = "M1 steps R4 and R5 run the query engine; until then only the test-only fake \
                  formats do"
    )
)]

mod arg;
mod bounds;
mod cache;
pub(crate) mod compile;
#[cfg(any(test, feature = "test-support"))]
mod fake;
pub(crate) mod inspect;
mod matches;
mod plan;
#[cfg(test)]
mod tests;
pub(crate) mod walk;

pub use arg::{IntoQuery, QueryArg};
pub use litchi_core::selector::Query;
pub use matches::Matches;
pub use plan::{Diagnostic, DiagnosticKind, Plan};

use crate::Result;
use inspect::Inspect;
use walk::Walk;

/// Runs `selector` over the document the adapter `nav` reads: the entry
/// point `File::query` dispatches to (M1 step R5).
pub(crate) fn run<'q, 'f, N>(nav: N, selector: impl IntoQuery<'q>) -> Result<Walk<'q, 'f, N>>
where
    N: Inspect<'f>,
    N::Node: Clone,
{
    let plan = selector.into_query()?.into_plan::<N>()?;
    Walk::new(nav, plan)
}

/// Test-only access to the query engine over the fake formats, for tests
/// and fuzz targets outside the crate (the counting-allocator test, the
/// plan-compilation fuzz target). Not public API.
#[cfg(feature = "test-support")]
#[doc(hidden)]
pub mod __fake {
    use std::fmt::Write as _;

    use super::compile::compile;
    use super::fake::{Doc, FakeKind, Item, Nav, SheetKind, SlideKind, Table, WordKind};
    use super::walk::{PlanRef, Walk};
    use super::{Plan, Query};
    use crate::Result;
    use crate::node::Value;
    use crate::value::{Color, Length};

    /// The fake format families.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub enum Family {
        /// The Word-like format.
        Word,
        /// The presentation format.
        Presentation,
        /// The workbook format.
        Workbook,
    }

    /// A fake document.
    #[derive(Debug, Clone)]
    pub struct Document(Inner);

    #[derive(Debug, Clone)]
    enum Inner {
        Word(Doc<WordKind>),
        Presentation(Doc<SlideKind>),
        Workbook(Doc<SheetKind>),
    }

    impl Document {
        /// A report: a header, one 3×3 table, then `paragraphs` body
        /// paragraphs of two runs each (every tenth a `Heading1` whose first
        /// run is bold and says "Revenue").
        #[must_use]
        pub fn word(paragraphs: usize) -> Self {
            let body = Item::new(WordKind::Body)
                .child(word_table(3, 3))
                .children((0..paragraphs).map(word_paragraph));
            let header = Item::new(WordKind::Header)
                .prop("type", Value::Token("default"))
                .child(Item::new(WordKind::Paragraph).text("Quarterly report"));
            Self(Inner::Word(Doc::new(
                Item::new(WordKind::Document).child(body).child(header),
            )))
        }

        /// A deck: `slides` slides, each with a title, `shapes` text boxes
        /// and a picture without alternative text.
        #[must_use]
        pub fn deck(slides: usize, shapes: usize) -> Self {
            let slide = |index: usize| {
                let title = Item::new(SlideKind::Shape)
                    .flavour("title")
                    .text(format!("Slide {}", index + 1))
                    .prop("name", "Title 1")
                    .prop("isTitle", true);
                let boxes = (0..shapes).map(|shape| {
                    Item::new(SlideKind::Shape)
                        .flavour("textbox")
                        .text(if shape == 0 { "Revenue grew" } else { "notes" })
                        .prop("name", format!("TextBox {}", shape + 2))
                        .prop("isTitle", false)
                        .prop("width", Length::from_emu(3_810_000))
                        .prop("fill", Color::rgb(0xFF, 0, 0))
                });
                Item::new(SlideKind::Slide)
                    .child(title)
                    .children(boxes)
                    .child(Item::new(SlideKind::Picture).prop("name", "Picture"))
            };
            Self(Inner::Presentation(Doc::new(
                Item::new(SlideKind::Presentation).children((0..slides).map(slide)),
            )))
        }

        /// A workbook: a sheet `Data` holding a `Sales` table (header row 1,
        /// `rows` data rows of Region, Sales and Status) and a sheet `Notes`.
        #[must_use]
        pub fn workbook(rows: usize) -> Self {
            let mut data = Item::new(SheetKind::Sheet).name("Data");
            for row in 1..=rows + 1 {
                data = data.child(Item::new(SheetKind::Row).number(row));
                for (column, text) in sales_row(row) {
                    let cell = Item::new(SheetKind::Cell)
                        .cell(&format!("{column}{row}"))
                        .text(text.clone());
                    let cell = match text.parse::<f64>() {
                        Ok(number) => cell.raw(Value::Number(number)),
                        Err(_) => cell,
                    };
                    data = data.child(cell);
                }
            }
            let notes = Item::new(SheetKind::Sheet).name("Notes");
            let last_row = u32::try_from(rows + 1).unwrap_or(u32::MAX);
            Self(Inner::Workbook(
                Doc::new(Item::new(SheetKind::Workbook).child(data).child(notes)).table(Table {
                    sheet: 0,

                    columns: vec!["Region".into(), "Sales".into(), "Status".into()],
                    first_column: 1,
                    header_row: 1,
                    last_row,
                    totals: false,
                }),
            ))
        }

        /// The document's family.
        #[must_use]
        pub fn family(&self) -> Family {
            match self.0 {
                Inner::Word(_) => Family::Word,
                Inner::Presentation(_) => Family::Presentation,
                Inner::Workbook(_) => Family::Workbook,
            }
        }
    }

    fn word_paragraph(index: usize) -> Item<WordKind> {
        let heading = index % 10 == 0;
        let first = Item::new(WordKind::Run)
            .text(if heading { "Revenue " } else { "Body " })
            .prop("bold", heading);
        let second = Item::new(WordKind::Run).text(format!("paragraph {index}"));
        Item::new(WordKind::Paragraph)
            .text(format!(
                "{}paragraph {index}",
                if heading { "Revenue " } else { "Body " }
            ))
            .prop("style", if heading { "Heading1" } else { "Normal" })
            .child(first)
            .child(second)
    }

    fn word_table(rows: usize, columns: usize) -> Item<WordKind> {
        Item::new(WordKind::Table).children((0..rows).map(|row| {
            Item::new(WordKind::Row).children((0..columns).map(|column| {
                Item::new(WordKind::Cell).child(
                    Item::new(WordKind::Paragraph)
                        .text(format!("cell {row},{column}"))
                        .child(Item::new(WordKind::Run).text(format!("cell {row},{column}"))),
                )
            }))
        }))
    }

    fn sales_row(row: usize) -> [(char, String); 3] {
        if row == 1 {
            return [
                ('A', "Region".to_owned()),
                ('B', "Sales".to_owned()),
                ('C', "Status".to_owned()),
            ];
        }
        let region = ["West", "East", "North", "South"][row % 4];
        let sales = (row * 37 % 2000).to_string();
        let status = if row % 3 == 0 { "open" } else { "closed" };
        [
            ('A', region.to_owned()),
            ('B', sales),
            ('C', status.to_owned()),
        ]
    }

    /// Compiles `query` for `family`.
    ///
    /// # Errors
    ///
    /// As [`Plan::compile`] does for a real format.
    pub fn compile_for<'q>(family: Family, query: &Query<'q>) -> Result<Plan<'q>> {
        match family {
            Family::Word => compile::<&Nav<'_, WordKind>>(query),
            Family::Presentation => compile::<&Nav<'_, SlideKind>>(query),
            Family::Workbook => compile::<&Nav<'_, SheetKind>>(query),
        }
    }

    /// What one run of a plan produced.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub struct Outcome {
        /// The matches yielded.
        pub matches: usize,
        /// The nodes the walk visited.
        pub visited: u64,
    }

    /// Runs `plan` over `document`, yielding at most `limit` matches.
    ///
    /// # Errors
    ///
    /// `invalid_argument` when the plan was compiled for another family,
    /// and any error the walk reports.
    pub fn run(document: &Document, plan: &Plan<'_>, limit: usize) -> Result<Outcome> {
        match &document.0 {
            Inner::Word(doc) => run_on(doc, plan, limit, None),
            Inner::Presentation(doc) => run_on(doc, plan, limit, None),
            Inner::Workbook(doc) => run_on(doc, plan, limit, None),
        }
    }

    /// Runs `plan` over `document` and writes each match's path to `out`,
    /// one per line.
    ///
    /// # Errors
    ///
    /// As [`run`].
    pub fn paths(
        document: &Document,
        plan: &Plan<'_>,
        limit: usize,
        out: &mut String,
    ) -> Result<Outcome> {
        match &document.0 {
            Inner::Word(doc) => run_on(doc, plan, limit, Some(out)),
            Inner::Presentation(doc) => run_on(doc, plan, limit, Some(out)),
            Inner::Workbook(doc) => run_on(doc, plan, limit, Some(out)),
        }
    }

    fn run_on<K: FakeKind>(
        doc: &Doc<K>,
        plan: &Plan<'_>,
        limit: usize,
        mut out: Option<&mut String>,
    ) -> Result<Outcome> {
        let nav = Nav::new(doc);
        let mut walk = Walk::new(&nav, PlanRef::Borrowed(plan))?.limit(limit);
        let mut matches = 0;
        while let Some(hit) = walk.next_hit() {
            let hit = hit?;
            if let Some(out) = out.as_deref_mut() {
                let _ = writeln!(out, "{}", hit.path());
            }
            matches += 1;
        }
        Ok(Outcome {
            matches,
            visited: walk.visited(),
        })
    }
}

/// Explicit native identity access for queries.
pub mod raw {
    /// Opts into native identity predicates and native identity scope paths.
    ///
    /// # Errors
    /// Returns the input's parsing error.
    pub fn query<'q>(query: impl super::IntoQuery<'q>) -> crate::Result<super::QueryArg<'q>> {
        let mut arg = query.into_query()?;
        arg.raw = true;
        Ok(arg)
    }
}
