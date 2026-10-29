//! The seam between the query walker and a format adapter (record 0010
//! §5.3, §5.4): what the walker reads from a node beyond navigation, and the
//! header-row hooks a workbook adapter supplies.
//!
//! A format adapter implements [`Navigate`] for path resolution (M1 step R2)
//! and [`Inspect`] for queries. Every method is called with static dispatch;
//! the readers mirror the node view's (`node::Project`), so an adapter whose
//! node implements the view delegates to it.

use std::borrow::Cow;
use std::fmt;

use litchi_core::selector::{CellRef, Position};

use crate::error::{Code, Error, Result};
use crate::node::spec::FormatSpec;
use crate::node::{Format, PathStep, Value};
use crate::path::navigate::{ElementKind, Navigate};

/// A format adapter's read access for the query walker.
///
/// The walker asks [`Navigate::children`] for the children of each node it
/// enters, so the adapter decides how lazily each axis is read; the methods
/// here read one node and never decode another part.
pub(crate) trait Inspect<'f>: Navigate<'f> {
    /// Whether property predicates should use the adapter's individual
    /// property hook instead of materializing the cached aggregate format.
    /// Object-tree iWork adapters opt in for lazy chart properties; all other
    /// adapters retain the aggregate path and its allocation behavior.
    const INDIVIDUAL_PROPERTIES: bool = false;

    /// The format's spec table (record 0010 §7.2). The walker prunes with
    /// each element's `children` list and compiles predicate keys against
    /// its `props`. The element of a kind sits at the kind's
    /// [`ElementKind::ordinal`].
    const SPEC: &'static FormatSpec<Self::Kind>;

    /// The containers whose direct children `*` selects in a query without a
    /// scope: Word's `body`, a slide, a worksheet (OfficeCLI's top-level
    /// elements, `WordHandler.Query.cs` 1748-1774).
    const TOP: &'static [Self::Kind];

    /// The element kind of `node`.
    fn kind(&self, node: &Self::Node) -> Self::Kind;

    /// The OfficeCLI `type` of `node` (`paragraph`, `title`, `textbox`), the
    /// value of a selector's `type` key.
    fn type_name(&self, node: &Self::Node) -> &'static str;

    /// The node's text, for `:contains`, `:empty` and the `text` key; `None`
    /// when the element has no text.
    fn text<'n>(&'n self, node: &'n Self::Node) -> Result<Option<Cow<'n, str>>>;

    /// The node's `style` member (Word paragraphs), the fallback of the
    /// `style` key when the kind declares no `style` property.
    fn style<'n>(&'n self, node: &'n Self::Node) -> Result<Option<Cow<'n, str>>> {
        let _ = node;
        Ok(None)
    }

    /// The children of `parent` whose kinds are in `kinds`, for a walk that
    /// is likely to stop after a few of them: one with a `limit` or `first()`,
    /// or whose every alternative selects its `[N]`th match (record 0267).
    ///
    /// An adapter may stream them instead of reading the parent whole: a
    /// worksheet adapter reads its cells through the leaf's cell cursor,
    /// which parses only the leading rows the walk visits. The children and
    /// their order are those of [`Navigate::children`]; the default is that
    /// call.
    fn early_children(
        &self,
        parent: &Self::Node,
        kinds: crate::path::navigate::KindSet<Self::Kind>,
    ) -> Result<Self::Children> {
        self.children(parent, kinds)
    }

    /// Whether `node` has no children whose kinds are in `kinds`, known
    /// without building its child iterator. The walker then does not enter
    /// it: a matched node is yielded without a clone, a frame or a child
    /// read (record 0280). The default, `false`, always enters; an adapter
    /// answers `true` only when [`Navigate::children`] with `kinds` would be
    /// empty.
    fn childless(
        &self,
        node: &Self::Node,
        kinds: crate::path::navigate::KindSet<Self::Kind>,
    ) -> bool {
        let _ = (node, kinds);
        false
    }

    /// The node's local properties, in the order of its kind's property
    /// table. Read once per candidate that reaches a property predicate.
    fn format<'n>(&'n self, node: &'n Self::Node) -> Result<Format<'n>>;

    /// The value of the property at `index` of the node's property table:
    /// what [`Inspect::format`] gives there, read without the rest of the
    /// set when the adapter can. The walker reads a condition's one property
    /// this way when no other test of the compound reads properties.
    fn prop<'n>(&'n self, node: &'n Self::Node, index: usize) -> Result<Option<Value<'n>>> {
        Ok(self.format(node)?.get(index))
    }

    /// The number a number-indexed element carries: a worksheet row or
    /// column number (1-based), a numbering level, a note id. Selectors
    /// compare `row[N]`, `col[B]` and `level[L]` with it (record 0010 §5.2).
    fn number(&self, node: &Self::Node) -> Option<usize> {
        let _ = node;
        None
    }

    /// Whether the format walks presentation placeholders as `shape` nodes
    /// (PPTX `p:sp` with a `p:ph` binding, ODP shapes with a
    /// `presentation:class`). A bare `placeholder` selector then compiles to
    /// `shape` with the [`Inspect::is_placeholder`] test, so it matches the
    /// same nodes `shape` does, at their `/shape[N]` addresses. A format
    /// that walks placeholders as their own kind (PPT) keeps `false`.
    const PLACEHOLDER_SHAPES: bool = false;

    /// Whether `node` is a presentation placeholder, for formats with
    /// [`Inspect::PLACEHOLDER_SHAPES`].
    fn is_placeholder(&self, node: &Self::Node) -> bool {
        let _ = node;
        false
    }

    /// The A1 coordinate of a cell node, for the column (`B`) and range
    /// (`A1:C5`) element filters of workbook selectors.
    fn cell(&self, node: &Self::Node) -> Option<CellRef> {
        let _ = node;
        None
    }

    /// The zero-based `(column, row)` of [`Inspect::cell`], which an adapter
    /// that holds the coordinate gives without spelling it out.
    fn cell_at(&self, node: &Self::Node) -> Option<(u32, u32)> {
        self.cell(node).map(|cell| (cell.column(), cell.row()))
    }

    /// Writes the path step of `node`, the child at `ordinal` among its
    /// parent's children of kind `kind` (record 0010 §2.6): `p[3]` by
    /// default. A workbook writes a sheet's name, a cell's A1 reference and
    /// a row's number; an element with a stable identity may write
    /// `p[@paraId=…]`.
    fn write_step(
        &self,
        node: &Self::Node,
        kind: Self::Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        let _ = node;
        fmt::Display::fmt(&PathStep::at(kind.name(), ordinal), out)
    }

    /// Binds the columns of a header-row predicate (`row[Salary>5000]`,
    /// record 0010 §5.4) to the one table in `scope` (the workbook root or a
    /// worksheet) whose header names every key: a table part (ListObject)
    /// first, else a detected table. Called once per query.
    ///
    /// # Errors
    ///
    /// `not_found` when no table owns every column (with the tables' column
    /// names as the suggestion), `invalid_selector` when several do, and
    /// `unsupported_capability` (the default) when the format has no table
    /// binding.
    // Errors must omit authored column/table names and predicate values.
    fn bind_header(&self, scope: &Self::Node, keys: &[HeaderKey<'_>]) -> Result<HeaderBinding> {
        let _ = (scope, keys);
        Err(Error::new(
            Code::UnsupportedCapability,
            format!(
                "header-row predicates are not supported for {:?} documents",
                Self::FORMAT
            ),
        )
        .with_format(Self::FORMAT)
        .with_suggestion("filter rows by row properties, or select cells by column letter"))
    }

    /// Whether a table in `scope` has a column named `key`, which makes a
    /// bare row-property key (`row[hidden=true]`) ambiguous (OfficeCLI's
    /// `RowKeyCollidesWithColumn`).
    ///
    /// # Errors
    ///
    /// A leaf error while reading the table metadata.
    fn names_column(&self, scope: &Self::Node, key: &str) -> Result<bool> {
        let _ = (scope, key);
        Ok(false)
    }

    /// The value of the cell in `column` (1-based) of the worksheet row
    /// `row`: the raw stored value when `raw` is set (a number for a numeric
    /// cell), else the displayed text. `None` when the row has no cell
    /// there, which compares as empty text.
    ///
    /// # Errors
    ///
    /// A leaf error while reading the cell.
    fn row_cell<'n>(
        &'n self,
        row: &'n Self::Node,
        column: u32,
        raw: bool,
    ) -> Result<Option<Value<'n>>> {
        let _ = (row, column, raw);
        Ok(None)
    }
}

/// A shared reference reads as the adapter it points to (see the
/// forwarding [`Navigate`] implementation).
impl<'f, N: Inspect<'f>> Inspect<'f> for &N {
    const SPEC: &'static FormatSpec<Self::Kind> = N::SPEC;
    const TOP: &'static [Self::Kind] = N::TOP;
    const PLACEHOLDER_SHAPES: bool = N::PLACEHOLDER_SHAPES;
    const INDIVIDUAL_PROPERTIES: bool = N::INDIVIDUAL_PROPERTIES;

    fn kind(&self, node: &Self::Node) -> Self::Kind {
        (**self).kind(node)
    }

    fn is_placeholder(&self, node: &Self::Node) -> bool {
        (**self).is_placeholder(node)
    }

    fn type_name(&self, node: &Self::Node) -> &'static str {
        (**self).type_name(node)
    }

    fn text<'n>(&'n self, node: &'n Self::Node) -> Result<Option<Cow<'n, str>>> {
        (**self).text(node)
    }

    fn style<'n>(&'n self, node: &'n Self::Node) -> Result<Option<Cow<'n, str>>> {
        (**self).style(node)
    }

    fn early_children(
        &self,
        parent: &Self::Node,
        kinds: crate::path::navigate::KindSet<Self::Kind>,
    ) -> Result<Self::Children> {
        (**self).early_children(parent, kinds)
    }

    fn childless(
        &self,
        node: &Self::Node,
        kinds: crate::path::navigate::KindSet<Self::Kind>,
    ) -> bool {
        (**self).childless(node, kinds)
    }

    fn format<'n>(&'n self, node: &'n Self::Node) -> Result<Format<'n>> {
        (**self).format(node)
    }

    fn prop<'n>(&'n self, node: &'n Self::Node, index: usize) -> Result<Option<Value<'n>>> {
        (**self).prop(node, index)
    }

    fn number(&self, node: &Self::Node) -> Option<usize> {
        (**self).number(node)
    }

    fn cell(&self, node: &Self::Node) -> Option<CellRef> {
        (**self).cell(node)
    }

    fn cell_at(&self, node: &Self::Node) -> Option<(u32, u32)> {
        (**self).cell_at(node)
    }

    fn write_step(
        &self,
        node: &Self::Node,
        kind: Self::Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        (**self).write_step(node, kind, ordinal, out)
    }

    fn bind_header(&self, scope: &Self::Node, keys: &[HeaderKey<'_>]) -> Result<HeaderBinding> {
        (**self).bind_header(scope, keys)
    }

    fn names_column(&self, scope: &Self::Node, key: &str) -> Result<bool> {
        (**self).names_column(scope, key)
    }

    fn row_cell<'n>(
        &'n self,
        row: &'n Self::Node,
        column: u32,
        raw: bool,
    ) -> Result<Option<Value<'n>>> {
        (**self).row_cell(row, column, raw)
    }
}

/// One table column a header-row predicate names (record 0010 §5.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct HeaderKey<'k> {
    /// The header name or column letters, without the `col.`/`column.`
    /// prefix.
    pub(crate) name: &'k str,
    /// Whether a `col.`/`column.` prefix forced the column reading.
    pub(crate) forced: bool,
}

/// The table a header-row predicate binds to: its worksheet, its data rows
/// and the column of each key.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct HeaderBinding {
    /// The worksheet's position among the workbook's sheets.
    pub(crate) sheet: Position,
    /// The first and last data rows, 1-based and inclusive: the header and
    /// totals rows are excluded.
    pub(crate) rows: (u32, u32),
    /// The 1-based worksheet column of each key, in the order the keys were
    /// given.
    pub(crate) columns: Vec<u32>,
}
