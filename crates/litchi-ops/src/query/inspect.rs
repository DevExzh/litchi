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

    /// The node's local properties, in the order of its kind's property
    /// table. Read once per candidate that reaches a property predicate.
    fn format<'n>(&'n self, node: &'n Self::Node) -> Result<Format<'n>>;

    /// The number a number-indexed element carries: a worksheet row or
    /// column number (1-based), a numbering level, a note id. Selectors
    /// compare `row[N]`, `col[B]` and `level[L]` with it (record 0010 §5.2).
    fn number(&self, node: &Self::Node) -> Option<usize> {
        let _ = node;
        None
    }

    /// The A1 coordinate of a cell node, for the column (`B`) and range
    /// (`A1:C5`) element filters of workbook selectors.
    fn cell(&self, node: &Self::Node) -> Option<CellRef> {
        let _ = node;
        None
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

    fn kind(&self, node: &Self::Node) -> Self::Kind {
        (**self).kind(node)
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

    fn format<'n>(&'n self, node: &'n Self::Node) -> Result<Format<'n>> {
        (**self).format(node)
    }

    fn number(&self, node: &Self::Node) -> Option<usize> {
        (**self).number(node)
    }

    fn cell(&self, node: &Self::Node) -> Option<CellRef> {
        (**self).cell(node)
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
