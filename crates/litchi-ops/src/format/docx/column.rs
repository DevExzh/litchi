//! Table grid columns (record 0345): `/body/tbl[N]/col[C]`.
//!
//! OfficeCLI's `column` is a virtual element synthesized from the table's
//! `w:tblGrid` and the cell at each row's column slot. A column node is the
//! table's container entry and the column's position in the grid: `get`
//! reads its `width` (the `w:gridCol` width in twips), `set` writes it
//! (and the matching cells' `w:tcW`), `add` inserts a grid column with a
//! cell in every row (`width`, `text`), and `remove` drops the grid column
//! and its cells, through the leaf's grid edits.

use std::borrow::Cow;

use litchi_core::{FileFormat, Position};

use super::element::{element_path, path_with, refusal};
use super::nested::{parent_node, sibling_position};
use super::node::{Block, Node, Step, Value};
use super::setters::Setter;
use super::spec::Kind;
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::node::spec::GapSpec;
use crate::path::navigate::{Lookup, Select};
use crate::props::Props;
use crate::registry::Operation;
use crate::value::Scalar;

/// The grid widths of a table entry, one per `w:gridCol` of its
/// `w:tblGrid`. Only the table's start is read: the scan stops at the
/// grid's end or the first row, so resolving a column does not parse the
/// table's properties or rows.
pub(super) fn grid(block: &Block) -> Result<Vec<Option<u32>>> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    use quick_xml::events::Event;
    let indexed = block.indexed()?;
    let mut reader = quick_xml::Reader::from_reader(indexed.xml_bytes());
    let mut depth = 0usize;
    let mut in_grid = false;
    let mut widths = Vec::new();
    loop {
        let event = reader.read_event().map_err(|_| {
            Error::new(Code::CorruptFile, "a DOCX table is not well-formed")
                .with_format(FileFormat::Docx)
        })?;
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if in_grid && depth == 1 {
                    return Ok(widths);
                }
                continue;
            },
            Event::Eof => return Ok(widths),
            _ => continue,
        };
        let local = element.local_name();
        match (depth, local.as_ref()) {
            (1, b"tblGrid") if empty => return Ok(widths),
            (1, b"tblGrid") => in_grid = true,
            (1, b"tr") => return Ok(widths),
            (2, b"gridCol") if in_grid => widths.push(
                element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .find(|attribute| attribute.key.local_name().as_ref() == b"w")
                    .and_then(|attribute| {
                        std::str::from_utf8(&attribute.value)
                            .ok()?
                            .trim()
                            .parse()
                            .ok()
                    }),
            ),
            _ => {},
        }
        if !empty {
            depth += 1;
        }
    }
}

/// The column `select` names under the table node `parent`.
pub(super) fn pick_column<'f>(
    docx: &'f Docx,
    parent: &Node<'f>,
    block: &Block,
    select: &Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let count = grid(block)?.len();
    let index = match *select {
        Select::None => (count > 0).then_some(0),
        Select::Ordinal(position) => (position.get() < count).then(|| position.get()),
        Select::Last => count.checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(match index {
        Some(index) => {
            let ordinal = u32::try_from(index).unwrap_or(u32::MAX);
            Lookup::Found(Node::new(
                docx,
                Kind::Column,
                Value::Column(block.clone(), ordinal),
                parent
                    .child_address()
                    .with(Step::new(Kind::Column, ordinal)),
            ))
        },
        None => Lookup::Missing,
    })
}

/// The width of column `column` of a table entry, in twips.
pub(super) fn width(block: &Block, column: u32) -> Result<Option<u32>> {
    Ok(grid(block)?
        .get(usize::try_from(column).unwrap_or(usize::MAX))
        .copied()
        .flatten())
}

/// The value of property `index` of a column whose width is `width`.
pub(super) fn column_value<'a>(width: Option<u32>, index: usize) -> Option<crate::node::Value<'a>> {
    const WIDTH: usize = super::spec::index_of(super::spec::COLUMN_PROPS, "width");
    match index {
        WIDTH => width.map(|width| crate::node::Value::Text(Cow::Owned(width.to_string()))),
        _ => None,
    }
}

fn text_of(value: Option<Scalar<'_>>) -> String {
    match value {
        Some(Scalar::Text(text) | Scalar::Token(text)) => text.to_owned(),
        Some(scalar) => scalar.to_string(),
        None => String::new(),
    }
}

/// A column width: twips (an optional `dxa` suffix) or a length.
fn twips(text: &str) -> Result<u32> {
    super::nested::twips_length(text.trim())
        .filter(|twips| *twips > 0)
        .ok_or_else(|| {
            Error::new(
                Code::InvalidValue,
                "a column width is a positive number of twips or a length such as 3cm",
            )
        })
}

/// The table node a column node belongs to, and the column's position.
fn table_of<'f>(docx: &'f Docx, node: &Node<'_>) -> Result<(Node<'f>, Position)> {
    let Value::Column(_, column) = &node.value else {
        return Err(refusal("the node is not a table column"));
    };
    let table = parent_node(docx, node)?;
    Ok((
        table,
        Position::new(usize::try_from(*column).unwrap_or(usize::MAX)),
    ))
}

impl Docx {
    /// `add(<table>, column, {width, text})` (record 0345).
    pub(super) fn add_column(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, litchi_core::selector::Path<'static>)> {
        let Value::Table(block) = &parent.value else {
            return Err(refusal("the column parent holds no table"));
        };
        let mut width = None;
        let mut text = String::new();
        for (key, value) in node.props().iter() {
            let (_, entry) = super::spec::COLUMN_REGISTRY
                .resolve(key.as_str())
                .map_err(|error| error.with_property(key.as_str().to_owned()))?;
            entry
                .prepare(Operation::Add, value, |setter, value| {
                    match setter {
                        Setter::ColumnWidth => width = Some(twips(&text_of(value))?),
                        Setter::ColumnText => text = text_of(value),
                        _ => {
                            return Err(Error::new(
                                Code::UnsupportedProperty,
                                "the property does not seed a DOCX table column",
                            ));
                        },
                    }
                    Ok(())
                })
                .map_err(|error| {
                    error
                        .with_property(key.as_str().to_owned())
                        .with_format(FileFormat::Docx)
                })?;
        }
        let table = element_path(parent)?;
        let count = grid(block)?.len();
        let at = sibling_position(self, parent, Kind::Column, count, node.placement())?;
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        let position = edit.insert_column_at(&table, at.map(Position::new), width, &text)?;
        let commit = edit.commit(self)?;
        Ok((commit, path_with(parent, Some((Kind::Column, position)))))
    }

    /// A `set` whose targets are table columns (record 0345): `width`.
    /// `Err` hands the editor back for other targets.
    pub(super) fn try_set_column(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty() || nodes.iter().any(|node| node.element_kind() != Kind::Column) {
            return Ok(Err(editor));
        }
        let mut width = None;
        for (key, value) in props.iter() {
            let (_, entry) = super::spec::COLUMN_REGISTRY
                .resolve(key.as_str())
                .map_err(|error| error.with_property(key.as_str().to_owned()))?;
            entry
                .prepare(Operation::Set, value, |setter, value| {
                    match setter {
                        Setter::ColumnWidth => width = Some(twips(&text_of(value))?),
                        _ => {
                            return Err(Error::new(
                                Code::UnsupportedProperty,
                                "the property does not set a DOCX table column",
                            ));
                        },
                    }
                    Ok(())
                })
                .map_err(|error| {
                    error
                        .with_property(key.as_str().to_owned())
                        .with_format(FileFormat::Docx)
                })?;
        }
        let mut targets = Vec::with_capacity(nodes.len());
        for node in nodes {
            let (table, column) = table_of(self, node)?;
            targets.push((element_path(&table)?, column));
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        if let Some(width) = width {
            for (table, column) in &targets {
                edit.set_column_width_at(table, *column, width)?;
            }
        }
        edit.commit(self).map(Ok)
    }

    /// Remove table columns (record 0345), later columns of a table first
    /// so no removal shifts another.
    pub(super) fn remove_columns(&self, editor: DocxEditor, nodes: &[Node<'_>]) -> Result<Commit> {
        let mut targets = Vec::with_capacity(nodes.len());
        for node in nodes {
            let (table, column) = table_of(self, node)?;
            targets.push((
                super::nested::order_key(&element_path(&table)?),
                element_path(&table)?,
                column,
            ));
        }
        targets.sort_by(|left, right| (&right.0, right.2).cmp(&(&left.0, left.2)));
        targets.dedup_by(|left, right| left.0 == right.0 && left.2 == right.2);
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        for (_, table, column) in &targets {
            edit.remove_column_at(table, *column)?;
        }
        edit.commit(self)
    }
}
