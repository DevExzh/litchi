//! Inert ODB table presentation and schema declarations.

/// The ODF declaration from which a table was read.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum TableKind {
    /// A `db:table-representation` presentation declaration.
    Representation,
    /// A `db:table-definition` schema declaration.
    Definition,
}

/// One named database column declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Column {
    name: String,
}

impl Column {
    pub(crate) fn parsed(name: String) -> Self {
        Self { name }
    }

    /// Returns the producer-visible column name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// One named table schema or presentation declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Table {
    name: String,
    kind: TableKind,
    columns: Vec<Column>,
}

impl Table {
    pub(crate) fn parsed(name: String, kind: TableKind) -> Self {
        Self {
            name,
            kind,
            columns: Vec::new(),
        }
    }

    pub(crate) fn push_column(&mut self, column: Column) {
        self.columns.push(column);
    }

    /// Returns the producer-visible table name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns whether this is a schema definition or a presentation declaration.
    #[must_use]
    pub const fn kind(&self) -> TableKind {
        self.kind
    }

    /// Returns columns in their ODF source order.
    #[must_use]
    pub fn columns(&self) -> &[Column] {
        &self.columns
    }
}
