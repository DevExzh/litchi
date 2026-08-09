//! Inert ODB table presentation and schema declarations.

use litchi_core::{Error, Result};

/// The ODF declaration from which a table was read.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum TableKind {
    /// A `db:table-representation` presentation declaration.
    Representation,
    /// A `db:table-definition` schema declaration.
    Definition,
}

/// A database column type from the ODF `db:data-type` vocabulary.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DataType {
    Bit,
    Boolean,
    TinyInt,
    SmallInt,
    Integer,
    BigInt,
    Float,
    Real,
    Double,
    Numeric,
    Decimal,
    Char,
    VarChar,
    LongVarChar,
    Date,
    Time,
    Timestamp,
    Binary,
    VarBinary,
    LongVarBinary,
    SqlNull,
    Other,
    Object,
    Distinct,
    Struct,
    Array,
    Blob,
    Clob,
    Ref,
}

impl DataType {
    /// Returns the exact ODF token, including the normative `timestmp` spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Bit => "bit",
            Self::Boolean => "boolean",
            Self::TinyInt => "tinyint",
            Self::SmallInt => "smallint",
            Self::Integer => "integer",
            Self::BigInt => "bigint",
            Self::Float => "float",
            Self::Real => "real",
            Self::Double => "double",
            Self::Numeric => "numeric",
            Self::Decimal => "decimal",
            Self::Char => "char",
            Self::VarChar => "varchar",
            Self::LongVarChar => "longvarchar",
            Self::Date => "date",
            Self::Time => "time",
            Self::Timestamp => "timestmp",
            Self::Binary => "binary",
            Self::VarBinary => "varbinary",
            Self::LongVarBinary => "longvarbinary",
            Self::SqlNull => "sqlnull",
            Self::Other => "other",
            Self::Object => "object",
            Self::Distinct => "distinct",
            Self::Struct => "struct",
            Self::Array => "array",
            Self::Blob => "blob",
            Self::Clob => "clob",
            Self::Ref => "ref",
        }
    }
}

/// One named database column declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Column {
    name: String,
    data_type: Option<DataType>,
    type_name: Option<String>,
    precision: Option<u64>,
    scale: Option<u64>,
    nullable: Option<bool>,
    empty_allowed: Option<bool>,
    autoincrement: Option<bool>,
}

impl Column {
    pub(crate) fn parsed(name: String, schema: ColumnSchema) -> Self {
        Self {
            name,
            data_type: schema.data_type,
            type_name: schema.type_name,
            precision: schema.precision,
            scale: schema.scale,
            nullable: schema.nullable,
            empty_allowed: schema.empty_allowed,
            autoincrement: schema.autoincrement,
        }
    }

    /// Returns the producer-visible column name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the ODF database type token for a schema column, if declared.
    #[must_use]
    pub const fn data_type(&self) -> Option<DataType> {
        self.data_type
    }

    /// Returns the producer-specific database type name, if declared.
    #[must_use]
    pub fn type_name(&self) -> Option<&str> {
        self.type_name.as_deref()
    }

    /// Returns the declared positive precision, if present.
    #[must_use]
    pub const fn precision(&self) -> Option<u64> {
        self.precision
    }

    /// Returns the declared positive scale, if present.
    #[must_use]
    pub const fn scale(&self) -> Option<u64> {
        self.scale
    }

    /// Returns schema nullability (`true` for `nullable`), if declared.
    #[must_use]
    pub const fn nullable(&self) -> Option<bool> {
        self.nullable
    }

    /// Returns whether empty values are allowed, if declared.
    #[must_use]
    pub const fn empty_allowed(&self) -> Option<bool> {
        self.empty_allowed
    }

    /// Returns whether this is an auto-increment column, if declared.
    #[must_use]
    pub const fn autoincrement(&self) -> Option<bool> {
        self.autoincrement
    }
}

#[derive(Default)]
pub(crate) struct ColumnSchema {
    pub(crate) data_type: Option<DataType>,
    pub(crate) type_name: Option<String>,
    pub(crate) precision: Option<u64>,
    pub(crate) scale: Option<u64>,
    pub(crate) nullable: Option<bool>,
    pub(crate) empty_allowed: Option<bool>,
    pub(crate) autoincrement: Option<bool>,
}

/// The constraint category of an ODF database key.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum KeyKind {
    /// A primary key.
    Primary,
    /// A uniqueness constraint.
    Unique,
    /// A foreign key.
    Foreign,
}

/// A referential action declared for a foreign key.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ReferentialAction {
    Cascade,
    Restrict,
    SetNull,
    NoAction,
    SetDefault,
}

/// One column mapping in a database key.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct KeyColumn {
    name: Option<String>,
    related_column: Option<String>,
}

impl KeyColumn {
    pub(crate) const fn parsed(name: Option<String>, related_column: Option<String>) -> Self {
        Self {
            name,
            related_column,
        }
    }

    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    #[must_use]
    pub fn related_column(&self) -> Option<&str> {
        self.related_column.as_deref()
    }
}

/// A primary, unique, or foreign database key declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Key {
    name: Option<String>,
    kind: KeyKind,
    referenced_table: Option<String>,
    update_rule: Option<ReferentialAction>,
    delete_rule: Option<ReferentialAction>,
    columns: Vec<KeyColumn>,
}

impl Key {
    pub(crate) const fn parsed(
        name: Option<String>,
        kind: KeyKind,
        referenced_table: Option<String>,
        update_rule: Option<ReferentialAction>,
        delete_rule: Option<ReferentialAction>,
    ) -> Self {
        Self {
            name,
            kind,
            referenced_table,
            update_rule,
            delete_rule,
            columns: Vec::new(),
        }
    }

    pub(crate) fn try_push_column(&mut self, column: KeyColumn) -> Result<()> {
        self.columns
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB key columns",
                source,
            })?;
        self.columns.push(column);
        Ok(())
    }

    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    #[must_use]
    pub const fn kind(&self) -> KeyKind {
        self.kind
    }

    #[must_use]
    pub fn referenced_table(&self) -> Option<&str> {
        self.referenced_table.as_deref()
    }

    #[must_use]
    pub const fn update_rule(&self) -> Option<ReferentialAction> {
        self.update_rule
    }

    #[must_use]
    pub const fn delete_rule(&self) -> Option<ReferentialAction> {
        self.delete_rule
    }

    #[must_use]
    pub fn columns(&self) -> &[KeyColumn] {
        &self.columns
    }
}

/// One column ordering declaration in an index.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct IndexColumn {
    name: String,
    ascending: Option<bool>,
}

impl IndexColumn {
    pub(crate) const fn parsed(name: String, ascending: Option<bool>) -> Self {
        Self { name, ascending }
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub const fn ascending(&self) -> Option<bool> {
        self.ascending
    }
}

/// A database index declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Index {
    name: String,
    unique: Option<bool>,
    clustered: Option<bool>,
    columns: Vec<IndexColumn>,
}

impl Index {
    pub(crate) const fn parsed(
        name: String,
        unique: Option<bool>,
        clustered: Option<bool>,
    ) -> Self {
        Self {
            name,
            unique,
            clustered,
            columns: Vec::new(),
        }
    }

    pub(crate) fn try_push_column(&mut self, column: IndexColumn) -> Result<()> {
        self.columns
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB index columns",
                source,
            })?;
        self.columns.push(column);
        Ok(())
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub const fn unique(&self) -> Option<bool> {
        self.unique
    }

    #[must_use]
    pub const fn clustered(&self) -> Option<bool> {
        self.clustered
    }

    #[must_use]
    pub fn columns(&self) -> &[IndexColumn] {
        &self.columns
    }
}

/// One named table schema or presentation declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Table {
    name: String,
    kind: TableKind,
    columns: Vec<Column>,
    keys: Vec<Key>,
    indices: Vec<Index>,
}

impl Table {
    pub(crate) fn parsed(name: String, kind: TableKind) -> Self {
        Self {
            name,
            kind,
            columns: Vec::new(),
            keys: Vec::new(),
            indices: Vec::new(),
        }
    }

    pub(crate) fn try_push_column(&mut self, column: Column) -> Result<()> {
        self.columns
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB table columns",
                source,
            })?;
        self.columns.push(column);
        Ok(())
    }

    pub(crate) fn try_push_key(&mut self, key: Key) -> Result<()> {
        self.keys
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB table keys",
                source,
            })?;
        self.keys.push(key);
        Ok(())
    }

    pub(crate) fn try_push_index(&mut self, index: Index) -> Result<()> {
        self.indices
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB table indices",
                source,
            })?;
        self.indices.push(index);
        Ok(())
    }

    pub(crate) fn keys_mut(&mut self) -> &mut [Key] {
        &mut self.keys
    }

    pub(crate) fn indices_mut(&mut self) -> &mut [Index] {
        &mut self.indices
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

    /// Returns schema key declarations in source order.
    #[must_use]
    pub fn keys(&self) -> &[Key] {
        &self.keys
    }

    /// Returns schema index declarations in source order.
    #[must_use]
    pub fn indices(&self) -> &[Index] {
        &self.indices
    }
}
