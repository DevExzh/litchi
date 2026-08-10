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
    default_value: Option<String>,
}

impl Column {
    /// Creates a named column declaration.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self::parsed(name.into(), ColumnSchema::default())
    }

    /// Sets the standard ODF database type.
    #[must_use]
    pub const fn with_data_type(mut self, value: Option<DataType>) -> Self {
        self.data_type = value;
        self
    }

    /// Sets the producer-specific type name.
    #[must_use]
    pub fn with_type_name(mut self, value: impl Into<String>) -> Self {
        self.type_name = Some(value.into());
        self
    }

    /// Sets the positive precision.
    #[must_use]
    pub const fn with_precision(mut self, value: Option<u64>) -> Self {
        self.precision = value;
        self
    }

    /// Sets the positive scale.
    #[must_use]
    pub const fn with_scale(mut self, value: Option<u64>) -> Self {
        self.scale = value;
        self
    }

    /// Sets schema nullability.
    #[must_use]
    pub const fn with_nullable(mut self, value: Option<bool>) -> Self {
        self.nullable = value;
        self
    }

    /// Sets whether empty values are accepted.
    #[must_use]
    pub const fn with_empty_allowed(mut self, value: Option<bool>) -> Self {
        self.empty_allowed = value;
        self
    }

    /// Sets the auto-increment declaration.
    #[must_use]
    pub const fn with_autoincrement(mut self, value: Option<bool>) -> Self {
        self.autoincrement = value;
        self
    }

    /// Sets the inert producer-encoded default value.
    #[must_use]
    pub fn with_default_value(mut self, value: impl Into<String>) -> Self {
        self.default_value = Some(value.into());
        self
    }

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
            default_value: schema.default_value,
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

    /// Returns the inert producer-encoded default value, if declared.
    #[must_use]
    pub fn default_value(&self) -> Option<&str> {
        self.default_value.as_deref()
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
    pub(crate) default_value: Option<String>,
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
    /// Creates a key column mapping.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: Some(name.into()),
            related_column: None,
        }
    }

    /// Sets the referenced column for a foreign-key mapping.
    #[must_use]
    pub fn with_related_column(mut self, value: impl Into<String>) -> Self {
        self.related_column = Some(value.into());
        self
    }

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

/// A foreign-key relation projected from an ODF `db:key` declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Relation {
    table: String,
    key: Option<String>,
    referenced_table: String,
    update_rule: Option<ReferentialAction>,
    delete_rule: Option<ReferentialAction>,
    columns: Vec<KeyColumn>,
}

impl Relation {
    pub(crate) fn from_key(table: &str, key: &Key) -> Option<Self> {
        let referenced_table = key.referenced_table()?.to_owned();
        Some(Self {
            table: table.to_owned(),
            key: key.name.clone(),
            referenced_table,
            update_rule: key.update_rule,
            delete_rule: key.delete_rule,
            columns: key.columns.clone(),
        })
    }

    /// Returns the table that owns the foreign key.
    #[must_use]
    pub fn table(&self) -> &str {
        &self.table
    }

    /// Returns the foreign-key name, when declared.
    #[must_use]
    pub fn key(&self) -> Option<&str> {
        self.key.as_deref()
    }

    /// Returns the related table name.
    #[must_use]
    pub fn referenced_table(&self) -> &str {
        &self.referenced_table
    }

    /// Returns the update rule.
    #[must_use]
    pub const fn update_rule(&self) -> Option<ReferentialAction> {
        self.update_rule
    }

    /// Returns the delete rule.
    #[must_use]
    pub const fn delete_rule(&self) -> Option<ReferentialAction> {
        self.delete_rule
    }

    /// Returns the local-to-related column mappings.
    #[must_use]
    pub fn columns(&self) -> &[KeyColumn] {
        &self.columns
    }
}

impl Key {
    /// Creates a named key declaration.
    #[must_use]
    pub fn new(name: impl Into<String>, kind: KeyKind) -> Self {
        Self {
            name: Some(name.into()),
            kind,
            referenced_table: None,
            update_rule: None,
            delete_rule: None,
            columns: Vec::new(),
        }
    }

    /// Sets the referenced table for a foreign key.
    #[must_use]
    pub fn with_referenced_table(mut self, value: impl Into<String>) -> Self {
        self.referenced_table = Some(value.into());
        self
    }

    /// Sets the update rule.
    #[must_use]
    pub const fn with_update_rule(mut self, value: Option<ReferentialAction>) -> Self {
        self.update_rule = value;
        self
    }

    /// Sets the delete rule.
    #[must_use]
    pub const fn with_delete_rule(mut self, value: Option<ReferentialAction>) -> Self {
        self.delete_rule = value;
        self
    }

    /// Appends a key-column mapping.
    #[must_use]
    pub fn with_column(mut self, value: KeyColumn) -> Self {
        self.columns.push(value);
        self
    }

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
    /// Creates a named index-column declaration.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            ascending: None,
        }
    }

    /// Sets the optional ascending-order declaration.
    #[must_use]
    pub const fn with_ascending(mut self, value: Option<bool>) -> Self {
        self.ascending = value;
        self
    }

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
    /// Creates a named index declaration.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            unique: None,
            clustered: None,
            columns: Vec::new(),
        }
    }

    /// Sets the optional uniqueness declaration.
    #[must_use]
    pub const fn with_unique(mut self, value: Option<bool>) -> Self {
        self.unique = value;
        self
    }

    /// Sets the optional clustered declaration.
    #[must_use]
    pub const fn with_clustered(mut self, value: Option<bool>) -> Self {
        self.clustered = value;
        self
    }

    /// Appends an index-column declaration.
    #[must_use]
    pub fn with_column(mut self, value: IndexColumn) -> Self {
        self.columns.push(value);
        self
    }

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
    filter_statement: Option<String>,
    order_statement: Option<String>,
}

impl Table {
    /// Creates a named presentation or schema table declaration.
    #[must_use]
    pub fn new(name: impl Into<String>, kind: TableKind) -> Self {
        Self::parsed(name.into(), kind)
    }

    /// Appends a column declaration.
    #[must_use]
    pub fn with_column(mut self, value: Column) -> Self {
        self.columns.push(value);
        self
    }

    /// Appends a key declaration.
    #[must_use]
    pub fn with_key(mut self, value: Key) -> Self {
        self.keys.push(value);
        self
    }

    /// Appends an index declaration.
    #[must_use]
    pub fn with_index(mut self, value: Index) -> Self {
        self.indices.push(value);
        self
    }

    /// Sets the inert filter command for a table presentation.
    #[must_use]
    pub fn with_filter_statement(mut self, value: impl Into<String>) -> Self {
        self.filter_statement = Some(value.into());
        self
    }

    /// Sets the inert ordering command for a table presentation.
    #[must_use]
    pub fn with_order_statement(mut self, value: impl Into<String>) -> Self {
        self.order_statement = Some(value.into());
        self
    }

    pub(crate) fn parsed(name: String, kind: TableKind) -> Self {
        Self {
            name,
            kind,
            columns: Vec::new(),
            keys: Vec::new(),
            indices: Vec::new(),
            filter_statement: None,
            order_statement: None,
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

    pub(crate) fn set_filter_statement(&mut self, value: String) -> Result<()> {
        set_statement(&mut self.filter_statement, value, "filter")
    }

    pub(crate) fn set_order_statement(&mut self, value: String) -> Result<()> {
        set_statement(&mut self.order_statement, value, "order")
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

    /// Finds the first exact named column.
    #[must_use]
    pub fn column(&self, name: &str) -> Option<&Column> {
        self.columns.iter().find(|column| column.name() == name)
    }

    /// Finds the first exact named constraint.
    #[must_use]
    pub fn key(&self, name: &str) -> Option<&Key> {
        self.keys.iter().find(|key| key.name() == Some(name))
    }

    /// Finds the first exact named index.
    #[must_use]
    pub fn index(&self, name: &str) -> Option<&Index> {
        self.indices.iter().find(|index| index.name() == name)
    }

    /// Returns the first primary-key constraint, if declared.
    #[must_use]
    pub fn primary_key(&self) -> Option<&Key> {
        self.keys.iter().find(|key| key.kind() == KeyKind::Primary)
    }

    /// Iterates foreign-key constraints in declaration order.
    pub fn foreign_keys(&self) -> impl Iterator<Item = &Key> {
        self.keys
            .iter()
            .filter(|key| key.kind() == KeyKind::Foreign)
    }

    /// Returns the inert filter command, if declared.
    #[must_use]
    pub fn filter_statement(&self) -> Option<&str> {
        self.filter_statement.as_deref()
    }

    /// Returns the inert ordering command, if declared.
    #[must_use]
    pub fn order_statement(&self) -> Option<&str> {
        self.order_statement.as_deref()
    }
}

fn set_statement(target: &mut Option<String>, value: String, kind: &str) -> Result<()> {
    if target.replace(value).is_some() {
        return Err(Error::InvalidFormat(format!(
            "ODB table contains duplicate {kind} statements"
        )));
    }
    Ok(())
}
