//! Typed inert `db:schema-definition` metadata.

use super::document::{
    DatabaseContent, DatabaseDocument, DatabaseElement, parse_database_content,
    validate_database_root,
};
use super::query::{OdfDatabaseColumnValue, validate_date, validate_double, validate_duration};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const DB: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MAX_VALUE: usize = 1024 * 1024;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_TABLES: usize = 4096;
const MAX_COLUMNS: usize = 65_536;
const MAX_KEYS: usize = 65_536;
const MAX_INDICES: usize = 65_536;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfDatabaseSchemaPositiveInteger(String);
impl OdfDatabaseSchemaPositiveInteger {
    pub fn new(value: &str) -> Result<Self> {
        let value = value.trim_matches(|c| matches!(c, ' ' | '\t' | '\r' | '\n'));
        let digits = value.strip_prefix('+').unwrap_or(value);
        if digits.is_empty()
            || digits.len() > 4096
            || !digits.bytes().all(|v| v.is_ascii_digit())
            || digits.bytes().all(|v| v == b'0')
        {
            return invalid("invalid database schema positive integer");
        }
        Ok(Self(digits.trim_start_matches('0').to_string()))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDatabaseDataType {
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
impl OdfDatabaseDataType {
    fn parse(v: &str) -> Result<Self> {
        Ok(match v {
            "bit" => Self::Bit,
            "boolean" => Self::Boolean,
            "tinyint" => Self::TinyInt,
            "smallint" => Self::SmallInt,
            "integer" => Self::Integer,
            "bigint" => Self::BigInt,
            "float" => Self::Float,
            "real" => Self::Real,
            "double" => Self::Double,
            "numeric" => Self::Numeric,
            "decimal" => Self::Decimal,
            "char" => Self::Char,
            "varchar" => Self::VarChar,
            "longvarchar" => Self::LongVarChar,
            "date" => Self::Date,
            "time" => Self::Time,
            "timestmp" => Self::Timestamp,
            "binary" => Self::Binary,
            "varbinary" => Self::VarBinary,
            "longvarbinary" => Self::LongVarBinary,
            "sqlnull" => Self::SqlNull,
            "other" => Self::Other,
            "object" => Self::Object,
            "distinct" => Self::Distinct,
            "struct" => Self::Struct,
            "array" => Self::Array,
            "blob" => Self::Blob,
            "clob" => Self::Clob,
            "ref" => Self::Ref,
            _ => return invalid(format!("invalid database data type '{v}'")),
        })
    }
    fn as_str(self) -> &'static str {
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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDatabaseNullable {
    NoNulls,
    Nullable,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDatabaseKeyType {
    Primary,
    Unique,
    Foreign,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDatabaseReferentialRule {
    Cascade,
    Restrict,
    SetNull,
    NoAction,
    SetDefault,
}
impl OdfDatabaseReferentialRule {
    fn parse(v: &str) -> Result<Self> {
        Ok(match v {
            "cascade" => Self::Cascade,
            "restrict" => Self::Restrict,
            "set-null" => Self::SetNull,
            "no-action" => Self::NoAction,
            "set-default" => Self::SetDefault,
            _ => return invalid(format!("invalid referential rule '{v}'")),
        })
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::Cascade => "cascade",
            Self::Restrict => "restrict",
            Self::SetNull => "set-null",
            Self::NoAction => "no-action",
            Self::SetDefault => "set-default",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseColumnDefinition {
    pub name: String,
    pub data_type: Option<OdfDatabaseDataType>,
    pub type_name: Option<String>,
    pub precision: Option<OdfDatabaseSchemaPositiveInteger>,
    pub scale: Option<OdfDatabaseSchemaPositiveInteger>,
    pub nullable: Option<OdfDatabaseNullable>,
    pub empty_allowed: Option<bool>,
    pub autoincrement: Option<bool>,
    pub default_value: Option<OdfDatabaseColumnValue>,
}
impl OdfDatabaseColumnDefinition {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            data_type: None,
            type_name: None,
            precision: None,
            scale: None,
            nullable: None,
            empty_allowed: None,
            autoincrement: None,
            default_value: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseKeyColumn {
    pub name: Option<String>,
    pub related_column_name: Option<String>,
}
impl OdfDatabaseKeyColumn {
    pub fn new(name: impl Into<String>) -> Self {
        Self { name: Some(name.into()), related_column_name: None }
    }
    pub fn foreign(local: impl Into<String>, related: impl Into<String>) -> Self {
        Self { name: Some(local.into()), related_column_name: Some(related.into()) }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseKey {
    pub name: Option<String>,
    pub key_type: OdfDatabaseKeyType,
    pub referenced_table_name: Option<String>,
    pub update_rule: Option<OdfDatabaseReferentialRule>,
    pub delete_rule: Option<OdfDatabaseReferentialRule>,
    pub column_groups: Vec<Vec<OdfDatabaseKeyColumn>>,
}
impl OdfDatabaseKey {
    pub fn primary(name: Option<String>, columns: Vec<String>) -> Self {
        Self { name, key_type: OdfDatabaseKeyType::Primary, referenced_table_name: None,
            update_rule: None, delete_rule: None,
            column_groups: vec![columns.into_iter().map(OdfDatabaseKeyColumn::new).collect()] }
    }
    pub fn unique(name: Option<String>, columns: Vec<String>) -> Self {
        let mut value = Self::primary(name, columns);
        value.key_type = OdfDatabaseKeyType::Unique;
        value
    }
    pub fn foreign(name: Option<String>, table: impl Into<String>, columns: Vec<(String, String)>) -> Self {
        Self { name, key_type: OdfDatabaseKeyType::Foreign, referenced_table_name: Some(table.into()),
            update_rule: None, delete_rule: None,
            column_groups: vec![columns.into_iter().map(|(a, b)| OdfDatabaseKeyColumn::foreign(a, b)).collect()] }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseIndexColumn {
    pub name: String,
    pub ascending: Option<bool>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseIndex {
    pub name: String,
    pub catalog_name: Option<String>,
    pub unique: Option<bool>,
    pub clustered: Option<bool>,
    pub column_groups: Vec<Vec<OdfDatabaseIndexColumn>>,
}
impl OdfDatabaseIndex {
    pub fn new(name: impl Into<String>, columns: Vec<String>) -> Self {
        Self { name: name.into(), catalog_name: None, unique: None, clustered: None,
            column_groups: vec![columns.into_iter().map(|name| OdfDatabaseIndexColumn { name, ascending: None }).collect()] }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseTableDefinition {
    pub name: String,
    pub catalog_name: Option<String>,
    pub schema_name: Option<String>,
    pub table_type: Option<String>,
    pub columns: Vec<OdfDatabaseColumnDefinition>,
    pub keys: Option<Vec<OdfDatabaseKey>>,
    pub indices: Option<Vec<OdfDatabaseIndex>>,
}
impl OdfDatabaseTableDefinition {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            catalog_name: None,
            schema_name: None,
            table_type: None,
            columns: Vec::new(),
            keys: None,
            indices: None,
        }
    }
    pub fn new_view(name: impl Into<String>) -> Self {
        let mut value = Self::new(name);
        value.table_type = Some("VIEW".to_string());
        value
    }
    pub fn is_view(&self) -> bool {
        self.table_type.as_deref().is_some_and(|value| value.eq_ignore_ascii_case("VIEW"))
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseSchemaDefinition {
    pub tables: Vec<OdfDatabaseTableDefinition>,
}
impl OdfDatabaseSchemaDefinition {
    pub fn validate(&self) -> Result<()> {
        validate_schema(self)
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut b = Budget::default();
        let mut x = format!(
            r#"<db:schema-definition xmlns:db="{DB}" xmlns:office="{OFFICE}"><db:table-definitions>"#
        );
        if self.tables.len() > MAX_TABLES {
            return invalid("too many schema tables");
        }
        for v in &self.tables {
            write_table(&mut x, v, &mut b)?;
        }
        x.push_str("</db:table-definitions></db:schema-definition>");
        Ok(x)
    }
}

impl DatabaseDocument {
    pub fn schema_definition(&self) -> Result<Option<OdfDatabaseSchemaDefinition>> {
        schema_from_root(self.database())
    }
}
pub fn parse_database_schema_definition_xml(
    xml: &str,
) -> Result<Option<OdfDatabaseSchemaDefinition>> {
    preflight_schema_xml(xml)?;
    let root = parse_database_content(xml)?;
    validate_database_root(&root)?;
    validate_schema_root_order(&root)?;
    schema_from_root(&root)
}
pub fn set_database_schema_definition_xml(
    xml: &str,
    value: Option<&OdfDatabaseSchemaDefinition>,
) -> Result<String> {
    parse_database_schema_definition_xml(xml)?;
    let span = locate(xml)?;
    let (start, end) = span.schema.unwrap_or((span.close, span.close));
    let replacement = value
        .map(OdfDatabaseSchemaDefinition::to_xml_fragment)
        .transpose()?
        .unwrap_or_default();
    let mut out = String::with_capacity(xml.len() - (end - start) + replacement.len());
    out.push_str(&xml[..start]);
    out.push_str(&replacement);
    out.push_str(&xml[end..]);
    Ok(out)
}

fn preflight_schema_xml(xml: &str) -> Result<()> {
    if xml.len() > 64 * 1024 * 1024 { return invalid("database schema XML is too large"); }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid database schema XML: {error}")))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let resolved = namespace(&resolved)?;
                let local = owned(element.local_name().as_ref())?;
                if resolved.as_deref() == Some("urn:oasis:names:tc:opendocument:xmlns:script:1.0")
                    || (resolved.as_deref() == Some(OFFICE) && local == "event-listeners")
                    || (resolved.as_deref() == Some(OFFICE) && local == "scripts" && matches!(event, Event::Start(_)))
                { return invalid("active content is forbidden in database schema XML"); }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                    if depth > 256 { return invalid("database schema XML is too deeply nested"); }
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => return invalid("DTD, entity references, and processing instructions are forbidden in database schema XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn validate_schema_root_order(root: &DatabaseElement) -> Result<()> {
    let mut previous = 0usize;
    let mut any = false;
    let mut seen = [false; 6];
    for child in children(root)? {
        if child.namespace_uri() != Some(DB) { continue; }
        let rank = match child.local_name() {
            "data-source" => 0, "forms" => 1, "reports" => 2, "queries" => 3,
            "table-representations" => 4, "schema-definition" => 5,
            name => return invalid(format!("unexpected db:{name} child in office:database")),
        };
        if seen[rank] { return invalid(format!("duplicate db:{} child", child.local_name())); }
        if any && rank < previous { return invalid("database direct children are out of schema order"); }
        seen[rank] = true;
        previous = rank;
        any = true;
    }
    if !seen[0] { return invalid("database has no data source"); }
    Ok(())
}

#[derive(Default)]
struct Budget {
    strings: usize,
    columns: usize,
    keys: usize,
    indices: usize,
}
impl Budget {
    fn string(&mut self, label: &str, v: &str) -> Result<String> {
        if v.len() > MAX_VALUE {
            return invalid(format!("schema {label} is too large"));
        }
        self.strings = self
            .strings
            .checked_add(v.len())
            .ok_or_else(|| Error::InvalidFormat("schema string budget overflow".into()))?;
        if self.strings > MAX_AGGREGATE {
            return invalid("schema strings are too large");
        }
        Ok(v.into())
    }
    fn column(&mut self) -> Result<()> {
        self.columns += 1;
        if self.columns > MAX_COLUMNS {
            invalid("too many schema columns")
        } else {
            Ok(())
        }
    }
    fn key(&mut self) -> Result<()> {
        self.keys += 1;
        if self.keys > MAX_KEYS {
            invalid("too many schema keys")
        } else {
            Ok(())
        }
    }
    fn index(&mut self) -> Result<()> {
        self.indices += 1;
        if self.indices > MAX_INDICES {
            invalid("too many schema indices")
        } else {
            Ok(())
        }
    }
}

fn schema_from_root(root: &DatabaseElement) -> Result<Option<OdfDatabaseSchemaDefinition>> {
    let Some(schema) = root
        .children()
        .find(|v| v.namespace_uri() == Some(DB) && v.local_name() == "schema-definition")
    else {
        return Ok(None);
    };
    expect(schema, "schema-definition")?;
    attrs(schema, &[])?;
    let kids = children(schema)?;
    if kids.len() != 1 {
        return invalid("db:schema-definition requires one db:table-definitions");
    }
    expect(kids[0], "table-definitions")?;
    attrs(kids[0], &[])?;
    let tables = children(kids[0])?;
    if tables.len() > MAX_TABLES {
        return invalid("too many schema tables");
    }
    let mut b = Budget::default();
    let value = OdfDatabaseSchemaDefinition {
        tables: tables
            .into_iter()
            .map(|v| parse_table(v, &mut b))
            .collect::<Result<_>>()?,
    };
    value.validate()?;
    Ok(Some(value))
}

fn validate_schema(schema: &OdfDatabaseSchemaDefinition) -> Result<()> {
    if schema.tables.len() > MAX_TABLES { return invalid("too many schema tables"); }
    let mut budget = Budget::default();
    let mut table_names = HashSet::new();
    for table in &schema.tables {
        required_identity(&table.name, "table name", &mut budget)?;
        if !table_names.insert(table.name.as_str()) {
            return invalid(format!("duplicate schema table or view '{}'", table.name));
        }
        optional_string(table.catalog_name.as_deref(), "catalog name", &mut budget)?;
        optional_string(table.schema_name.as_deref(), "schema name", &mut budget)?;
        optional_string(table.table_type.as_deref(), "table type", &mut budget)?;
        if table.columns.is_empty() { return invalid(format!("schema table '{}' has no columns", table.name)); }
        let mut columns = HashSet::new();
        for column in &table.columns {
            budget.column()?;
            required_identity(&column.name, "column name", &mut budget)?;
            if !columns.insert(column.name.as_str()) {
                return invalid(format!("duplicate column '{}' in table '{}'", column.name, table.name));
            }
            optional_string(column.type_name.as_deref(), "column type name", &mut budget)?;
            validate_column_semantics(column)?;
        }
        let mut key_names = HashSet::new();
        let mut primary = false;
        if let Some(keys) = &table.keys {
            if keys.is_empty() { return invalid("db:keys must not be empty"); }
            for key in keys {
                budget.key()?;
                if let Some(name) = &key.name {
                    required_identity(name, "key name", &mut budget)?;
                    if !key_names.insert(name.as_str()) { return invalid(format!("duplicate key '{}' in table '{}'", name, table.name)); }
                }
                if key.key_type == OdfDatabaseKeyType::Primary {
                    if primary { return invalid(format!("table '{}' has multiple primary keys", table.name)); }
                    primary = true;
                }
                validate_key_shape(table, key, &columns, &mut budget)?;
            }
        }
        let mut index_names = HashSet::new();
        if let Some(indices) = &table.indices {
            if indices.is_empty() { return invalid("db:indices must not be empty"); }
            for index in indices {
                budget.index()?;
                required_identity(&index.name, "index name", &mut budget)?;
                if !index_names.insert(index.name.as_str()) { return invalid(format!("duplicate index '{}' in table '{}'", index.name, table.name)); }
                optional_string(index.catalog_name.as_deref(), "index catalog", &mut budget)?;
                if index.column_groups.is_empty() { return invalid("db:index requires index-columns"); }
                let mut used = HashSet::new();
                for group in &index.column_groups {
                    if group.is_empty() { return invalid("db:index-columns must not be empty"); }
                    for column in group {
                        required_identity(&column.name, "index column", &mut budget)?;
                        if !columns.contains(column.name.as_str()) { return invalid(format!("index '{}' references missing column '{}'", index.name, column.name)); }
                        if !used.insert(column.name.as_str()) { return invalid(format!("index '{}' repeats column '{}'", index.name, column.name)); }
                    }
                }
            }
        }
    }
    for table in &schema.tables {
        for key in table.keys.iter().flatten() {
            if key.key_type != OdfDatabaseKeyType::Foreign { continue; }
            let target_name = key.referenced_table_name.as_deref().ok_or_else(|| Error::InvalidFormat(format!("foreign key in '{}' has no referenced table", table.name)))?;
            let target = schema.tables.iter().find(|candidate| candidate.name == target_name)
                .ok_or_else(|| Error::InvalidFormat(format!("foreign key in '{}' references missing table '{}'", table.name, target_name)))?;
            let target_columns: HashSet<&str> = target.columns.iter().map(|column| column.name.as_str()).collect();
            for group in &key.column_groups {
                for pair in group {
                    let related = pair.related_column_name.as_deref().ok_or_else(|| Error::InvalidFormat(format!("foreign key in '{}' has an unpaired related column", table.name)))?;
                    if !target_columns.contains(related) { return invalid(format!("foreign key in '{}' references missing column '{}.{}'", table.name, target.name, related)); }
                }
            }
        }
    }
    Ok(())
}

fn validate_key_shape(table: &OdfDatabaseTableDefinition, key: &OdfDatabaseKey, columns: &HashSet<&str>, budget: &mut Budget) -> Result<()> {
    if key.column_groups.is_empty() { return invalid("db:key requires key-columns"); }
    let foreign = key.key_type == OdfDatabaseKeyType::Foreign;
    if foreign {
        let target = key.referenced_table_name.as_deref().ok_or_else(|| Error::InvalidFormat("foreign key requires referenced table".into()))?;
        required_identity(target, "referenced table", budget)?;
    } else if key.referenced_table_name.is_some() || key.update_rule.is_some() || key.delete_rule.is_some() {
        return invalid("referential attributes are valid only on foreign keys");
    }
    let mut used = HashSet::new();
    for group in &key.column_groups {
        if group.is_empty() { return invalid("db:key-columns must not be empty"); }
        for pair in group {
            let local = pair.name.as_deref().ok_or_else(|| Error::InvalidFormat("key column requires a local name".into()))?;
            required_identity(local, "key column", budget)?;
            if !columns.contains(local) { return invalid(format!("key in table '{}' references missing local column '{}'", table.name, local)); }
            if !used.insert(local) { return invalid(format!("key in table '{}' repeats local column '{}'", table.name, local)); }
            match (foreign, pair.related_column_name.as_deref()) {
                (true, Some(related)) => required_identity(related, "related column", budget)?,
                (true, None) => return invalid("foreign key columns require paired related columns"),
                (false, Some(_)) => return invalid("related columns are valid only on foreign keys"),
                (false, None) => {},
            }
        }
    }
    Ok(())
}

fn validate_column_semantics(column: &OdfDatabaseColumnDefinition) -> Result<()> {
    if column.autoincrement == Some(true) {
        if !matches!(column.data_type, Some(OdfDatabaseDataType::TinyInt | OdfDatabaseDataType::SmallInt | OdfDatabaseDataType::Integer | OdfDatabaseDataType::BigInt)) {
            return invalid(format!("autoincrement column '{}' must have an integer type", column.name));
        }
        if column.default_value.is_some() || column.nullable == Some(OdfDatabaseNullable::Nullable) {
            return invalid(format!("autoincrement column '{}' cannot be nullable or have a default", column.name));
        }
    }
    if column.empty_allowed == Some(true) && !matches!(column.data_type, None | Some(OdfDatabaseDataType::Char | OdfDatabaseDataType::VarChar | OdfDatabaseDataType::LongVarChar | OdfDatabaseDataType::Clob)) {
        return invalid(format!("non-character column '{}' cannot allow empty strings", column.name));
    }
    if let (Some(data_type), Some(value)) = (column.data_type, column.default_value.as_ref()) {
        let compatible = match data_type {
            OdfDatabaseDataType::Bit | OdfDatabaseDataType::Boolean => matches!(value, OdfDatabaseColumnValue::Boolean(_)),
            OdfDatabaseDataType::TinyInt | OdfDatabaseDataType::SmallInt | OdfDatabaseDataType::Integer | OdfDatabaseDataType::BigInt | OdfDatabaseDataType::Float | OdfDatabaseDataType::Real | OdfDatabaseDataType::Double | OdfDatabaseDataType::Numeric | OdfDatabaseDataType::Decimal => matches!(value, OdfDatabaseColumnValue::Float(_) | OdfDatabaseColumnValue::Percentage(_) | OdfDatabaseColumnValue::Currency { .. }),
            OdfDatabaseDataType::Char | OdfDatabaseDataType::VarChar | OdfDatabaseDataType::LongVarChar | OdfDatabaseDataType::Clob => matches!(value, OdfDatabaseColumnValue::String(_)),
            OdfDatabaseDataType::Date | OdfDatabaseDataType::Timestamp => matches!(value, OdfDatabaseColumnValue::Date(_)),
            OdfDatabaseDataType::Time => matches!(value, OdfDatabaseColumnValue::Time(_)),
            _ => false,
        };
        if !compatible { return invalid(format!("default value does not match the type of column '{}'", column.name)); }
    }
    Ok(())
}

fn required_identity(value: &str, label: &str, budget: &mut Budget) -> Result<()> {
    if value.is_empty() { return invalid(format!("schema {label} must not be empty")); }
    budget.string(label, value)?;
    if !value.chars().all(xml_char) { return invalid(format!("schema {label} contains an invalid XML character")); }
    Ok(())
}
fn optional_string(value: Option<&str>, label: &str, budget: &mut Budget) -> Result<()> {
    if let Some(value) = value { required_identity(value, label, budget)?; }
    Ok(())
}
fn xml_char(ch: char) -> bool {
    matches!(ch, '\u{9}' | '\u{A}' | '\u{D}') || ('\u{20}'..='\u{D7FF}').contains(&ch)
        || ('\u{E000}'..='\u{FFFD}').contains(&ch) || ('\u{10000}'..='\u{10FFFF}').contains(&ch)
}

fn parse_table(e: &DatabaseElement, b: &mut Budget) -> Result<OdfDatabaseTableDefinition> {
    expect(e, "table-definition")?;
    attrs(
        e,
        &[
            (DB, "name"),
            (DB, "catalog-name"),
            (DB, "schema-name"),
            (DB, "type"),
        ],
    )?;
    let kids = ordered(e, &["column-definitions", "keys", "indices"])?;
    if kids
        .first()
        .is_none_or(|v| v.local_name() != "column-definitions")
    {
        return invalid("table definition requires column definitions");
    }
    let cols = children(kids[0])?;
    if cols.is_empty() {
        return invalid("column definitions must not be empty");
    }
    let mut columns = Vec::new();
    for v in cols {
        b.column()?;
        columns.push(parse_column(v, b)?);
    }
    let keys = child(&kids, "keys").map(|v| parse_keys(v, b)).transpose()?;
    let indices = child(&kids, "indices")
        .map(|v| parse_indices(v, b))
        .transpose()?;
    Ok(OdfDatabaseTableDefinition {
        name: req(e, "name", b)?,
        catalog_name: opt(e, "catalog-name", b)?,
        schema_name: opt(e, "schema-name", b)?,
        table_type: opt(e, "type", b)?,
        columns,
        keys,
        indices,
    })
}

fn parse_column(e: &DatabaseElement, b: &mut Budget) -> Result<OdfDatabaseColumnDefinition> {
    expect(e, "column-definition")?;
    attrs(
        e,
        &[
            (DB, "name"),
            (DB, "data-type"),
            (DB, "type-name"),
            (DB, "precision"),
            (DB, "scale"),
            (DB, "is-nullable"),
            (DB, "is-empty-allowed"),
            (DB, "is-autoincrement"),
            (OFFICE, "value-type"),
            (OFFICE, "value"),
            (OFFICE, "currency"),
            (OFFICE, "date-value"),
            (OFFICE, "time-value"),
            (OFFICE, "boolean-value"),
            (OFFICE, "string-value"),
        ],
    )?;
    empty(e)?;
    Ok(OdfDatabaseColumnDefinition {
        name: req(e, "name", b)?,
        data_type: e
            .attribute(Some(DB), "data-type")
            .map(OdfDatabaseDataType::parse)
            .transpose()?,
        type_name: opt(e, "type-name", b)?,
        precision: positive(e, "precision")?,
        scale: positive(e, "scale")?,
        nullable: e
            .attribute(Some(DB), "is-nullable")
            .map(|v| match v {
                "no-nulls" => Ok(OdfDatabaseNullable::NoNulls),
                "nullable" => Ok(OdfDatabaseNullable::Nullable),
                _ => invalid("invalid database nullable value"),
            })
            .transpose()?,
        empty_allowed: bool_attr(e, "is-empty-allowed")?,
        autoincrement: bool_attr(e, "is-autoincrement")?,
        default_value: parse_value(e, b)?,
    })
}

fn parse_keys(e: &DatabaseElement, b: &mut Budget) -> Result<Vec<OdfDatabaseKey>> {
    expect(e, "keys")?;
    attrs(e, &[])?;
    let kids = children(e)?;
    if kids.is_empty() {
        return invalid("db:keys must not be empty");
    }
    kids.into_iter()
        .map(|v| {
            b.key()?;
            parse_key(v, b)
        })
        .collect()
}
fn parse_key(e: &DatabaseElement, b: &mut Budget) -> Result<OdfDatabaseKey> {
    expect(e, "key")?;
    attrs(
        e,
        &[
            (DB, "name"),
            (DB, "type"),
            (DB, "referenced-table-name"),
            (DB, "update-rule"),
            (DB, "delete-rule"),
        ],
    )?;
    let groups = children(e)?;
    if groups.is_empty() {
        return invalid("db:key requires key-columns");
    }
    let mut column_groups = Vec::new();
    for g in groups {
        expect(g, "key-columns")?;
        attrs(g, &[])?;
        let cols = children(g)?;
        if cols.is_empty() {
            return invalid("db:key-columns must not be empty");
        }
        column_groups.push(
            cols.into_iter()
                .map(|v| {
                    expect(v, "key-column")?;
                    attrs(v, &[(DB, "name"), (DB, "related-column-name")])?;
                    empty(v)?;
                    Ok(OdfDatabaseKeyColumn {
                        name: opt(v, "name", b)?,
                        related_column_name: opt(v, "related-column-name", b)?,
                    })
                })
                .collect::<Result<_>>()?,
        );
    }
    let key_type = match e.attribute(Some(DB), "type") {
        Some("primary") => OdfDatabaseKeyType::Primary,
        Some("unique") => OdfDatabaseKeyType::Unique,
        Some("foreign") => OdfDatabaseKeyType::Foreign,
        _ => return invalid("invalid or missing database key type"),
    };
    Ok(OdfDatabaseKey {
        name: opt(e, "name", b)?,
        key_type,
        referenced_table_name: opt(e, "referenced-table-name", b)?,
        update_rule: e
            .attribute(Some(DB), "update-rule")
            .map(OdfDatabaseReferentialRule::parse)
            .transpose()?,
        delete_rule: e
            .attribute(Some(DB), "delete-rule")
            .map(OdfDatabaseReferentialRule::parse)
            .transpose()?,
        column_groups,
    })
}

fn parse_indices(e: &DatabaseElement, b: &mut Budget) -> Result<Vec<OdfDatabaseIndex>> {
    expect(e, "indices")?;
    attrs(e, &[])?;
    let kids = children(e)?;
    if kids.is_empty() {
        return invalid("db:indices must not be empty");
    }
    kids.into_iter()
        .map(|v| {
            b.index()?;
            parse_index(v, b)
        })
        .collect()
}
fn parse_index(e: &DatabaseElement, b: &mut Budget) -> Result<OdfDatabaseIndex> {
    expect(e, "index")?;
    attrs(
        e,
        &[
            (DB, "name"),
            (DB, "catalog-name"),
            (DB, "is-unique"),
            (DB, "is-clustered"),
        ],
    )?;
    let groups = children(e)?;
    if groups.is_empty() {
        return invalid("db:index requires index-columns");
    }
    let mut column_groups = Vec::new();
    for g in groups {
        expect(g, "index-columns")?;
        attrs(g, &[])?;
        let cols = children(g)?;
        if cols.is_empty() {
            return invalid("db:index-columns must not be empty");
        }
        column_groups.push(
            cols.into_iter()
                .map(|v| {
                    expect(v, "index-column")?;
                    attrs(v, &[(DB, "name"), (DB, "is-ascending")])?;
                    empty(v)?;
                    Ok(OdfDatabaseIndexColumn {
                        name: req(v, "name", b)?,
                        ascending: bool_attr(v, "is-ascending")?,
                    })
                })
                .collect::<Result<_>>()?,
        );
    }
    Ok(OdfDatabaseIndex {
        name: req(e, "name", b)?,
        catalog_name: opt(e, "catalog-name", b)?,
        unique: bool_attr(e, "is-unique")?,
        clustered: bool_attr(e, "is-clustered")?,
        column_groups,
    })
}

fn parse_value(e: &DatabaseElement, b: &mut Budget) -> Result<Option<OdfDatabaseColumnValue>> {
    let Some(kind) = e.attribute(Some(OFFICE), "value-type") else {
        if [
            "value",
            "currency",
            "date-value",
            "time-value",
            "boolean-value",
            "string-value",
        ]
        .iter()
        .any(|v| e.attribute(Some(OFFICE), v).is_some())
        {
            return invalid("schema default value has no value type");
        }
        return Ok(None);
    };
    let v = match kind {
        "float" => {
            let v = req_ns(e, OFFICE, "value", b)?;
            validate_double(&v)?;
            OdfDatabaseColumnValue::Float(v)
        },
        "percentage" => {
            let v = req_ns(e, OFFICE, "value", b)?;
            validate_double(&v)?;
            OdfDatabaseColumnValue::Percentage(v)
        },
        "currency" => {
            let value = req_ns(e, OFFICE, "value", b)?;
            validate_double(&value)?;
            OdfDatabaseColumnValue::Currency {
                value,
                currency: opt_ns(e, OFFICE, "currency", b)?,
            }
        },
        "date" => {
            let v = req_ns(e, OFFICE, "date-value", b)?;
            validate_date(&v)?;
            OdfDatabaseColumnValue::Date(v)
        },
        "time" => {
            let v = req_ns(e, OFFICE, "time-value", b)?;
            validate_duration(&v)?;
            OdfDatabaseColumnValue::Time(v)
        },
        "boolean" => {
            OdfDatabaseColumnValue::Boolean(strict_bool(&req_ns(e, OFFICE, "boolean-value", b)?)?)
        },
        "string" => OdfDatabaseColumnValue::String(opt_ns(e, OFFICE, "string-value", b)?),
        _ => return invalid("invalid schema default value type"),
    };
    let allowed = match kind {
        "float" | "percentage" => &["value-type", "value"][..],
        "currency" => &["value-type", "value", "currency"],
        "date" => &["value-type", "date-value"],
        "time" => &["value-type", "time-value"],
        "boolean" => &["value-type", "boolean-value"],
        _ => &["value-type", "string-value"],
    };
    for a in e
        .attributes()
        .iter()
        .filter(|v| v.namespace_uri() == Some(OFFICE))
    {
        if !allowed.contains(&a.local_name()) {
            return invalid("incompatible schema default value attributes");
        }
    }
    Ok(Some(v))
}

fn write_table(x: &mut String, v: &OdfDatabaseTableDefinition, b: &mut Budget) -> Result<()> {
    x.push_str("<db:table-definition");
    out(x, "db:name", &v.name, "table name", b)?;
    opt_out(
        x,
        "db:catalog-name",
        v.catalog_name.as_deref(),
        "catalog",
        b,
    )?;
    opt_out(x, "db:schema-name", v.schema_name.as_deref(), "schema", b)?;
    opt_out(x, "db:type", v.table_type.as_deref(), "table type", b)?;
    x.push_str("><db:column-definitions>");
    if v.columns.is_empty() {
        return invalid("column definitions must not be empty");
    }
    for c in &v.columns {
        b.column()?;
        write_column(x, c, b)?;
    }
    x.push_str("</db:column-definitions>");
    if let Some(keys) = &v.keys {
        if keys.is_empty() {
            return invalid("db:keys must not be empty");
        }
        x.push_str("<db:keys>");
        for k in keys {
            b.key()?;
            write_key(x, k, b)?;
        }
        x.push_str("</db:keys>");
    }
    if let Some(indices) = &v.indices {
        if indices.is_empty() {
            return invalid("db:indices must not be empty");
        }
        x.push_str("<db:indices>");
        for i in indices {
            b.index()?;
            write_index(x, i, b)?;
        }
        x.push_str("</db:indices>");
    }
    x.push_str("</db:table-definition>");
    Ok(())
}
fn write_column(x: &mut String, v: &OdfDatabaseColumnDefinition, b: &mut Budget) -> Result<()> {
    x.push_str("<db:column-definition");
    out(x, "db:name", &v.name, "column name", b)?;
    if let Some(t) = v.data_type {
        lit(x, "db:data-type", t.as_str());
    }
    opt_out(x, "db:type-name", v.type_name.as_deref(), "type name", b)?;
    if let Some(n) = &v.precision {
        lit(x, "db:precision", n.as_str());
    }
    if let Some(n) = &v.scale {
        lit(x, "db:scale", n.as_str());
    }
    if let Some(n) = v.nullable {
        lit(
            x,
            "db:is-nullable",
            if n == OdfDatabaseNullable::NoNulls {
                "no-nulls"
            } else {
                "nullable"
            },
        );
    }
    bool_out(x, "db:is-empty-allowed", v.empty_allowed);
    bool_out(x, "db:is-autoincrement", v.autoincrement);
    if let Some(value) = &v.default_value {
        write_value(x, value, b)?;
    }
    x.push_str("/>");
    Ok(())
}
fn write_key(x: &mut String, v: &OdfDatabaseKey, b: &mut Budget) -> Result<()> {
    x.push_str("<db:key");
    opt_out(x, "db:name", v.name.as_deref(), "key name", b)?;
    lit(
        x,
        "db:type",
        match v.key_type {
            OdfDatabaseKeyType::Primary => "primary",
            OdfDatabaseKeyType::Unique => "unique",
            OdfDatabaseKeyType::Foreign => "foreign",
        },
    );
    opt_out(
        x,
        "db:referenced-table-name",
        v.referenced_table_name.as_deref(),
        "referenced table",
        b,
    )?;
    if let Some(r) = v.update_rule {
        lit(x, "db:update-rule", r.as_str())
    }
    if let Some(r) = v.delete_rule {
        lit(x, "db:delete-rule", r.as_str())
    }
    if v.column_groups.is_empty() {
        return invalid("db:key requires key-columns");
    }
    x.push('>');
    for g in &v.column_groups {
        if g.is_empty() {
            return invalid("db:key-columns must not be empty");
        }
        x.push_str("<db:key-columns>");
        for c in g {
            x.push_str("<db:key-column");
            opt_out(x, "db:name", c.name.as_deref(), "key column", b)?;
            opt_out(
                x,
                "db:related-column-name",
                c.related_column_name.as_deref(),
                "related column",
                b,
            )?;
            x.push_str("/>");
        }
        x.push_str("</db:key-columns>");
    }
    x.push_str("</db:key>");
    Ok(())
}
fn write_index(x: &mut String, v: &OdfDatabaseIndex, b: &mut Budget) -> Result<()> {
    x.push_str("<db:index");
    out(x, "db:name", &v.name, "index name", b)?;
    opt_out(
        x,
        "db:catalog-name",
        v.catalog_name.as_deref(),
        "index catalog",
        b,
    )?;
    bool_out(x, "db:is-unique", v.unique);
    bool_out(x, "db:is-clustered", v.clustered);
    if v.column_groups.is_empty() {
        return invalid("db:index requires index-columns");
    }
    x.push('>');
    for g in &v.column_groups {
        if g.is_empty() {
            return invalid("db:index-columns must not be empty");
        }
        x.push_str("<db:index-columns>");
        for c in g {
            x.push_str("<db:index-column");
            out(x, "db:name", &c.name, "index column", b)?;
            bool_out(x, "db:is-ascending", c.ascending);
            x.push_str("/>");
        }
        x.push_str("</db:index-columns>");
    }
    x.push_str("</db:index>");
    Ok(())
}
fn write_value(x: &mut String, v: &OdfDatabaseColumnValue, b: &mut Budget) -> Result<()> {
    match v {
        OdfDatabaseColumnValue::Float(v) => {
            validate_double(v)?;
            lit(x, "office:value-type", "float");
            out(x, "office:value", v, "default float", b)?
        },
        OdfDatabaseColumnValue::Percentage(v) => {
            validate_double(v)?;
            lit(x, "office:value-type", "percentage");
            out(x, "office:value", v, "default percentage", b)?
        },
        OdfDatabaseColumnValue::Currency { value, currency } => {
            validate_double(value)?;
            lit(x, "office:value-type", "currency");
            out(x, "office:value", value, "default currency", b)?;
            opt_out(x, "office:currency", currency.as_deref(), "currency", b)?
        },
        OdfDatabaseColumnValue::Date(v) => {
            validate_date(v)?;
            lit(x, "office:value-type", "date");
            out(x, "office:date-value", v, "default date", b)?
        },
        OdfDatabaseColumnValue::Time(v) => {
            validate_duration(v)?;
            lit(x, "office:value-type", "time");
            out(x, "office:time-value", v, "default time", b)?
        },
        OdfDatabaseColumnValue::Boolean(v) => {
            lit(x, "office:value-type", "boolean");
            lit(x, "office:boolean-value", if *v { "true" } else { "false" })
        },
        OdfDatabaseColumnValue::String(v) => {
            lit(x, "office:value-type", "string");
            opt_out(x, "office:string-value", v.as_deref(), "default string", b)?
        },
    }
    Ok(())
}

fn children(e: &DatabaseElement) -> Result<Vec<&DatabaseElement>> {
    let mut r = Vec::new();
    for v in e.content() {
        match v {
            DatabaseContent::Element(v) => r.push(v),
            DatabaseContent::Text(v) if v.trim().is_empty() => {},
            DatabaseContent::Text(_) => return invalid(format!("text in db:{}", e.local_name())),
        }
    }
    Ok(r)
}
fn ordered<'a>(e: &'a DatabaseElement, names: &[&str]) -> Result<Vec<&'a DatabaseElement>> {
    let v = children(e)?;
    let (mut prior, mut first) = (0, true);
    let mut seen = vec![false; names.len()];
    for c in &v {
        if c.namespace_uri() != Some(DB) {
            return invalid("foreign schema child");
        }
        let rank = names
            .iter()
            .position(|n| *n == c.local_name())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("unexpected schema child {}", c.local_name()))
            })?;
        if seen[rank] || !first && rank < prior {
            return invalid("schema children are duplicated or out of order");
        }
        seen[rank] = true;
        prior = rank;
        first = false;
    }
    Ok(v)
}
fn child<'a>(v: &[&'a DatabaseElement], name: &str) -> Option<&'a DatabaseElement> {
    v.iter().copied().find(|v| v.local_name() == name)
}
fn expect(e: &DatabaseElement, n: &str) -> Result<()> {
    if e.namespace_uri() == Some(DB) && e.local_name() == n {
        Ok(())
    } else {
        invalid(format!("expected db:{n}"))
    }
}
fn empty(e: &DatabaseElement) -> Result<()> {
    if children(e)?.is_empty() {
        Ok(())
    } else {
        invalid(format!("db:{} must be empty", e.local_name()))
    }
}
fn attrs(e: &DatabaseElement, a: &[(&str, &str)]) -> Result<()> {
    for v in e.attributes() {
        if !a
            .iter()
            .any(|(n, l)| v.namespace_uri() == Some(*n) && v.local_name() == *l)
        {
            return invalid(format!("unexpected attribute {}", v.local_name()));
        }
    }
    Ok(())
}
fn req(e: &DatabaseElement, n: &str, b: &mut Budget) -> Result<String> {
    req_ns(e, DB, n, b)
}
fn opt(e: &DatabaseElement, n: &str, b: &mut Budget) -> Result<Option<String>> {
    opt_ns(e, DB, n, b)
}
fn req_ns(e: &DatabaseElement, ns: &str, n: &str, b: &mut Budget) -> Result<String> {
    let v = e
        .attribute(Some(ns), n)
        .ok_or_else(|| Error::InvalidFormat(format!("missing schema {n}")))?;
    b.string(n, v)
}
fn opt_ns(e: &DatabaseElement, ns: &str, n: &str, b: &mut Budget) -> Result<Option<String>> {
    e.attribute(Some(ns), n).map(|v| b.string(n, v)).transpose()
}
fn bool_attr(e: &DatabaseElement, n: &str) -> Result<Option<bool>> {
    e.attribute(Some(DB), n).map(strict_bool).transpose()
}
fn strict_bool(v: &str) -> Result<bool> {
    match v {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => invalid("invalid strict schema boolean"),
    }
}
fn positive(e: &DatabaseElement, n: &str) -> Result<Option<OdfDatabaseSchemaPositiveInteger>> {
    e.attribute(Some(DB), n)
        .map(OdfDatabaseSchemaPositiveInteger::new)
        .transpose()
}
fn out(x: &mut String, n: &str, v: &str, label: &str, b: &mut Budget) -> Result<()> {
    b.string(label, v)?;
    lit(x, n, v);
    Ok(())
}
fn opt_out(x: &mut String, n: &str, v: Option<&str>, label: &str, b: &mut Budget) -> Result<()> {
    if let Some(v) = v {
        out(x, n, v, label, b)?
    }
    Ok(())
}
fn bool_out(x: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        lit(x, n, if v { "true" } else { "false" })
    }
}
fn lit(x: &mut String, n: &str, v: &str) {
    x.push(' ');
    x.push_str(n);
    x.push_str("=\"");
    x.push_str(
        &v.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;"),
    );
    x.push('"')
}

#[derive(Default)]
struct Span {
    close: usize,
    schema: Option<(usize, usize)>,
}
fn locate(xml: &str) -> Result<Span> {
    let mut r = NsReader::from_str(xml);
    let mut b = Vec::new();
    let mut stack: Vec<(Option<String>, String, usize)> = Vec::new();
    let mut depth = None;
    let mut span = Span::default();
    loop {
        let start = r.buffer_position() as usize;
        let (ns, event) = r
            .read_resolved_event_into(&mut b)
            .map_err(|e| Error::InvalidFormat(format!("invalid schema XML: {e}")))?;
        let resolved = namespace(&ns)?;
        drop(ns);
        let end = r.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if resolved.as_deref() == Some(OFFICE) && local == "database" {
                    depth = Some(stack.len())
                }
                stack.push((resolved, local, start));
            },
            Event::Empty(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if depth.is_some_and(|d| stack.len() == d + 1)
                    && resolved.as_deref() == Some(DB)
                    && local == "schema-definition"
                {
                    span.schema = Some((start, end));
                }
            },
            Event::End(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if let Some((ns, opened, at)) = stack.pop() {
                    if opened != local {
                        return invalid("mismatched schema XML");
                    }
                    if depth.is_some_and(|d| stack.len() == d + 1)
                        && ns.as_deref() == Some(DB)
                        && opened == "schema-definition"
                    {
                        span.schema = Some((at, end));
                    }
                    if depth == Some(stack.len())
                        && ns.as_deref() == Some(OFFICE)
                        && opened == "database"
                    {
                        span.close = start;
                        depth = None
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        b.clear();
    }
    if span.close == 0 {
        invalid("could not locate database close")
    } else {
        Ok(span)
    }
}
fn namespace(v: &ResolveResult<'_>) -> Result<Option<String>> {
    match v {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(v)) => std::str::from_utf8(v)
            .map(|v| Some(v.into()))
            .map_err(|_| Error::InvalidFormat("non-UTF-8 namespace".into())),
        ResolveResult::Unknown(v) => {
            invalid(format!("unknown prefix {}", String::from_utf8_lossy(v)))
        },
    }
}
fn owned(v: &[u8]) -> Result<String> {
    std::str::from_utf8(v)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat("non-UTF-8 name".into()))
}
fn invalid<T>(v: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(v.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn wrap(schema: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:d="{DB}" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource x:type="simple" x:href="db"/></d:connection-data></d:data-source><!--keep--><d:queries/>{schema}</o:database></o:body></o:document-content>"#
        )
    }
    #[test]
    fn complete_schema_roundtrip_and_lossless_mutation() {
        let xml = r#"<d:schema-definition><d:table-definitions><d:table-definition d:name="orders" d:schema-name="public" d:type="TABLE"><d:column-definitions><d:column-definition d:name="id" d:data-type="integer" d:precision="+00010" d:is-nullable="no-nulls" d:is-autoincrement="true"/><d:column-definition d:name="created" d:data-type="date" o:value-type="date" o:date-value="2024-02-29"/></d:column-definitions><d:keys><d:key d:name="pk" d:type="primary"><d:key-columns><d:key-column d:name="id"/></d:key-columns></d:key></d:keys><d:indices><d:index d:name="created_idx" d:is-unique="false"><d:index-columns><d:index-column d:name="created" d:is-ascending="false"/></d:index-columns></d:index></d:indices></d:table-definition></d:table-definitions></d:schema-definition>"#;
        let parsed = parse_database_schema_definition_xml(&wrap(xml))
            .unwrap()
            .unwrap();
        assert_eq!(
            parsed.tables[0].columns[0]
                .precision
                .as_ref()
                .unwrap()
                .as_str(),
            "10"
        );
        let canonical = parsed.to_xml_fragment().unwrap();
        assert_eq!(
            parse_database_schema_definition_xml(&wrap(&canonical))
                .unwrap()
                .unwrap(),
            parsed
        );
        let inserted = set_database_schema_definition_xml(&wrap(""), Some(&parsed)).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("<db:schema-definition"));
        assert!(
            set_database_schema_definition_xml(&inserted, None)
                .unwrap()
                .contains("<!--keep--><d:queries/>")
        );
    }
    #[test]
    fn rejects_order_cardinality_lexicals_and_empty_groups() {
        let bad = [
            r#"<d:schema-definition/>"#,
            r#"<d:schema-definition><d:table-definitions><d:table-definition d:name="t"><d:column-definitions/></d:table-definition></d:table-definitions></d:schema-definition>"#,
            r#"<d:schema-definition><d:table-definitions><d:table-definition d:name="t"><d:column-definitions><d:column-definition d:name="c" d:data-type="timestamp"/></d:column-definitions></d:table-definition></d:table-definitions></d:schema-definition>"#,
            r#"<d:schema-definition><d:table-definitions><d:table-definition d:name="t"><d:column-definitions><d:column-definition d:name="c" d:precision="0"/></d:column-definitions><d:keys/></d:table-definition></d:table-definitions></d:schema-definition>"#,
            r#"<d:schema-definition><d:table-definitions><d:table-definition d:name="t"><d:column-definitions><d:column-definition d:name="c"/></d:column-definitions><d:indices/><d:keys><d:key d:type="primary"><d:key-columns><d:key-column/></d:key-columns></d:key></d:keys></d:table-definition></d:table-definitions></d:schema-definition>"#,
        ];
        for v in bad {
            assert!(
                parse_database_schema_definition_xml(&wrap(v)).is_err(),
                "accepted {v}"
            );
        }
    }
}
