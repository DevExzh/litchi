//! Typed, inert ODF database queries and table representations.

use super::document::{
    DatabaseContent, DatabaseDocument, DatabaseElement, parse_database_content,
    validate_database_root,
};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const DB: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_VALUE: usize = 1024 * 1024;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_ITEMS: usize = 4096;
const MAX_COLUMNS: usize = 65_536;
const MAX_COLLECTION_DEPTH: usize = 64;
const MAX_XML_DEPTH: usize = 256;

/// An inert `db:order-statement` or `db:filter-statement` command.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseStatement {
    pub command: String,
    pub apply_command: Option<bool>,
}

impl OdfDatabaseStatement {
    pub fn new(command: impl Into<String>) -> Self {
        Self {
            command: command.into(),
            apply_command: None,
        }
    }
}

/// The optional update target associated with a stored query.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseUpdateTable {
    pub name: String,
    pub catalog_name: Option<String>,
    pub schema_name: Option<String>,
}

impl OdfDatabaseUpdateTable {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            catalog_name: None,
            schema_name: None,
        }
    }
}

/// A schema-typed default value carried by `db:column`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfDatabaseColumnValue {
    Float(String),
    Percentage(String),
    Currency {
        value: String,
        currency: Option<String>,
    },
    Date(String),
    Time(String),
    Boolean(bool),
    String(Option<String>),
}

/// Display and default-value metadata for a query or table column.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseColumn {
    pub name: String,
    pub title: Option<String>,
    pub description: Option<String>,
    pub visible: Option<bool>,
    pub style_name: Option<String>,
    pub default_cell_style_name: Option<String>,
    pub value: Option<OdfDatabaseColumnValue>,
}

impl OdfDatabaseColumn {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            title: None,
            description: None,
            visible: None,
            style_name: None,
            default_cell_style_name: None,
            value: None,
        }
    }
}

/// A stored, inert query definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseQuery {
    pub name: String,
    pub title: Option<String>,
    pub description: Option<String>,
    pub command: String,
    pub escape_processing: Option<bool>,
    pub style_name: Option<String>,
    pub default_row_style_name: Option<String>,
    pub order_statement: Option<OdfDatabaseStatement>,
    pub filter_statement: Option<OdfDatabaseStatement>,
    pub columns: Option<Vec<OdfDatabaseColumn>>,
    pub update_table: Option<OdfDatabaseUpdateTable>,
}

impl OdfDatabaseQuery {
    pub fn new(name: impl Into<String>, command: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            title: None,
            description: None,
            command: command.into(),
            escape_processing: None,
            style_name: None,
            default_row_style_name: None,
            order_statement: None,
            filter_statement: None,
            columns: None,
            update_table: None,
        }
    }
}

/// A recursive member of `db:queries` or `db:query-collection`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(clippy::large_enum_variant)] // public API; boxing would break callers
pub enum OdfDatabaseQueryItem {
    Query(OdfDatabaseQuery),
    Collection(OdfDatabaseQueryCollection),
}

/// A named recursive query collection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseQueryCollection {
    pub name: String,
    pub title: Option<String>,
    pub description: Option<String>,
    pub items: Vec<OdfDatabaseQueryItem>,
}

impl OdfDatabaseQueryCollection {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            title: None,
            description: None,
            items: Vec::new(),
        }
    }
}

/// The optional `db:queries` subtree.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseQueries {
    pub items: Vec<OdfDatabaseQueryItem>,
}

impl OdfDatabaseQueries {
    pub fn to_xml_fragment(&self) -> Result<String> {
        let mut budget = Budget::default();
        let mut xml = format!(r#"<db:queries xmlns:db="{DB}" xmlns:office="{OFFICE}">"#);
        for item in &self.items {
            write_query_item(&mut xml, item, 1, &mut budget)?;
        }
        xml.push_str("</db:queries>");
        Ok(xml)
    }
}

/// A saved table presentation with inert filter and order commands.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseTableRepresentation {
    pub name: String,
    pub catalog_name: Option<String>,
    pub schema_name: Option<String>,
    pub title: Option<String>,
    pub description: Option<String>,
    pub style_name: Option<String>,
    pub default_row_style_name: Option<String>,
    pub order_statement: Option<OdfDatabaseStatement>,
    pub filter_statement: Option<OdfDatabaseStatement>,
    pub columns: Option<Vec<OdfDatabaseColumn>>,
}

impl OdfDatabaseTableRepresentation {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            catalog_name: None,
            schema_name: None,
            title: None,
            description: None,
            style_name: None,
            default_row_style_name: None,
            order_statement: None,
            filter_statement: None,
            columns: None,
        }
    }
}

/// The optional `db:table-representations` subtree.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseTableRepresentations {
    pub items: Vec<OdfDatabaseTableRepresentation>,
}

impl OdfDatabaseTableRepresentations {
    pub fn to_xml_fragment(&self) -> Result<String> {
        if self.items.len() > MAX_ITEMS {
            return invalid("too many database table representations");
        }
        let mut budget = Budget::default();
        let mut xml =
            format!(r#"<db:table-representations xmlns:db="{DB}" xmlns:office="{OFFICE}">"#);
        for table in &self.items {
            write_table(&mut xml, table, &mut budget)?;
        }
        xml.push_str("</db:table-representations>");
        Ok(xml)
    }
}

/// Both adjacent query-related direct children of `office:database`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseQueryModel {
    pub queries: Option<OdfDatabaseQueries>,
    pub table_representations: Option<OdfDatabaseTableRepresentations>,
}

impl DatabaseDocument {
    pub fn queries(&self) -> Result<Option<OdfDatabaseQueries>> {
        Ok(model_from_root(self.database())?.queries)
    }
    pub fn table_representations(&self) -> Result<Option<OdfDatabaseTableRepresentations>> {
        Ok(model_from_root(self.database())?.table_representations)
    }
    pub fn query_model(&self) -> Result<OdfDatabaseQueryModel> {
        model_from_root(self.database())
    }
}

pub fn parse_database_query_model_xml(xml: &str) -> Result<OdfDatabaseQueryModel> {
    preflight(xml)?;
    let root = parse_database_content(xml)?;
    validate_database_root(&root)?;
    model_from_root(&root)
}

pub fn parse_database_queries_xml(xml: &str) -> Result<Option<OdfDatabaseQueries>> {
    Ok(parse_database_query_model_xml(xml)?.queries)
}

pub fn parse_database_table_representations_xml(
    xml: &str,
) -> Result<Option<OdfDatabaseTableRepresentations>> {
    Ok(parse_database_query_model_xml(xml)?.table_representations)
}

pub fn set_database_queries_xml(xml: &str, value: Option<&OdfDatabaseQueries>) -> Result<String> {
    mutate(
        xml,
        Target::Queries,
        value.map(OdfDatabaseQueries::to_xml_fragment).transpose()?,
    )
}

pub fn set_database_table_representations_xml(
    xml: &str,
    value: Option<&OdfDatabaseTableRepresentations>,
) -> Result<String> {
    mutate(
        xml,
        Target::Tables,
        value
            .map(OdfDatabaseTableRepresentations::to_xml_fragment)
            .transpose()?,
    )
}

#[derive(Default)]
struct Budget {
    aggregate: usize,
    items: usize,
    columns: usize,
}

impl Budget {
    fn string(&mut self, label: &str, value: &str) -> Result<String> {
        if value.len() > MAX_VALUE {
            return invalid(format!("database {label} is too large"));
        }
        self.aggregate = self.aggregate.checked_add(value.len()).ok_or_else(|| {
            Error::InvalidFormat("database query string budget overflow".to_string())
        })?;
        if self.aggregate > MAX_AGGREGATE {
            return invalid("database query strings are too large");
        }
        Ok(value.to_string())
    }
    fn item(&mut self) -> Result<()> {
        self.items += 1;
        if self.items > MAX_ITEMS {
            invalid("too many database queries or collections")
        } else {
            Ok(())
        }
    }
    fn column(&mut self) -> Result<()> {
        self.columns += 1;
        if self.columns > MAX_COLUMNS {
            invalid("too many database query columns")
        } else {
            Ok(())
        }
    }
}

fn model_from_root(root: &DatabaseElement) -> Result<OdfDatabaseQueryModel> {
    let children = strict_root_children(root)?;
    let mut budget = Budget::default();
    let queries = children
        .iter()
        .find(|v| v.local_name() == "queries")
        .map(|v| parse_queries(v, &mut budget))
        .transpose()?;
    let table_representations = children
        .iter()
        .find(|v| v.local_name() == "table-representations")
        .map(|v| parse_tables(v, &mut budget))
        .transpose()?;
    Ok(OdfDatabaseQueryModel {
        queries,
        table_representations,
    })
}

fn strict_root_children(root: &DatabaseElement) -> Result<Vec<&DatabaseElement>> {
    let mut result = Vec::new();
    let mut previous = 0;
    let mut seen = [false; 6];
    for child in children(root)? {
        if child.namespace_uri() != Some(DB) {
            return invalid("foreign child in office:database");
        }
        let rank = match child.local_name() {
            "data-source" => 0,
            "forms" => 1,
            "reports" => 2,
            "queries" => 3,
            "table-representations" => 4,
            "schema-definition" => 5,
            name => return invalid(format!("unexpected db:{name} child in office:database")),
        };
        if seen[rank] {
            return invalid(format!("duplicate db:{} child", child.local_name()));
        }
        if !result.is_empty() && rank < previous {
            return invalid("database direct children are out of schema order");
        }
        seen[rank] = true;
        previous = rank;
        result.push(child);
    }
    if !seen[0] {
        return invalid("database has no data source");
    }
    Ok(result)
}

fn parse_queries(element: &DatabaseElement, budget: &mut Budget) -> Result<OdfDatabaseQueries> {
    expect(element, "queries")?;
    allow_attrs(element, &[])?;
    let mut items = Vec::new();
    for child in children(element)? {
        items.push(parse_item(child, 1, budget)?);
    }
    Ok(OdfDatabaseQueries { items })
}

fn parse_item(
    element: &DatabaseElement,
    depth: usize,
    budget: &mut Budget,
) -> Result<OdfDatabaseQueryItem> {
    if depth > MAX_COLLECTION_DEPTH {
        return invalid("database query collections are too deeply nested");
    }
    budget.item()?;
    match element.local_name() {
        "query" => Ok(OdfDatabaseQueryItem::Query(parse_query(element, budget)?)),
        "query-collection" => {
            expect(element, "query-collection")?;
            allow_attrs(element, &[(DB, "name"), (DB, "title"), (DB, "description")])?;
            let mut items = Vec::new();
            for child in children(element)? {
                items.push(parse_item(child, depth + 1, budget)?);
            }
            Ok(OdfDatabaseQueryItem::Collection(
                OdfDatabaseQueryCollection {
                    name: required(element, DB, "name", budget)?,
                    title: optional(element, DB, "title", budget)?,
                    description: optional(element, DB, "description", budget)?,
                    items,
                },
            ))
        },
        name => invalid(format!("unexpected db:{name} child in db:queries")),
    }
}

fn parse_query(element: &DatabaseElement, budget: &mut Budget) -> Result<OdfDatabaseQuery> {
    expect(element, "query")?;
    allow_attrs(
        element,
        &[
            (DB, "command"),
            (DB, "escape-processing"),
            (DB, "name"),
            (DB, "title"),
            (DB, "description"),
            (DB, "style-name"),
            (DB, "default-row-style-name"),
        ],
    )?;
    let kids = ordered(
        element,
        &[
            "order-statement",
            "filter-statement",
            "columns",
            "update-table",
        ],
    )?;
    Ok(OdfDatabaseQuery {
        name: required(element, DB, "name", budget)?,
        title: optional(element, DB, "title", budget)?,
        description: optional(element, DB, "description", budget)?,
        command: required(element, DB, "command", budget)?,
        escape_processing: optional_bool(element, DB, "escape-processing")?,
        style_name: optional(element, DB, "style-name", budget)?,
        default_row_style_name: optional(element, DB, "default-row-style-name", budget)?,
        order_statement: child(&kids, "order-statement")
            .map(|v| parse_statement(v, "order-statement", budget))
            .transpose()?,
        filter_statement: child(&kids, "filter-statement")
            .map(|v| parse_statement(v, "filter-statement", budget))
            .transpose()?,
        columns: child(&kids, "columns")
            .map(|v| parse_columns(v, budget))
            .transpose()?,
        update_table: child(&kids, "update-table")
            .map(|v| parse_update(v, budget))
            .transpose()?,
    })
}

fn parse_statement(
    element: &DatabaseElement,
    local: &str,
    budget: &mut Budget,
) -> Result<OdfDatabaseStatement> {
    expect(element, local)?;
    allow_attrs(element, &[(DB, "command"), (DB, "apply-command")])?;
    empty(element)?;
    Ok(OdfDatabaseStatement {
        command: required(element, DB, "command", budget)?,
        apply_command: optional_bool(element, DB, "apply-command")?,
    })
}

fn parse_update(element: &DatabaseElement, budget: &mut Budget) -> Result<OdfDatabaseUpdateTable> {
    expect(element, "update-table")?;
    allow_attrs(
        element,
        &[(DB, "name"), (DB, "catalog-name"), (DB, "schema-name")],
    )?;
    empty(element)?;
    Ok(OdfDatabaseUpdateTable {
        name: required(element, DB, "name", budget)?,
        catalog_name: optional(element, DB, "catalog-name", budget)?,
        schema_name: optional(element, DB, "schema-name", budget)?,
    })
}

fn parse_tables(
    element: &DatabaseElement,
    budget: &mut Budget,
) -> Result<OdfDatabaseTableRepresentations> {
    expect(element, "table-representations")?;
    allow_attrs(element, &[])?;
    let kids = children(element)?;
    if kids.len() > MAX_ITEMS {
        return invalid("too many database table representations");
    }
    Ok(OdfDatabaseTableRepresentations {
        items: kids
            .into_iter()
            .map(|v| parse_table(v, budget))
            .collect::<Result<_>>()?,
    })
}

fn parse_table(
    element: &DatabaseElement,
    budget: &mut Budget,
) -> Result<OdfDatabaseTableRepresentation> {
    expect(element, "table-representation")?;
    allow_attrs(
        element,
        &[
            (DB, "name"),
            (DB, "catalog-name"),
            (DB, "schema-name"),
            (DB, "title"),
            (DB, "description"),
            (DB, "style-name"),
            (DB, "default-row-style-name"),
        ],
    )?;
    let kids = ordered(element, &["order-statement", "filter-statement", "columns"])?;
    Ok(OdfDatabaseTableRepresentation {
        name: required(element, DB, "name", budget)?,
        catalog_name: optional(element, DB, "catalog-name", budget)?,
        schema_name: optional(element, DB, "schema-name", budget)?,
        title: optional(element, DB, "title", budget)?,
        description: optional(element, DB, "description", budget)?,
        style_name: optional(element, DB, "style-name", budget)?,
        default_row_style_name: optional(element, DB, "default-row-style-name", budget)?,
        order_statement: child(&kids, "order-statement")
            .map(|v| parse_statement(v, "order-statement", budget))
            .transpose()?,
        filter_statement: child(&kids, "filter-statement")
            .map(|v| parse_statement(v, "filter-statement", budget))
            .transpose()?,
        columns: child(&kids, "columns")
            .map(|v| parse_columns(v, budget))
            .transpose()?,
    })
}

fn parse_columns(element: &DatabaseElement, budget: &mut Budget) -> Result<Vec<OdfDatabaseColumn>> {
    expect(element, "columns")?;
    allow_attrs(element, &[])?;
    let kids = children(element)?;
    if kids.is_empty() {
        return invalid("db:columns must contain at least one db:column");
    }
    let mut values = Vec::with_capacity(kids.len());
    for value in kids {
        budget.column()?;
        values.push(parse_column(value, budget)?);
    }
    Ok(values)
}

fn parse_column(element: &DatabaseElement, budget: &mut Budget) -> Result<OdfDatabaseColumn> {
    expect(element, "column")?;
    allow_attrs(
        element,
        &[
            (DB, "name"),
            (DB, "title"),
            (DB, "description"),
            (DB, "visible"),
            (DB, "style-name"),
            (DB, "default-cell-style-name"),
            (OFFICE, "value-type"),
            (OFFICE, "value"),
            (OFFICE, "currency"),
            (OFFICE, "date-value"),
            (OFFICE, "time-value"),
            (OFFICE, "boolean-value"),
            (OFFICE, "string-value"),
        ],
    )?;
    empty(element)?;
    Ok(OdfDatabaseColumn {
        name: required(element, DB, "name", budget)?,
        title: optional(element, DB, "title", budget)?,
        description: optional(element, DB, "description", budget)?,
        visible: optional_bool(element, DB, "visible")?,
        style_name: optional(element, DB, "style-name", budget)?,
        default_cell_style_name: optional(element, DB, "default-cell-style-name", budget)?,
        value: parse_value(element, budget)?,
    })
}

fn parse_value(
    element: &DatabaseElement,
    budget: &mut Budget,
) -> Result<Option<OdfDatabaseColumnValue>> {
    let Some(kind) = element.attribute(Some(OFFICE), "value-type") else {
        if [
            "value",
            "currency",
            "date-value",
            "time-value",
            "boolean-value",
            "string-value",
        ]
        .iter()
        .any(|v| element.attribute(Some(OFFICE), v).is_some())
        {
            return invalid("database column value has no office:value-type");
        }
        return Ok(None);
    };
    let result = match kind {
        "float" => {
            value_attrs(element, &["value-type", "value"])?;
            let v = required(element, OFFICE, "value", budget)?;
            validate_double(&v)?;
            OdfDatabaseColumnValue::Float(v)
        },
        "percentage" => {
            value_attrs(element, &["value-type", "value"])?;
            let v = required(element, OFFICE, "value", budget)?;
            validate_double(&v)?;
            OdfDatabaseColumnValue::Percentage(v)
        },
        "currency" => {
            value_attrs(element, &["value-type", "value", "currency"])?;
            let value = required(element, OFFICE, "value", budget)?;
            validate_double(&value)?;
            OdfDatabaseColumnValue::Currency {
                value,
                currency: optional(element, OFFICE, "currency", budget)?,
            }
        },
        "date" => {
            value_attrs(element, &["value-type", "date-value"])?;
            let v = required(element, OFFICE, "date-value", budget)?;
            validate_date(&v)?;
            OdfDatabaseColumnValue::Date(v)
        },
        "time" => {
            value_attrs(element, &["value-type", "time-value"])?;
            let v = required(element, OFFICE, "time-value", budget)?;
            validate_duration(&v)?;
            OdfDatabaseColumnValue::Time(v)
        },
        "boolean" => {
            value_attrs(element, &["value-type", "boolean-value"])?;
            OdfDatabaseColumnValue::Boolean(parse_bool(&required(
                element,
                OFFICE,
                "boolean-value",
                budget,
            )?)?)
        },
        "string" => {
            value_attrs(element, &["value-type", "string-value"])?;
            OdfDatabaseColumnValue::String(optional(element, OFFICE, "string-value", budget)?)
        },
        value => return invalid(format!("unsupported database column value type '{value}'")),
    };
    Ok(Some(result))
}

fn write_query_item(
    xml: &mut String,
    item: &OdfDatabaseQueryItem,
    depth: usize,
    budget: &mut Budget,
) -> Result<()> {
    if depth > MAX_COLLECTION_DEPTH {
        return invalid("database query collections are too deeply nested");
    }
    budget.item()?;
    match item {
        OdfDatabaseQueryItem::Query(v) => write_query(xml, v, budget),
        OdfDatabaseQueryItem::Collection(v) => {
            xml.push_str("<db:query-collection");
            required_out(xml, "db:name", &v.name, "collection name", budget)?;
            optional_out(
                xml,
                "db:title",
                v.title.as_deref(),
                "collection title",
                budget,
            )?;
            optional_out(
                xml,
                "db:description",
                v.description.as_deref(),
                "collection description",
                budget,
            )?;
            if v.items.is_empty() {
                xml.push_str("/>");
            } else {
                xml.push('>');
                for child in &v.items {
                    write_query_item(xml, child, depth + 1, budget)?;
                }
                xml.push_str("</db:query-collection>");
            }
            Ok(())
        },
    }
}

fn write_query(xml: &mut String, v: &OdfDatabaseQuery, budget: &mut Budget) -> Result<()> {
    xml.push_str("<db:query");
    required_out(xml, "db:command", &v.command, "query command", budget)?;
    bool_out(xml, "db:escape-processing", v.escape_processing);
    required_out(xml, "db:name", &v.name, "query name", budget)?;
    optional_out(xml, "db:title", v.title.as_deref(), "query title", budget)?;
    optional_out(
        xml,
        "db:description",
        v.description.as_deref(),
        "query description",
        budget,
    )?;
    optional_out(
        xml,
        "db:style-name",
        v.style_name.as_deref(),
        "query style",
        budget,
    )?;
    optional_out(
        xml,
        "db:default-row-style-name",
        v.default_row_style_name.as_deref(),
        "query row style",
        budget,
    )?;
    if v.order_statement.is_none()
        && v.filter_statement.is_none()
        && v.columns.is_none()
        && v.update_table.is_none()
    {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    if let Some(s) = &v.order_statement {
        write_statement(xml, "order-statement", s, budget)?;
    }
    if let Some(s) = &v.filter_statement {
        write_statement(xml, "filter-statement", s, budget)?;
    }
    if let Some(c) = &v.columns {
        write_columns(xml, c, budget)?;
    }
    if let Some(t) = &v.update_table {
        write_update(xml, t, budget)?;
    }
    xml.push_str("</db:query>");
    Ok(())
}

fn write_table(
    xml: &mut String,
    v: &OdfDatabaseTableRepresentation,
    budget: &mut Budget,
) -> Result<()> {
    xml.push_str("<db:table-representation");
    required_out(xml, "db:name", &v.name, "table name", budget)?;
    optional_out(
        xml,
        "db:catalog-name",
        v.catalog_name.as_deref(),
        "table catalog",
        budget,
    )?;
    optional_out(
        xml,
        "db:schema-name",
        v.schema_name.as_deref(),
        "table schema",
        budget,
    )?;
    optional_out(xml, "db:title", v.title.as_deref(), "table title", budget)?;
    optional_out(
        xml,
        "db:description",
        v.description.as_deref(),
        "table description",
        budget,
    )?;
    optional_out(
        xml,
        "db:style-name",
        v.style_name.as_deref(),
        "table style",
        budget,
    )?;
    optional_out(
        xml,
        "db:default-row-style-name",
        v.default_row_style_name.as_deref(),
        "table row style",
        budget,
    )?;
    if v.order_statement.is_none() && v.filter_statement.is_none() && v.columns.is_none() {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    if let Some(s) = &v.order_statement {
        write_statement(xml, "order-statement", s, budget)?;
    }
    if let Some(s) = &v.filter_statement {
        write_statement(xml, "filter-statement", s, budget)?;
    }
    if let Some(c) = &v.columns {
        write_columns(xml, c, budget)?;
    }
    xml.push_str("</db:table-representation>");
    Ok(())
}

fn write_statement(
    xml: &mut String,
    local: &str,
    v: &OdfDatabaseStatement,
    budget: &mut Budget,
) -> Result<()> {
    xml.push_str("<db:");
    xml.push_str(local);
    required_out(xml, "db:command", &v.command, "statement command", budget)?;
    bool_out(xml, "db:apply-command", v.apply_command);
    xml.push_str("/>");
    Ok(())
}

fn write_update(xml: &mut String, v: &OdfDatabaseUpdateTable, budget: &mut Budget) -> Result<()> {
    xml.push_str("<db:update-table");
    required_out(xml, "db:name", &v.name, "update table name", budget)?;
    optional_out(
        xml,
        "db:catalog-name",
        v.catalog_name.as_deref(),
        "update catalog",
        budget,
    )?;
    optional_out(
        xml,
        "db:schema-name",
        v.schema_name.as_deref(),
        "update schema",
        budget,
    )?;
    xml.push_str("/>");
    Ok(())
}

fn write_columns(
    xml: &mut String,
    values: &[OdfDatabaseColumn],
    budget: &mut Budget,
) -> Result<()> {
    if values.is_empty() {
        return invalid("db:columns must contain at least one db:column");
    }
    xml.push_str("<db:columns>");
    for value in values {
        budget.column()?;
        write_column(xml, value, budget)?;
    }
    xml.push_str("</db:columns>");
    Ok(())
}

fn write_column(xml: &mut String, v: &OdfDatabaseColumn, budget: &mut Budget) -> Result<()> {
    xml.push_str("<db:column");
    required_out(xml, "db:name", &v.name, "column name", budget)?;
    optional_out(xml, "db:title", v.title.as_deref(), "column title", budget)?;
    optional_out(
        xml,
        "db:description",
        v.description.as_deref(),
        "column description",
        budget,
    )?;
    bool_out(xml, "db:visible", v.visible);
    optional_out(
        xml,
        "db:style-name",
        v.style_name.as_deref(),
        "column style",
        budget,
    )?;
    optional_out(
        xml,
        "db:default-cell-style-name",
        v.default_cell_style_name.as_deref(),
        "column cell style",
        budget,
    )?;
    if let Some(value) = &v.value {
        write_value(xml, value, budget)?;
    }
    xml.push_str("/>");
    Ok(())
}

fn write_value(xml: &mut String, v: &OdfDatabaseColumnValue, budget: &mut Budget) -> Result<()> {
    match v {
        OdfDatabaseColumnValue::Float(v) => {
            validate_double(v)?;
            literal(xml, "office:value-type", "float");
            required_out(xml, "office:value", v, "column float", budget)?;
        },
        OdfDatabaseColumnValue::Percentage(v) => {
            validate_double(v)?;
            literal(xml, "office:value-type", "percentage");
            required_out(xml, "office:value", v, "column percentage", budget)?;
        },
        OdfDatabaseColumnValue::Currency { value, currency } => {
            validate_double(value)?;
            literal(xml, "office:value-type", "currency");
            required_out(xml, "office:value", value, "column currency value", budget)?;
            optional_out(
                xml,
                "office:currency",
                currency.as_deref(),
                "column currency",
                budget,
            )?;
        },
        OdfDatabaseColumnValue::Date(v) => {
            validate_date(v)?;
            literal(xml, "office:value-type", "date");
            required_out(xml, "office:date-value", v, "column date", budget)?;
        },
        OdfDatabaseColumnValue::Time(v) => {
            validate_duration(v)?;
            literal(xml, "office:value-type", "time");
            required_out(xml, "office:time-value", v, "column time", budget)?;
        },
        OdfDatabaseColumnValue::Boolean(v) => {
            literal(xml, "office:value-type", "boolean");
            literal(
                xml,
                "office:boolean-value",
                if *v { "true" } else { "false" },
            );
        },
        OdfDatabaseColumnValue::String(v) => {
            literal(xml, "office:value-type", "string");
            optional_out(
                xml,
                "office:string-value",
                v.as_deref(),
                "column string",
                budget,
            )?;
        },
    }
    Ok(())
}

fn children(element: &DatabaseElement) -> Result<Vec<&DatabaseElement>> {
    let mut result = Vec::new();
    for value in element.content() {
        match value {
            DatabaseContent::Element(v) => result.push(v),
            DatabaseContent::Text(v)
                if v.trim_matches(|c| matches!(c, ' ' | '\t' | '\r' | '\n'))
                    .is_empty() => {},
            DatabaseContent::Text(_) => {
                return invalid(format!(
                    "text is not allowed in db:{}",
                    element.local_name()
                ));
            },
        }
    }
    Ok(result)
}

fn ordered<'a>(element: &'a DatabaseElement, order: &[&str]) -> Result<Vec<&'a DatabaseElement>> {
    let values = children(element)?;
    let mut prior = 0;
    let mut first = true;
    let mut seen = vec![false; order.len()];
    for value in &values {
        if value.namespace_uri() != Some(DB) {
            return invalid(format!("foreign child in db:{}", element.local_name()));
        }
        let rank = order
            .iter()
            .position(|name| *name == value.local_name())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "unexpected db:{} child in db:{}",
                    value.local_name(),
                    element.local_name()
                ))
            })?;
        if seen[rank] {
            return invalid(format!("duplicate db:{} child", value.local_name()));
        }
        if !first && rank < prior {
            return invalid(format!(
                "children of db:{} are out of schema order",
                element.local_name()
            ));
        }
        seen[rank] = true;
        prior = rank;
        first = false;
    }
    Ok(values)
}

fn child<'a>(values: &[&'a DatabaseElement], local: &str) -> Option<&'a DatabaseElement> {
    values.iter().copied().find(|v| v.local_name() == local)
}
fn expect(element: &DatabaseElement, local: &str) -> Result<()> {
    if element.namespace_uri() == Some(DB) && element.local_name() == local {
        Ok(())
    } else {
        invalid(format!("expected db:{local}"))
    }
}
fn empty(element: &DatabaseElement) -> Result<()> {
    if children(element)?.is_empty() {
        Ok(())
    } else {
        invalid(format!("db:{} must be empty", element.local_name()))
    }
}

fn allow_attrs(element: &DatabaseElement, allowed: &[(&str, &str)]) -> Result<()> {
    for attr in element.attributes() {
        if !allowed
            .iter()
            .any(|(ns, local)| attr.namespace_uri() == Some(*ns) && attr.local_name() == *local)
        {
            return invalid(format!(
                "unexpected attribute {} on db:{}",
                attr.local_name(),
                element.local_name()
            ));
        }
    }
    Ok(())
}

fn value_attrs(element: &DatabaseElement, allowed: &[&str]) -> Result<()> {
    for attr in element
        .attributes()
        .iter()
        .filter(|v| v.namespace_uri() == Some(OFFICE))
    {
        if !allowed.contains(&attr.local_name()) {
            return invalid("incompatible office value attributes on database column");
        }
    }
    Ok(())
}

fn required(
    element: &DatabaseElement,
    ns: &str,
    local: &str,
    budget: &mut Budget,
) -> Result<String> {
    let value = element.attribute(Some(ns), local).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "db:{} has no required {local} attribute",
            element.local_name()
        ))
    })?;
    budget.string(local, value)
}
fn optional(
    element: &DatabaseElement,
    ns: &str,
    local: &str,
    budget: &mut Budget,
) -> Result<Option<String>> {
    element
        .attribute(Some(ns), local)
        .map(|v| budget.string(local, v))
        .transpose()
}
fn optional_bool(element: &DatabaseElement, ns: &str, local: &str) -> Result<Option<bool>> {
    element
        .attribute(Some(ns), local)
        .map(parse_bool)
        .transpose()
}
fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => invalid(format!("invalid strict database boolean '{value}'")),
    }
}

pub(super) fn validate_double(value: &str) -> Result<()> {
    if matches!(value, "INF" | "-INF" | "NaN") {
        return Ok(());
    }
    let bytes = value.as_bytes();
    let mut i = usize::from(matches!(bytes.first(), Some(b'+') | Some(b'-')));
    let start = i;
    while bytes.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
    }
    let mut digits = i - start;
    if bytes.get(i) == Some(&b'.') {
        i += 1;
        let start = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        digits += i - start;
    }
    if digits == 0 {
        return invalid(format!("invalid XML Schema double '{value}'"));
    }
    if matches!(bytes.get(i), Some(b'e') | Some(b'E')) {
        i += 1;
        if matches!(bytes.get(i), Some(b'+') | Some(b'-')) {
            i += 1;
        }
        let start = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        if i == start {
            return invalid(format!("invalid XML Schema double '{value}'"));
        }
    }
    if i == bytes.len() {
        Ok(())
    } else {
        invalid(format!("invalid XML Schema double '{value}'"))
    }
}

pub(super) fn validate_date(value: &str) -> Result<()> {
    fn temporal_error<T>(value: &str) -> Result<T> {
        invalid(format!("invalid XML Schema date or dateTime '{value}'"))
    }
    fn two_digits(value: &[u8]) -> Option<u8> {
        if value.len() == 2 && value.iter().all(u8::is_ascii_digit) {
            Some((value[0] - b'0') * 10 + value[1] - b'0')
        } else {
            None
        }
    }
    fn timezone(value: &[u8]) -> bool {
        if value.is_empty() || value == b"Z" {
            return true;
        }
        if value.len() != 6 || !matches!(value[0], b'+' | b'-') || value[3] != b':' {
            return false;
        }
        let (Some(hour), Some(minute)) = (two_digits(&value[1..3]), two_digits(&value[4..6]))
        else {
            return false;
        };
        hour < 14 && minute < 60 || hour == 14 && minute == 0
    }

    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'-'));
    let year_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    let year = &bytes[year_start..index];
    if year.len() < 4
        || year.iter().all(|byte| *byte == b'0')
        || (year.len() > 4 && year.first() == Some(&b'0'))
        || bytes.get(index) != Some(&b'-')
    {
        return temporal_error(value);
    }
    index += 1;
    let Some(month) = bytes.get(index..index + 2).and_then(two_digits) else {
        return temporal_error(value);
    };
    index += 2;
    if bytes.get(index) != Some(&b'-') {
        return temporal_error(value);
    }
    index += 1;
    let Some(day) = bytes.get(index..index + 2).and_then(two_digits) else {
        return temporal_error(value);
    };
    index += 2;

    let mut year_mod_400 = 0u16;
    for byte in year {
        year_mod_400 = (year_mod_400 * 10 + u16::from(byte - b'0')) % 400;
    }
    let leap =
        year_mod_400.is_multiple_of(4) && (!year_mod_400.is_multiple_of(100) || year_mod_400 == 0);
    let max_day = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return temporal_error(value),
    };
    if day == 0 || day > max_day {
        return temporal_error(value);
    }

    if bytes.get(index) != Some(&b'T') {
        return if timezone(&bytes[index..]) {
            Ok(())
        } else {
            temporal_error(value)
        };
    }
    index += 1;
    let Some(hour) = bytes.get(index..index + 2).and_then(two_digits) else {
        return temporal_error(value);
    };
    index += 2;
    if bytes.get(index) != Some(&b':') {
        return temporal_error(value);
    }
    index += 1;
    let Some(minute) = bytes.get(index..index + 2).and_then(two_digits) else {
        return temporal_error(value);
    };
    index += 2;
    if bytes.get(index) != Some(&b':') {
        return temporal_error(value);
    }
    index += 1;
    let Some(second) = bytes.get(index..index + 2).and_then(two_digits) else {
        return temporal_error(value);
    };
    index += 2;
    let mut nonzero_fraction = false;
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        let fraction_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            nonzero_fraction |= bytes[index] != b'0';
            index += 1;
        }
        if index == fraction_start {
            return temporal_error(value);
        }
    }
    if minute > 59
        || second > 59
        || hour > 24
        || hour == 24 && (minute != 0 || second != 0 || nonzero_fraction)
        || !timezone(&bytes[index..])
    {
        return temporal_error(value);
    }
    Ok(())
}

pub(super) fn validate_duration(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'-'));
    if bytes.get(index) != Some(&b'P') {
        return invalid("invalid XML Schema duration");
    }
    index += 1;
    let mut in_time = false;
    let mut prior_rank = 0u8;
    let mut components = 0usize;
    let mut time_components = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'T' {
            if in_time {
                return invalid("invalid XML Schema duration");
            }
            in_time = true;
            prior_rank = 0;
            index += 1;
            continue;
        }
        let number_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if index == number_start {
            return invalid("invalid XML Schema duration");
        }
        let mut fractional = false;
        if bytes.get(index) == Some(&b'.') {
            fractional = true;
            index += 1;
            let fraction_start = index;
            while bytes.get(index).is_some_and(u8::is_ascii_digit) {
                index += 1;
            }
            if index == fraction_start {
                return invalid("invalid XML Schema duration");
            }
        }
        let Some(designator) = bytes.get(index).copied() else {
            return invalid("invalid XML Schema duration");
        };
        index += 1;
        let rank = match (in_time, designator) {
            (false, b'Y') | (true, b'H') => 1,
            (false, b'M') | (true, b'M') => 2,
            (false, b'D') | (true, b'S') => 3,
            _ => return invalid("invalid XML Schema duration"),
        };
        if rank <= prior_rank || fractional && !(in_time && designator == b'S') {
            return invalid("invalid XML Schema duration");
        }
        prior_rank = rank;
        components += 1;
        if in_time {
            time_components += 1;
        }
    }
    if components == 0 || in_time && time_components == 0 {
        invalid("invalid XML Schema duration")
    } else {
        Ok(())
    }
}

fn required_out(
    xml: &mut String,
    name: &str,
    value: &str,
    label: &str,
    budget: &mut Budget,
) -> Result<()> {
    budget.string(label, value)?;
    literal(xml, name, value);
    Ok(())
}
fn optional_out(
    xml: &mut String,
    name: &str,
    value: Option<&str>,
    label: &str,
    budget: &mut Budget,
) -> Result<()> {
    if let Some(value) = value {
        required_out(xml, name, value, label, budget)?;
    }
    Ok(())
}
fn bool_out(xml: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        literal(xml, name, if value { "true" } else { "false" });
    }
}
fn literal(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    xml.push_str(&escape(value));
    xml.push('"');
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[derive(Clone, Copy)]
enum Target {
    Queries,
    Tables,
}
#[derive(Default)]
struct Spans {
    close: usize,
    queries: Option<(usize, usize)>,
    tables: Option<(usize, usize)>,
    schema: Option<(usize, usize)>,
}

fn mutate(xml: &str, target: Target, replacement: Option<String>) -> Result<String> {
    parse_database_query_model_xml(xml)?;
    let spans = locate(xml)?;
    let current = match target {
        Target::Queries => spans.queries,
        Target::Tables => spans.tables,
    };
    let (start, end) = current.unwrap_or_else(|| {
        let at = match target {
            Target::Queries => spans
                .tables
                .map(|v| v.0)
                .or(spans.schema.map(|v| v.0))
                .unwrap_or(spans.close),
            Target::Tables => spans.schema.map(|v| v.0).unwrap_or(spans.close),
        };
        (at, at)
    });
    let replacement = replacement.unwrap_or_default();
    let mut out = String::with_capacity(xml.len() - (end - start) + replacement.len());
    out.push_str(&xml[..start]);
    out.push_str(&replacement);
    out.push_str(&xml[end..]);
    Ok(out)
}

fn locate(xml: &str) -> Result<Spans> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Option<String>, String, usize)> = Vec::new();
    let mut db_depth = None;
    let mut spans = Spans::default();
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| Error::InvalidFormat(format!("invalid database query XML: {e}")))?;
        let resolved_namespace = namespace_value(&namespace)?;
        drop(namespace);
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if resolved_namespace.as_deref() == Some(OFFICE) && local == "database" {
                    db_depth = Some(stack.len());
                }
                stack.push((resolved_namespace, local, start));
            },
            Event::Empty(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if db_depth.is_some_and(|d| stack.len() == d + 1)
                    && resolved_namespace.as_deref() == Some(DB)
                {
                    set_span(&mut spans, &local, start, end);
                }
            },
            Event::End(ref e) => {
                let local = owned(e.local_name().as_ref())?;
                if let Some((ns, opened, opened_at)) = stack.pop() {
                    if opened != local {
                        return invalid("mismatched database query element");
                    }
                    if db_depth.is_some_and(|d| stack.len() == d + 1) && ns.as_deref() == Some(DB) {
                        set_span(&mut spans, &opened, opened_at, end);
                    }
                    if db_depth == Some(stack.len())
                        && ns.as_deref() == Some(OFFICE)
                        && opened == "database"
                    {
                        spans.close = start;
                        db_depth = None;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if spans.close == 0 {
        invalid("could not locate office:database closing tag")
    } else {
        Ok(spans)
    }
}

fn set_span(spans: &mut Spans, local: &str, start: usize, end: usize) {
    match local {
        "queries" => spans.queries = Some((start, end)),
        "table-representations" => spans.tables = Some((start, end)),
        "schema-definition" => spans.schema = Some((start, end)),
        _ => {},
    }
}

fn preflight(xml: &str) -> Result<()> {
    if xml.len() > MAX_XML {
        return invalid("database query XML is too large");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| Error::InvalidFormat(format!("invalid database query XML: {e}")))?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let ns = namespace_value(&namespace)?;
                let local = owned(e.local_name().as_ref())?;
                if ns.as_deref() == Some(SCRIPT)
                    || (ns.as_deref() == Some(OFFICE)
                        && matches!(local.as_str(), "scripts" | "event-listeners"))
                {
                    return invalid("active content is forbidden in database query XML");
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                    if depth > MAX_XML_DEPTH {
                        return invalid("database query XML is too deeply nested");
                    }
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return invalid(
                    "DTD, entity references, and processing instructions are forbidden in database query XML",
                );
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn namespace_value(value: &ResolveResult<'_>) -> Result<Option<String>> {
    match value {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(v)) => std::str::from_utf8(v)
            .map(|v| Some(v.to_string()))
            .map_err(|_| Error::InvalidFormat("non-UTF-8 namespace".to_string())),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unknown namespace prefix {}",
            String::from_utf8_lossy(prefix)
        )),
    }
}
fn owned(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat("non-UTF-8 element name".to_string()))
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn wrap(queries: &str, tables: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:d="{DB}" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource x:type="simple" x:href="db"/></d:connection-data></d:data-source><!--keep--><d:forms/>{queries}{tables}<d:schema-definition/></o:database></o:body></o:document-content>"#
        )
    }

    #[test]
    fn recursive_queries_and_table_representations_roundtrip() {
        let q = r#"<d:queries><d:query-collection d:name="Reports"><d:query d:command="SELECT &lt;x>" d:name="Recent"><d:order-statement d:command="id DESC"/><d:filter-statement d:command="active" d:apply-command="true"/><d:columns><d:column d:name="amount" o:value-type="currency" o:value="12.5" o:currency="EUR"/><d:column d:name="label" o:value-type="string" o:string-value="A&amp;B"/></d:columns><d:update-table d:name="records"/></d:query></d:query-collection></d:queries>"#;
        let t = r#"<d:table-representations><d:table-representation d:name="records"><d:columns><d:column d:name="created" o:value-type="date" o:date-value="2026-07-19T10:11:12+08:00"/><d:column d:name="delay" o:value-type="time" o:time-value="PT1H2M3.5S"/></d:columns></d:table-representation></d:table-representations>"#;
        let parsed = parse_database_query_model_xml(&wrap(q, t)).unwrap();
        let q2 = parsed.queries.as_ref().unwrap().to_xml_fragment().unwrap();
        let t2 = parsed
            .table_representations
            .as_ref()
            .unwrap()
            .to_xml_fragment()
            .unwrap();
        assert!(q2.find("order-statement").unwrap() < q2.find("filter-statement").unwrap());
        assert_eq!(
            parse_database_query_model_xml(&wrap(&q2, &t2)).unwrap(),
            parsed
        );
    }

    #[test]
    fn rejects_order_cardinality_lexical_active_and_bounds() {
        let bad = [
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns><d:column d:name="c"/></d:columns><d:filter-statement d:command="x"/></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:order-statement d:command="x"/><d:order-statement d:command="y"/></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x" d:escape-processing="1"/></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns/></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns><d:column d:name="c" o:value-type="float" o:value="1x"/></d:columns></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns><d:column d:name="c" o:value-type="date" o:date-value="2026-02-29"/></d:columns></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns><d:column d:name="c" o:value-type="date" o:date-value="2024-04-31T24:00:00.1Z"/></d:columns></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns><d:column d:name="c" o:value-type="time" o:time-value="P1Y2Y"/></d:columns></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns><d:column d:name="c" o:value-type="time" o:time-value="P1D2M"/></d:columns></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><d:columns><d:column d:name="c" o:value-type="time" o:time-value="P1DT"/></d:columns></d:query></d:queries>"#,
            r#"<d:queries><d:query d:name="q" d:command="x"><o:event-listeners/></d:query></d:queries>"#,
        ];
        for value in bad {
            assert!(
                parse_database_query_model_xml(&wrap(value, "")).is_err(),
                "accepted {value}"
            );
        }
        let huge = "x".repeat(MAX_VALUE + 1);
        assert!(
            parse_database_query_model_xml(&wrap(
                &format!(r#"<d:queries><d:query d:name="q" d:command="{huge}"/></d:queries>"#),
                ""
            ))
            .is_err()
        );
        assert!(parse_database_query_model_xml(&format!("<!DOCTYPE x>{}", wrap("", ""))).is_err());
    }

    #[test]
    fn setters_preserve_unrelated_bytes_and_schema_order() {
        let original = wrap(
            "",
            r#"<d:table-representations><d:table-representation d:name="old"/></d:table-representations>"#,
        );
        let queries = OdfDatabaseQueries {
            items: vec![OdfDatabaseQueryItem::Query(OdfDatabaseQuery::new(
                "q",
                "SELECT * FROM t",
            ))],
        };
        let inserted = set_database_queries_xml(&original, Some(&queries)).unwrap();
        assert!(inserted.contains("<!--keep--><d:forms/><db:queries"));
        assert!(
            inserted.find("<db:queries").unwrap()
                < inserted.find("<d:table-representations").unwrap()
        );
        let tables = OdfDatabaseTableRepresentations {
            items: vec![OdfDatabaseTableRepresentation::new("new")],
        };
        let replaced = set_database_table_representations_xml(&inserted, Some(&tables)).unwrap();
        assert!(replaced.contains("<!--keep-->") && !replaced.contains("d:name=\"old\""));
        let removed = set_database_queries_xml(&replaced, None).unwrap();
        assert!(!removed.contains("<db:queries") && removed.contains("<db:table-representations"));
        assert_eq!(
            parse_database_query_model_xml(&removed)
                .unwrap()
                .table_representations
                .unwrap(),
            tables
        );
    }
}
