//! Stored query semantics.

use super::table::Column;
use litchi_core::{Error, Result};

const MAX_COMMAND_BYTES: usize = 1024 * 1024;
const MAX_COMMAND_TOKENS: usize = 65_536;

/// The placeholder syntax of one inert stored-query parameter occurrence.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum QueryParameterKind {
    /// A colon-prefixed identifier.
    Named,
    /// A question-mark placeholder.
    Positional,
}

/// One source-ordered parameter occurrence in a stored query command.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryParameter {
    kind: QueryParameterKind,
    name: Option<String>,
    ordinal: usize,
}

impl QueryParameter {
    /// Returns whether the occurrence is named (`:name`) or positional (`?`).
    #[must_use]
    pub const fn kind(&self) -> QueryParameterKind {
        self.kind
    }

    /// Returns the named marker without its colon, if applicable.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the one-based occurrence ordinal.
    #[must_use]
    pub const fn ordinal(&self) -> usize {
        self.ordinal
    }
}

/// A common explicit SQL join form found without executing the command.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum QueryJoinKind {
    /// `JOIN` or `INNER JOIN`.
    Inner,
    /// `LEFT [OUTER] JOIN`.
    Left,
    /// `RIGHT [OUTER] JOIN`.
    Right,
    /// `FULL [OUTER] JOIN`.
    Full,
    /// `CROSS JOIN`.
    Cross,
    /// `NATURAL JOIN`.
    Natural,
}

/// One inert joined-relation reference.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryJoin {
    kind: QueryJoinKind,
    relation: String,
}

impl QueryJoin {
    /// Returns the explicit join form.
    #[must_use]
    pub const fn kind(&self) -> QueryJoinKind {
        self.kind
    }

    /// Returns the source relation token, including a schema qualifier.
    #[must_use]
    pub fn relation(&self) -> &str {
        &self.relation
    }
}

/// Bounded lexical inventory of common query parameters and explicit joins.
///
/// This is deliberately not a SQL execution plan or dialect validator.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct QueryCommandInventory {
    parameters: Vec<QueryParameter>,
    joins: Vec<QueryJoin>,
}

impl QueryCommandInventory {
    /// Returns parameter occurrences in command order.
    #[must_use]
    pub fn parameters(&self) -> &[QueryParameter] {
        &self.parameters
    }

    /// Returns common explicit joins in command order.
    #[must_use]
    pub fn joins(&self) -> &[QueryJoin] {
        &self.joins
    }
}

/// An inert table target updated by a stored query.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryUpdateTarget {
    name: String,
    schema: Option<String>,
    catalog: Option<String>,
}

impl QueryUpdateTarget {
    /// Creates a named update target.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            schema: None,
            catalog: None,
        }
    }

    /// Sets the optional schema qualifier.
    #[must_use]
    pub fn with_schema_name(mut self, value: impl Into<String>) -> Self {
        self.schema = Some(value.into());
        self
    }

    /// Sets the optional catalog qualifier.
    #[must_use]
    pub fn with_catalog_name(mut self, value: impl Into<String>) -> Self {
        self.catalog = Some(value.into());
        self
    }

    pub(crate) fn parsed(
        name: String,
        schema_name: Option<String>,
        catalog_name: Option<String>,
    ) -> Self {
        Self {
            name,
            schema: schema_name,
            catalog: catalog_name,
        }
    }

    /// Returns the target table name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the schema qualifier, if declared.
    #[must_use]
    pub fn schema_name(&self) -> Option<&str> {
        self.schema.as_deref()
    }

    /// Returns the catalog qualifier, if declared.
    #[must_use]
    pub fn catalog_name(&self) -> Option<&str> {
        self.catalog.as_deref()
    }
}

/// A stored database query.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Query {
    name: String,
    command: String,
    escape_processing: Option<bool>,
    columns: Vec<Column>,
    filter_statement: Option<String>,
    order_statement: Option<String>,
    update_target: Option<QueryUpdateTarget>,
}

impl Query {
    /// Creates an inert stored-query declaration.
    #[must_use]
    pub fn new(name: impl Into<String>, command: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            command: command.into(),
            escape_processing: None,
            columns: Vec::new(),
            filter_statement: None,
            order_statement: None,
            update_target: None,
        }
    }

    /// Sets the optional ODF escape-processing declaration.
    #[must_use]
    pub const fn with_escape_processing(mut self, value: Option<bool>) -> Self {
        self.escape_processing = value;
        self
    }

    /// Appends one inert query result-column presentation declaration.
    #[must_use]
    pub fn with_column(mut self, value: Column) -> Self {
        self.columns.push(value);
        self
    }

    /// Sets the inert filter command metadata.
    #[must_use]
    pub fn with_filter_statement(mut self, value: impl Into<String>) -> Self {
        self.filter_statement = Some(value.into());
        self
    }

    /// Sets the inert ordering command metadata.
    #[must_use]
    pub fn with_order_statement(mut self, value: impl Into<String>) -> Self {
        self.order_statement = Some(value.into());
        self
    }

    /// Sets the inert update-table target.
    #[must_use]
    pub fn with_update_target(mut self, value: QueryUpdateTarget) -> Self {
        self.update_target = Some(value);
        self
    }

    pub(crate) fn parsed(name: String, command: String, escape_processing: Option<bool>) -> Self {
        Self {
            name,
            command,
            escape_processing,
            columns: Vec::new(),
            filter_statement: None,
            order_statement: None,
            update_target: None,
        }
    }

    pub(crate) fn try_push_column(&mut self, value: Column) -> Result<()> {
        self.columns
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB query columns",
                source,
            })?;
        self.columns.push(value);
        Ok(())
    }

    pub(crate) fn set_filter_statement(&mut self, value: String) -> Result<()> {
        set_once(&mut self.filter_statement, value, "filter statement")
    }

    pub(crate) fn set_order_statement(&mut self, value: String) -> Result<()> {
        set_once(&mut self.order_statement, value, "order statement")
    }

    pub(crate) fn set_update_target(&mut self, value: QueryUpdateTarget) -> Result<()> {
        if self.update_target.replace(value).is_some() {
            return Err(Error::InvalidFormat(
                "ODB query contains duplicate update targets".to_string(),
            ));
        }
        Ok(())
    }

    /// Returns the query name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the query command text.
    #[must_use]
    pub fn command(&self) -> &str {
        &self.command
    }

    /// Returns the ODF escape-processing declaration, if the producer stored one.
    ///
    /// This metadata is descriptive only. Litchi never parses, connects to, or
    /// executes the command.
    #[must_use]
    pub const fn escape_processing(&self) -> Option<bool> {
        self.escape_processing
    }

    /// Returns inert query result-column declarations in source order.
    #[must_use]
    pub fn columns(&self) -> &[Column] {
        &self.columns
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

    /// Returns the inert update-table target, if declared.
    #[must_use]
    pub const fn update_target(&self) -> Option<&QueryUpdateTarget> {
        self.update_target.as_ref()
    }

    /// Inventories common named/positional parameters and explicit joins
    /// without connecting to a database or executing the command.
    ///
    /// Quoted strings and SQL comments are skipped, so marker-like text inside
    /// them is not reported.
    ///
    /// # Errors
    ///
    /// Returns an error for an oversized command or an unterminated quoted
    /// string, quoted identifier, or block comment.
    pub fn command_inventory(&self) -> Result<QueryCommandInventory> {
        command_inventory(&self.command)
    }
}

fn set_once(target: &mut Option<String>, value: String, kind: &str) -> Result<()> {
    if target.replace(value).is_some() {
        return Err(Error::InvalidFormat(format!(
            "ODB query contains duplicate {kind}s"
        )));
    }
    Ok(())
}

enum SqlToken {
    Word(String),
    QuotedIdentifier(String),
    Dot,
}

fn command_inventory(command: &str) -> Result<QueryCommandInventory> {
    if command.len() > MAX_COMMAND_BYTES {
        return Err(Error::InvalidFormat(
            "ODB query command exceeds the inventory byte limit".to_string(),
        ));
    }
    let bytes = command.as_bytes();
    let mut cursor = 0usize;
    let mut tokens = Vec::new();
    let mut parameters = Vec::new();
    while cursor < bytes.len() {
        if bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        } else if bytes[cursor..].starts_with(b"--") {
            cursor = bytes[cursor..]
                .iter()
                .position(|byte| *byte == b'\n')
                .map_or(bytes.len(), |offset| cursor + offset + 1);
        } else if bytes[cursor..].starts_with(b"/*") {
            cursor = bytes[cursor + 2..]
                .windows(2)
                .position(|window| window == b"*/")
                .map(|offset| cursor + offset + 4)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODB query has an unterminated block comment".to_string())
                })?;
        } else if bytes[cursor] == b'\'' {
            cursor = quoted_end(bytes, cursor, b'\'', "string")?;
        } else if matches!(bytes[cursor], b'"' | b'`') {
            let quote = bytes[cursor];
            let end = quoted_end(bytes, cursor, quote, "identifier")?;
            push_sql_token(
                &mut tokens,
                SqlToken::QuotedIdentifier(command[cursor + 1..end - 1].to_owned()),
            )?;
            cursor = end;
        } else if bytes[cursor] == b'[' {
            let end = bracketed_identifier_end(bytes, cursor)?;
            push_sql_token(
                &mut tokens,
                SqlToken::QuotedIdentifier(command[cursor + 1..end - 1].to_owned()),
            )?;
            cursor = end;
        } else if bytes[cursor] == b':'
            && (cursor == 0 || bytes[cursor - 1] != b':')
            && bytes
                .get(cursor + 1)
                .is_some_and(|byte| is_identifier_start(*byte))
        {
            let start = cursor + 1;
            cursor = identifier_end(bytes, start);
            let ordinal = parameters.len() + 1;
            push_parameter(
                &mut parameters,
                QueryParameter {
                    kind: QueryParameterKind::Named,
                    name: Some(command[start..cursor].to_owned()),
                    ordinal,
                },
            )?;
        } else if bytes[cursor] == b'?' {
            let ordinal = parameters.len() + 1;
            push_parameter(
                &mut parameters,
                QueryParameter {
                    kind: QueryParameterKind::Positional,
                    name: None,
                    ordinal,
                },
            )?;
            cursor += 1;
        } else if is_identifier_start(bytes[cursor]) {
            let start = cursor;
            cursor = identifier_end(bytes, cursor);
            push_sql_token(
                &mut tokens,
                SqlToken::Word(command[start..cursor].to_owned()),
            )?;
        } else {
            if bytes[cursor] == b'.' {
                push_sql_token(&mut tokens, SqlToken::Dot)?;
            }
            cursor += 1;
        }
    }
    let joins = joined_relations(&tokens)?;
    Ok(QueryCommandInventory { parameters, joins })
}

fn quoted_end(bytes: &[u8], start: usize, quote: u8, kind: &str) -> Result<usize> {
    let mut cursor = start + 1;
    while cursor < bytes.len() {
        if bytes[cursor] == quote {
            if bytes.get(cursor + 1) == Some(&quote) {
                cursor += 2;
            } else {
                return Ok(cursor + 1);
            }
        } else {
            cursor += 1;
        }
    }
    Err(Error::InvalidFormat(format!(
        "ODB query has an unterminated quoted {kind}"
    )))
}

fn bracketed_identifier_end(bytes: &[u8], start: usize) -> Result<usize> {
    let mut cursor = start + 1;
    while cursor < bytes.len() {
        if bytes[cursor] != b']' {
            cursor += 1;
        } else if bytes.get(cursor + 1) == Some(&b']') {
            cursor += 2;
        } else {
            return Ok(cursor + 1);
        }
    }
    Err(Error::InvalidFormat(
        "ODB query has an unterminated bracketed identifier".to_string(),
    ))
}

fn identifier_end(bytes: &[u8], start: usize) -> usize {
    let mut cursor = start + 1;
    while bytes
        .get(cursor)
        .is_some_and(|byte| is_identifier_continue(*byte))
    {
        cursor += 1;
    }
    cursor
}

fn push_sql_token(tokens: &mut Vec<SqlToken>, token: SqlToken) -> Result<()> {
    if tokens.len() >= MAX_COMMAND_TOKENS {
        return Err(Error::InvalidFormat(
            "ODB query command exceeds the token limit".to_string(),
        ));
    }
    tokens.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODB query command tokens",
        source,
    })?;
    tokens.push(token);
    Ok(())
}

fn push_parameter(parameters: &mut Vec<QueryParameter>, parameter: QueryParameter) -> Result<()> {
    if parameters.len() >= MAX_COMMAND_TOKENS {
        return Err(Error::InvalidFormat(
            "ODB query command exceeds the parameter limit".to_string(),
        ));
    }
    parameters
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "ODB query parameter inventory",
            source,
        })?;
    parameters.push(parameter);
    Ok(())
}

fn joined_relations(tokens: &[SqlToken]) -> Result<Vec<QueryJoin>> {
    let mut joins = Vec::new();
    for (index, token) in tokens.iter().enumerate() {
        if !identifier_is(token, "join") {
            continue;
        }
        let kind = join_kind(tokens, index);
        let mut cursor = index + 1;
        if tokens
            .get(cursor)
            .is_some_and(|token| identifier_is(token, "lateral"))
        {
            cursor += 1;
        }
        let Some(first) = relation_identifier(tokens.get(cursor)) else {
            continue;
        };
        let mut relation = first.to_owned();
        while matches!(tokens.get(cursor + 1), Some(SqlToken::Dot)) {
            let Some(next) = relation_identifier(tokens.get(cursor + 2)) else {
                break;
            };
            relation.push('.');
            relation.push_str(next);
            cursor += 2;
        }
        joins.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODB query join inventory",
            source,
        })?;
        joins.push(QueryJoin { kind, relation });
    }
    Ok(joins)
}

fn join_kind(tokens: &[SqlToken], join: usize) -> QueryJoinKind {
    let previous = join.checked_sub(1).and_then(|index| tokens.get(index));
    let modifier = if previous.is_some_and(|token| identifier_is(token, "outer")) {
        join.checked_sub(2).and_then(|index| tokens.get(index))
    } else {
        previous
    };
    if modifier.is_some_and(|token| identifier_is(token, "left")) {
        QueryJoinKind::Left
    } else if modifier.is_some_and(|token| identifier_is(token, "right")) {
        QueryJoinKind::Right
    } else if modifier.is_some_and(|token| identifier_is(token, "full")) {
        QueryJoinKind::Full
    } else if modifier.is_some_and(|token| identifier_is(token, "cross")) {
        QueryJoinKind::Cross
    } else if modifier.is_some_and(|token| identifier_is(token, "natural")) {
        QueryJoinKind::Natural
    } else {
        QueryJoinKind::Inner
    }
}

fn identifier_is(token: &SqlToken, expected: &str) -> bool {
    matches!(token, SqlToken::Word(value) if value.eq_ignore_ascii_case(expected))
}

fn relation_identifier(token: Option<&SqlToken>) -> Option<&str> {
    match token? {
        SqlToken::Word(value) | SqlToken::QuotedIdentifier(value) => Some(value),
        SqlToken::Dot => None,
    }
}

fn is_identifier_start(byte: u8) -> bool {
    byte.is_ascii_alphabetic() || byte == b'_'
}

fn is_identifier_continue(byte: u8) -> bool {
    is_identifier_start(byte) || byte.is_ascii_digit() || byte == b'$'
}
