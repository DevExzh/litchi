//! Named ranges and named expressions in `OpenDocument` spreadsheets.
//!
//! ODF permits named definitions at document scope and at sheet scope.  A
//! named range identifies a cell range and can additionally declare one or
//! more special uses.  A named expression associates a name with an
//! `OpenFormula` expression.

use super::structure::validate_cell_range_addresses;
use litchi_core::{Error, Result, xml::escape_xml};
use std::collections::{HashMap, HashSet};

const OPENFORMULA_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
const OPENOFFICE_CALC_NAMESPACE: &str = "http://openoffice.org/2004/calc";
const MAX_NAMED_DEFINITIONS: usize = 262_144;
const MAX_NAMED_VALUE_BYTES: usize = 65_536;
const MAX_NAMED_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// Scope in which a named range or expression is visible.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Scope {
    /// The definition is visible throughout the spreadsheet.
    Global,
    /// The definition is local to the named sheet.
    Sheet(String),
}

impl Scope {
    /// Construct a sheet-local scope.
    pub fn sheet(name: impl Into<String>) -> Self {
        Self::Sheet(name.into())
    }

    /// Return the local sheet name, or `None` for a global definition.
    #[must_use]
    pub fn sheet_name(&self) -> Option<&str> {
        match self {
            Self::Global => None,
            Self::Sheet(name) => Some(name),
        }
    }
}

/// A special use declared by `table:range-usable-as`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Usage {
    /// The range can be used as a print range.
    PrintRange,
    /// The range can be used as filter criteria.
    Filter,
    /// The range can identify rows repeated when printing.
    RepeatRow,
    /// The range can identify columns repeated when printing.
    RepeatColumn,
}

impl Usage {
    /// The ODF token used in `table:range-usable-as`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::PrintRange => "print-range",
            Self::Filter => "filter",
            Self::RepeatRow => "repeat-row",
            Self::RepeatColumn => "repeat-column",
        }
    }

    pub(crate) fn parse(token: &str) -> Result<Self> {
        match token {
            "print-range" => Ok(Self::PrintRange),
            "filter" => Ok(Self::Filter),
            "repeat-row" => Ok(Self::RepeatRow),
            "repeat-column" => Ok(Self::RepeatColumn),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:range-usable-as token: {token}"
            ))),
        }
    }
}

/// A named cell range (`table:named-range`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range {
    /// Name by which formulas can refer to the range.
    pub name: String,
    /// ODF cell-range address, for example `$Sheet1.$A$1:.$B$5`.
    pub cell_range_address: String,
    /// Optional base cell used to resolve relative addresses.
    pub base_cell_address: Option<String>,
    /// Special uses permitted for this range.
    pub usable_as: Vec<Usage>,
    /// Visibility of the definition.
    pub scope: Scope,
}

impl Range {
    /// Create a named range with no special usage or base cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new(
        name: impl Into<String>,
        cell_range_address: impl Into<String>,
        scope: Scope,
    ) -> Result<Self> {
        let range = Self {
            name: name.into(),
            cell_range_address: cell_range_address.into(),
            base_cell_address: None,
            usable_as: Vec::new(),
            scope,
        };
        range.validate()?;
        Ok(range)
    }

    /// Set the base cell used to resolve relative references.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn with_base_cell(mut self, address: impl Into<String>) -> Result<Self> {
        self.base_cell_address = Some(address.into());
        self.validate()?;
        Ok(self)
    }

    /// Add a special use, ignoring a duplicate use already present.
    #[must_use]
    pub fn with_usage(mut self, usage: Usage) -> Self {
        if !self.usable_as.contains(&usage) {
            self.usable_as.push(usage);
        }
        self
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_common(&self.name, &self.scope)?;
        validate_nonempty("named range cell address", &self.cell_range_address)?;
        if self.cell_range_address != "#REF!" {
            validate_cell_range_addresses(std::slice::from_ref(&self.cell_range_address))?;
        }
        if let Some(base) = &self.base_cell_address {
            validate_nonempty("named range base cell address", base)?;
            if base != "#REF!" {
                validate_cell_range_addresses(std::slice::from_ref(base))?;
            }
        }
        Ok(())
    }

    pub(crate) fn write_xml(&self, out: &mut String) {
        out.push_str("<table:named-range table:name=\"");
        out.push_str(&escape_xml(&self.name));
        out.push_str("\" table:cell-range-address=\"");
        out.push_str(&escape_xml(&self.cell_range_address));
        out.push('"');
        if let Some(base) = &self.base_cell_address {
            out.push_str(" table:base-cell-address=\"");
            out.push_str(&escape_xml(base));
            out.push('"');
        }
        if !self.usable_as.is_empty() {
            out.push_str(" table:range-usable-as=\"");
            for (index, usage) in self.usable_as.iter().enumerate() {
                if index != 0 {
                    out.push(' ');
                }
                out.push_str(usage.as_str());
            }
            out.push('"');
        }
        out.push_str("/>");
    }
}

/// A named `OpenFormula` expression (`table:named-expression`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Expression {
    /// Name by which formulas can refer to the expression.
    pub name: String,
    /// The `OpenFormula` expression, including its namespace prefix when present.
    pub expression: String,
    /// Optional base cell used to resolve relative references.
    pub base_cell_address: Option<String>,
    /// Namespace binding used by the expression's formula prefix.
    pub formula_namespace: Option<Namespace>,
    /// Visibility of the definition.
    pub scope: Scope,
}

/// Formula-specific vocabulary used by named expressions and conditions.
pub mod formula {
    /// Namespace binding for a qualified formula expression.
    #[derive(Debug, Clone, PartialEq, Eq)]
    pub struct Namespace {
        /// Prefix appearing before `:` in the expression.
        pub prefix: String,
        /// Namespace URI bound to the prefix.
        pub uri: String,
    }
}

use formula::Namespace;

impl Expression {
    /// Create a named expression with no base cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new(
        name: impl Into<String>,
        expression: impl Into<String>,
        scope: Scope,
    ) -> Result<Self> {
        let expression = expression.into();
        let formula_namespace = match formula_prefix(&expression) {
            Some("of") => Some(Namespace {
                prefix: "of".to_string(),
                uri: OPENFORMULA_NAMESPACE.to_string(),
            }),
            Some("oooc") => Some(Namespace {
                prefix: "oooc".to_string(),
                uri: OPENOFFICE_CALC_NAMESPACE.to_string(),
            }),
            _ => None,
        };
        let expression = Self {
            name: name.into(),
            expression,
            base_cell_address: None,
            formula_namespace,
            scope,
        };
        expression.validate()?;
        Ok(expression)
    }

    /// Create an expression with an explicit formula namespace binding.
    ///
    /// Use this for custom formula prefixes. The prefix is taken from the
    /// expression itself, such as `calc` in `calc:=SUM([.A1:.A3])`.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new_with_namespace(
        name: impl Into<String>,
        expression: impl Into<String>,
        namespace_uri: impl Into<String>,
        scope: Scope,
    ) -> Result<Self> {
        let expression = expression.into();
        let prefix = formula_prefix(&expression).ok_or_else(|| {
            Error::InvalidFormat(
                "an explicit formula namespace requires a qualified expression".to_string(),
            )
        })?;
        let value = Self {
            name: name.into(),
            formula_namespace: Some(Namespace {
                prefix: prefix.to_string(),
                uri: namespace_uri.into(),
            }),
            expression,
            base_cell_address: None,
            scope,
        };
        value.validate()?;
        Ok(value)
    }

    /// Set the base cell used to resolve relative references.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn with_base_cell(mut self, address: impl Into<String>) -> Result<Self> {
        self.base_cell_address = Some(address.into());
        self.validate()?;
        Ok(self)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_common(&self.name, &self.scope)?;
        validate_nonempty("named expression", &self.expression)?;
        if let Some(base) = &self.base_cell_address {
            validate_nonempty("named expression base cell address", base)?;
            if base != "#REF!" {
                validate_cell_range_addresses(std::slice::from_ref(base))?;
            }
        }
        let expression_prefix = formula_prefix(&self.expression);
        match (&self.formula_namespace, expression_prefix) {
            (Some(namespace), Some(prefix)) => {
                validate_xml_prefix(&namespace.prefix)?;
                validate_nonempty("formula namespace URI", &namespace.uri)?;
                if namespace.prefix != prefix {
                    return Err(Error::InvalidFormat(format!(
                        "formula namespace prefix '{}' does not match expression prefix '{prefix}'",
                        namespace.prefix
                    )));
                }
            },
            (Some(_), None) => {
                return Err(Error::InvalidFormat(
                    "formula namespace supplied for an unqualified expression".to_string(),
                ));
            },
            (None, Some(prefix)) => {
                return Err(Error::InvalidFormat(format!(
                    "formula prefix '{prefix}' has no namespace binding"
                )));
            },
            (None, None) => {},
        }
        Ok(())
    }

    pub(crate) fn write_xml(&self, out: &mut String) {
        out.push_str("<table:named-expression table:name=\"");
        out.push_str(&escape_xml(&self.name));
        if let Some(namespace) = &self.formula_namespace {
            out.push_str("\" xmlns:");
            out.push_str(&namespace.prefix);
            out.push_str("=\"");
            out.push_str(&escape_xml(&namespace.uri));
        }
        out.push_str("\" table:expression=\"");
        out.push_str(&escape_xml(&self.expression));
        out.push('"');
        if let Some(base) = &self.base_cell_address {
            out.push_str(" table:base-cell-address=\"");
            out.push_str(&escape_xml(base));
            out.push('"');
        }
        out.push_str("/>");
    }
}

/// Either kind of named spreadsheet definition, preserving document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Definition {
    /// A named cell range.
    Range(Range),
    /// A named `OpenFormula` expression.
    Expression(Expression),
}

impl Definition {
    /// Definition name.
    #[must_use]
    pub fn name(&self) -> &str {
        match self {
            Self::Range(value) => &value.name,
            Self::Expression(value) => &value.name,
        }
    }

    /// Definition scope.
    #[must_use]
    pub fn scope(&self) -> &Scope {
        match self {
            Self::Range(value) => &value.scope,
            Self::Expression(value) => &value.scope,
        }
    }

    pub(crate) fn validate(&self) -> Result<()> {
        match self {
            Self::Range(value) => value.validate(),
            Self::Expression(value) => value.validate(),
        }
    }

    pub(crate) fn write_xml(&self, out: &mut String) {
        match self {
            Self::Range(value) => value.write_xml(out),
            Self::Expression(value) => value.write_xml(out),
        }
    }
}

impl From<Range> for Definition {
    fn from(value: Range) -> Self {
        Self::Range(value)
    }
}

impl From<Expression> for Definition {
    fn from(value: Expression) -> Self {
        Self::Expression(value)
    }
}

pub(crate) fn write_definitions<'a>(
    out: &mut String,
    definitions: impl Iterator<Item = &'a Definition>,
) {
    let mut definitions = definitions.peekable();
    if definitions.peek().is_none() {
        return;
    }
    out.push_str("<table:named-expressions>");
    for definition in definitions {
        definition.write_xml(out);
    }
    out.push_str("</table:named-expressions>");
}

/// # Errors
///
/// Returns an error when the value cannot be serialized.
pub fn write_definition_fragment(definition: &Definition) -> Result<String> {
    definition.validate()?;
    let mut output = String::with_capacity(256);
    match definition {
        Definition::Range(value) => {
            output.push_str("<table:named-range xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" table:name=\"");
            output.push_str(&escape_xml(&value.name));
            output.push_str("\" table:cell-range-address=\"");
            output.push_str(&escape_xml(&value.cell_range_address));
            output.push('"');
            if let Some(base) = &value.base_cell_address {
                output.push_str(" table:base-cell-address=\"");
                output.push_str(&escape_xml(base));
                output.push('"');
            }
            if !value.usable_as.is_empty() {
                output.push_str(" table:range-usable-as=\"");
                for (index, usage) in value.usable_as.iter().enumerate() {
                    if index != 0 {
                        output.push(' ');
                    }
                    output.push_str(usage.as_str());
                }
                output.push('"');
            }
            output.push_str("/>");
        },
        Definition::Expression(value) => {
            output.push_str("<table:named-expression xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" table:name=\"");
            output.push_str(&escape_xml(&value.name));
            if let Some(namespace) = &value.formula_namespace {
                output.push_str("\" xmlns:");
                output.push_str(&namespace.prefix);
                output.push_str("=\"");
                output.push_str(&escape_xml(&namespace.uri));
            }
            output.push_str("\" table:expression=\"");
            output.push_str(&escape_xml(&value.expression));
            output.push('"');
            if let Some(base) = &value.base_cell_address {
                output.push_str(" table:base-cell-address=\"");
                output.push_str(&escape_xml(base));
                output.push('"');
            }
            output.push_str("/>");
        },
    }
    Ok(output)
}

pub(crate) fn validate_collection(definitions: &[Definition]) -> Result<()> {
    if definitions.len() > MAX_NAMED_DEFINITIONS {
        return Err(Error::InvalidFormat(format!(
            "named definition count exceeds {MAX_NAMED_DEFINITIONS}"
        )));
    }
    let mut names = HashSet::with_capacity(definitions.len());
    let mut aggregate = 0usize;
    for definition in definitions {
        definition.validate()?;
        if !names.insert((definition.scope().clone(), definition.name().to_string())) {
            return Err(Error::InvalidFormat(format!(
                "duplicate named definition '{}' in {:?}",
                definition.name(),
                definition.scope()
            )));
        }
        let values: Vec<&str> = match definition {
            Definition::Range(value) => vec![
                value.name.as_str(),
                value.cell_range_address.as_str(),
                value.base_cell_address.as_deref().unwrap_or(""),
            ],
            Definition::Expression(value) => vec![
                value.name.as_str(),
                value.expression.as_str(),
                value.base_cell_address.as_deref().unwrap_or(""),
                value
                    .formula_namespace
                    .as_ref()
                    .map_or("", |namespace| namespace.uri.as_str()),
            ],
        };
        for value in values {
            if value.len() > MAX_NAMED_VALUE_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "named definition value exceeds {MAX_NAMED_VALUE_BYTES} bytes"
                )));
            }
            aggregate = aggregate.checked_add(value.len()).ok_or_else(|| {
                Error::InvalidFormat("named definition text size overflow".to_string())
            })?;
        }
    }
    if aggregate > MAX_NAMED_AGGREGATE_BYTES {
        return Err(Error::InvalidFormat(
            "named definition text exceeds 16 MiB".to_string(),
        ));
    }
    validate_named_dependencies(definitions)
}

#[must_use]
pub fn expression_references_name(expression: &str, name: &str) -> bool {
    formula_identifiers(expression)
        .into_iter()
        .any(|identifier| identifier == name)
}

fn validate_named_dependencies(definitions: &[Definition]) -> Result<()> {
    let mut indexes = HashMap::with_capacity(definitions.len());
    for (index, definition) in definitions.iter().enumerate() {
        indexes.insert((definition.scope().clone(), definition.name()), index);
    }
    let mut edges = vec![Vec::new(); definitions.len()];
    for (index, definition) in definitions.iter().enumerate() {
        let Definition::Expression(expression) = definition else {
            continue;
        };
        for identifier in formula_identifiers(&expression.expression) {
            let local = (expression.scope.clone(), identifier);
            let global = (Scope::Global, identifier);
            if let Some(target) = indexes.get(&local).or_else(|| indexes.get(&global))
                && !edges[index].contains(target)
            {
                edges[index].push(*target);
            }
        }
    }
    let mut state = vec![0u8; definitions.len()];
    fn visit(index: usize, edges: &[Vec<usize>], state: &mut [u8]) -> Result<()> {
        if state[index] == 1 {
            return Err(Error::InvalidFormat(
                "named expression dependency cycle".to_string(),
            ));
        }
        if state[index] == 2 {
            return Ok(());
        }
        state[index] = 1;
        for &target in &edges[index] {
            visit(target, edges, state)?;
        }
        state[index] = 2;
        Ok(())
    }
    for index in 0..definitions.len() {
        visit(index, &edges, &mut state)?;
    }
    Ok(())
}

fn formula_identifiers(expression: &str) -> Vec<&str> {
    let bytes = expression.as_bytes();
    let mut identifiers = Vec::new();
    let mut index = 0usize;
    let mut quoted = false;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            if quoted && bytes.get(index + 1) == Some(&b'"') {
                index += 2;
                continue;
            }
            quoted = !quoted;
            index += 1;
            continue;
        }
        if quoted || !(bytes[index].is_ascii_alphabetic() || bytes[index] == b'_') {
            index += 1;
            continue;
        }
        let start = index;
        index += 1;
        while index < bytes.len()
            && (bytes[index].is_ascii_alphanumeric() || matches!(bytes[index], b'_' | b'.'))
        {
            index += 1;
        }
        let mut following = index;
        while following < bytes.len() && bytes[following].is_ascii_whitespace() {
            following += 1;
        }
        let is_function = bytes.get(following) == Some(&b'(');
        let is_namespace =
            bytes.get(following) == Some(&b':') && bytes.get(following + 1) == Some(&b'=');
        if !is_function && !is_namespace {
            identifiers.push(&expression[start..index]);
        }
    }
    identifiers
}

/// # Errors
///
/// Returns an error when a duplicate named definition is found.
pub fn ensure_unique(definitions: &[Definition], candidate: &Definition) -> Result<()> {
    if definitions.iter().any(|existing| {
        existing.name() == candidate.name() && existing.scope() == candidate.scope()
    }) {
        return Err(Error::InvalidFormat(format!(
            "duplicate named definition '{}' in {:?}",
            candidate.name(),
            candidate.scope()
        )));
    }
    Ok(())
}

fn validate_common(name: &str, scope: &Scope) -> Result<()> {
    validate_nonempty("named definition name", name)?;
    if let Scope::Sheet(sheet) = scope {
        validate_nonempty("named definition sheet scope", sheet)?;
    }
    Ok(())
}

fn validate_nonempty(label: &str, value: &str) -> Result<()> {
    if value.trim().is_empty() {
        Err(Error::InvalidFormat(format!("{label} must not be empty")))
    } else if value.len() > MAX_NAMED_VALUE_BYTES {
        Err(Error::InvalidFormat(format!(
            "{label} exceeds {MAX_NAMED_VALUE_BYTES} bytes"
        )))
    } else {
        Ok(())
    }
}

fn formula_prefix(expression: &str) -> Option<&str> {
    let (prefix, remainder) = expression.split_once(':')?;
    if prefix.is_empty() || !remainder.starts_with('=') {
        None
    } else {
        Some(prefix)
    }
}

fn validate_xml_prefix(prefix: &str) -> Result<()> {
    let mut bytes = prefix.bytes();
    let Some(first) = bytes.next() else {
        return Err(Error::InvalidFormat(
            "formula namespace prefix must not be empty".to_string(),
        ));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return Err(Error::InvalidFormat(format!(
            "invalid formula namespace prefix '{prefix}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_all_range_usages_and_escapes_attributes() {
        let range = Range::new("Revenue&Tax", "$'Sales & Tax'.$A$1:.$B$2", Scope::Global)
            .unwrap()
            .with_base_cell("$'Sales & Tax'.$A$1")
            .unwrap()
            .with_usage(Usage::PrintRange)
            .with_usage(Usage::Filter)
            .with_usage(Usage::Filter);

        let mut xml = String::new();
        range.write_xml(&mut xml);
        assert!(xml.contains("table:name=\"Revenue&amp;Tax\""));
        assert!(xml.contains("table:range-usable-as=\"print-range filter\""));
        assert_eq!(xml.matches("filter").count(), 1);
    }

    #[test]
    fn rejects_empty_required_values() {
        assert!(Range::new("", "$Sheet1.$A$1", Scope::Global).is_err());
        assert!(Expression::new("Total", " ", Scope::Global).is_err());
        assert!(Expression::new("Total", "of:=SUM([.A1:.A2])", Scope::sheet("")).is_err());
    }

    #[test]
    fn explicit_formula_namespace_is_safely_serialized() {
        let expression = Expression::new_with_namespace(
            "Custom",
            "calc:=SUM([.A1:.A2])",
            "urn:example:calc&formula",
            Scope::Global,
        )
        .unwrap();
        let mut xml = String::new();
        expression.write_xml(&mut xml);
        assert!(xml.contains("xmlns:calc=\"urn:example:calc&amp;formula\""));
        assert!(
            Expression::new_with_namespace(
                "Unsafe",
                "bad prefix:=1",
                "urn:example",
                Scope::Global,
            )
            .is_err()
        );
    }
}
