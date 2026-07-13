//! Named ranges and named expressions in OpenDocument spreadsheets.
//!
//! ODF permits named definitions at document scope and at sheet scope.  A
//! named range identifies a cell range and can additionally declare one or
//! more special uses.  A named expression associates a name with an
//! OpenFormula expression.

use litchi_core::{Error, Result, xml::escape_xml};

const OPENFORMULA_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
const OPENOFFICE_CALC_NAMESPACE: &str = "http://openoffice.org/2004/calc";

/// Scope in which a named range or expression is visible.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum NamedDefinitionScope {
    /// The definition is visible throughout the spreadsheet.
    Global,
    /// The definition is local to the named sheet.
    Sheet(String),
}

impl NamedDefinitionScope {
    /// Construct a sheet-local scope.
    pub fn sheet(name: impl Into<String>) -> Self {
        Self::Sheet(name.into())
    }

    /// Return the local sheet name, or `None` for a global definition.
    pub fn sheet_name(&self) -> Option<&str> {
        match self {
            Self::Global => None,
            Self::Sheet(name) => Some(name),
        }
    }
}

/// A special use declared by `table:range-usable-as`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NamedRangeUsage {
    /// The range can be used as a print range.
    PrintRange,
    /// The range can be used as filter criteria.
    Filter,
    /// The range can identify rows repeated when printing.
    RepeatRow,
    /// The range can identify columns repeated when printing.
    RepeatColumn,
}

impl NamedRangeUsage {
    /// The ODF token used in `table:range-usable-as`.
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
pub struct NamedRange {
    /// Name by which formulas can refer to the range.
    pub name: String,
    /// ODF cell-range address, for example `$Sheet1.$A$1:.$B$5`.
    pub cell_range_address: String,
    /// Optional base cell used to resolve relative addresses.
    pub base_cell_address: Option<String>,
    /// Special uses permitted for this range.
    pub usable_as: Vec<NamedRangeUsage>,
    /// Visibility of the definition.
    pub scope: NamedDefinitionScope,
}

impl NamedRange {
    /// Create a named range with no special usage or base cell.
    pub fn new(
        name: impl Into<String>,
        cell_range_address: impl Into<String>,
        scope: NamedDefinitionScope,
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
    pub fn with_base_cell(mut self, address: impl Into<String>) -> Result<Self> {
        self.base_cell_address = Some(address.into());
        self.validate()?;
        Ok(self)
    }

    /// Add a special use, ignoring a duplicate use already present.
    pub fn with_usage(mut self, usage: NamedRangeUsage) -> Self {
        if !self.usable_as.contains(&usage) {
            self.usable_as.push(usage);
        }
        self
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_common(&self.name, &self.scope)?;
        validate_nonempty("named range cell address", &self.cell_range_address)?;
        if let Some(base) = &self.base_cell_address {
            validate_nonempty("named range base cell address", base)?;
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

/// A named OpenFormula expression (`table:named-expression`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedExpression {
    /// Name by which formulas can refer to the expression.
    pub name: String,
    /// The OpenFormula expression, including its namespace prefix when present.
    pub expression: String,
    /// Optional base cell used to resolve relative references.
    pub base_cell_address: Option<String>,
    /// Namespace binding used by the expression's formula prefix.
    pub formula_namespace: Option<FormulaNamespace>,
    /// Visibility of the definition.
    pub scope: NamedDefinitionScope,
}

/// Namespace binding for a qualified formula expression.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaNamespace {
    /// Prefix appearing before `:` in the expression.
    pub prefix: String,
    /// Namespace URI bound to the prefix.
    pub uri: String,
}

impl NamedExpression {
    /// Create a named expression with no base cell.
    pub fn new(
        name: impl Into<String>,
        expression: impl Into<String>,
        scope: NamedDefinitionScope,
    ) -> Result<Self> {
        let expression = expression.into();
        let formula_namespace = match formula_prefix(&expression) {
            Some("of") => Some(FormulaNamespace {
                prefix: "of".to_string(),
                uri: OPENFORMULA_NAMESPACE.to_string(),
            }),
            Some("oooc") => Some(FormulaNamespace {
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
    pub fn new_with_namespace(
        name: impl Into<String>,
        expression: impl Into<String>,
        namespace_uri: impl Into<String>,
        scope: NamedDefinitionScope,
    ) -> Result<Self> {
        let expression = expression.into();
        let prefix = formula_prefix(&expression).ok_or_else(|| {
            Error::InvalidFormat(
                "an explicit formula namespace requires a qualified expression".to_string(),
            )
        })?;
        let value = Self {
            name: name.into(),
            formula_namespace: Some(FormulaNamespace {
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
pub enum NamedDefinition {
    /// A named cell range.
    Range(NamedRange),
    /// A named OpenFormula expression.
    Expression(NamedExpression),
}

impl NamedDefinition {
    /// Definition name.
    pub fn name(&self) -> &str {
        match self {
            Self::Range(value) => &value.name,
            Self::Expression(value) => &value.name,
        }
    }

    /// Definition scope.
    pub fn scope(&self) -> &NamedDefinitionScope {
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

impl From<NamedRange> for NamedDefinition {
    fn from(value: NamedRange) -> Self {
        Self::Range(value)
    }
}

impl From<NamedExpression> for NamedDefinition {
    fn from(value: NamedExpression) -> Self {
        Self::Expression(value)
    }
}

pub(crate) fn write_named_definitions<'a>(
    out: &mut String,
    definitions: impl Iterator<Item = &'a NamedDefinition>,
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

pub(crate) fn ensure_unique(
    definitions: &[NamedDefinition],
    candidate: &NamedDefinition,
) -> Result<()> {
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

fn validate_common(name: &str, scope: &NamedDefinitionScope) -> Result<()> {
    validate_nonempty("named definition name", name)?;
    if let NamedDefinitionScope::Sheet(sheet) = scope {
        validate_nonempty("named definition sheet scope", sheet)?;
    }
    Ok(())
}

fn validate_nonempty(label: &str, value: &str) -> Result<()> {
    if value.trim().is_empty() {
        Err(Error::InvalidFormat(format!("{label} must not be empty")))
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
        let range = NamedRange::new(
            "Revenue&Tax",
            "$'Sales & Tax'.$A$1:.$B$2",
            NamedDefinitionScope::Global,
        )
        .unwrap()
        .with_base_cell("$'Sales & Tax'.$A$1")
        .unwrap()
        .with_usage(NamedRangeUsage::PrintRange)
        .with_usage(NamedRangeUsage::Filter)
        .with_usage(NamedRangeUsage::Filter);

        let mut xml = String::new();
        range.write_xml(&mut xml);
        assert!(xml.contains("table:name=\"Revenue&amp;Tax\""));
        assert!(xml.contains("table:range-usable-as=\"print-range filter\""));
        assert_eq!(xml.matches("filter").count(), 1);
    }

    #[test]
    fn rejects_empty_required_values() {
        assert!(NamedRange::new("", "$Sheet1.$A$1", NamedDefinitionScope::Global).is_err());
        assert!(NamedExpression::new("Total", " ", NamedDefinitionScope::Global).is_err());
        assert!(
            NamedExpression::new(
                "Total",
                "of:=SUM([.A1:.A2])",
                NamedDefinitionScope::sheet("")
            )
            .is_err()
        );
    }

    #[test]
    fn explicit_formula_namespace_is_safely_serialized() {
        let expression = NamedExpression::new_with_namespace(
            "Custom",
            "calc:=SUM([.A1:.A2])",
            "urn:example:calc&formula",
            NamedDefinitionScope::Global,
        )
        .unwrap();
        let mut xml = String::new();
        expression.write_xml(&mut xml);
        assert!(xml.contains("xmlns:calc=\"urn:example:calc&amp;formula\""));
        assert!(
            NamedExpression::new_with_namespace(
                "Unsafe",
                "bad prefix:=1",
                "urn:example",
                NamedDefinitionScope::Global,
            )
            .is_err()
        );
    }
}
