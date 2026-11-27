//! Checked constructors for Content `MathML`.

use litchi_core::Result;

use crate::codec;
use crate::model::{ContentSymbol, Element};

/// The lexical representation of a `cn` number token.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberType {
    Real,
    Integer,
    ENotation,
    Rational,
    ComplexCartesian,
    ComplexPolar,
    Constant,
}

impl NumberType {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Real => "real",
            Self::Integer => "integer",
            Self::ENotation => "e-notation",
            Self::Rational => "rational",
            Self::ComplexCartesian => "complex-cartesian",
            Self::ComplexPolar => "complex-polar",
            Self::Constant => "constant",
        }
    }
}

/// Endpoint inclusion for a Content `MathML` interval.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Closure {
    Open,
    Closed,
    OpenClosed,
    ClosedOpen,
}

impl Closure {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Open => "open",
            Self::Closed => "closed",
            Self::OpenClosed => "open-closed",
            Self::ClosedOpen => "closed-open",
        }
    }
}

/// One single-expression Content `MathML` qualifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Qualifier {
    Condition,
    Degree,
    DomainOfApplication,
    LogBase,
    LowLimit,
    MomentAbout,
    UpLimit,
}

impl Qualifier {
    const fn local_name(self) -> &'static str {
        match self {
            Self::Condition => "condition",
            Self::Degree => "degree",
            Self::DomainOfApplication => "domainofapplication",
            Self::LogBase => "logbase",
            Self::LowLimit => "lowlimit",
            Self::MomentAbout => "momentabout",
            Self::UpLimit => "uplimit",
        }
    }
}

/// Create a checked Content `MathML` application.
///
/// Qualifiers and arguments are supplied in schema order after the operator.
///
/// # Errors
///
/// Returns an error when the operator arity, qualifier order, or argument
/// content model is invalid.
pub fn apply(operator: ContentSymbol, arguments: Vec<Element>) -> Result<Element> {
    let mut children = Vec::with_capacity(arguments.len().saturating_add(1));
    children.push(symbol(operator));
    children.extend(arguments);
    checked("apply", children)
}

/// Create a checked bound-variable declaration with optional degree.
///
/// # Errors
///
/// Returns an error when `identifier` is not a content expression or the
/// resulting bound-variable model is invalid.
pub fn bound_variable(identifier: Element, degree: Option<Element>) -> Result<Element> {
    let mut children = vec![identifier];
    if let Some(value) = degree {
        children.push(qualifier(Qualifier::Degree, value)?);
    }
    checked("bvar", children)
}

/// Create a checked Content `MathML` declaration.
///
/// Both the declared object and its optional value are Content `MathML`
/// expressions. This permits declarations of identifiers, operators, and
/// computed values as specified by `MathML` 2.
///
/// # Errors
///
/// Returns an error when either supplied node is not a Content `MathML`
/// expression.
pub fn declaration(object: Element, value: Option<Element>) -> Result<Element> {
    let mut children = vec![object];
    if let Some(declared_value) = value {
        children.push(declared_value);
    }
    checked("declare", children)
}

/// Create a checked Content `MathML` function wrapper.
///
/// # Errors
///
/// Returns an error when `expression` is not a `MathML` expression.
pub fn function(expression: Element) -> Result<Element> {
    checked("fn", vec![expression])
}

/// Create a Content `MathML` identifier token.
#[must_use]
pub fn identifier(text: &str) -> Element {
    token("ci", text)
}

/// Create a checked closed or open interval.
///
/// # Errors
///
/// Returns an error when either endpoint is not a Content `MathML` expression.
pub fn interval(lower: Element, upper: Element, closure: Closure) -> Result<Element> {
    let mut result = checked("interval", vec![lower, upper])?;
    result.set_fixed_attribute("closure", closure.as_str());
    Ok(result)
}

/// Create a checked generated interval from bound variables and a condition.
///
/// # Errors
///
/// Returns an error unless at least one checked bound variable precedes one
/// checked condition qualifier.
pub fn generated_interval(
    bound_variables: Vec<Element>,
    condition: Element,
    closure: Closure,
) -> Result<Element> {
    let mut children = bound_variables;
    children.push(condition);
    let mut result = checked("interval", children)?;
    result.set_fixed_attribute("closure", closure.as_str());
    Ok(result)
}

/// Create a checked lambda expression.
///
/// `domain` is an optional interval or domain qualifier placed between the
/// bound variables and body.
///
/// # Errors
///
/// Returns an error when ordering or any child content model is invalid.
pub fn lambda(
    bound_variables: Vec<Element>,
    domain: Option<Element>,
    body: Element,
) -> Result<Element> {
    let mut children = bound_variables;
    if let Some(value) = domain {
        children.push(value);
    }
    children.push(body);
    checked("lambda", children)
}

/// Create a checked Content `MathML` list.
///
/// # Errors
///
/// Returns an error when its enumerated or generated content model is invalid.
pub fn list(children: Vec<Element>) -> Result<Element> {
    checked("list", children)
}

/// Create a checked enumerated or generated matrix.
///
/// # Errors
///
/// Returns an error when enumerated rows are empty or ragged, or when a
/// generated matrix has invalid qualifiers or content.
pub fn matrix(children: Vec<Element>) -> Result<Element> {
    checked("matrix", children)
}

/// Create a checked matrix row.
///
/// # Errors
///
/// Returns an error when the row is empty or a cell is not a Content `MathML`
/// expression.
pub fn matrix_row(cells: Vec<Element>) -> Result<Element> {
    checked("matrixrow", cells)
}

/// Create a checked single-part Content `MathML` number.
///
/// E-notation, rational, and complex types require [`number_pair`] instead.
///
/// # Errors
///
/// Returns an error when `number_type` requires two components.
pub fn number(text: &str, number_type: NumberType) -> Result<Element> {
    let mut result = token("cn", text);
    result.set_fixed_attribute("type", number_type.as_str());
    codec::validate_subtree(&result)?;
    Ok(result)
}

/// Create a checked E-notation, rational, or complex two-part number.
///
/// # Errors
///
/// Returns an error when `number_type` is not a two-part representation.
pub fn number_pair(left: &str, right: &str, number_type: NumberType) -> Result<Element> {
    let mut result = element("cn");
    result.set_fixed_attribute("type", number_type.as_str());
    result.push_text(left);
    result.push_child(element("sep"));
    result.push_text(right);
    codec::validate_subtree(&result)?;
    Ok(result)
}

/// Create a checked piecewise value.
///
/// # Errors
///
/// Returns an error when pieces are malformed or `otherwise` is not the final
/// optional child.
pub fn piecewise(pieces: Vec<(Element, Element)>, otherwise: Option<Element>) -> Result<Element> {
    let mut children = Vec::with_capacity(
        pieces
            .len()
            .saturating_add(usize::from(otherwise.is_some())),
    );
    for (value, condition) in pieces {
        children.push(checked("piece", vec![value, condition])?);
    }
    if let Some(value) = otherwise {
        children.push(checked("otherwise", vec![value])?);
    }
    checked("piecewise", children)
}

/// Create a checked one-expression qualifier.
///
/// # Errors
///
/// Returns an error when the expression does not satisfy the qualifier's
/// schema-specific content model.
pub fn qualifier(kind: Qualifier, expression: Element) -> Result<Element> {
    checked(kind.local_name(), vec![expression])
}

/// Create a checked Content `MathML` relation application.
///
/// # Errors
///
/// Returns an error when `relation` is not a relation symbol or its argument
/// count/content model is invalid.
pub fn relation(relation: ContentSymbol, arguments: Vec<Element>) -> Result<Element> {
    let mut children = Vec::with_capacity(arguments.len().saturating_add(1));
    children.push(symbol(relation));
    children.extend(arguments);
    checked("reln", children)
}

/// Create a checked Content `MathML` set.
///
/// # Errors
///
/// Returns an error when its enumerated or generated content model is invalid.
pub fn set(children: Vec<Element>) -> Result<Element> {
    checked("set", children)
}

/// Create a named empty Content `MathML` symbol.
#[must_use]
pub fn symbol(symbol: ContentSymbol) -> Element {
    element(symbol.as_str())
}

/// Create an opaque Content `MathML` `csymbol` token.
#[must_use]
pub fn symbol_token(text: &str) -> Element {
    token("csymbol", text)
}

/// Create a checked Content `MathML` vector.
///
/// # Errors
///
/// Returns an error when its enumerated or generated content model is invalid.
pub fn vector(children: Vec<Element>) -> Result<Element> {
    checked("vector", children)
}

fn checked(local_name: &'static str, children: Vec<Element>) -> Result<Element> {
    let mut result = element(local_name);
    for child in children {
        result.push_child(child);
    }
    codec::validate_subtree(&result)?;
    Ok(result)
}

fn element(local_name: &'static str) -> Element {
    Element::fixed_mathml(local_name)
}

fn token(local_name: &'static str, text: &str) -> Element {
    let mut result = element(local_name);
    result.push_text(text);
    result
}
