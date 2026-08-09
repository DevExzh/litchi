//! Typed constructors for common `MathML` schemata.
//!
//! These free functions build [`Element`] subtrees with the element
//! structure `MathML` 2 expects (for example, `mfrac` with exactly two
//! children) and expose enumerated attribute values as typed enums. The
//! result is ordinary inert tree data: it can be edited further through the
//! [`Element`] mutation API and installed through the [`crate::Formula`]
//! facade.

use crate::facade::StarMathVersion;
use crate::model::Element;

/// The `MathML` `mathvariant` attribute value family (`MathML` 2 §3.2.2).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Variant {
    Normal,
    Bold,
    Italic,
    BoldItalic,
    DoubleStruck,
    BoldFraktur,
    Script,
    BoldScript,
    Fraktur,
    SansSerif,
    BoldSansSerif,
    SansSerifItalic,
    SansSerifBoldItalic,
    Monospace,
}

impl Variant {
    /// The `MathML` attribute spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::Bold => "bold",
            Self::Italic => "italic",
            Self::BoldItalic => "bold-italic",
            Self::DoubleStruck => "double-struck",
            Self::BoldFraktur => "bold-fraktur",
            Self::Script => "script",
            Self::BoldScript => "bold-script",
            Self::Fraktur => "fraktur",
            Self::SansSerif => "sans-serif",
            Self::BoldSansSerif => "bold-sans-serif",
            Self::SansSerifItalic => "sans-serif-italic",
            Self::SansSerifBoldItalic => "sans-serif-bold-italic",
            Self::Monospace => "monospace",
        }
    }
}

/// The `MathML` `display` attribute value of a `math` root.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Display {
    /// Displayed formula (`display="block"`).
    Block,
    /// Inline formula (`display="inline"`).
    Inline,
}

impl Display {
    /// The `MathML` attribute spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Block => "block",
            Self::Inline => "inline",
        }
    }
}

fn element(local_name: &'static str) -> Element {
    Element::fixed_mathml(local_name)
}

fn token(local_name: &'static str, text: &str) -> Element {
    let mut element = element(local_name);
    element.push_text(text);
    element
}

/// An `mi` identifier token.
#[must_use]
pub fn identifier(text: &str) -> Element {
    token("mi", text)
}

/// An `mi` identifier token with an explicit `mathvariant`.
#[must_use]
pub fn identifier_with_variant(text: &str, variant: Variant) -> Element {
    let mut element = identifier(text);
    element.set_fixed_attribute("mathvariant", variant.as_str());
    element
}

/// An `mn` number token.
#[must_use]
pub fn number(text: &str) -> Element {
    token("mn", text)
}

/// An `mo` operator token.
#[must_use]
pub fn operator(text: &str) -> Element {
    token("mo", text)
}

/// An `mtext` literal text token.
#[must_use]
pub fn literal_text(text: &str) -> Element {
    token("mtext", text)
}

/// An `ms` string literal token with the given quote characters.
#[must_use]
pub fn string_literal(text: &str, left_quote: &str, right_quote: &str) -> Element {
    let mut element = token("ms", text);
    element.set_fixed_attribute("lquote", left_quote);
    element.set_fixed_attribute("rquote", right_quote);
    element
}

/// An `mrow` grouping the given children in order.
#[must_use]
pub fn row(children: Vec<Element>) -> Element {
    let mut element = element("mrow");
    for child in children {
        element.push_child(child);
    }
    element
}

/// An `mfrac` with exactly numerator and denominator children.
#[must_use]
pub fn fraction(numerator: Element, denominator: Element) -> Element {
    row_schemata("mfrac", [numerator, denominator])
}

/// An `msqrt` around the radicand.
#[must_use]
pub fn square_root(radicand: Element) -> Element {
    row_schemata("msqrt", [radicand])
}

/// An `mroot` with the radicand first and the index second.
#[must_use]
pub fn root(radicand: Element, index: Element) -> Element {
    row_schemata("mroot", [radicand, index])
}

/// An `msub` with base and subscript.
#[must_use]
pub fn subscript(base: Element, sub: Element) -> Element {
    row_schemata("msub", [base, sub])
}

/// An `msup` with base and superscript.
#[must_use]
pub fn superscript(base: Element, sup: Element) -> Element {
    row_schemata("msup", [base, sup])
}

/// An `msubsup` with base, subscript, and superscript.
#[must_use]
pub fn sub_superscript(base: Element, sub: Element, sup: Element) -> Element {
    row_schemata("msubsup", [base, sub, sup])
}

/// A `munder` with base and underscript.
#[must_use]
pub fn under(base: Element, underscript: Element) -> Element {
    row_schemata("munder", [base, underscript])
}

/// A `mover` with base and overscript.
#[must_use]
pub fn over(base: Element, overscript: Element) -> Element {
    row_schemata("mover", [base, overscript])
}

/// A `munderover` with base, underscript, and overscript.
#[must_use]
pub fn under_over(base: Element, underscript: Element, overscript: Element) -> Element {
    row_schemata("munderover", [base, underscript, overscript])
}

fn row_schemata<const N: usize>(local_name: &'static str, children: [Element; N]) -> Element {
    let mut element = element(local_name);
    for child in children {
        element.push_child(child);
    }
    element
}

/// An `mfenced` with explicit open/close characters and separators.
#[must_use]
pub fn fenced(children: Vec<Element>, open: &str, close: &str, separators: &str) -> Element {
    let mut element = element("mfenced");
    for child in children {
        element.push_child(child);
    }
    element.set_fixed_attribute("open", open);
    element.set_fixed_attribute("close", close);
    element.set_fixed_attribute("separators", separators);
    element
}

/// An `mtable` built from rows of cells.
#[must_use]
pub fn table(rows: Vec<Vec<Element>>) -> Element {
    let mut table = element("mtable");
    for cells in rows {
        let mut row = element("mtr");
        for cell in cells {
            let mut td = element("mtd");
            td.push_child(cell);
            row.push_child(td);
        }
        table.push_child(row);
    }
    table
}

/// A `semantics` wrapper pairing presentation content with an optional
/// `StarMath` annotation (the `math:annotation` encoding `OpenOffice` writes).
#[must_use]
pub fn semantics(content: Element, starmath_source: Option<&str>) -> Element {
    semantics_with_starmath(
        content,
        starmath_source.map(|source| (StarMathVersion::V5, source)),
    )
}

/// A `semantics` wrapper with an explicitly versioned `LibreOffice` `StarMath`
/// source annotation.
#[must_use]
pub fn semantics_with_starmath(
    content: Element,
    starmath: Option<(StarMathVersion, &str)>,
) -> Element {
    let mut wrapper = element("semantics");
    wrapper.push_child(content);
    if let Some((version, source)) = starmath {
        let mut annotation = element("annotation");
        annotation.set_fixed_attribute("encoding", version.encoding());
        annotation.push_text(source);
        wrapper.push_child(annotation);
    }
    wrapper
}

/// A `math` root element wrapping the body with the given display style.
#[must_use]
pub fn document_root(body: Element, display: Display) -> Element {
    let mut element = element("math");
    element.set_fixed_attribute("display", display.as_str());
    element.push_child(body);
    element
}
