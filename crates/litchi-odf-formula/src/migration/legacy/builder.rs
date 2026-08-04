//! Typed constructors for common MathML schemata.
//!
//! These free functions build [`MathElement`] subtrees with the element
//! structure MathML 2 expects (for example, `mfrac` with exactly two
//! children) and expose enumerated attribute values as typed enums. The
//! result is ordinary inert tree data: it can be edited further through the
//! [`MathElement`] mutation API and installed into a formula document with
//! [`super::FormulaDocument::set_math`].

use super::document::MathElement;

/// The MathML `mathvariant` attribute value family (MathML 2 §3.2.2).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathVariant {
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
    Initial,
    Tailed,
    Looped,
    Stretched,
}

impl MathVariant {
    /// The MathML attribute spelling.
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
            Self::Initial => "initial",
            Self::Tailed => "tailed",
            Self::Looped => "looped",
            Self::Stretched => "stretched",
        }
    }
}

/// The MathML `display` attribute value of a `math` root.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathDisplay {
    /// Displayed formula (`display="block"`).
    Block,
    /// Inline formula (`display="inline"`).
    Inline,
}

impl MathDisplay {
    /// The MathML attribute spelling.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Block => "block",
            Self::Inline => "inline",
        }
    }
}

fn element(local_name: &str) -> MathElement {
    MathElement::new(local_name).expect("builder element names are valid")
}

fn token(local_name: &str, text: &str) -> MathElement {
    let mut element = element(local_name);
    element.push_text(text);
    element
}

/// An `mi` identifier token.
pub fn identifier(text: &str) -> MathElement {
    token("mi", text)
}

/// An `mi` identifier token with an explicit `mathvariant`.
pub fn identifier_with_variant(text: &str, variant: MathVariant) -> MathElement {
    let mut element = identifier(text);
    element
        .set_attribute(None, "mathvariant", variant.as_str())
        .expect("fixed attribute name is valid");
    element
}

/// An `mn` number token.
pub fn number(text: &str) -> MathElement {
    token("mn", text)
}

/// An `mo` operator token.
pub fn operator(text: &str) -> MathElement {
    token("mo", text)
}

/// An `mtext` literal text token.
pub fn literal_text(text: &str) -> MathElement {
    token("mtext", text)
}

/// An `ms` string literal token with the given quote characters.
pub fn string_literal(text: &str, left_quote: &str, right_quote: &str) -> MathElement {
    let mut element = token("ms", text);
    element
        .set_attribute(None, "lquote", left_quote)
        .expect("fixed attribute name is valid");
    element
        .set_attribute(None, "rquote", right_quote)
        .expect("fixed attribute name is valid");
    element
}

/// An `mrow` grouping the given children in order.
pub fn row(children: Vec<MathElement>) -> MathElement {
    let mut element = element("mrow");
    for child in children {
        element.push_child(child);
    }
    element
}

/// An `mfrac` with exactly numerator and denominator children.
pub fn fraction(numerator: MathElement, denominator: MathElement) -> MathElement {
    row_schemata("mfrac", [numerator, denominator])
}

/// An `msqrt` around the radicand.
pub fn square_root(radicand: MathElement) -> MathElement {
    row_schemata("msqrt", [radicand])
}

/// An `mroot` with the radicand first and the index second.
pub fn root(radicand: MathElement, index: MathElement) -> MathElement {
    row_schemata("mroot", [radicand, index])
}

/// An `msub` with base and subscript.
pub fn subscript(base: MathElement, sub: MathElement) -> MathElement {
    row_schemata("msub", [base, sub])
}

/// An `msup` with base and superscript.
pub fn superscript(base: MathElement, sup: MathElement) -> MathElement {
    row_schemata("msup", [base, sup])
}

/// An `msubsup` with base, subscript, and superscript.
pub fn sub_superscript(base: MathElement, sub: MathElement, sup: MathElement) -> MathElement {
    row_schemata("msubsup", [base, sub, sup])
}

/// A `munder` with base and underscript.
pub fn under(base: MathElement, underscript: MathElement) -> MathElement {
    row_schemata("munder", [base, underscript])
}

/// A `mover` with base and overscript.
pub fn over(base: MathElement, overscript: MathElement) -> MathElement {
    row_schemata("mover", [base, overscript])
}

/// A `munderover` with base, underscript, and overscript.
pub fn under_over(
    base: MathElement,
    underscript: MathElement,
    overscript: MathElement,
) -> MathElement {
    row_schemata("munderover", [base, underscript, overscript])
}

fn row_schemata<const N: usize>(local_name: &str, children: [MathElement; N]) -> MathElement {
    let mut element = element(local_name);
    for child in children {
        element.push_child(child);
    }
    element
}

/// An `mfenced` with explicit open/close characters and separators.
pub fn fenced(
    children: Vec<MathElement>,
    open: &str,
    close: &str,
    separators: &str,
) -> MathElement {
    let mut element = element("mfenced");
    for child in children {
        element.push_child(child);
    }
    element
        .set_attribute(None, "open", open)
        .expect("fixed attribute name is valid");
    element
        .set_attribute(None, "close", close)
        .expect("fixed attribute name is valid");
    element
        .set_attribute(None, "separators", separators)
        .expect("fixed attribute name is valid");
    element
}

/// An `mtable` built from rows of cells.
pub fn table(rows: Vec<Vec<MathElement>>) -> MathElement {
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
/// StarMath annotation (the `math:annotation` encoding OpenOffice writes).
pub fn semantics(content: MathElement, starmath_source: Option<&str>) -> MathElement {
    let mut wrapper = element("semantics");
    wrapper.push_child(content);
    if let Some(source) = starmath_source {
        let mut annotation = element("annotation");
        annotation
            .set_attribute(None, "encoding", "StarMath 5.0")
            .expect("fixed attribute name is valid");
        annotation.push_text(source);
        wrapper.push_child(annotation);
    }
    wrapper
}

/// A `math` root element wrapping the body with the given display style.
pub fn document_root(body: MathElement, display: MathDisplay) -> MathElement {
    let mut element = element("math");
    element
        .set_attribute(None, "display", display.as_str())
        .expect("fixed attribute name is valid");
    element.push_child(body);
    element
}

#[cfg(test)]
mod tests {
    use super::super::document::parse_mathml;
    use super::*;

    #[test]
    fn builds_schemata_with_typed_attributes() {
        let formula = document_root(
            semantics(
                row(vec![
                    superscript(
                        identifier_with_variant("x", MathVariant::Italic),
                        number("2"),
                    ),
                    operator("+"),
                    fraction(number("1"), root(identifier("x"), number("3"))),
                    under_over(operator("∑"), identifier("i"), number("n")),
                ]),
                Some("x^2 + 1 over root{3}{x}"),
            ),
            MathDisplay::Block,
        );
        let xml = formula.to_xml();
        let reparsed = parse_mathml(&xml).unwrap();
        assert_eq!(reparsed, formula);
        assert_eq!(reparsed.attribute(None, "display"), Some("block"));
        let semantics = reparsed.children().next().unwrap();
        assert_eq!(
            semantics.kind(),
            super::super::document::MathElementKind::Semantics
        );
        let annotation = semantics.children().nth(1).unwrap();
        assert_eq!(annotation.attribute(None, "encoding"), Some("StarMath 5.0"));
        let row = semantics.children().next().unwrap();
        let kinds: Vec<_> = row.children().map(MathElement::kind).collect();
        assert!(kinds.contains(&super::super::document::MathElementKind::Superscript));
        assert!(kinds.contains(&super::super::document::MathElementKind::Fraction));
        assert!(kinds.contains(&super::super::document::MathElementKind::UnderOver));
        assert!(xml.contains("mathvariant=\"italic\""));
    }

    #[test]
    fn builds_fenced_table_and_literal_tokens() {
        let fenced = fenced(vec![identifier("a"), identifier("b")], "[", "]", ",");
        assert_eq!(fenced.local_name(), "mfenced");
        assert_eq!(fenced.attribute(None, "open"), Some("["));
        assert_eq!(fenced.attribute(None, "separators"), Some(","));

        let table = table(vec![
            vec![number("1"), number("2")],
            vec![number("3"), number("4")],
        ]);
        assert_eq!(table.children().count(), 2);
        let first_cell = table.children().next().unwrap().children().next().unwrap();
        assert_eq!(
            first_cell.kind(),
            super::super::document::MathElementKind::TableCell
        );

        let literal = string_literal("hello", "«", "»");
        assert_eq!(literal.attribute(None, "lquote"), Some("«"));
        assert_eq!(literal.all_text(), "hello");
        assert_eq!(
            literal.namespace_uri(),
            Some(super::super::document::MATHML_NAMESPACE)
        );
    }
}
