//! The LaTeX of an OpenDocument formula object, shared by the ODT and ODP
//! adapters: its MathML converted through `litchi-equation`.

/// The LaTeX of the first MathML `math` element in `xml`: an object's inline
/// MathML, or the `content.xml` of its formula subdocument. `None` without
/// the `equation` feature, when `xml` holds no `math` element, and when the
/// MathML does not convert.
pub(crate) fn latex(xml: &str) -> Option<String> {
    #[cfg(feature = "equation")]
    {
        litchi_equation::mathml_to_latex(math_element(xml)?, &litchi_equation::Limits::DEFAULT).ok()
    }
    #[cfg(not(feature = "equation"))]
    {
        let _ = xml;
        None
    }
}

/// The source text of the first element whose local name is `math`.
#[cfg(feature = "equation")]
fn math_element(xml: &str) -> Option<&str> {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut open: Option<(usize, usize)> = None;
    let mut depth = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position()).ok()?;
        let event = reader.read_event().ok()?;
        let end = usize::try_from(reader.buffer_position()).ok()?;
        match event {
            Event::Start(tag) => {
                depth += 1;
                if open.is_none() && tag.local_name().as_ref() == b"math" {
                    open = Some((start, depth));
                }
            },
            Event::Empty(tag) if open.is_none() && tag.local_name().as_ref() == b"math" => {
                return xml.get(start..end);
            },
            Event::End(_) => {
                if let Some((from, level)) = open
                    && level == depth
                {
                    return xml.get(from..end);
                }
                depth = depth.checked_sub(1)?;
            },
            Event::Eof => return None,
            _ => {},
        }
    }
}

#[cfg(all(test, feature = "equation"))]
mod tests {
    use super::latex;

    #[test]
    fn formula_subdocuments_and_inline_math_convert() {
        let math = r#"<math:math xmlns:math="http://www.w3.org/1998/Math/MathML"><math:semantics><math:msup><math:mi>x</math:mi><math:mn>2</math:mn></math:msup><math:annotation math:encoding="StarMath 5.0">x^2</math:annotation></math:semantics></math:math>"#;
        assert_eq!(latex(math).as_deref(), Some("x^{2}"));
        let content = format!(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body>{math}</office:body></office:document-content>"#
        );
        assert_eq!(latex(&content).as_deref(), Some("x^{2}"));
        assert_eq!(latex("<office:document-content/>"), None);
    }
}
