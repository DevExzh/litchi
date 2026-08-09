#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
#![expect(
    clippy::needless_pass_by_value,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Typed, inert Office Math (OMML) fragments for Word documents.
//!
//! Office Math is a rich XML vocabulary. This module validates the XML
//! envelope and OMML root, then deliberately preserves the fragment instead
//! of reducing it to a lossy equation AST. Callers can round-trip every
//! schema-supported math construct while layout, equation evaluation, and
//! deep schema validation remain the responsibility of Office-compatible
//! renderers.

use crate::error::{Error, Result};
use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::{
    OMML_NAMESPACE_URI, decode_xml_reference, extract_omml_formulas, is_omml_name,
};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as FmtWrite;

/// Largest accepted Office Math fragment.
///
/// This keeps raw-fragment authoring bounded while leaving ample room for
/// equations with deeply nested matrices, delimiters, and annotations.
const MAX_OFFICE_MATH_FRAGMENT_BYTES: usize = 4 * 1024 * 1024;

/// A validated single OMML `<m:oMath>` equation.
///
/// The XML is deliberately retained verbatim (apart from adding a local
/// transitional binding when its OMML root namespace relied on an ancestor
/// declaration). `WordprocessingML` namespaces used inside the math content
/// continue to resolve from the containing DOCX document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OfficeMath {
    xml: String,
}

impl OfficeMath {
    /// Parse one standalone `<m:oMath>` element.
    ///
    /// The fragment must be well-formed XML with exactly one Office Math root.
    /// DTDs, processing instructions, and custom entity references are
    /// rejected because they are not valid DOCX math content and can make raw
    /// XML injection unsafe. When a fragment inherits an OMML namespace from
    /// its original document, that binding is made local using the
    /// transitional OMML namespace; provide an explicit strict binding to
    /// retain strict conformance. Other inherited namespace prefixes must be
    /// locally declared by the fragment or available in its output context.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<String>) -> Result<Self> {
        Ok(Self {
            xml: normalize_fragment(xml.into(), OfficeMathRoot::Equation)?,
        })
    }

    /// Create a plain-text Office Math equation.
    ///
    /// Use [`Self::from_xml`] for fractions, radicals, matrices, and other
    /// richer OMML constructs.
    pub fn text(text: impl AsRef<str>) -> Self {
        let mut xml = String::new();
        write!(
            xml,
            r#"<m:oMath xmlns:m="{OMML_NAMESPACE_URI}"><m:r><m:t>{}</m:t></m:r></m:oMath>"#,
            escape_xml(text.as_ref())
        )
        .expect("writing Office Math XML to a String cannot fail");
        Self { xml }
    }

    /// Return the validated OMML element.
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Consume this equation and return its validated OMML element.
    #[must_use]
    pub fn into_xml(self) -> String {
        self.xml
    }
}

/// A validated display-math `<m:oMathPara>` element.
///
/// A math paragraph owns one or more [`OfficeMath`] equations and can retain
/// arbitrary supported OMML paragraph properties without the library having to
/// interpret their layout semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OfficeMathParagraph {
    xml: String,
}

impl OfficeMathParagraph {
    /// Parse one standalone `<m:oMathPara>` element containing at least one
    /// Office Math equation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<String>) -> Result<Self> {
        let xml = normalize_fragment(xml.into(), OfficeMathRoot::Paragraph)?;
        if extract_omml_formulas(xml.as_bytes())?.is_empty() {
            return Err(invalid(
                "Office Math paragraph must contain an oMath equation",
            ));
        }
        Ok(Self { xml })
    }

    /// Wrap a single equation as a display-math paragraph.
    #[must_use]
    pub fn from_equation(equation: OfficeMath) -> Self {
        let mut xml = String::new();
        write!(
            xml,
            r#"<m:oMathPara xmlns:m="{OMML_NAMESPACE_URI}">{}</m:oMathPara>"#,
            equation.xml()
        )
        .expect("writing Office Math XML to a String cannot fail");
        Self { xml }
    }

    /// Create a display-math paragraph from one or more equations.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_equations(equations: impl IntoIterator<Item = OfficeMath>) -> Result<Self> {
        let mut equations = equations.into_iter();
        let first = equations
            .next()
            .ok_or_else(|| invalid("Office Math paragraph requires at least one equation"))?;
        let mut xml = String::new();
        write!(xml, r#"<m:oMathPara xmlns:m="{OMML_NAMESPACE_URI}">"#)
            .expect("writing Office Math XML to a String cannot fail");
        xml.push_str(first.xml());
        for equation in equations {
            xml.push_str(equation.xml());
        }
        xml.push_str("</m:oMathPara>");
        Self::from_xml(xml)
    }

    /// Return the validated OMML math-paragraph element.
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Return the display equations in document order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn equations(&self) -> Result<Vec<OfficeMath>> {
        extract_omml_formulas(self.xml.as_bytes())?
            .into_iter()
            .map(OfficeMath::from_xml)
            .collect()
    }

    /// Consume this math paragraph and return its validated OMML element.
    #[must_use]
    pub fn into_xml(self) -> String {
        self.xml
    }
}

#[derive(Clone, Copy)]
enum OfficeMathRoot {
    Equation,
    Paragraph,
}

impl OfficeMathRoot {
    const fn local_name(self) -> &'static [u8] {
        match self {
            Self::Equation => b"oMath",
            Self::Paragraph => b"oMathPara",
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Equation => "oMath",
            Self::Paragraph => "oMathPara",
        }
    }
}

struct RootRange {
    start: usize,
    end: usize,
    namespace_binding: Option<InheritedNamespaceBinding>,
}

enum InheritedNamespaceBinding {
    Prefix(String),
    Default,
}

fn normalize_fragment(mut xml: String, expected_root: OfficeMathRoot) -> Result<String> {
    if xml.len() > MAX_OFFICE_MATH_FRAGMENT_BYTES {
        return Err(invalid(format!(
            "Office Math fragment exceeds {MAX_OFFICE_MATH_FRAGMENT_BYTES} bytes"
        )));
    }

    let root = validate_fragment(&xml, expected_root)?;
    if let Some(binding) = root.namespace_binding.as_ref() {
        let insertion = namespace_insertion_offset(&xml, &root)?;
        let declaration = match binding {
            InheritedNamespaceBinding::Prefix(prefix) => {
                format!(r#" xmlns:{prefix}="{OMML_NAMESPACE_URI}""#)
            },
            InheritedNamespaceBinding::Default => {
                format!(r#" xmlns="{OMML_NAMESPACE_URI}""#)
            },
        };
        xml.insert_str(insertion, &declaration);
    }
    Ok(xml)
}

fn validate_fragment(xml: &str, expected_root: OfficeMathRoot) -> Result<RootRange> {
    enum FragmentEvent {
        Start {
            matches_expected_root: bool,
            namespace_binding: Option<InheritedNamespaceBinding>,
        },
        Empty {
            matches_expected_root: bool,
            namespace_binding: Option<InheritedNamespaceBinding>,
        },
        End,
        Text {
            whitespace: bool,
        },
        Comment,
        GeneralReference,
        ForbiddenDeclaration,
        Eof,
    }

    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut root = None;
    let mut depth = 0usize;
    let mut root_complete = false;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| invalid("Office Math XML offset does not fit usize"))?;
        let event = {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    validate_element_attributes(&element, decoder)?;
                    let (matches_expected_root, namespace_binding) =
                        classify_root(&namespace, &element, expected_root.local_name())?;
                    FragmentEvent::Start {
                        matches_expected_root,
                        namespace_binding,
                    }
                },
                Event::Empty(element) => {
                    validate_element_attributes(&element, decoder)?;
                    let (matches_expected_root, namespace_binding) =
                        classify_root(&namespace, &element, expected_root.local_name())?;
                    FragmentEvent::Empty {
                        matches_expected_root,
                        namespace_binding,
                    }
                },
                Event::End(_) => FragmentEvent::End,
                Event::Text(text) => FragmentEvent::Text {
                    whitespace: is_xml_whitespace(text.as_ref()),
                },
                Event::CData(text) => FragmentEvent::Text {
                    whitespace: is_xml_whitespace(text.as_ref()),
                },
                Event::Comment(_) => FragmentEvent::Comment,
                Event::GeneralRef(reference) => {
                    decode_xml_reference(&reference)?;
                    FragmentEvent::GeneralReference
                },
                Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                    FragmentEvent::ForbiddenDeclaration
                },
                Event::Eof => FragmentEvent::Eof,
            }
        };
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| invalid("Office Math XML offset does not fit usize"))?;

        match event {
            FragmentEvent::Start {
                matches_expected_root,
                namespace_binding,
            } => {
                if depth == 0 {
                    if root.is_some() || root_complete {
                        return Err(invalid(
                            "Office Math fragment contains multiple root elements",
                        ));
                    }
                    if !matches_expected_root {
                        return Err(unexpected_root(expected_root));
                    }
                    root = Some(RootRange {
                        start: event_start,
                        end: event_end,
                        namespace_binding,
                    });
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("Office Math XML nesting is too deep"))?;
            },
            FragmentEvent::Empty {
                matches_expected_root,
                namespace_binding,
            } => {
                if depth == 0 {
                    if root.is_some() || root_complete {
                        return Err(invalid(
                            "Office Math fragment contains multiple root elements",
                        ));
                    }
                    if !matches_expected_root {
                        return Err(unexpected_root(expected_root));
                    }
                    root = Some(RootRange {
                        start: event_start,
                        end: event_end,
                        namespace_binding,
                    });
                    root_complete = true;
                }
            },
            FragmentEvent::End => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Office Math XML has an unexpected closing tag"))?;
                if depth == 0 {
                    root_complete = true;
                }
            },
            FragmentEvent::Text { whitespace } => {
                if depth <= 1 && !whitespace {
                    return Err(invalid(
                        "Office Math fragment has text outside or directly inside its root element",
                    ));
                }
            },
            FragmentEvent::Comment => {
                if depth == 0 {
                    return Err(invalid(
                        "Office Math fragment has content outside its root element",
                    ));
                }
            },
            FragmentEvent::GeneralReference => {
                if depth <= 1 {
                    return Err(invalid(
                        "Office Math fragment has content outside or directly inside its root element",
                    ));
                }
            },
            FragmentEvent::ForbiddenDeclaration => {
                return Err(invalid(
                    "Office Math fragments cannot contain declarations, DTDs, or processing instructions",
                ));
            },
            FragmentEvent::Eof => break,
        }
    }

    if depth != 0 {
        return Err(invalid("Office Math fragment has an unclosed element"));
    }
    if !root_complete {
        return Err(invalid(format!(
            "Office Math fragment requires one {} root element",
            expected_root.label()
        )));
    }
    root.ok_or_else(|| {
        invalid(format!(
            "Office Math fragment requires one {} root element",
            expected_root.label()
        ))
    })
}

fn unexpected_root(expected_root: OfficeMathRoot) -> Error {
    invalid(format!(
        "Office Math fragment root must be an OMML {} element",
        expected_root.label()
    ))
}

/// Validate every attribute value without changing the source XML.
///
/// `quick-xml` leaves attribute entity references lazy. Decoding them here
/// rejects custom entities (which cannot be declared because DTDs are
/// prohibited) while preserving the original raw spelling for round-tripping.
fn validate_element_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}

/// Decide whether an element can be the requested OMML root and, when its
/// namespace binding was inherited from an enclosing document, record the
/// binding that must be made local before the fragment is written elsewhere.
///
/// An isolated formula such as `<q:oMath>` or `<oMath>` cannot reveal whether
/// its original namespace was transitional or strict. The mutable DOCX writer
/// emits transitional `WordprocessingML`, so inherited bindings are normalized
/// to the transitional OMML URI. An explicit namespace declaration remains
/// intact and is checked by `is_omml_name` instead.
fn classify_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_local_name: &[u8],
) -> Result<(bool, Option<InheritedNamespaceBinding>)> {
    let name = element.name();
    if name.local_name().as_ref() != expected_local_name {
        return Ok((false, None));
    }

    match namespace {
        ResolveResult::Unknown(prefix) => {
            let Some(name_prefix) = name.prefix() else {
                return Ok((false, None));
            };
            if name_prefix.as_ref() != prefix.as_slice() || prefix.is_empty() {
                return Ok((false, None));
            }
            if root_declares_namespace_binding(element, Some(prefix.as_slice()))? {
                return Err(invalid(
                    "Office Math root namespace cannot be explicitly unbound",
                ));
            }
            let prefix = std::str::from_utf8(prefix)
                .map_err(|_source_error| invalid("Office Math namespace prefix is not UTF-8"))?;
            if !is_safe_namespace_prefix(prefix) {
                return Err(invalid("Office Math namespace prefix is invalid"));
            }
            Ok((
                true,
                Some(InheritedNamespaceBinding::Prefix(prefix.to_owned())),
            ))
        },
        ResolveResult::Unbound => {
            if name.prefix().is_some() {
                return Ok((false, None));
            }
            if root_declares_namespace_binding(element, None)? {
                return Err(invalid(
                    "Office Math root namespace cannot be explicitly unbound",
                ));
            }
            Ok((true, Some(InheritedNamespaceBinding::Default)))
        },
        ResolveResult::Bound(_) => Ok((is_omml_name(namespace, name, expected_local_name), None)),
    }
}

/// Detect an explicit root namespace declaration before synthesizing one.
///
/// An unresolved prefix/default namespace normally means the declaration was
/// inherited from an outer document. An explicit empty declaration instead
/// means that the root is deliberately unbound; adding another `xmlns`
/// attribute would create invalid XML, so that form is rejected.
fn root_declares_namespace_binding(
    element: &BytesStart<'_>,
    prefix: Option<&[u8]>,
) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        match attribute.key.as_namespace_binding() {
            Some(PrefixDeclaration::Default) if prefix.is_none() => return Ok(true),
            Some(PrefixDeclaration::Named(declared_prefix))
                if prefix.is_some_and(|prefix| prefix == declared_prefix) =>
            {
                return Ok(true);
            },
            _ => {},
        }
    }
    Ok(false)
}

/// Check the subset of XML namespace prefixes that can safely be replayed in
/// a synthesized `xmlns:prefix` attribute.
///
/// Prefixes used by OOXML are ASCII. Keeping inherited prefixes to that safe
/// subset avoids turning a permissive XML token parser into an attribute-name
/// injection path when a fragment is made self-contained.
fn is_safe_namespace_prefix(prefix: &str) -> bool {
    let bytes = prefix.as_bytes();
    let Some((&first, rest)) = bytes.split_first() else {
        return false;
    };
    matches!(first, b'A'..=b'Z' | b'a'..=b'z' | b'_')
        && rest.iter().all(
            |byte| matches!(byte, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'-' | b'.'),
        )
}

fn namespace_insertion_offset(xml: &str, root: &RootRange) -> Result<usize> {
    let tag = xml
        .as_bytes()
        .get(root.start..root.end)
        .ok_or_else(|| invalid("Office Math root range is invalid"))?;
    let trailing = if tag.ends_with(b"/>") {
        2
    } else if tag.ends_with(b">") {
        1
    } else {
        return Err(invalid("Office Math root start tag is invalid"));
    };
    let insertion = root
        .end
        .checked_sub(trailing)
        .ok_or_else(|| invalid("Office Math root range is invalid"))?;
    if !xml.is_char_boundary(insertion) {
        return Err(invalid("Office Math root is not aligned to UTF-8"));
    }
    Ok(insertion)
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn equation_fragments_are_validated_and_bind_inherited_omml_prefixes() {
        let equation =
            OfficeMath::from_xml("<m:oMath><m:r><m:t>x &amp; y</m:t></m:r></m:oMath>").unwrap();
        assert!(
            equation
                .xml()
                .starts_with(&format!(r#"<m:oMath xmlns:m="{OMML_NAMESPACE_URI}">"#))
        );
        assert!(equation.xml().contains("x &amp; y"));
        assert!(OfficeMath::from_xml("<m:oMath xmlns:m=\"urn:foreign\"/>").is_err());
        assert!(OfficeMath::from_xml("<oMath xmlns=\"\"/>").is_err());
        assert!(OfficeMath::from_xml("<q:oMath xmlns:q=\"\"/>").is_err());
        assert!(OfficeMath::from_xml("<m:oMath/><m:oMath/>").is_err());
        assert!(OfficeMath::from_xml("<!DOCTYPE math><m:oMath/>").is_err());
        assert!(OfficeMath::from_xml("<m:oMath>&custom;</m:oMath>").is_err());
        assert!(OfficeMath::from_xml("<m:oMath data-value=\"&custom;\"/>").is_err());
        assert!(OfficeMath::from_xml("<m:oMath>not-a-math-child</m:oMath>").is_err());
        let inherited_alias = OfficeMath::from_xml("<q:oMath><q:r/></q:oMath>").unwrap();
        assert!(
            inherited_alias
                .xml()
                .starts_with(&format!(r#"<q:oMath xmlns:q="{OMML_NAMESPACE_URI}">"#))
        );
        let inherited_default = OfficeMath::from_xml("<oMath><r/></oMath>").unwrap();
        assert!(
            inherited_default
                .xml()
                .starts_with(&format!(r#"<oMath xmlns="{OMML_NAMESPACE_URI}">"#))
        );
        assert!(OfficeMath::from_xml(
            "<math:oMath xmlns:math=\"http://purl.oclc.org/ooxml/officeDocument/math\"><math:r/></math:oMath>"
        )
        .is_ok());
    }

    #[test]
    fn text_and_display_equations_preserve_their_typed_boundaries() {
        let first = OfficeMath::text("x < y");
        let second = OfficeMath::from_xml("<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>").unwrap();
        let display = OfficeMathParagraph::from_equations([first.clone(), second.clone()]).unwrap();

        assert_eq!(display.equations().unwrap(), vec![first, second]);
        assert!(!display.xml().contains(r#"\n"#));
        assert!(OfficeMathParagraph::from_xml("<m:oMathPara/>").is_err());

        let inherited_prefix =
            OfficeMathParagraph::from_xml("<q:oMathPara><q:oMath><q:r/></q:oMath></q:oMathPara>")
                .unwrap();
        assert_eq!(
            inherited_prefix.equations().unwrap()[0].xml(),
            format!(r#"<q:oMath xmlns:q="{OMML_NAMESPACE_URI}"><q:r/></q:oMath>"#)
        );
    }

    #[test]
    fn only_safe_inherited_prefixes_are_replayed() {
        assert!(is_safe_namespace_prefix("m"));
        assert!(is_safe_namespace_prefix("math_2"));
        assert!(!is_safe_namespace_prefix(""));
        assert!(!is_safe_namespace_prefix("2math"));
        assert!(!is_safe_namespace_prefix("math:extra"));
        assert!(!is_safe_namespace_prefix("math\"attribute"));
    }
}
