//! Inert, namespace-aware MathML access for OpenDocument Formula packages.

use crate::{OpenDocumentFamily, OpenDocumentPackage, PackageWriter, constants};
use litchi_core::{Error, Metadata, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

pub(crate) const MATHML_NAMESPACE: &str = "http://www.w3.org/1998/Math/MathML";
const MAX_MATH_DEPTH: usize = 128;
const MAX_MATH_NODES: usize = 65_536;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 32 * 1_048_576;

/// A commonly used MathML element kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum MathElementKind {
    Math,
    Semantics,
    Annotation,
    AnnotationXml,
    Row,
    Identifier,
    Number,
    Operator,
    Text,
    Space,
    StringLiteral,
    Glyph,
    Fraction,
    SquareRoot,
    Root,
    Style,
    Error,
    Padded,
    Phantom,
    Fenced,
    Enclose,
    Subscript,
    Superscript,
    SubSuperscript,
    Under,
    Over,
    UnderOver,
    MultiScripts,
    Table,
    TableRow,
    TableCell,
    AlignGroup,
    AlignMark,
    /// A future MathML element or a vendor element in another namespace.
    Other,
}

/// One decoded attribute with its expanded namespace name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathAttribute {
    namespace_uri: Option<String>,
    local_name: String,
    value: String,
}

impl MathAttribute {
    pub(crate) fn from_parts(
        namespace_uri: Option<String>,
        local_name: String,
        value: String,
    ) -> Self {
        Self {
            namespace_uri,
            local_name,
            value,
        }
    }

    /// Return the expanded namespace URI, or `None` for an unqualified attribute.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the XML local name.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Return the decoded and normalized XML attribute value.
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Ordered mixed content within a MathML element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MathContent {
    /// Decoded character content, including CDATA and character references.
    ///
    /// Named references other than XML's five predefined entities are retained
    /// in `&name;` notation because MathML 2 documents may declare them in a
    /// document type definition that is intentionally not evaluated here.
    Text(String),
    /// A child element.
    Element(MathElement),
}

/// A complete element in the formula's MathML subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathElement {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<MathAttribute>,
    content: Vec<MathContent>,
}

impl MathElement {
    pub(crate) fn from_parts(
        namespace_uri: Option<String>,
        local_name: String,
        attributes: Vec<MathAttribute>,
        content: Vec<MathContent>,
    ) -> Self {
        Self {
            namespace_uri,
            local_name,
            attributes,
            content,
        }
    }

    pub(crate) fn attributes_mut(&mut self) -> &mut Vec<MathAttribute> {
        &mut self.attributes
    }

    pub(crate) fn content_mut(&mut self) -> &mut Vec<MathContent> {
        &mut self.content
    }

    /// Return the element's expanded namespace URI.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the element's XML local name.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Classify common MathML elements without discarding unknown ones.
    pub fn kind(&self) -> MathElementKind {
        if self.namespace_uri() != Some(MATHML_NAMESPACE) {
            return MathElementKind::Other;
        }
        match self.local_name.as_str() {
            "math" => MathElementKind::Math,
            "semantics" => MathElementKind::Semantics,
            "annotation" => MathElementKind::Annotation,
            "annotation-xml" => MathElementKind::AnnotationXml,
            "mrow" => MathElementKind::Row,
            "mi" => MathElementKind::Identifier,
            "mn" => MathElementKind::Number,
            "mo" => MathElementKind::Operator,
            "mtext" => MathElementKind::Text,
            "mspace" => MathElementKind::Space,
            "ms" => MathElementKind::StringLiteral,
            "mglyph" => MathElementKind::Glyph,
            "mfrac" => MathElementKind::Fraction,
            "msqrt" => MathElementKind::SquareRoot,
            "mroot" => MathElementKind::Root,
            "mstyle" => MathElementKind::Style,
            "merror" => MathElementKind::Error,
            "mpadded" => MathElementKind::Padded,
            "mphantom" => MathElementKind::Phantom,
            "mfenced" => MathElementKind::Fenced,
            "menclose" => MathElementKind::Enclose,
            "msub" => MathElementKind::Subscript,
            "msup" => MathElementKind::Superscript,
            "msubsup" => MathElementKind::SubSuperscript,
            "munder" => MathElementKind::Under,
            "mover" => MathElementKind::Over,
            "munderover" => MathElementKind::UnderOver,
            "mmultiscripts" => MathElementKind::MultiScripts,
            "mtable" => MathElementKind::Table,
            "mtr" | "mlabeledtr" => MathElementKind::TableRow,
            "mtd" => MathElementKind::TableCell,
            "maligngroup" => MathElementKind::AlignGroup,
            "malignmark" => MathElementKind::AlignMark,
            _ => MathElementKind::Other,
        }
    }

    /// Return all decoded attributes in document order.
    pub fn attributes(&self) -> &[MathAttribute] {
        &self.attributes
    }

    /// Find an attribute by expanded name.
    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(MathAttribute::value)
    }

    /// Return ordered mixed content.
    pub fn content(&self) -> &[MathContent] {
        &self.content
    }

    /// Iterate direct child elements.
    pub fn children(&self) -> impl Iterator<Item = &MathElement> {
        self.content.iter().filter_map(|content| match content {
            MathContent::Element(element) => Some(element),
            MathContent::Text(_) => None,
        })
    }

    /// Compose all descendant character content in exact element/text order.
    pub fn all_text(&self) -> String {
        fn append(element: &MathElement, output: &mut String) {
            for content in &element.content {
                match content {
                    MathContent::Text(text) => output.push_str(text),
                    MathContent::Element(child) => append(child, output),
                }
            }
        }
        let mut output = String::new();
        append(self, &mut output);
        output
    }

    fn collect_annotations<'a>(&'a self, output: &mut Vec<&'a MathElement>) {
        if matches!(
            self.kind(),
            MathElementKind::Annotation | MathElementKind::AnnotationXml
        ) {
            output.push(self);
        }
        for child in self.children() {
            child.collect_annotations(output);
        }
    }
}

/// A validated OpenDocument Formula (`.odf`) or formula template (`.otf`).
///
/// Formula markup and annotations are inert data and are never evaluated.
pub struct FormulaDocument {
    package: OpenDocumentPackage,
    math: MathElement,
}

impl FormulaDocument {
    /// Create a standard OpenDocument Formula package from validated MathML.
    ///
    /// The supplied XML must have a MathML `math` root. Its markup and any
    /// annotations are stored as inert document data and are never evaluated.
    pub fn create(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_with_mimetype(mathml.as_ref(), constants::ODF_FORMULA)
    }

    /// Create an OpenDocument Formula template package from validated MathML.
    ///
    /// This is equivalent to [`Self::create`] but emits the standard formula
    /// template MIME type used by `.otf` files. Formula markup remains inert.
    pub fn create_template(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_with_mimetype(mathml.as_ref(), constants::ODF_FORMULA_TEMPLATE)
    }

    fn create_with_mimetype(mathml: &str, mimetype: &str) -> Result<Self> {
        // Validate before emitting any package data so malformed MathML never
        // becomes a partially authored formula package.
        parse_mathml(mathml)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype)?;
        writer.add_file(constants::ODF_CONTENT, mathml.as_bytes())?;
        Self::from_bytes(writer.finish_to_bytes()?)
    }

    /// Open a formula package from a path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read a formula package from a stream.
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    /// Validate a formula package from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OpenDocumentPackage::from_bytes(bytes)?;
        if package.family() != OpenDocumentFamily::Formula {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument formula: MIME type is '{}'",
                package.mimetype()
            )));
        }
        let math = parse_mathml(&package.content_xml()?)?;
        Ok(Self { package, math })
    }

    /// Whether this package uses the formula-template MIME type.
    pub fn is_template(&self) -> bool {
        self.package.is_template()
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &str {
        self.package.mimetype()
    }

    /// Return the complete MathML root.
    pub fn math(&self) -> &MathElement {
        &self.math
    }

    /// Replace the document's MathML tree and repackage the formula.
    ///
    /// `math` must be a MathML `math` element. The tree is serialized,
    /// re-validated through the same bounded parser used for reading, and
    /// stored as the new `content.xml`; remaining package files are copied
    /// unchanged. The document is left untouched when validation fails.
    pub fn set_math(&mut self, math: MathElement) -> Result<()> {
        if math.namespace_uri() != Some(MATHML_NAMESPACE) || math.local_name() != "math" {
            return Err(Error::InvalidFormat(
                "replacement formula tree must have a MathML math root".to_string(),
            ));
        }
        let xml = math.to_xml();
        // Re-validate the serialized form so a package can never contain
        // MathML the reader itself would reject.
        let validated = parse_mathml(&xml)?;

        let mut writer = PackageWriter::new();
        writer.set_mimetype(self.package.mimetype())?;
        writer.add_file(constants::ODF_CONTENT, xml.as_bytes())?;
        if self.package.has_file(constants::ODF_META)? {
            let bytes = self.package.get_file(constants::ODF_META)?;
            writer.add_file(constants::ODF_META, &bytes)?;
        }
        writer.copy_auxiliary_files_from(self.package.owned_package())?;
        let package = OpenDocumentPackage::from_bytes(writer.finish_to_bytes()?)?;
        self.package = package;
        self.math = validated;
        Ok(())
    }

    /// Return every MathML annotation in document order.
    pub fn annotations(&self) -> Vec<&MathElement> {
        let mut annotations = Vec::new();
        self.math.collect_annotations(&mut annotations);
        annotations
    }

    /// Return the first StarMath annotation source, when present.
    pub fn starmath_source(&self) -> Option<String> {
        self.annotations().into_iter().find_map(|annotation| {
            annotation
                .attribute(None, "encoding")
                .is_some_and(|encoding| encoding.eq_ignore_ascii_case("StarMath 5.0"))
                .then(|| annotation.all_text())
        })
    }

    /// Extract common package metadata.
    pub fn metadata(&self) -> Result<Metadata> {
        self.package.metadata()
    }

    /// Extract complete OpenDocument metadata.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.package.odf_metadata()
    }

    /// Return the exact original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the exact original package bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.to_bytes()
    }

    /// Consume the document and return the exact original package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Save without reconstructing the MathML or ZIP package.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

pub(crate) fn parse_mathml(xml: &str) -> Result<MathElement> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root = None;
    let mut root_closed = false;
    let mut node_count = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid formula MathML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if root_closed {
                    return Err(Error::InvalidFormat(
                        "formula contains multiple root elements".to_string(),
                    ));
                }
                let resolved_namespace_uri = namespace_uri(&namespace)?;
                let node = make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                if stack.is_empty()
                    && (node.namespace_uri() != Some(MATHML_NAMESPACE)
                        || node.local_name() != "math")
                {
                    return Err(Error::InvalidFormat(
                        "formula content must have a MathML math root".to_string(),
                    ));
                }
                stack.push(node);
                if stack.len() > MAX_MATH_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "MathML nesting exceeds {MAX_MATH_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref element) => {
                if stack.is_empty() {
                    if root_closed {
                        return Err(Error::InvalidFormat(
                            "formula contains multiple root elements".to_string(),
                        ));
                    }
                    let resolved_namespace_uri = namespace_uri(&namespace)?;
                    let node =
                        make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                    if node.namespace_uri() != Some(MATHML_NAMESPACE) || node.local_name() != "math"
                    {
                        return Err(Error::InvalidFormat(
                            "formula content must have a MathML math root".to_string(),
                        ));
                    }
                    root = Some(node);
                    root_closed = true;
                    buffer.clear();
                    continue;
                }
                let resolved_namespace_uri = namespace_uri(&namespace)?;
                let node = make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                stack
                    .last_mut()
                    .expect("parent exists")
                    .content
                    .push(MathContent::Element(node));
            },
            Event::End(_) => {
                let node = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("MathML element stack underflow".to_string())
                })?;
                if let Some(parent) = stack.last_mut() {
                    parent.content.push(MathContent::Element(node));
                } else {
                    if root.is_some() {
                        return Err(Error::InvalidFormat(
                            "formula contains multiple MathML roots".to_string(),
                        ));
                    }
                    root = Some(node);
                    root_closed = true;
                }
            },
            Event::Text(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid MathML text: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid MathML CDATA: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => {
                push_text(
                    stack.last_mut().expect("element exists"),
                    decode_reference(reference)?,
                    &mut text_bytes,
                )?;
            },
            Event::Text(ref text) if !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the MathML root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if stack.is_empty() => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the MathML root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || !root_closed {
        return Err(Error::InvalidFormat(
            "formula contains incomplete MathML".to_string(),
        ));
    }
    root.ok_or_else(|| Error::InvalidFormat("formula has no MathML root".to_string()))
}

fn make_element(
    reader: &NsReader<&[u8]>,
    resolved_namespace_uri: Option<String>,
    element: &BytesStart<'_>,
    node_count: &mut usize,
) -> Result<MathElement> {
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("MathML node count overflow".to_string()))?;
    if *node_count > MAX_MATH_NODES {
        return Err(Error::InvalidFormat(format!(
            "formula exceeds {MAX_MATH_NODES} MathML elements"
        )));
    }
    if element.attributes().count() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "MathML element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    let local_name = decode_utf8(element.local_name().as_ref(), "element name")?;
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid MathML attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_uri(&namespace)?;
        let local_name = decode_utf8(local.as_ref(), "attribute name")?;
        if attributes.iter().any(|existing: &MathAttribute| {
            existing.namespace_uri == namespace_uri && existing.local_name == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded MathML attribute '{local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid MathML attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "MathML attribute exceeds 1 MiB".to_string(),
            ));
        }
        attributes.push(MathAttribute {
            namespace_uri,
            local_name,
            value,
        });
    }
    Ok(MathElement {
        namespace_uri: resolved_namespace_uri,
        local_name,
        attributes,
        content: Vec::new(),
    })
}

fn push_text(element: &mut MathElement, value: String, total: &mut usize) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("MathML text size overflow".to_string()))?;
    if *total > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(
            "formula exceeds 32 MiB of MathML text".to_string(),
        ));
    }
    if let Some(MathContent::Text(existing)) = element.content.last_mut() {
        existing.push_str(&value);
    } else {
        element.content.push(MathContent::Text(value));
    }
    Ok(())
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_utf8(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown MathML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode_utf8(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 MathML {kind}")))
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid MathML character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid MathML entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Ok(format!("&{name};")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;
    use std::io::Cursor;

    fn package(mimetype: &str, content: &str) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn formula_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8"?>
<math xmlns="http://www.w3.org/1998/Math/MathML" xmlns:v="urn:vendor:math" display="block">
 <semantics>
  <mrow><mi mathvariant="italic">f</mi><mo stretchy="false">(</mo><mi>x</mi><mo>)</mo><mo>=</mo><mfrac><mn>1</mn><mrow><mi>x</mi><mo>+</mo><mn>2</mn></mrow></mfrac><mtext><![CDATA[ exact <text> ]]></mtext><v:hint v:mode="safe">extension</v:hint></mrow>
  <annotation encoding="StarMath 5.0">{ f(x) = 1 over {x+2} &amp; "exact" }</annotation>
  <annotation-xml encoding="application/mathml-content+xml"><apply xmlns="http://www.w3.org/1998/Math/MathML"><divide/><cn>1</cn><ci>x</ci></apply></annotation-xml>
 </semantics>
</math>"#
    }

    #[test]
    fn parses_libreoffice_style_mathml_and_annotations_losslessly() {
        let bytes = package(constants::ODF_FORMULA, formula_xml());
        let document = FormulaDocument::from_bytes(bytes.clone()).unwrap();
        assert!(!document.is_template());
        assert_eq!(document.math().kind(), MathElementKind::Math);
        assert_eq!(document.math().attribute(None, "display"), Some("block"));
        let semantics = document.math().children().next().unwrap();
        assert_eq!(semantics.kind(), MathElementKind::Semantics);
        let row = semantics.children().next().unwrap();
        assert_eq!(row.kind(), MathElementKind::Row);
        let kinds: Vec<_> = row.children().map(MathElement::kind).collect();
        assert!(kinds.contains(&MathElementKind::Fraction));
        assert!(kinds.contains(&MathElementKind::Text));
        assert!(kinds.contains(&MathElementKind::Other));
        assert!(row.all_text().contains("f(x)=1x+2 exact <text> extension"));
        assert_eq!(document.annotations().len(), 2);
        assert_eq!(
            document.starmath_source().as_deref(),
            Some("{ f(x) = 1 over {x+2} & \"exact\" }")
        );
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn accepts_formula_templates_readers_and_empty_math() {
        let bytes = package(
            constants::ODF_FORMULA_TEMPLATE,
            r#"<m:math xmlns:m="http://www.w3.org/1998/Math/MathML"/>"#,
        );
        let document = FormulaDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
        assert!(document.is_template());
        assert!(document.math().content().is_empty());
        assert_eq!(document.into_bytes(), bytes);
    }

    #[test]
    fn creates_formula_and_template_packages_from_validated_mathml() {
        let formula = FormulaDocument::create(formula_xml()).unwrap();
        assert!(!formula.is_template());
        assert_eq!(formula.mimetype(), constants::ODF_FORMULA);
        assert_eq!(formula.math().kind(), MathElementKind::Math);
        let package = crate::OpenDocumentPackage::from_bytes(formula.to_bytes()).unwrap();
        assert_eq!(package.content_xml().unwrap(), formula_xml());

        let template = FormulaDocument::create_template(
            r#"<m:math xmlns:m="http://www.w3.org/1998/Math/MathML"><m:mi>x</m:mi></m:math>"#,
        )
        .unwrap();
        assert!(template.is_template());
        assert_eq!(template.mimetype(), constants::ODF_FORMULA_TEMPLATE);
        assert_eq!(
            template.math().children().next().unwrap().kind(),
            MathElementKind::Identifier
        );

        assert!(FormulaDocument::create("<math/>").is_err());
        assert!(FormulaDocument::create_template("not XML").is_err());
    }

    #[test]
    fn retains_mixed_content_in_exact_semantic_order() {
        let xml = r#"<!DOCTYPE math [<!ENTITY ApplyFunction "&#x2061;">]><math xmlns="http://www.w3.org/1998/Math/MathML"><mrow>before<mi>a</mi>&ApplyFunction;between<mo>+</mo><mi>b</mi>after</mrow></math>"#;
        let document = FormulaDocument::from_bytes(package(constants::ODF_FORMULA, xml)).unwrap();
        let row = document.math().children().next().unwrap();
        assert_eq!(row.all_text(), "beforea&ApplyFunction;between+bafter");
        assert!(matches!(row.content()[0], MathContent::Text(ref value) if value == "before"));
        assert!(matches!(row.content()[1], MathContent::Element(_)));
        assert!(
            matches!(row.content()[2], MathContent::Text(ref value) if value == "&ApplyFunction;between")
        );
    }

    #[test]
    fn rejects_other_families_wrong_roots_and_incomplete_xml() {
        assert!(FormulaDocument::from_bytes(package(constants::ODF_CHART, formula_xml())).is_err());
        for xml in [
            r#"<math/>"#,
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi>"#,
            r#"<math xmlns="http://www.w3.org/1998/Math/MathML"/><math xmlns="http://www.w3.org/1998/Math/MathML"/>"#,
            r#"outside<math xmlns="http://www.w3.org/1998/Math/MathML"/>"#,
        ] {
            assert!(
                FormulaDocument::from_bytes(package(constants::ODF_FORMULA, xml)).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn set_math_replaces_the_tree_and_repackages_atomically() {
        let bytes = package(constants::ODF_FORMULA, formula_xml());
        let mut document = FormulaDocument::from_bytes(bytes.clone()).unwrap();
        let original_math = document.math().clone();

        // Non-math roots are rejected without touching the package.
        let foreign = MathElement::new("mrow").unwrap();
        assert!(document.set_math(foreign).is_err());
        assert_eq!(document.as_bytes(), bytes.as_slice());

        let replacement = crate::formula::builder::document_root(
            crate::formula::builder::fraction(
                crate::formula::builder::number("1"),
                crate::formula::builder::identifier("x"),
            ),
            crate::formula::builder::MathDisplay::Inline,
        );
        document.set_math(replacement.clone()).unwrap();
        assert_eq!(document.math(), &replacement);
        assert!(document.to_bytes() != bytes);

        // The repackaged bytes reopen cleanly and preserve the family.
        let reopened = FormulaDocument::from_bytes(document.to_bytes()).unwrap();
        assert_eq!(reopened.math(), &replacement);
        assert_eq!(reopened.mimetype(), constants::ODF_FORMULA);
        assert!(!reopened.is_template());
        assert_ne!(original_math, replacement);
    }

    #[test]
    fn set_math_preserves_package_metadata() {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_FORMULA).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, formula_xml().as_bytes())
            .unwrap();
        writer
            .add_file(constants::ODF_META, b"<office:document-meta/>")
            .unwrap();
        let mut document = FormulaDocument::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

        let replacement = crate::formula::builder::document_root(
            crate::formula::builder::identifier("y"),
            crate::formula::builder::MathDisplay::Block,
        );
        document.set_math(replacement).unwrap();
        let package = crate::OpenDocumentPackage::from_bytes(document.to_bytes()).unwrap();
        assert!(package.has_file(constants::ODF_META).unwrap());
        assert_eq!(
            package.get_file(constants::ODF_META).unwrap(),
            b"<office:document-meta/>"
        );
    }

    #[test]
    fn rejects_duplicate_expanded_attributes_and_excessive_nesting() {
        let duplicate = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" xmlns:a="urn:test" xmlns:b="urn:test" a:x="one" b:x="two"/>"#;
        assert!(FormulaDocument::from_bytes(package(constants::ODF_FORMULA, duplicate)).is_err());

        let nested = "<mrow>".repeat(MAX_MATH_DEPTH) + &"</mrow>".repeat(MAX_MATH_DEPTH);
        let deep = format!(r#"<math xmlns="{MATHML_NAMESPACE}">{nested}</math>"#);
        assert!(FormulaDocument::from_bytes(package(constants::ODF_FORMULA, &deep)).is_err());
    }
}
