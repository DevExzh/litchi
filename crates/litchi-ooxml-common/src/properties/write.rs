use super::{
    Dialect, Keywords, MAX_PROPERTY_TEXT, MAX_XML_BYTES, MAX_XML_EVENTS, Props, graph,
    keyword::Item,
};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::rel::TargetMode;
use litchi_opc::{OpcPackage, PackURI};

const CANONICAL_PART: &str = "/docProps/core.xml";
const MAX_PART_CANDIDATES: u32 = 10_000;
// Reserve enough reader events for the declaration, root, every other
// property, and EOF. A nonempty structured keyword value costs Start, Text,
// and End events, so three events per item is the safe worst case.
const MAX_KEYWORD_ITEMS: usize = (MAX_XML_EVENTS - 64) / 3;
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const CORE_OPEN_PREFIX: &str = r#"<cp:coreProperties xmlns:cp=""#;
const CORE_OPEN_SUFFIX: &str = r#"" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#;
const CORE_CLOSE: &str = "</cp:coreProperties>";

/// Consumes and writes a present core-properties value.
///
/// Existing nonstandard target paths, relationship IDs, and Strict versus
/// Transitional dialects are retained. Returns whether package bytes changed.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn write(package: &mut OpcPackage, props: Props) -> Result<bool> {
    sync(package, &props)
}

/// Removes the validated core-properties graph.
///
/// Absence is a successful no-op. Shared or ambiguous parts are rejected before
/// mutation. Signatures are removed only after a real package mutation.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn clear(package: &mut OpcPackage) -> Result<bool> {
    let graph = graph::inspect(package)?;
    let Some(part_name) = graph.part else {
        return Ok(false);
    };
    let relationship_id = graph.relationship_id.ok_or_else(|| {
        Error::Relationship("core-properties part has no owning relationship".to_owned())
    })?;
    graph::ensure_clear_safe(package, &part_name, &relationship_id)?;

    let mut relationships = package.rels().clone();
    if relationships.remove(&relationship_id).is_none() {
        return Err(Error::Relationship(format!(
            "core-properties relationship '{relationship_id}' disappeared during staging"
        )));
    }

    if !package.remove_part(&part_name) {
        return Err(Error::Missing(format!(
            "core-properties part '{}' disappeared during removal",
            part_name.as_str()
        )));
    }
    *package.rels_mut() = relationships;
    package.unsign();
    Ok(true)
}

pub(super) fn sync(package: &mut OpcPackage, props: &Props) -> Result<bool> {
    let graph = graph::inspect(package)?;
    match (graph.part, graph.dialect) {
        (Some(part_name), Some(dialect)) => {
            let current = package.get_part(&part_name)?;
            let current_xml = std::str::from_utf8(current.blob()).map_err(|error| {
                Error::Xml(format!("invalid UTF-8 in core properties: {error}"))
            })?;
            let (current_props, actual_dialect) = super::read::decode(current_xml)?;
            if actual_dialect != dialect {
                return Err(Error::Invalid(
                    "core-properties XML namespace does not match its relationship dialect"
                        .to_owned(),
                ));
            }
            if current_props == *props {
                return Ok(false);
            }
            let xml = encode(props, dialect)?.into_bytes();
            if current.blob() == xml.as_slice() {
                return Ok(false);
            }
            package.get_part_mut(&part_name)?.set_blob(xml);
            package.unsign();
            Ok(true)
        },
        (None, None) => {
            let part_name = available_part_name(package)?;
            let xml = encode(props, Dialect::Transitional)?.into_bytes();
            let target = part_name.as_str().trim_start_matches('/').to_owned();
            let mut relationships = package.rels().clone();
            let relationship_id = next_relationship_id(&relationships)?;
            relationships.try_add_relationship(
                rt::CORE_PROPERTIES.to_owned(),
                target,
                relationship_id,
                TargetMode::Internal,
            )?;
            package.validate_new_part_name(&part_name)?;
            let part = BlobPart::new(part_name, ct::OPC_CORE_PROPERTIES.to_owned(), xml);

            package.add_part(Box::new(part));
            *package.rels_mut() = relationships;
            package.unsign();
            Ok(true)
        },
        _ => Err(Error::Relationship(
            "core-properties graph has inconsistent part and relationship state".to_owned(),
        )),
    }
}

pub(super) fn encode(props: &Props, dialect: Dialect) -> Result<String> {
    let encoded_bytes = preflight(props, dialect)?;
    let mut xml = BoundedXml::with_capacity(encoded_bytes)?;
    xml.push_str(XML_DECLARATION)?;
    xml.push_str(CORE_OPEN_PREFIX)?;
    xml.push_str(dialect.namespace())?;
    xml.push_str(CORE_OPEN_SUFFIX)?;

    push_text(&mut xml, "dc:title", props.title.as_deref())?;
    push_text(&mut xml, "dc:subject", props.subject.as_deref())?;
    push_text(&mut xml, "dc:creator", props.creator.as_deref())?;
    push_keywords(&mut xml, props.keywords.as_ref())?;
    push_text(&mut xml, "dc:description", props.description.as_deref())?;
    push_text(&mut xml, "dc:identifier", props.identifier.as_deref())?;
    push_text(
        &mut xml,
        "cp:lastModifiedBy",
        props.last_modified_by.as_deref(),
    )?;
    push_text(&mut xml, "cp:category", props.category.as_deref())?;
    push_text(
        &mut xml,
        "cp:contentStatus",
        props.content_status.as_deref(),
    )?;
    push_text(&mut xml, "cp:revision", props.revision.as_deref())?;
    push_text(&mut xml, "cp:version", props.version.as_deref())?;
    push_text(&mut xml, "dc:language", props.language.as_deref())?;
    if let Some(created) = props.created.as_ref() {
        validate_text(created.as_str(), "dcterms:created")?;
        xml.push_str(r#"<dcterms:created xsi:type="dcterms:W3CDTF">"#)?;
        push_escaped(&mut xml, created.as_str())?;
        xml.push_str("</dcterms:created>")?;
    }
    if let Some(modified) = props.modified.as_ref() {
        validate_text(modified.as_str(), "dcterms:modified")?;
        xml.push_str(r#"<dcterms:modified xsi:type="dcterms:W3CDTF">"#)?;
        push_escaped(&mut xml, modified.as_str())?;
        xml.push_str("</dcterms:modified>")?;
    }
    if let Some(last_printed) = props.last_printed.as_ref() {
        push_text(&mut xml, "cp:lastPrinted", Some(last_printed.as_str()))?;
    }
    xml.push_str(CORE_CLOSE)?;
    if xml.len() != encoded_bytes {
        return Err(Error::Invalid(
            "core-properties encoded-size invariant failed".to_owned(),
        ));
    }
    xml.finish()
}

fn preflight(props: &Props, dialect: Dialect) -> Result<usize> {
    let mut bytes = 0usize;
    add_xml_bytes(&mut bytes, XML_DECLARATION.len())?;
    add_xml_bytes(&mut bytes, CORE_OPEN_PREFIX.len())?;
    add_xml_bytes(&mut bytes, dialect.namespace().len())?;
    add_xml_bytes(&mut bytes, CORE_OPEN_SUFFIX.len())?;

    for (element, value) in [
        ("dc:title", props.title.as_deref()),
        ("dc:subject", props.subject.as_deref()),
        ("dc:creator", props.creator.as_deref()),
        ("dc:description", props.description.as_deref()),
        ("dc:identifier", props.identifier.as_deref()),
        ("cp:lastModifiedBy", props.last_modified_by.as_deref()),
        ("cp:category", props.category.as_deref()),
        ("cp:contentStatus", props.content_status.as_deref()),
        ("cp:revision", props.revision.as_deref()),
        ("cp:version", props.version.as_deref()),
        ("dc:language", props.language.as_deref()),
        (
            "cp:lastPrinted",
            props
                .last_printed
                .as_ref()
                .map(super::time::DateTime::as_str),
        ),
    ] {
        add_text_element_bytes(&mut bytes, element, value)?;
    }
    if let Some(keywords) = props.keywords.as_ref() {
        add_xml_bytes(&mut bytes, keyword_xml_bytes(keywords)?)?;
    }
    if let Some(created) = props.created.as_ref() {
        add_typed_time_bytes(
            &mut bytes,
            "dcterms:created",
            created.as_str(),
            r#"<dcterms:created xsi:type="dcterms:W3CDTF">"#,
            "</dcterms:created>",
        )?;
    }
    if let Some(modified) = props.modified.as_ref() {
        add_typed_time_bytes(
            &mut bytes,
            "dcterms:modified",
            modified.as_str(),
            r#"<dcterms:modified xsi:type="dcterms:W3CDTF">"#,
            "</dcterms:modified>",
        )?;
    }
    add_xml_bytes(&mut bytes, CORE_CLOSE.len())?;
    Ok(bytes)
}

fn add_text_element_bytes(bytes: &mut usize, element: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    add_xml_bytes(bytes, element.len().saturating_mul(2).saturating_add(5))?;
    add_xml_bytes(bytes, escaped_xml_bytes(value, element)?)
}

fn add_typed_time_bytes(
    bytes: &mut usize,
    property: &str,
    value: &str,
    open: &str,
    close: &str,
) -> Result<()> {
    add_xml_bytes(bytes, open.len())?;
    add_xml_bytes(bytes, escaped_xml_bytes(value, property)?)?;
    add_xml_bytes(bytes, close.len())
}

fn keyword_xml_bytes(keywords: &Keywords) -> Result<usize> {
    if keywords.items.len() > MAX_KEYWORD_ITEMS {
        return Err(Error::Limit {
            resource: "core keyword items",
            max: MAX_KEYWORD_ITEMS,
            actual: keywords.items.len(),
        });
    }

    let mut encoded = 0usize;
    let mut text = 0usize;
    add_xml_bytes(&mut encoded, "<cp:keywords".len())?;
    if let Some(language) = keywords.lang.as_ref() {
        add_xml_bytes(&mut encoded, r#" xml:lang=""#.len())?;
        add_xml_bytes(
            &mut encoded,
            escaped_xml_bytes(language.as_str(), "cp:keywords xml:lang")?,
        )?;
        add_xml_bytes(&mut encoded, 1)?;
    }
    add_xml_bytes(&mut encoded, 1)?;

    for item in &keywords.items {
        match item {
            Item::Text(value) => {
                add_keyword_bytes(&mut text, value.len())?;
                add_xml_bytes(&mut encoded, escaped_xml_bytes(value, "cp:keywords")?)?;
            },
            Item::Value(value) => {
                add_keyword_bytes(&mut text, value.text.len())?;
                add_xml_bytes(&mut encoded, "<cp:value".len())?;
                if let Some(language) = value.lang.as_ref() {
                    add_xml_bytes(&mut encoded, r#" xml:lang=""#.len())?;
                    add_xml_bytes(
                        &mut encoded,
                        escaped_xml_bytes(language.as_str(), "cp:value xml:lang")?,
                    )?;
                    add_xml_bytes(&mut encoded, 1)?;
                }
                add_xml_bytes(&mut encoded, 1)?;
                add_xml_bytes(&mut encoded, escaped_xml_bytes(&value.text, "cp:value")?)?;
                add_xml_bytes(&mut encoded, "</cp:value>".len())?;
            },
        }
    }
    add_xml_bytes(&mut encoded, "</cp:keywords>".len())?;
    Ok(encoded)
}

fn escaped_xml_bytes(value: &str, property: &str) -> Result<usize> {
    if value.len() > MAX_PROPERTY_TEXT {
        return Err(Error::Limit {
            resource: "core property text bytes",
            max: MAX_PROPERTY_TEXT,
            actual: value.len(),
        });
    }
    let mut bytes = 0usize;
    for character in value.chars() {
        if !is_xml_10_char(character) {
            return Err(Error::Invalid(format!(
                "{property} contains XML 1.0-forbidden character U+{:04X}",
                u32::from(character)
            )));
        }
        let additional = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        bytes = bytes.checked_add(additional).ok_or(Error::Limit {
            resource: "core-properties XML bytes",
            max: MAX_XML_BYTES,
            actual: usize::MAX,
        })?;
    }
    Ok(bytes)
}

fn add_xml_bytes(total: &mut usize, additional: usize) -> Result<()> {
    let actual = total.checked_add(additional).ok_or(Error::Limit {
        resource: "core-properties XML bytes",
        max: MAX_XML_BYTES,
        actual: usize::MAX,
    })?;
    if actual > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "core-properties XML bytes",
            max: MAX_XML_BYTES,
            actual,
        });
    }
    *total = actual;
    Ok(())
}

/// A fallible XML output sink bounded by the core-properties part limit.
///
/// The sink checks the resulting length before reserving or appending bytes.
/// `encode` preflights the exact output length, so the initial reservation is
/// both bounded and sufficient for the complete document.
struct BoundedXml {
    bytes: Vec<u8>,
}

impl BoundedXml {
    fn with_capacity(capacity: usize) -> Result<Self> {
        if capacity > MAX_XML_BYTES {
            return Err(Error::Limit {
                resource: "core-properties XML bytes",
                max: MAX_XML_BYTES,
                actual: capacity,
            });
        }
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(capacity).map_err(|error| {
            Error::Invalid(format!(
                "core-properties XML output allocation failed: {error}"
            ))
        })?;
        Ok(Self { bytes })
    }

    fn push_str(&mut self, value: &str) -> Result<()> {
        self.push_bytes(value.as_bytes())
    }

    fn push_bytes(&mut self, value: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or(Error::Limit {
                resource: "core-properties XML bytes",
                max: MAX_XML_BYTES,
                actual: usize::MAX,
            })?;
        if length > MAX_XML_BYTES {
            return Err(Error::Limit {
                resource: "core-properties XML bytes",
                max: MAX_XML_BYTES,
                actual: length,
            });
        }
        self.bytes.try_reserve_exact(value.len()).map_err(|error| {
            Error::Invalid(format!(
                "core-properties XML output allocation failed: {error}"
            ))
        })?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        let mut encoded = [0; 4];
        let length = value.encode_utf8(&mut encoded).len();
        self.push_bytes(&encoded[..length])
    }

    fn len(&self) -> usize {
        self.bytes.len()
    }

    fn finish(self) -> Result<String> {
        String::from_utf8(self.bytes).map_err(|error| {
            Error::Invalid(format!("core-properties XML output was not UTF-8: {error}"))
        })
    }
}

fn push_escaped(xml: &mut BoundedXml, value: &str) -> Result<()> {
    let mut start = 0;
    for (index, character) in value.char_indices() {
        let replacement = match character {
            '&' => Some("&amp;"),
            '<' => Some("&lt;"),
            '>' => Some("&gt;"),
            '"' => Some("&quot;"),
            '\'' => Some("&apos;"),
            _ => None,
        };
        let Some(replacement) = replacement else {
            continue;
        };
        xml.push_str(&value[start..index])?;
        xml.push_str(replacement)?;
        start = index + character.len_utf8();
    }
    xml.push_str(&value[start..])
}

fn push_keywords(xml: &mut BoundedXml, keywords: Option<&Keywords>) -> Result<()> {
    let Some(keywords) = keywords else {
        return Ok(());
    };
    xml.push_str("<cp:keywords")?;
    if let Some(language) = keywords.lang.as_ref() {
        validate_text(language.as_str(), "cp:keywords xml:lang")?;
        xml.push_str(r#" xml:lang=""#)?;
        push_escaped(xml, language.as_str())?;
        xml.push_char('"')?;
    }
    xml.push_char('>')?;

    let mut bytes = 0usize;
    for item in &keywords.items {
        match item {
            Item::Text(text) => {
                add_keyword_bytes(&mut bytes, text.len())?;
                validate_xml_text(text, "cp:keywords")?;
                push_escaped(xml, text)?;
            },
            Item::Value(value) => {
                add_keyword_bytes(&mut bytes, value.text.len())?;
                validate_xml_text(&value.text, "cp:value")?;
                xml.push_str("<cp:value")?;
                if let Some(language) = value.lang.as_ref() {
                    validate_text(language.as_str(), "cp:value xml:lang")?;
                    xml.push_str(r#" xml:lang=""#)?;
                    push_escaped(xml, language.as_str())?;
                    xml.push_char('"')?;
                }
                xml.push_char('>')?;
                push_escaped(xml, &value.text)?;
                xml.push_str("</cp:value>")?;
            },
        }
    }
    xml.push_str("</cp:keywords>")?;
    Ok(())
}

fn add_keyword_bytes(total: &mut usize, additional: usize) -> Result<()> {
    *total = total.checked_add(additional).ok_or(Error::Limit {
        resource: "core property text bytes",
        max: MAX_PROPERTY_TEXT,
        actual: usize::MAX,
    })?;
    if *total > MAX_PROPERTY_TEXT {
        return Err(Error::Limit {
            resource: "core property text bytes",
            max: MAX_PROPERTY_TEXT,
            actual: *total,
        });
    }
    Ok(())
}

fn push_text(xml: &mut BoundedXml, element: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    validate_text(value, element)?;
    xml.push_char('<')?;
    xml.push_str(element)?;
    xml.push_char('>')?;
    push_escaped(xml, value)?;
    xml.push_str("</")?;
    xml.push_str(element)?;
    xml.push_char('>')?;
    Ok(())
}

fn validate_text(value: &str, property: &str) -> Result<()> {
    if value.len() > MAX_PROPERTY_TEXT {
        return Err(Error::Limit {
            resource: "core property text bytes",
            max: MAX_PROPERTY_TEXT,
            actual: value.len(),
        });
    }
    validate_xml_text(value, property)
}

fn validate_xml_text(value: &str, property: &str) -> Result<()> {
    if let Some(character) = value.chars().find(|character| !is_xml_10_char(*character)) {
        return Err(Error::Invalid(format!(
            "{property} contains XML 1.0-forbidden character U+{:04X}",
            u32::from(character)
        )));
    }
    Ok(())
}

fn is_xml_10_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(u32::from(character), 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn available_part_name(package: &OpcPackage) -> Result<PackURI> {
    for index in 1..=MAX_PART_CANDIDATES {
        let candidate = if index == 1 {
            CANONICAL_PART.to_owned()
        } else {
            format!("/docProps/core{index}.xml")
        };
        let candidate = PackURI::new(&candidate).map_err(Error::Uri)?;
        if package.validate_new_part_name(&candidate).is_ok() {
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        resource: "core-properties part-name candidates",
        max: MAX_PART_CANDIDATES as usize,
        actual: MAX_PART_CANDIDATES as usize,
    })
}

fn next_relationship_id(relationships: &litchi_opc::Relationships) -> Result<String> {
    let mut index = 1u32;
    loop {
        let candidate = format!("rId{index}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
        index = index.checked_add(1).ok_or(Error::Limit {
            resource: "package relationship identifiers",
            max: u32::MAX as usize,
            actual: usize::MAX,
        })?;
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::properties::{STRICT_CORE_PROPERTIES_RELATIONSHIP, Slot, read};
    use litchi_opc::part::Part;

    fn package_with_core(path: &str, relationship_type: &str, xml: Vec<u8>) -> OpcPackage {
        let mut package = OpcPackage::new();
        let name = PackURI::new(path).unwrap();
        package.add_part(Box::new(BlobPart::new(
            name,
            ct::OPC_CORE_PROPERTIES.to_owned(),
            xml,
        )));
        package.relate_to(path.trim_start_matches('/'), relationship_type);
        package
    }

    #[test]
    fn writes_reads_and_clears_nonstandard_target() {
        let original = encode(&Props::new().title("Before"), Dialect::Transitional)
            .unwrap()
            .into_bytes();
        let mut package = package_with_core("/odd/Metadata.XML", rt::CORE_PROPERTIES, original);
        let relationship_id = package
            .rels()
            .iter()
            .find(|relationship| graph::is_core_relationship(relationship.reltype()))
            .unwrap()
            .r_id()
            .to_owned();

        assert!(write(&mut package, Props::new().title("After")).unwrap());
        assert_eq!(
            read(&package).unwrap().unwrap().title.as_deref(),
            Some("After")
        );
        let relationship = package.rels().get(&relationship_id).unwrap();
        assert_eq!(relationship.target_ref(), "odd/Metadata.XML");
        assert!(
            package
                .iter_parts()
                .all(|part| part.partname().as_str() != CANONICAL_PART)
        );

        assert!(clear(&mut package).unwrap());
        assert!(!clear(&mut package).unwrap());
        assert!(read(&package).unwrap().is_none());
    }

    #[test]
    fn preserves_strict_dialect_and_byte_identical_write_is_a_noop() {
        let props = Props::new().title("Strict");
        let original = format!(
            "<?xml version=\"1.0\"?>\n<cp:coreProperties xmlns:cp=\"{}\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n  <dc:title>Strict</dc:title>\n</cp:coreProperties>",
            Dialect::Strict.namespace()
        )
        .into_bytes();
        let mut package = package_with_core(
            "/strict/core.xml",
            STRICT_CORE_PROPERTIES_RELATIONSHIP,
            original.clone(),
        );

        assert!(!write(&mut package, props.clone()).unwrap());
        assert_eq!(
            package
                .get_part(&PackURI::new("/strict/core.xml").unwrap())
                .unwrap()
                .blob(),
            original
        );
        assert!(write(&mut package, props.title("Changed")).unwrap());
        let xml = package
            .get_part(&PackURI::new("/strict/core.xml").unwrap())
            .unwrap()
            .blob();
        assert!(
            xml.windows(Dialect::Strict.namespace().len())
                .any(|window| window == Dialect::Strict.namespace().as_bytes())
        );
    }

    #[test]
    fn equal_put_through_slot_preserves_exact_bytes_and_signature() {
        let original = format!(
            "<?xml version=\"1.0\"?>\n<cp:coreProperties xmlns:cp=\"{}\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Same</dc:title></cp:coreProperties>",
            Dialect::Transitional.namespace()
        )
        .into_bytes();
        let mut package =
            package_with_core("/metadata/core.xml", rt::CORE_PROPERTIES, original.clone());
        add_signature_origin(&mut package);
        let mut slot = Slot::load(&package).unwrap();
        slot.put(Props::new().title("Same"));

        assert!(!slot.flush(&mut package).unwrap());
        assert!(package.is_signed());
        assert_eq!(
            package
                .get_part(&PackURI::new("/metadata/core.xml").unwrap())
                .unwrap()
                .blob(),
            original
        );

        slot.get_mut().unwrap().title = Some("Changed".to_owned());
        assert!(slot.flush(&mut package).unwrap());
        assert!(!package.is_signed());
    }

    #[test]
    fn distinguishes_present_empty_from_absence() {
        let absent = OpcPackage::new();
        assert!(read(&absent).unwrap().is_none());

        let xml = encode(&Props::new(), Dialect::Transitional)
            .unwrap()
            .into_bytes();
        let present = package_with_core("/docProps/core.xml", rt::CORE_PROPERTIES, xml);
        assert_eq!(read(&present).unwrap(), Some(Props::new()));
    }

    #[test]
    fn revision_writer_validates_even_empty_and_opaque_lexical_values() {
        for value in ["", "seven", "0007", " 7 "] {
            let xml = Props::new().revision(value).xml().unwrap();
            assert!(xml.contains("<cp:revision>"));
            assert_eq!(
                read::decode(&xml).unwrap().0.revision.as_deref(),
                Some(value)
            );
        }

        assert!(matches!(
            Props::new().revision("bad\0value").xml(),
            Err(Error::Invalid(_))
        ));
        let oversized = "x".repeat(MAX_PROPERTY_TEXT + 1);
        assert!(matches!(
            Props::new().revision(oversized).xml(),
            Err(Error::Limit {
                resource: "core property text bytes",
                max: MAX_PROPERTY_TEXT,
                actual,
            }) if actual == MAX_PROPERTY_TEXT + 1
        ));
    }

    #[test]
    fn rejects_keyword_item_and_encoded_xml_budgets_before_package_mutation() {
        let maximum = Keywords {
            lang: None,
            items: vec![Item::Value(super::super::keyword::Value::new("x")); MAX_KEYWORD_ITEMS],
        };
        let maximum_xml = Props::new().keywords(maximum).xml().unwrap();
        assert_eq!(
            read::decode(&maximum_xml)
                .unwrap()
                .0
                .keywords
                .unwrap()
                .items
                .len(),
            MAX_KEYWORD_ITEMS
        );

        let keywords = Keywords {
            lang: None,
            items: vec![Item::Value(super::super::keyword::Value::new("x")); MAX_KEYWORD_ITEMS + 1],
        };
        assert!(matches!(
            Props::new().keywords(keywords).xml(),
            Err(Error::Limit {
                resource: "core keyword items",
                max: MAX_KEYWORD_ITEMS,
                actual,
            }) if actual == MAX_KEYWORD_ITEMS + 1
        ));

        let original = encode(&Props::new().title("Before"), Dialect::Transitional)
            .unwrap()
            .into_bytes();
        let mut package =
            package_with_core("/docProps/core.xml", rt::CORE_PROPERTIES, original.clone());
        add_signature_origin(&mut package);
        let escaped = "&".repeat(MAX_PROPERTY_TEXT);
        let oversized = Props::new()
            .title(escaped.clone())
            .subject(escaped.clone())
            .creator(escaped.clone())
            .description(escaped);
        assert!(matches!(
            write(&mut package, oversized),
            Err(Error::Limit {
                resource: "core-properties XML bytes",
                max: MAX_XML_BYTES,
                ..
            })
        ));
        assert!(package.is_signed());
        assert_eq!(
            package
                .get_part(&PackURI::new("/docProps/core.xml").unwrap())
                .unwrap()
                .blob(),
            original
        );
    }

    #[test]
    fn bounded_output_rejects_oversized_append_before_mutation() {
        let mut xml = BoundedXml::with_capacity(MAX_XML_BYTES - 1).unwrap();
        xml.bytes.resize(MAX_XML_BYTES - 1, b'x');

        let result = xml.push_str("xx");

        assert!(matches!(
            result,
            Err(Error::Limit {
                resource: "core-properties XML bytes",
                max: MAX_XML_BYTES,
                actual,
            }) if actual == MAX_XML_BYTES + 1
        ));
        assert_eq!(xml.len(), MAX_XML_BYTES - 1);
    }

    #[test]
    fn unknown_payload_blocks_write_but_graph_safe_clear_sanitizes_it() {
        let unknown = format!(
            "<cp:coreProperties xmlns:cp=\"{}\"><cp:unknown>opaque</cp:unknown></cp:coreProperties>",
            Dialect::Transitional.namespace()
        )
        .into_bytes();
        let mut package =
            package_with_core("/docProps/core.xml", rt::CORE_PROPERTIES, unknown.clone());
        add_signature_origin(&mut package);

        assert!(matches!(
            write(&mut package, Props::new().title("Known")),
            Err(Error::Invalid(_))
        ));
        assert!(package.is_signed());
        assert_eq!(
            package
                .get_part(&PackURI::new("/docProps/core.xml").unwrap())
                .unwrap()
                .blob(),
            unknown
        );

        assert!(clear(&mut package).unwrap());
        assert!(!package.is_signed());
        assert!(read(&package).unwrap().is_none());
    }

    #[test]
    fn shared_inbound_allows_read_and_update_but_blocks_clear_before_mutation() {
        let original = encode(&Props::new().title("Before"), Dialect::Transitional)
            .unwrap()
            .into_bytes();
        let mut package =
            package_with_core("/docProps/core.xml", rt::CORE_PROPERTIES, original.clone());
        let source_name = PackURI::new("/word/document.xml").unwrap();
        let mut source = BlobPart::new(source_name, ct::WML_DOCUMENT_MAIN.to_owned(), Vec::new());
        source.relate_to("../docProps/core.xml", "urn:test:shared");
        package.add_part(Box::new(source));

        assert_eq!(
            read(&package).unwrap().unwrap().title.as_deref(),
            Some("Before")
        );
        assert!(write(&mut package, Props::new().title("After")).unwrap());
        assert_eq!(
            read(&package).unwrap().unwrap().title.as_deref(),
            Some("After")
        );

        add_signature_origin(&mut package);
        assert!(matches!(clear(&mut package), Err(Error::Relationship(_))));
        assert!(package.is_signed());
        assert!(
            package
                .get_part(&PackURI::new("/docProps/core.xml").unwrap())
                .is_ok()
        );
        assert!(
            package
                .rels()
                .iter()
                .any(|relationship| graph::is_core_relationship(relationship.reltype()))
        );
    }

    #[test]
    fn update_preserves_outbound_extensions_and_clear_leaves_extension_parts() {
        const EXTENSION_RELATIONSHIP: &str = "urn:test:core-properties-extension";
        let core_name = PackURI::new("/metadata/core.xml").unwrap();
        let extension_name = PackURI::new("/metadata/extension.xml").unwrap();
        let original = encode(&Props::new().title("Before"), Dialect::Transitional)
            .unwrap()
            .into_bytes();
        let mut package = package_with_core(core_name.as_str(), rt::CORE_PROPERTIES, original);
        package.add_part(Box::new(BlobPart::new(
            extension_name.clone(),
            "application/xml".to_owned(),
            b"<extension>retained</extension>".to_vec(),
        )));
        let extension_id = package
            .get_part_mut(&core_name)
            .unwrap()
            .relate_to("extension.xml", EXTENSION_RELATIONSHIP);

        assert!(write(&mut package, Props::new().title("After")).unwrap());
        let extension_relationship = package
            .get_part(&core_name)
            .unwrap()
            .rels()
            .get(&extension_id)
            .unwrap();
        assert_eq!(extension_relationship.reltype(), EXTENSION_RELATIONSHIP);
        assert_eq!(
            extension_relationship.target_partname().unwrap(),
            extension_name
        );
        assert_eq!(
            package.get_part(&extension_name).unwrap().blob(),
            b"<extension>retained</extension>"
        );

        assert!(clear(&mut package).unwrap());
        assert!(package.get_part(&core_name).is_err());
        assert_eq!(
            package.get_part(&extension_name).unwrap().blob(),
            b"<extension>retained</extension>"
        );
        assert!(
            !package
                .rels()
                .iter()
                .any(|relationship| graph::is_core_relationship(relationship.reltype()))
        );
    }

    #[test]
    fn allocates_around_an_unrelated_canonical_part() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(CANONICAL_PART).unwrap(),
            "application/xml".to_owned(),
            b"<unrelated/>".to_vec(),
        )));
        assert!(write(&mut package, Props::new().title("New")).unwrap());
        let graph = graph::inspect(&package).unwrap();
        assert_eq!(graph.part.unwrap().as_str(), "/docProps/core2.xml");
        assert_eq!(
            package
                .get_part(&PackURI::new(CANONICAL_PART).unwrap())
                .unwrap()
                .blob(),
            b"<unrelated/>"
        );
    }

    fn add_signature_origin(package: &mut OpcPackage) {
        let name = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
        package.add_part(Box::new(BlobPart::new(
            name,
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            Vec::new(),
        )));
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(package.is_signed());
    }
}
