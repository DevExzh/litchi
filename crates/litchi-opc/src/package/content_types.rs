//! Source-preserving `[Content_Types].xml` structural edits.

use super::super::{OwnedElementEdit, OwnedElementUpdate, OwnedXmlPart, PackURI, ReadLimits};
use crate::Result;
use crate::error::OpcError;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::fmt::Write as _;

/// Remove `<Override>` elements for the supplied part names and their
/// relationship members. The source XML is edited through validated structural
/// spans, so producer comments and lexical ordering outside those elements are
/// retained exactly.
pub(crate) fn without_part_overrides(
    source: &OwnedXmlPart,
    parts: &[PackURI],
    max_output_bytes: usize,
) -> Result<OwnedXmlPart> {
    let limits = ReadLimits::default();
    let maximum = max_output_bytes.min(limits.max_content_types_bytes());
    if source.bytes().len() > maximum {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types source exceeds the requested output limit".into(),
        ));
    }
    if parts.len() > limits.max_content_type_mappings() {
        return Err(OpcError::InvalidContentTypesManifest(
            "too many content-types removal selectors".into(),
        ));
    }
    let mut selector_bytes = 0usize;
    for part in parts {
        // A relationship URI adds at most "/_rels" and ".rels". Charge the
        // upper bound before constructing another URI from a caller's name.
        selector_bytes = selector_bytes
            .checked_add(part.as_str().len())
            .and_then(|total| total.checked_add(part.as_str().len()))
            .and_then(|total| total.checked_add(27))
            .ok_or_else(|| {
                OpcError::InvalidContentTypesManifest(
                    "content-types removal selector size overflows".into(),
                )
            })?;
    }
    if selector_bytes > limits.max_content_types_bytes() {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types removal selectors exceed the metadata limit".into(),
        ));
    }
    if parts.is_empty() {
        return Ok(source.clone());
    }
    let mut removals = Vec::new();
    removals
        .try_reserve(parts.len().saturating_mul(2))
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types removal names",
            source,
        })?;
    for part in parts {
        removals.push(part.clone());
        let relationships = part.rels_uri().map_err(OpcError::InvalidPackUri)?;
        removals.push(relationships);
    }
    removals
        .sort_unstable_by(|left, right| cmp_ascii_case_insensitive(left.as_str(), right.as_str()));
    removals.dedup_by(|left, right| left.is_equivalent_to(right));

    // Production callers hold an OwnedContentTypes token whose ContentTypeMap
    // already validated the namespace and attribute vocabulary.
    let mut reader = NsReader::from_reader(source.bytes());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut updates = Vec::new();
    updates
        .try_reserve_exact(removals.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types removal spans",
            source,
        })?;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event()?;
        let end = reader.buffer_position() as usize;
        if end < start || end > source.bytes().len() {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types XML event range is invalid".into(),
            ));
        }
        match event {
            Event::Start(element) => {
                if depth == 1 && element.local_name().as_ref() == b"Override" {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "non-empty content-type Overrides are unsupported for source removal"
                            .into(),
                    ));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest(
                        "content-types XML depth overflows".into(),
                    )
                })?;
            },
            Event::Empty(element) if depth == 1 && element.local_name().as_ref() == b"Override" => {
                let mut part_name = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        OpcError::InvalidContentTypesManifest(error.to_string())
                    })?;
                    if attribute.key.as_ref() == b"PartName" {
                        let value = attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|error| {
                                OpcError::InvalidContentTypesManifest(error.to_string())
                            })?;
                        part_name = Some(
                            PackURI::new(value.into_owned()).map_err(OpcError::InvalidPackUri)?,
                        );
                    }
                }
                if let Some(part_name) = part_name
                    && removals
                        .binary_search_by(|candidate| {
                            cmp_ascii_case_insensitive(candidate.as_str(), part_name.as_str())
                        })
                        .is_ok()
                {
                    updates.push(OwnedElementUpdate {
                        start_tag: start..end,
                        edit: OwnedElementEdit::Remove,
                    });
                }
            },
            Event::Empty(_) => {},
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest(
                        "unmatched content-types XML closing element".into(),
                    )
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(OpcError::InvalidContentTypesManifest(
            "unclosed content-types XML element".into(),
        ));
    }
    if updates.is_empty() {
        return Ok(source.clone());
    }
    source.update_elements(&updates, maximum)
}

pub(crate) fn preflight_part_overrides(
    source: &OwnedXmlPart,
    overrides: &[(&PackURI, &str)],
    max_output_bytes: usize,
) -> Result<(usize, usize)> {
    let limits = ReadLimits::default();
    let maximum = max_output_bytes.min(limits.max_content_types_bytes());
    if source.bytes().len() > maximum {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types source exceeds the requested output limit".into(),
        ));
    }
    if overrides.len() > limits.max_content_type_mappings() {
        return Err(OpcError::InvalidContentTypesManifest(
            "too many content-type override selectors".into(),
        ));
    }
    let mut selector_bytes = 0usize;
    for (part, content_type) in overrides {
        selector_bytes = selector_bytes
            .checked_add(part.as_str().len())
            .and_then(|size| size.checked_add(content_type.len()))
            .and_then(|size| size.checked_add(64))
            .ok_or_else(|| {
                OpcError::InvalidContentTypesManifest("content-type override size overflows".into())
            })?;
        if selector_bytes > maximum {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-type override selectors exceed the metadata limit".into(),
            ));
        }
    }
    Ok((maximum, selector_bytes))
}

/// Append one or more source-preserving `<Override>` children to the OPC
/// `Types` element. The caller supplies already validated part names and
/// content types; this helper preserves the existing manifest bytes and
/// validates the resulting source part before returning it.
pub(crate) fn with_part_overrides(
    source: &OwnedXmlPart,
    overrides: &[(&PackURI, &str)],
    max_output_bytes: usize,
) -> Result<OwnedXmlPart> {
    let (maximum, selector_bytes) = preflight_part_overrides(source, overrides, max_output_bytes)?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(overrides.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types override names",
            source,
        })?;
    names.extend(overrides.iter().map(|(part, _)| (*part).clone()));
    names.sort_unstable_by(|left, right| cmp_ascii_case_insensitive(left.as_str(), right.as_str()));
    if names
        .windows(2)
        .any(|window| window[0].is_equivalent_to(&window[1]))
    {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-type override selectors contain duplicate part names".into(),
        ));
    }
    if overrides.is_empty() {
        return Ok(source.clone());
    }

    let mut reader = NsReader::from_reader(source.bytes());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_tag = None;
    let mut root_name = None;
    let mut root_close = None;
    let mut root_empty = false;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event()?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) if depth == 0 => {
                if element.local_name().as_ref() != b"Types" {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "content-types root must be Types".into(),
                    ));
                }
                root_tag = Some(start..end);
                root_name = Some(element.name().as_ref().to_vec());
                depth = 1;
            },
            Event::Empty(element) if depth == 0 => {
                if element.local_name().as_ref() != b"Types" {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "content-types root must be Types".into(),
                    ));
                }
                root_tag = Some(start..end);
                root_name = Some(element.name().as_ref().to_vec());
                root_empty = true;
                break;
            },
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest(
                        "content-types XML depth overflows".into(),
                    )
                })?;
            },
            Event::End(_) => {
                if depth == 1 {
                    root_close = Some(start..end);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest("unbalanced content-types XML".into())
                })?;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let root_tag = root_tag.ok_or_else(|| {
        OpcError::InvalidContentTypesManifest("content-types Types root is missing".into())
    })?;
    if !root_empty && root_close.is_none() {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types Types root is unclosed".into(),
        ));
    }
    let root_name = root_name.ok_or_else(|| {
        OpcError::InvalidContentTypesManifest("content-types Types root name is missing".into())
    })?;
    let element_name = root_name.iter().position(|byte| *byte == b':').map_or_else(
        || b"Override".to_vec(),
        |colon| {
            let mut name = root_name[..colon].to_vec();
            name.extend_from_slice(b":Override");
            name
        },
    );
    let mut fragment_bound = 0usize;
    for (part, content_type) in overrides {
        let fields = part
            .as_str()
            .len()
            .checked_add(content_type.len())
            .and_then(|size| size.checked_mul(6))
            .and_then(|size| size.checked_add(element_name.len().saturating_add(39)))
            .ok_or_else(|| {
                OpcError::InvalidContentTypesManifest("content-type override size overflows".into())
            })?;
        fragment_bound = fragment_bound.checked_add(fields).ok_or_else(|| {
            OpcError::InvalidContentTypesManifest("content-type override size overflows".into())
        })?;
    }
    if source
        .bytes()
        .len()
        .checked_add(fragment_bound)
        .is_none_or(|size| size > maximum)
    {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-type override output exceeds the metadata limit".into(),
        ));
    }
    let mut fragment = String::new();
    fragment
        .try_reserve_exact(fragment_bound.max(selector_bytes))
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types override XML",
            source,
        })?;
    let element_name = std::str::from_utf8(&element_name).map_err(|_| {
        OpcError::InvalidContentTypesManifest("content-types root prefix is not UTF-8".into())
    })?;
    for (part, content_type) in overrides {
        write!(
            fragment,
            "<{} PartName=\"{}\" ContentType=\"{}\"/>",
            element_name,
            litchi_core::xml::escape_xml(part.as_str()),
            litchi_core::xml::escape_xml(content_type),
        )
        .map_err(|_| {
            OpcError::InvalidContentTypesManifest("content-type override formatting failed".into())
        })?;
    }
    let (range, suffix) = if let Some(close) = root_close.as_ref() {
        (close.start..close.start, Vec::new())
    } else {
        let mut suffix = Vec::new();
        suffix
            .try_reserve(root_name.len().saturating_add(4))
            .map_err(|source| OpcError::Allocation {
                resource: "OPC content-types root close",
                source,
            })?;
        suffix.extend_from_slice(b">");
        suffix.extend_from_slice(b"</");
        suffix.extend_from_slice(&root_name);
        suffix.push(b'>');
        (root_tag.end.saturating_sub(2)..root_tag.end, suffix)
    };
    let size = source
        .bytes()
        .len()
        .checked_sub(range.len())
        .and_then(|size| size.checked_add(fragment.len()))
        .and_then(|size| size.checked_add(suffix.len()))
        .ok_or_else(|| {
            OpcError::InvalidContentTypesManifest("content-type output size overflows".into())
        })?;
    if size > maximum {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-type override output exceeds the metadata limit".into(),
        ));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(size)
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types override output",
            source,
        })?;
    bytes.extend_from_slice(&source.bytes()[..range.start]);
    if root_empty {
        bytes.extend_from_slice(&suffix[..1]);
        bytes.extend_from_slice(fragment.as_bytes());
        bytes.extend_from_slice(&suffix[1..]);
    } else {
        bytes.extend_from_slice(fragment.as_bytes());
    }
    bytes.extend_from_slice(&source.bytes()[range.end..]);
    OwnedXmlPart::capture(
        source.name.clone(),
        source.content_type.clone(),
        std::sync::Arc::new(bytes),
    )
}

fn cmp_ascii_case_insensitive(left: &str, right: &str) -> std::cmp::Ordering {
    for (left, right) in left.bytes().zip(right.bytes()) {
        let ordering = left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase());
        if ordering != std::cmp::Ordering::Equal {
            return ordering;
        }
    }
    left.len().cmp(&right.len())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn removes_part_and_relationship_overrides_without_losing_lexical_context() {
        let bytes = br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<!-- keep -->
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/keep.xml" ContentType="application/xml"/>
<!-- remove target -->
<Override PartName="/word/drop.xml" ContentType="application/xml"/>
<Override PartName="/word/_rels/drop.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>"#;
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(bytes.to_vec()),
        )
        .unwrap();
        let output = without_part_overrides(
            &source,
            &[PackURI::new("/word/drop.xml").unwrap()],
            1024 * 1024,
        )
        .unwrap();
        let output = std::str::from_utf8(output.bytes()).unwrap();
        assert!(output.contains("<!-- keep -->"));
        assert!(output.contains("<!-- remove target -->"));
        assert!(output.contains("/word/keep.xml"));
        assert!(!output.contains("/word/drop.xml\""));
        assert!(!output.contains("drop.xml.rels"));
    }

    #[test]
    fn matches_case_insensitive_part_names_and_refuses_nonempty_overrides() {
        let bytes = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/drop.xml" ContentType="application/xml"/></Types>"#;
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(bytes.to_vec()),
        )
        .unwrap();
        let output =
            without_part_overrides(&source, &[PackURI::new("/WORD/DROP.XML").unwrap()], 1024)
                .unwrap();
        assert!(
            !std::str::from_utf8(output.bytes())
                .unwrap()
                .contains("drop.xml")
        );

        let nonempty = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/drop.xml" ContentType="application/xml"><x/></Override></Types>"#.to_vec(),
            ),
        )
        .unwrap();
        assert!(
            without_part_overrides(&nonempty, &[PackURI::new("/word/drop.xml").unwrap()], 1024,)
                .is_err()
        );
    }

    #[test]
    fn adds_overrides_without_normalizing_unrelated_manifest_bytes() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><!-- retain --><Default Extension="xml" ContentType="application/xml"/></Types>"#.to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/ink/ink1.xml").unwrap();
        let output =
            with_part_overrides(&source, &[(&part, "application/inkml+xml")], 4096).unwrap();
        let output = std::str::from_utf8(output.bytes()).unwrap();
        assert!(output.contains("<!-- retain -->"));
        assert!(output.contains("PartName=\"/word/ink/ink1.xml\""));
        assert!(output.contains("ContentType=\"application/inkml+xml\""));
    }

    #[test]
    fn expands_self_closing_types_and_batches_overrides_as_one_root() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#
                    .to_vec(),
            ),
        )
        .unwrap();
        let first = PackURI::new("/word/ink/ink1.xml").unwrap();
        let second = PackURI::new("/word/media/ink1.png").unwrap();
        let output = with_part_overrides(
            &source,
            &[(&first, "application/inkml+xml"), (&second, "image/png")],
            4096,
        )
        .unwrap();
        let text = std::str::from_utf8(output.bytes()).unwrap();
        assert_eq!(text.matches("<Types").count(), 1);
        assert_eq!(text.matches("</Types>").count(), 1);
        assert_eq!(text.matches("<Override ").count(), 2);
        assert!(text.contains("PartName=\"/word/ink/ink1.xml\""));
        assert!(text.contains("PartName=\"/word/media/ink1.png\""));
    }

    #[test]
    fn self_closing_types_preserve_trailing_lexical_members() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/><!--tail--><?keep?>
"#
                    .to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/ink/ink1.xml").unwrap();
        let output =
            with_part_overrides(&source, &[(&part, "application/inkml+xml")], 4096).unwrap();
        let text = std::str::from_utf8(output.bytes()).unwrap();
        assert!(text.ends_with("<!--tail--><?keep?>\n"));
    }

    #[test]
    fn preserves_a_prefixed_types_namespace_when_batching_overrides() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><!--keep--></ct:Types>"#.to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/ink/ink1.xml").unwrap();
        let output =
            with_part_overrides(&source, &[(&part, "application/inkml+xml")], 4096).unwrap();
        let text = std::str::from_utf8(output.bytes()).unwrap();
        assert!(text.contains("<!--keep-->") && text.contains("<ct:Override "));
        assert_eq!(text.matches("<ct:Types").count(), 1);
    }
}
