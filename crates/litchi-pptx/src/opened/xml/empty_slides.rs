//! Range-preserving expansion of an existing empty slide-ID list.
use super::{Error, Result, clone_xml_name, invalid, is_presentation_namespace, position};
use litchi_core::xml::ReaderOrigin;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

/// The caller already validated the complete slide-list grammar and proved
/// that it contains no entries. Scan that same immutable input for its lexical
/// insertion boundary; never synthesize a missing list or rewrite its wrapper.
pub(super) fn insert(xml: &[u8], id: u32, relationship_id: &str) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut active = false;
    let mut slot = None;
    let mut namespace = None;
    let mut root_strict = None;
    loop {
        let start = position(&reader, origin)?;
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|e| Error::Xml(e.to_string()))?;
        let pml = is_presentation_namespace(&resolved);
        if depth == 0 && matches!(&event, Event::Start(_) | Event::Empty(_)) {
            if !pml
                || !matches!(&event, Event::Start(e) if e.local_name().as_ref() == b"presentation")
            {
                return Err(invalid("empty slide list requires a presentation root"));
            }
            root_strict = Some(
                matches!(&resolved, ResolveResult::Bound(Namespace(value)) if *value == b"http://purl.oclc.org/ooxml/presentationml/main"),
            );
        }
        let list = matches!(&event, Event::Start(e) | Event::Empty(e) if pml && depth == 1 && e.local_name().as_ref() == b"sldIdLst");
        let list_namespace = if list {
            match &resolved {
                ResolveResult::Bound(Namespace(value)) => Some(
                    if value == b"http://purl.oclc.org/ooxml/presentationml/main" {
                        "http://purl.oclc.org/ooxml/presentationml/main"
                    } else {
                        "http://schemas.openxmlformats.org/presentationml/2006/main"
                    },
                ),
                _ => return Err(invalid("slide-ID list has no namespace")),
            }
        } else {
            None
        };
        if let Some(ns) = list_namespace {
            if root_strict != Some(ns == "http://purl.oclc.org/ooxml/presentationml/main") {
                return Err(invalid("slide list dialect differs from presentation root"));
            }
        }
        let event = event.into_owned();
        drop(resolved);
        let end = position(&reader, origin)?;
        match event {
            Event::Start(_) => {
                if list {
                    active = true;
                    namespace = list_namespace;
                }
                depth += 1;
            },
            Event::Empty(element) if list => {
                namespace = list_namespace;
                slot = Some((
                    start..end,
                    Some(clone_xml_name(
                        element.name().as_ref(),
                        "empty slide-list QName",
                    )?),
                ));
            },
            Event::End(_) => {
                if active && depth == 2 {
                    slot = Some((start..start, None));
                    active = false;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("slide list depth underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let (span, empty_name) = slot.ok_or_else(|| invalid("empty slide-ID list is missing"))?;
    let namespace = namespace.ok_or_else(|| invalid("empty slide-ID list namespace is missing"))?;
    let relationships = if namespace == "http://purl.oclc.org/ooxml/presentationml/main" {
        "http://purl.oclc.org/ooxml/officeDocument/relationships"
    } else {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    };
    // Bind fresh prefixes locally, independently of producer prefix choices.
    let entry_capacity = namespace
        .len()
        .checked_add(relationships.len())
        .and_then(|size| size.checked_add(relationship_id.len()))
        .and_then(|size| size.checked_add(96))
        .ok_or_else(|| invalid("empty slide-list entry size overflow"))?;
    let mut entry = String::new();
    entry
        .try_reserve_exact(entry_capacity)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation empty slide entry",
            source,
        })?;
    use std::fmt::Write;
    write!(entry, "<p:sldId xmlns:p=\"{namespace}\" xmlns:r=\"{relationships}\" id=\"{id}\" r:id=\"{relationship_id}\"/>")
        .map_err(|_| invalid("empty slide-list entry formatting failed"))?;
    let extra = entry
        .len()
        .checked_add(empty_name.as_ref().map_or(0, |name| name.len() + 3))
        .ok_or_else(|| invalid("empty slide-list expansion size overflow"))?;
    let capacity = xml
        .len()
        .checked_add(extra)
        .ok_or_else(|| invalid("empty slide-list output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation empty slide list",
            source,
        })?;
    output.extend_from_slice(&xml[..span.start]);
    if let Some(name) = empty_name {
        let opening = &xml[span.clone()];
        let opening = opening
            .strip_suffix(b"/>")
            .ok_or_else(|| invalid("empty slide-list close token is missing"))?;
        output.extend_from_slice(opening);
        output.push(b'>');
        output.extend_from_slice(entry.as_bytes());
        output.extend_from_slice(b"</");
        output.extend_from_slice(&name);
        output.push(b'>');
    } else {
        output.extend_from_slice(entry.as_bytes());
    }
    output.extend_from_slice(&xml[span.end..]);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use crate::Result;
    use crate::opened::xml::{insert_slide_binding, validate_slide_bindings};

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";

    #[test]
    fn first_binding_preserves_empty_list_wrappers_and_surrounding_bytes() -> Result<()> {
        for ns in [P, PS] {
            for (list, closing) in [
                ("<deck:sldIdLst foreign='value' />", "</deck:sldIdLst>"),
                (
                    "<deck:sldIdLst foreign='value'> \r\n<!--keep--><?keep x?></deck:sldIdLst>",
                    "</deck:sldIdLst>",
                ),
                (
                    "<sldIdLst xmlns='PLACEHOLDER' foreign='value'/>",
                    "</sldIdLst>",
                ),
            ] {
                let list = list.replace("PLACEHOLDER", ns);
                let prefix = format!(
                    "\u{feff}<?xml version='1.0'?><deck:presentation xmlns:deck='{ns}' xmlns:p='urn:foreign-p' xmlns:r='urn:foreign-r'><!--before-->"
                );
                let suffix = "<!--after--><deck:extLst/></deck:presentation>";
                let source = format!("{prefix}{list}{suffix}");
                let output =
                    insert_slide_binding(source.as_bytes(), std::iter::empty(), 0, 256, "rId9")?;
                assert!(output.starts_with(prefix.as_bytes()));
                assert!(output.ends_with(suffix.as_bytes()));
                let output_text =
                    std::str::from_utf8(&output).map_err(|e| crate::Error::Xml(e.to_string()))?;
                let start = output_text
                    .find("<p:sldId ")
                    .ok_or_else(|| crate::Error::Invalid("missing inserted binding".into()))?;
                let end = output_text[start..]
                    .find("/>")
                    .ok_or_else(|| crate::Error::Invalid("missing binding end".into()))?
                    + start
                    + 2;
                let without_entry = format!("{}{}", &output_text[..start], &output_text[end..]);
                let expanded = if list.ends_with("/>") {
                    format!("{}>{closing}", &list[..list.len() - 2])
                } else {
                    list.clone()
                };
                assert_eq!(without_entry, format!("{prefix}{expanded}{suffix}"));
                validate_slide_bindings(&output, &[(256, "rId9")], None)?;
                let second =
                    insert_slide_binding(&output, std::iter::once((256, "rId9")), 0, 257, "rId10")?;
                validate_slide_bindings(&second, &[(257, "rId10"), (256, "rId9")], None)?;
            }
        }
        Ok(())
    }

    #[test]
    fn empty_insertion_keeps_strict_existing_list_admission() {
        for body in [
            "",
            "<p:sldIdLst/><p:sldIdLst/>",
            "<p:extLst><p:sldIdLst/></p:extLst>",
            "<p:sldIdLst>text</p:sldIdLst>",
            "<p:sldIdLst><![CDATA[text]]></p:sldIdLst>",
            "<p:sldIdLst><foreign xmlns='urn:foreign'/></p:sldIdLst>",
            "<p:sldIdLst><p:sldId id='256' r:id='rId9'/></p:sldIdLst>",
        ] {
            let xml = format!(
                "<p:presentation xmlns:p='{P}' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>{body}</p:presentation>"
            );
            assert!(
                insert_slide_binding(xml.as_bytes(), std::iter::empty(), 0, 256, "rId9").is_err(),
                "{body}"
            );
        }
        let source = format!("<p:presentation xmlns:p='{P}'><p:sldIdLst/></p:presentation>");
        assert!(
            insert_slide_binding(source.as_bytes(), std::iter::empty(), 1, 256, "rId9").is_err()
        );
        assert!(
            insert_slide_binding(
                source.as_bytes(),
                std::iter::once((256, "rId9")),
                0,
                257,
                "rId10"
            )
            .is_err()
        );
    }

    #[test]
    fn empty_insertion_refuses_mixed_dialects_and_invalid_new_ids() {
        for (root, list) in [(P, PS), (PS, P)] {
            let xml = format!(
                "<p:presentation xmlns:p='{root}'><sldIdLst xmlns='{list}'/></p:presentation>"
            );
            assert!(
                insert_slide_binding(xml.as_bytes(), std::iter::empty(), 0, 256, "rId9").is_err()
            );
        }
        let xml = format!("<p:presentation xmlns:p='{P}'><p:sldIdLst/></p:presentation>");
        for id in ["", "9bad", "bad\"id", "bad\nid", "bad&value"] {
            assert!(insert_slide_binding(xml.as_bytes(), std::iter::empty(), 0, 256, id).is_err());
        }
        let huge = "r".repeat(crate::media_parts::MAX_STRING_BYTES + 1);
        assert!(matches!(
            insert_slide_binding(xml.as_bytes(), std::iter::empty(), 0, 256, &huge),
            Err(crate::Error::Limit { .. })
        ));
    }

    #[test]
    fn empty_insertion_retains_the_bounded_scan_before_allocation() {
        let source = format!(
            "<p:presentation xmlns:p='{P}'><p:sldIdLst/>{}{}</p:presentation>",
            "<foreign>".repeat(256),
            "</foreign>".repeat(256)
        );
        assert!(matches!(
            insert_slide_binding(source.as_bytes(), std::iter::empty(), 0, 256, "rId9"),
            Err(crate::Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                observed: 257,
                maximum: 256
            })
        ));
    }
}
