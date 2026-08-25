#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
use crate::error::{Error, Result};
use litchi_ooxml_common::xml_name::is_ncname;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use std::io::BufRead;

use super::model::{
    MAX_PAGE_BORDER_ART_SIZE, MAX_PAGE_BORDER_LINE_SIZE, MAX_PAGE_BORDER_SPACE, SectionProperties,
    Style,
};

pub(super) fn validate_header_footer_xml(xml: &str, header: bool) -> Result<()> {
    use quick_xml::reader::NsReader;
    let mut reader = NsReader::from_str(xml);
    let mut depth = 0usize;
    let mut root = false;
    let mut closed_root = false;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    loop {
        let (wordprocessing_namespace, event_namespace_unknown, event_namespace, event) = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            (
                crate::namespace::is_wordprocessing_namespace(&namespace),
                matches!(&namespace, ResolveResult::Unknown(_)),
                namespace_key(&namespace),
                event,
            )
        };
        match event {
            Event::Start(element) => {
                if event_namespace_unknown {
                    return Err(Error::InvalidFormat(
                        "header/footer XML uses an undeclared element prefix".to_string(),
                    ));
                }
                validate_attributes(&reader, &element)?;
                if depth == 0 {
                    let expected = if header {
                        b"hdr".as_slice()
                    } else {
                        b"ftr".as_slice()
                    };
                    if root
                        || closed_root
                        || !wordprocessing_namespace
                        || element.local_name().as_ref() != expected
                    {
                        return Err(Error::InvalidFormat(
                            "section header/footer XML has an invalid root".to_string(),
                        ));
                    }
                    root = true;
                }
                stack.push((element.name().as_ref().to_vec(), event_namespace.clone()));
                depth += 1;
            },
            Event::Empty(element) if depth == 0 => {
                if event_namespace_unknown {
                    return Err(Error::InvalidFormat(
                        "header/footer XML uses an undeclared element prefix".to_string(),
                    ));
                }
                validate_attributes(&reader, &element)?;
                let expected = if header {
                    b"hdr".as_slice()
                } else {
                    b"ftr".as_slice()
                };
                if root
                    || closed_root
                    || !wordprocessing_namespace
                    || element.local_name().as_ref() != expected
                {
                    return Err(Error::InvalidFormat(
                        "section header/footer XML has an invalid root".to_string(),
                    ));
                }
                root = true;
                closed_root = true;
            },
            Event::Empty(element) if depth > 0 => {
                if event_namespace_unknown {
                    return Err(Error::InvalidFormat(
                        "header/footer XML uses an undeclared element prefix".to_string(),
                    ));
                }
                validate_attributes(&reader, &element)?;
            },
            Event::End(element) => {
                if event_namespace_unknown {
                    return Err(Error::InvalidFormat(
                        "header/footer XML uses an undeclared element prefix".to_string(),
                    ));
                }
                let expected = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("invalid header/footer XML nesting".to_string())
                })?;
                if expected.0 != element.name().as_ref() || expected.1 != event_namespace {
                    return Err(Error::InvalidFormat(
                        "header/footer XML has mismatched end elements".to_string(),
                    ));
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid header/footer XML nesting".to_string())
                })?;
                if depth == 0 {
                    closed_root = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) if depth == 0 => {
                if root {
                    return Err(Error::InvalidFormat(
                        "header/footer XML has trailing markup".to_string(),
                    ));
                }
            },
            Event::Text(text) if depth == 0 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "header/footer XML has trailing non-whitespace text".to_string(),
                    ));
                }
            },
            Event::CData(_) | Event::DocType(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "header/footer XML has invalid content outside its root".to_string(),
                ));
            },
            Event::Text(_) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::CData(_) | Event::DocType(_) | Event::GeneralRef(_) | Event::Decl(_) => {
                return Err(Error::InvalidFormat(
                    "header/footer XML has unsupported content inside its root".to_string(),
                ));
            },
            Event::Empty(_) => {},
        }
    }
    if !root || !closed_root || depth != 0 || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "unterminated section header/footer XML".to_string(),
        ));
    }
    Ok(())
}

fn validate_attributes<R: BufRead>(
    reader: &quick_xml::reader::NsReader<R>,
    element: &BytesStart<'_>,
) -> Result<()> {
    let resolver = reader.resolver().clone();
    let mut seen = std::collections::HashSet::new();
    let mut seen_raw = std::collections::HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if !seen_raw.insert(attribute.key.as_ref().to_vec()) {
            return Err(Error::InvalidFormat(
                "header/footer XML contains duplicate attributes".to_string(),
            ));
        }
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if matches!(&namespace, ResolveResult::Unknown(_)) {
            return Err(Error::InvalidFormat(
                "header/footer XML uses an undeclared attribute prefix".to_string(),
            ));
        }
        let key = (
            namespace_key(&namespace),
            attribute.key.local_name().as_ref().to_vec(),
        );
        if !seen.insert(key) {
            return Err(Error::InvalidFormat(
                "header/footer XML contains duplicate attributes".to_string(),
            ));
        }
    }
    Ok(())
}

fn namespace_key(namespace: &ResolveResult<'_>) -> Vec<u8> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => value.to_vec(),
        ResolveResult::Unknown(prefix) => {
            let mut key = b"?".to_vec();
            key.extend_from_slice(prefix);
            key
        },
        ResolveResult::Unbound => Vec::new(),
    }
}

impl SectionProperties {
    pub(crate) fn validate(&self) -> Result<()> {
        if self.page_width == 0 || self.page_height == 0 {
            return Err(Error::InvalidFormat(
                "section page dimensions must be nonzero".to_string(),
            ));
        }
        let mut owned_part_xml = std::collections::HashMap::new();
        for references in [&self.headers, &self.footers] {
            for reference in references {
                if let Some(part) = &reference.part
                    && let Some(previous) =
                        owned_part_xml.insert(part.key.clone(), part.xml.clone())
                    && previous != part.xml
                {
                    return Err(Error::InvalidFormat(
                        "section header/footer references reuse a key with different XML"
                            .to_string(),
                    ));
                }
            }
        }
        for references in [&self.headers, &self.footers] {
            let mut kinds = std::collections::HashSet::new();
            for reference in references {
                if !kinds.insert(reference.kind) {
                    return Err(Error::InvalidFormat(
                        "section has duplicate header/footer reference type".to_string(),
                    ));
                }
                if reference
                    .relationship_id
                    .as_deref()
                    .is_some_and(|id| !is_ncname(id))
                {
                    return Err(Error::InvalidFormat(
                        "section header/footer relationship ID is not an XML NCName".to_string(),
                    ));
                }
                if reference.relationship_id.is_some() && reference.part.is_some() {
                    return Err(Error::InvalidFormat(
                        "section header/footer cannot be both existing and owned".to_string(),
                    ));
                }
                if let Some(part) = &reference.part
                    && (part.key.is_empty() || part.xml.is_empty())
                {
                    return Err(Error::InvalidFormat(
                        "section header/footer part key and XML must be non-empty".to_string(),
                    ));
                }
                if let Some(part) = &reference.part {
                    validate_header_footer_xml(
                        &part.xml,
                        std::ptr::eq(references, &raw const self.headers),
                    )?;
                }
            }
        }
        if let Some(columns) = &self.columns {
            if columns.count == 0 || columns.count > 45 {
                return Err(Error::InvalidFormat(
                    "section column count must be in 1..=45".to_string(),
                ));
            }
            if !columns.equal_width && usize::from(columns.count) != columns.columns.len() {
                return Err(Error::InvalidFormat(
                    "unequal section columns require one width per column".to_string(),
                ));
            }
        }
        if let Some(borders) = &self.page_borders {
            for border in [&borders.top, &borders.left, &borders.bottom, &borders.right]
                .into_iter()
                .flatten()
            {
                if let Some(size) = border.size {
                    let max = match border.style {
                        Style::Art(_) => MAX_PAGE_BORDER_ART_SIZE,
                        Style::Nil
                        | Style::None
                        | Style::Single
                        | Style::Thick
                        | Style::Double
                        | Style::Dotted
                        | Style::Dashed
                        | Style::DotDash
                        | Style::DotDotDash
                        | Style::Triple
                        | Style::ThinThickSmallGap
                        | Style::ThinThickMediumGap
                        | Style::ThinThickLargeGap
                        | Style::ThickThinSmallGap
                        | Style::ThickThinMediumGap
                        | Style::ThickThinLargeGap
                        | Style::ThinThickThinSmallGap
                        | Style::ThinThickThinMediumGap
                        | Style::ThinThickThinLargeGap
                        | Style::Wave
                        | Style::DoubleWave
                        | Style::DashSmallGap
                        | Style::DashDotStroked
                        | Style::ThreeDEmboss
                        | Style::ThreeDEngrave
                        | Style::Outset
                        | Style::Inset => MAX_PAGE_BORDER_LINE_SIZE,
                    };
                    if size > max {
                        return Err(Error::InvalidFormat(format!(
                            "page border size {size} exceeds the {max} limit"
                        )));
                    }
                }
                if let Some(space) = border.space
                    && space > MAX_PAGE_BORDER_SPACE
                {
                    return Err(Error::InvalidFormat(format!(
                        "page border space {space} exceeds the {MAX_PAGE_BORDER_SPACE} limit"
                    )));
                }
            }
        }
        if self
            .printer_settings_relationship_id
            .as_deref()
            .is_some_and(|id| !is_ncname(id))
        {
            return Err(Error::InvalidFormat(
                "section printer-settings relationship ID is not an XML NCName".to_string(),
            ));
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::super::model::{SectionHeaderFooterPart, SectionHeaderFooterReference};
    use super::SectionProperties;
    use crate::header_footer::Kind;

    #[test]
    fn direct_relationship_ids_must_be_ncname_values() {
        let mut header = SectionProperties::default();
        header.headers.push(SectionHeaderFooterReference {
            kind: Kind::Primary,
            relationship_id: Some("not valid".to_owned()),
            part: None,
        });
        assert!(header.validate().is_err());

        let mut printer = SectionProperties::default();
        printer.printer_settings_relationship_id = Some("not valid".to_owned());
        assert!(printer.validate().is_err());
    }

    #[test]
    fn header_footer_xml_requires_one_closed_matching_root() {
        let mut section = SectionProperties::default();
        section.headers.push(SectionHeaderFooterReference {
            kind: Kind::Primary,
            relationship_id: None,
            part: Some(SectionHeaderFooterPart {
                key: "header".to_owned(),
                xml: "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p></w:r></w:hdr>"
                    .to_owned(),
            }),
        });
        assert!(section.validate().is_err());
    }
}
