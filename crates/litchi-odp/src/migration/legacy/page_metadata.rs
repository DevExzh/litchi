//! Static ODF presentation page identifiers and layout references.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_PAGES: usize = 65_536;
const MAX_NAVIGATION_IDS: usize = 65_536;

/// Static attributes of one `draw:page` in slide order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageMetadata {
    /// Zero-based slide index.
    pub slide_index: usize,
    pub name: Option<String>,
    pub style_name: Option<String>,
    pub master_page_name: Option<String>,
    pub page_layout_name: Option<String>,
    /// Legacy ODF `draw:id`.
    pub draw_id: Option<String>,
    /// XML `xml:id`, which supersedes `draw:id` when both are present.
    pub xml_id: Option<String>,
    /// Inert page link. The library never resolves or fetches it.
    pub href: Option<String>,
    /// Ordered shape IDs from `draw:nav-order`.
    pub navigation_order: Vec<String>,
}

impl PageMetadata {
    /// Create empty metadata for a zero-based slide index.
    pub fn new(slide_index: usize) -> Self {
        Self {
            slide_index,
            name: None,
            style_name: None,
            master_page_name: None,
            page_layout_name: None,
            draw_id: None,
            xml_id: None,
            href: None,
            navigation_order: Vec::new(),
        }
    }

    fn validate(&self) -> Result<()> {
        for (value, description) in [
            (self.name.as_deref(), "draw:name"),
            (self.style_name.as_deref(), "draw:style-name"),
            (self.master_page_name.as_deref(), "draw:master-page-name"),
            (
                self.page_layout_name.as_deref(),
                "presentation:presentation-page-layout-name",
            ),
            (self.href.as_deref(), "xlink:href"),
        ] {
            if let Some(value) = value {
                validate_text(value, description, false)?;
            }
        }
        for (value, description) in [
            (self.draw_id.as_deref(), "draw:id"),
            (self.xml_id.as_deref(), "xml:id"),
        ] {
            if let Some(value) = value {
                validate_ncname(value, description)?;
            }
        }
        if let (Some(draw_id), Some(xml_id)) = (&self.draw_id, &self.xml_id)
            && draw_id != xml_id
        {
            return Err(invalid(
                "draw:id and xml:id on the same presentation page must match",
            ));
        }
        if self.navigation_order.len() > MAX_NAVIGATION_IDS {
            return Err(invalid("draw:nav-order exceeds 65536 shape IDs"));
        }
        let mut navigation_ids = HashSet::with_capacity(self.navigation_order.len());
        for id in &self.navigation_order {
            validate_ncname(id, "draw:nav-order item")?;
            if !navigation_ids.insert(id.as_str()) {
                return Err(invalid(format!("duplicate draw:nav-order ID '{id}'")));
            }
        }
        Ok(())
    }
}

/// Ordered static page metadata for an ODP presentation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PageMetadataCollection {
    pages: Vec<PageMetadata>,
}

impl PageMetadataCollection {
    /// Create and validate an ordered page metadata collection.
    pub fn new(pages: Vec<PageMetadata>) -> Result<Self> {
        let value = Self { pages };
        value.validate()?;
        Ok(value)
    }

    /// Return pages in slide order.
    pub fn pages(&self) -> &[PageMetadata] {
        &self.pages
    }

    /// Return metadata for a zero-based slide index.
    pub fn page(&self, slide_index: usize) -> Option<&PageMetadata> {
        self.pages
            .binary_search_by_key(&slide_index, |value| value.slide_index)
            .ok()
            .map(|index| &self.pages[index])
    }

    /// Validate ordering, identifiers, and per-page values.
    pub fn validate(&self) -> Result<()> {
        self.validate_for_slide_count(None)
    }

    /// Validate all page indices against a concrete slide count.
    pub fn validate_for_slides(&self, slide_count: usize) -> Result<()> {
        self.validate_for_slide_count(Some(slide_count))
    }

    fn validate_for_slide_count(&self, slide_count: Option<usize>) -> Result<()> {
        if self.pages.len() > MAX_PAGES {
            return Err(invalid("presentation page metadata exceeds 65536 pages"));
        }
        let mut previous = None;
        let mut ids = HashSet::new();
        for page in &self.pages {
            page.validate()?;
            if previous.is_some_and(|index| page.slide_index <= index) {
                return Err(invalid(
                    "presentation page metadata must use strictly increasing slide indices",
                ));
            }
            previous = Some(page.slide_index);
            if let Some(count) = slide_count
                && page.slide_index >= count
            {
                return Err(invalid(format!(
                    "presentation page metadata index {} exceeds slide count {count}",
                    page.slide_index
                )));
            }
            if let Some(id) = page.xml_id.as_deref().or(page.draw_id.as_deref())
                && !ids.insert(id)
            {
                return Err(invalid(format!("duplicate presentation page ID '{id}'")));
            }
        }
        Ok(())
    }

    /// Return whether no page metadata is retained.
    pub fn is_empty(&self) -> bool {
        self.pages.is_empty()
    }
}

/// Return the exact page names emitted for the current slide sequence.
pub(crate) fn effective_page_names(
    metadata: Option<&PageMetadataCollection>,
    slide_count: usize,
) -> Result<Vec<String>> {
    if slide_count > MAX_PAGES {
        return Err(invalid("presentation exceeds 65536 pages"));
    }
    if let Some(metadata) = metadata {
        metadata.validate_for_slides(slide_count)?;
    }
    (0..slide_count)
        .map(|index| {
            metadata
                .and_then(|value| value.page(index))
                .and_then(|value| value.name.clone())
                .map(Ok)
                .unwrap_or_else(|| fallback_page_name(index))
        })
        .collect()
}

/// Materialize stable names and insert an empty metadata record for a new page.
pub(crate) fn metadata_after_page_insert(
    metadata: Option<&PageMetadataCollection>,
    slide_count: usize,
    insert_index: usize,
) -> Result<PageMetadataCollection> {
    if insert_index > slide_count {
        return Err(invalid(
            "presentation page insertion index is out of bounds",
        ));
    }
    let new_count = slide_count
        .checked_add(1)
        .ok_or_else(|| invalid("presentation page count overflow"))?;
    if new_count > MAX_PAGES {
        return Err(invalid("presentation exceeds 65536 pages"));
    }
    let old_names = effective_page_names(metadata, slide_count)?;
    let used_names = old_names.iter().map(String::as_str).collect::<HashSet<_>>();
    let mut candidate = new_count;
    let new_name = loop {
        let name = format!("page{candidate}");
        if !used_names.contains(name.as_str()) {
            break name;
        }
        candidate = candidate
            .checked_add(1)
            .ok_or_else(|| invalid("presentation page name counter overflow"))?;
    };

    let mut pages = Vec::with_capacity(new_count);
    for (old_index, old_name) in old_names.iter().enumerate() {
        let new_index = if old_index >= insert_index {
            old_index
                .checked_add(1)
                .ok_or_else(|| invalid("presentation page index overflow"))?
        } else {
            old_index
        };
        let mut page = metadata
            .and_then(|value| value.page(old_index))
            .cloned()
            .unwrap_or_else(|| PageMetadata::new(new_index));
        page.slide_index = new_index;
        page.name.get_or_insert_with(|| old_name.clone());
        pages.push(page);
    }
    let mut inserted = PageMetadata::new(insert_index);
    inserted.name = Some(new_name);
    pages.push(inserted);
    pages.sort_by_key(|page| page.slide_index);
    PageMetadataCollection::new(pages)
}

/// Materialize stable names and remove one page metadata record.
pub(crate) fn metadata_after_page_remove(
    metadata: Option<&PageMetadataCollection>,
    slide_count: usize,
    remove_index: usize,
) -> Result<Option<PageMetadataCollection>> {
    if remove_index >= slide_count {
        return Err(invalid("presentation page removal index is out of bounds"));
    }
    let old_names = effective_page_names(metadata, slide_count)?;
    let mut pages = Vec::with_capacity(slide_count.saturating_sub(1));
    for (old_index, old_name) in old_names.iter().enumerate() {
        if old_index == remove_index {
            continue;
        }
        let new_index = if old_index > remove_index {
            old_index
                .checked_sub(1)
                .ok_or_else(|| invalid("presentation page index underflow"))?
        } else {
            old_index
        };
        let mut page = metadata
            .and_then(|value| value.page(old_index))
            .cloned()
            .unwrap_or_else(|| PageMetadata::new(new_index));
        page.slide_index = new_index;
        page.name.get_or_insert_with(|| old_name.clone());
        pages.push(page);
    }
    if pages.is_empty() {
        Ok(None)
    } else {
        PageMetadataCollection::new(pages).map(Some)
    }
}

fn fallback_page_name(slide_index: usize) -> Result<String> {
    let ordinal = slide_index
        .checked_add(1)
        .ok_or_else(|| invalid("presentation page name index overflow"))?;
    Ok(format!("page{ordinal}"))
}

/// Parse static metadata from direct `draw:page` children.
pub fn parse_page_metadata(xml: &str) -> Result<PageMetadataCollection> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("presentation page metadata XML exceeds 8 MiB"));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut presentation_depth = None;
    let mut found_presentation = false;
    let mut pages = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML nesting overflow"))?;
                if element_is(&reader, &element, OFFICE_NAMESPACE, b"presentation") {
                    if found_presentation {
                        return Err(invalid("duplicate office:presentation element"));
                    }
                    found_presentation = true;
                    presentation_depth = Some(depth);
                } else if element_is(&reader, &element, DRAW_NAMESPACE, b"page")
                    && presentation_depth == Some(depth - 1)
                {
                    if pages.len() >= MAX_PAGES {
                        return Err(invalid("presentation page metadata exceeds 65536 pages"));
                    }
                    pages.push(parse_page(&reader, &element, pages.len())?);
                }
            },
            Event::Empty(element)
                if element_is(&reader, &element, DRAW_NAMESPACE, b"page")
                    && presentation_depth == Some(depth) =>
            {
                if pages.len() >= MAX_PAGES {
                    return Err(invalid("presentation page metadata exceeds 65536 pages"));
                }
                pages.push(parse_page(&reader, &element, pages.len())?);
            },
            Event::End(element) => {
                if presentation_depth == Some(depth)
                    && end_is(&reader, &element, OFFICE_NAMESPACE, b"presentation")
                {
                    presentation_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced XML end element"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("active XML declarations are prohibited"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    PageMetadataCollection::new(pages)
}

pub(crate) fn write_page_attributes(
    metadata: Option<&PageMetadataCollection>,
    slide_index: usize,
    fallback_style_name: &str,
) -> Result<String> {
    let page = metadata.and_then(|value| value.page(slide_index));
    let mut output = String::with_capacity(128);
    let fallback_name = format!("page{}", slide_index + 1);
    write_attribute(
        &mut output,
        "draw:name",
        page.and_then(|value| value.name.as_deref())
            .unwrap_or(&fallback_name),
    );
    write_attribute(
        &mut output,
        "draw:style-name",
        page.and_then(|value| value.style_name.as_deref())
            .unwrap_or(fallback_style_name),
    );
    write_attribute(
        &mut output,
        "draw:master-page-name",
        page.and_then(|value| value.master_page_name.as_deref())
            .unwrap_or("Default"),
    );
    if let Some(page) = page {
        if let Some(value) = &page.page_layout_name {
            write_attribute(
                &mut output,
                "presentation:presentation-page-layout-name",
                value,
            );
        }
        if let Some(value) = &page.draw_id {
            write_attribute(&mut output, "draw:id", value);
        }
        if let Some(value) = &page.xml_id {
            write_attribute(&mut output, "xml:id", value);
        }
        if let Some(value) = &page.href {
            write_attribute(&mut output, "xlink:href", value);
        }
        if !page.navigation_order.is_empty() {
            write_attribute(
                &mut output,
                "draw:nav-order",
                &page.navigation_order.join(" "),
            );
        }
    }
    Ok(output)
}

fn parse_page(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    slide_index: usize,
) -> Result<PageMetadata> {
    let mut page = PageMetadata::new(slide_index);
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let slot = match (namespace, local.as_ref()) {
            (ResolveResult::Bound(found), b"name") if found == Namespace(DRAW_NAMESPACE) => {
                &mut page.name
            },
            (ResolveResult::Bound(found), b"style-name") if found == Namespace(DRAW_NAMESPACE) => {
                &mut page.style_name
            },
            (ResolveResult::Bound(found), b"master-page-name")
                if found == Namespace(DRAW_NAMESPACE) =>
            {
                &mut page.master_page_name
            },
            (ResolveResult::Bound(found), b"presentation-page-layout-name")
                if found == Namespace(PRESENTATION_NAMESPACE) =>
            {
                &mut page.page_layout_name
            },
            (ResolveResult::Bound(found), b"id") if found == Namespace(DRAW_NAMESPACE) => {
                &mut page.draw_id
            },
            (ResolveResult::Bound(found), b"id") if found == Namespace(XML_NAMESPACE) => {
                &mut page.xml_id
            },
            (ResolveResult::Bound(found), b"href") if found == Namespace(XLINK_NAMESPACE) => {
                &mut page.href
            },
            (ResolveResult::Bound(found), b"nav-order") if found == Namespace(DRAW_NAMESPACE) => {
                if !page.navigation_order.is_empty() {
                    return Err(invalid("duplicate draw:nav-order attribute"));
                }
                let value = decode_attribute(reader, &attribute)?;
                page.navigation_order = value.split_whitespace().map(str::to_string).collect();
                if page.navigation_order.is_empty() {
                    return Err(invalid("draw:nav-order cannot be empty"));
                }
                continue;
            },
            _ => continue,
        };
        if slot.is_some() {
            return Err(invalid("duplicate presentation page metadata attribute"));
        }
        *slot = Some(decode_attribute(reader, &attribute)?);
    }
    page.validate()?;
    Ok(page)
}

fn decode_attribute(
    reader: &NsReader<&[u8]>,
    attribute: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attribute
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
        .map_err(xml_error)
        .map(|value| value.into_owned())
}

fn validate_ncname(value: &str, description: &str) -> Result<()> {
    validate_text(value, description, false)?;
    let mut characters = value.chars();
    let Some(first) = characters.next() else {
        return Err(invalid(format!("{description} cannot be empty")));
    };
    if !(first == '_' || first.is_alphabetic())
        || characters.any(|character| {
            !(character == '_'
                || character == '-'
                || character == '.'
                || character.is_alphanumeric())
        })
    {
        return Err(invalid(format!("{description} is not an XML NCName")));
    }
    Ok(())
}

fn validate_text(value: &str, description: &str, allow_empty: bool) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("{description} exceeds 1 MiB")));
    }
    if !allow_empty && value.is_empty() {
        return Err(invalid(format!("{description} cannot be empty")));
    }
    if value.chars().any(|character| {
        character == '\0' || (character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Err(invalid(format!(
            "{description} contains invalid XML characters"
        )));
    }
    Ok(())
}

fn write_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

fn element_is(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn end_is(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesEnd<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!(
        "presentation page metadata XML parsing error: {error}"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{MutablePresentation, Presentation, Builder};

    const PREFIX: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:presentation>"#;
    const SUFFIX: &str = "</o:presentation></o:body></o:document-content>";

    fn metadata() -> PageMetadataCollection {
        PageMetadataCollection::new(vec![PageMetadata {
            slide_index: 0,
            name: Some("Quarterly Review".to_string()),
            style_name: Some("dp1".to_string()),
            master_page_name: Some("Corporate Master".to_string()),
            page_layout_name: Some("TitleAndBody".to_string()),
            draw_id: Some("page-id-1".to_string()),
            xml_id: Some("page-id-1".to_string()),
            href: Some("../Templates/master.odp#page1".to_string()),
            navigation_order: vec!["title1".to_string(), "body1".to_string()],
        }])
        .unwrap()
    }

    #[test]
    fn parses_complete_page_metadata_without_resolving_references() {
        let xml = format!(
            r#"{PREFIX}<d:page d:name="Quarterly Review" d:style-name="dp1" d:master-page-name="Corporate Master" p:presentation-page-layout-name="TitleAndBody" d:id="page-id-1" xml:id="page-id-1" x:href="../Templates/master.odp#page1" d:nav-order="title1 body1"/>{SUFFIX}"#
        );
        assert_eq!(parse_page_metadata(&xml).unwrap(), metadata());
    }

    #[test]
    fn builder_and_mutable_round_trip_page_metadata() {
        let metadata = metadata();
        let mut builder = Builder::new();
        builder.add_slide_with_title("Title", "Body").unwrap();
        builder.set_page_metadata(Some(metadata.clone())).unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(presentation.page_metadata().unwrap(), metadata);
        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        assert_eq!(mutable.page_metadata(), Some(&metadata));
        let reparsed = Presentation::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.page_metadata().unwrap(), metadata);
    }

    #[test]
    fn rejects_duplicate_ids_bad_navigation_and_active_xml() {
        for body in [
            r#"<d:page d:id="same"/><d:page xml:id="same"/>"#,
            r#"<d:page d:id="a" xml:id="b"/>"#,
            r#"<d:page d:nav-order="a a"/>"#,
            r#"<d:page d:nav-order="bad:id"/>"#,
            r#"<d:page d:name=""/>"#,
        ] {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(
                parse_page_metadata(&xml).is_err(),
                "accepted {xml}"
            );
        }
        let active = format!("{PREFIX}<!DOCTYPE x><d:page/>{SUFFIX}");
        assert!(parse_page_metadata(&active).is_err());
    }
}
