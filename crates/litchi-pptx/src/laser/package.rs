//! OPC graph ownership for persisted laser-trace extensions.

use super::{Conformance, Limits, Trace, TracePoint, read_with, validate, write};
use crate::{Error, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::Part;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;
const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";

/// Load bounded traces from one `PresentationML` slide part.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_slide_traces(
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut Limits,
) -> Result<Vec<Trace>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_owned(),
            actual: slide.content_type().to_owned(),
        });
    }
    read_with(slide_index, slide.blob(), limits)
}

/// Store one trace in a slide's extension list, preserving its dialect.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn store_slide_trace(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    points: &[TracePoint],
) -> Result<()> {
    validate(points)?;
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_owned(),
            actual: slide.content_type().to_owned(),
        });
    }
    if !load_slide_traces(0, slide, &mut Limits::default())?.is_empty() {
        return Err(Error::Invalid(
            "slide already contains a laser-trace extension; replacement is not supported"
                .to_owned(),
        ));
    }

    let conformance = Conformance::from_namespace(slide_dialect(slide.blob())?);
    let fragment = write(points, conformance)?;
    let updated = insert_extension_fragment(slide.blob(), &fragment)?;
    let probe = BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), updated.clone());
    let traces = load_slide_traces(0, &probe, &mut Limits::default())?;
    if traces.len() != 1 || traces[0].points().len() != points.len() {
        return Err(Error::Invalid(
            "laser-trace storage failed read-back validation".to_owned(),
        ));
    }
    package.get_part_mut(slide_name)?.set_blob(updated);
    Ok(())
}

fn slide_dialect(xml: &[u8]) -> Result<&'static str> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes", MAX_SLIDE_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !is_presentationml_name(&namespace, element.name(), b"sld") {
                    return Err(Error::Invalid(
                        "slide XML must have a PresentationML sld root element".to_owned(),
                    ));
                }
                return Ok(
                    if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STRICT_PML)
                    {
                        "http://purl.oclc.org/ooxml/presentationml/main"
                    } else {
                        "http://schemas.openxmlformats.org/presentationml/2006/main"
                    },
                );
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "DTDs and processing instructions are rejected".to_owned(),
                ));
            },
            Event::Eof => return Err(Error::Invalid("slide XML has no root element".to_owned())),
            _ => {},
        }
    }
}

fn insert_extension_fragment(xml: &[u8], fragment: &str) -> Result<Vec<u8>> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes", MAX_SLIDE_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_end = None;
    let mut ext_end = None;
    let mut empty_ext = None;

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_err| Error::Invalid("slide XML offset overflow".to_owned()))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                add_node(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth", MAX_XML_DEPTH));
                }
                if depth == 1 {
                    if root_seen || !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(Error::Invalid(
                            "slide XML must have one PresentationML sld root element".to_owned(),
                        ));
                    }
                    root_seen = true;
                }
                if depth == 2 && is_presentationml_name(&namespace, element.name(), b"extLst") {
                    if ext_end.is_some() || empty_ext.is_some() {
                        return Err(Error::Invalid(
                            "slide has multiple extension lists".to_owned(),
                        ));
                    }
                    ext_end = Some(usize::MAX);
                }
            },
            Event::Empty(element) => {
                add_node(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    if root_seen || !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(Error::Invalid(
                            "slide XML must have one PresentationML sld root element".to_owned(),
                        ));
                    }
                    root_seen = true;
                }
                if child_depth == 2 && is_presentationml_name(&namespace, element.name(), b"extLst")
                {
                    if ext_end.is_some() || empty_ext.is_some() {
                        return Err(Error::Invalid(
                            "slide has multiple extension lists".to_owned(),
                        ));
                    }
                    empty_ext = Some((
                        start,
                        usize::try_from(reader.buffer_position()).map_err(|_err| {
                            Error::Invalid("slide XML offset overflow".to_owned())
                        })?,
                    ));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(Error::Invalid("invalid slide XML nesting".to_owned()));
                }
                if depth == 2
                    && ext_end == Some(usize::MAX)
                    && is_presentationml_name(&namespace, element.name(), b"extLst")
                {
                    ext_end = Some(start);
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(Error::Invalid(
                            "slide root closing element is invalid".to_owned(),
                        ));
                    }
                    root_end = Some(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "DTDs and processing instructions are rejected".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen {
        return Err(Error::Invalid(
            "unterminated or missing PresentationML slide root".to_owned(),
        ));
    }
    let root_end =
        root_end.ok_or_else(|| Error::Invalid("slide is missing its end tag".to_owned()))?;

    if let Some((start, end)) = empty_ext {
        let element = xml[start..end].trim_ascii_end();
        let open = element
            .strip_suffix(b"/>")
            .ok_or_else(|| Error::Invalid("slide extension list is not empty".to_owned()))?;
        let removed = end
            .checked_sub(start)
            .ok_or_else(|| Error::Invalid("slide XML offsets are out of order".to_owned()))?;
        let replacement = open
            .len()
            .checked_add(1)
            .and_then(|size| size.checked_add(fragment.len()))
            .and_then(|size| size.checked_add(b"</p:extLst>".len()))
            .ok_or_else(|| limit("updated slide XML bytes", MAX_SLIDE_XML_BYTES))?;
        let size = checked_output_size(
            xml.len()
                .checked_sub(removed)
                .and_then(|size| size.checked_add(replacement)),
        )?;
        let mut output = Vec::with_capacity(size);
        output.extend_from_slice(&xml[..start]);
        output.extend_from_slice(open);
        output.extend_from_slice(b">");
        output.extend_from_slice(fragment.as_bytes());
        output.extend_from_slice(b"</p:extLst>");
        output.extend_from_slice(&xml[end..]);
        return Ok(output);
    }
    if let Some(position) = ext_end.filter(|position| *position != usize::MAX) {
        let size = checked_output_size(xml.len().checked_add(fragment.len()))?;
        let mut output = Vec::with_capacity(size);
        output.extend_from_slice(&xml[..position]);
        output.extend_from_slice(fragment.as_bytes());
        output.extend_from_slice(&xml[position..]);
        return Ok(output);
    }

    let dialect = slide_dialect(xml)?;
    let size = checked_output_size(
        xml.len()
            .checked_add(b"<p:extLst xmlns:p=\"".len())
            .and_then(|size| size.checked_add(dialect.len()))
            .and_then(|size| size.checked_add(b"\">".len()))
            .and_then(|size| size.checked_add(fragment.len()))
            .and_then(|size| size.checked_add(b"</p:extLst>".len())),
    )?;
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..root_end]);
    output.extend_from_slice(b"<p:extLst xmlns:p=\"");
    output.extend_from_slice(dialect.as_bytes());
    output.extend_from_slice(b"\">");
    output.extend_from_slice(fragment.as_bytes());
    output.extend_from_slice(b"</p:extLst>");
    output.extend_from_slice(&xml[root_end..]);
    Ok(output)
}

fn checked_output_size(size: Option<usize>) -> Result<usize> {
    let size = size.ok_or_else(|| limit("updated slide XML bytes", MAX_SLIDE_XML_BYTES))?;
    if size > MAX_SLIDE_XML_BYTES {
        return Err(limit("updated slide XML bytes", MAX_SLIDE_XML_BYTES));
    }
    Ok(size)
}

fn add_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("slide XML node count", MAX_XML_NODES))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("slide XML node count", MAX_XML_NODES));
    }
    Ok(())
}

fn is_presentationml_name(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    if name.local_name().as_ref() != local {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => *value == PML || *value == STRICT_PML,
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::time::Offset;
    use litchi_drawingml::coordinate::Coordinate;

    fn point() -> TracePoint {
        TracePoint::new(
            Offset::ZERO,
            Coordinate::emu(914_400).unwrap(),
            Coordinate::emu(457_200).unwrap(),
        )
    }

    fn package(xml: &str) -> (OpcPackage, PackURI) {
        let name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            name.clone(),
            ct::PML_SLIDE.into(),
            xml.as_bytes().to_vec(),
        )));
        (package, name)
    }

    #[test]
    fn package_writer_owns_empty_and_existing_extension_lists() {
        let pml = "http://schemas.openxmlformats.org/presentationml/2006/main";
        for tail in [
            "",
            "<p:extLst/>",
            "<p:extLst><p:ext uri=\"opaque\"/></p:extLst>",
        ] {
            let xml = format!(r#"<p:sld xmlns:p="{pml}"><p:cSld/><p:spTree/>{tail}</p:sld>"#);
            let (mut package, name) = package(&xml);
            store_slide_trace(&mut package, &name, &[point()]).unwrap();
            let slide = package.get_part(&name).unwrap();
            let traces = load_slide_traces(0, slide, &mut Limits::default()).unwrap();
            assert_eq!(traces.len(), 1);
            assert_eq!(traces[0].point_count(), 1);
        }
    }

    #[test]
    fn package_writer_preserves_strict_dialect_and_rejects_replacement() {
        let strict = "http://purl.oclc.org/ooxml/presentationml/main";
        let xml = format!(r#"<p:sld xmlns:p="{strict}"><p:cSld/></p:sld>"#);
        let (mut package, name) = package(&xml);
        store_slide_trace(&mut package, &name, &[point()]).unwrap();
        let slide = package.get_part(&name).unwrap();
        let output = std::str::from_utf8(slide.blob()).unwrap();
        assert!(output.contains(strict));
        assert!(store_slide_trace(&mut package, &name, &[point()]).is_err());
    }
}
