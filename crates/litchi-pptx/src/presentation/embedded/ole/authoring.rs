use super::codec::OLE_GRAPHIC_DATA_URI;
use super::model::{Authored, Frame, Kind};
use super::package::load_slide;
use super::{Limits, Mode};
use crate::presentation::embedded::{MAX_XML_DEPTH, invalid, is_presentationml_name, limit};
use crate::{Error, Result};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::fmt::Write;

const MAX_PAYLOAD_BYTES: usize = 32 * 1024 * 1024;
const MAX_PROG_ID_CHARS: usize = 39;
const MAX_NAME_CHARS: usize = 256;

/// Add an inert embedded OLE/package payload and verify it through the read
/// inventory before publishing the package mutation.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn add(
    package: &mut OpcPackage,
    slide_part_name: &str,
    kind: Kind,
    program_id: Option<&str>,
    name: Option<&str>,
    frame: Frame,
    payload: &[u8],
) -> Result<Authored> {
    if let Some(value) = program_id {
        validate_program_id(value)?;
    }
    if let Some(value) = name
        && (value.is_empty() || value.chars().count() > MAX_NAME_CHARS)
    {
        return Err(invalid("OLE object name is empty or too long"));
    }
    if frame.cx <= 0 || frame.cy <= 0 {
        return Err(invalid("OLE frame extents must be positive"));
    }
    if payload.is_empty() || payload.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("OLE payload bytes", MAX_PAYLOAD_BYTES));
    }
    let slide_uri = PackURI::new(slide_part_name).map_err(Error::Uri)?;
    let slide = package.get_part(&slide_uri)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_string(),
            actual: slide.content_type().to_string(),
        });
    }
    let slide_xml = slide.blob().to_vec();
    let shape_id = next_shape_id(&slide_xml)?;
    let index = next_part_index(package)?;
    let part_name =
        PackURI::new(format!("/ppt/embeddings/oleObject{index}.bin")).map_err(Error::Uri)?;
    let relationship_type = match kind {
        Kind::OleObject => rt::OLE_OBJECT,
        Kind::Package => rt::PACKAGE,
    };
    let content_type = match kind {
        Kind::OleObject => ct::OFC_OLE_OBJECT,
        Kind::Package => ct::OFC_PACKAGE,
    };
    let fragment = frame_xml(shape_id, name, program_id, "rIdPending", frame);
    let _ = insert_frame(&slide_xml, fragment.as_bytes())?;
    let target = part_name.relative_ref(slide_uri.base_uri());
    let relationship_id = package
        .get_part_mut(&slide_uri)?
        .relate_to(&target, relationship_type);
    let fragment = frame_xml(shape_id, name, program_id, &relationship_id, frame);
    let patched = insert_frame(&slide_xml, fragment.as_bytes())?;
    package.get_part_mut(&slide_uri)?.set_blob(patched);
    package.add_part(Box::new(BlobPart::new(
        part_name.clone(),
        content_type.to_string(),
        payload.to_vec(),
    )));
    let mut limits = Limits::default();
    let objects = load_slide(package, 0, package.get_part(&slide_uri)?, &mut limits)?;
    if !objects.iter().any(|object| {
        object.shape_id() == Some(shape_id)
            && object.mode() == Mode::Embedded
            && object.relationship_id() == Some(relationship_id.as_str())
            && object.kind() == Some(kind)
            && object.target().and_then(|target| target.part_name()) == Some(&part_name)
    }) {
        return Err(invalid("authored OLE object failed read-side verification"));
    }
    package.unsign();
    Ok(Authored {
        part_name,
        relationship_id,
        shape_id,
    })
}

fn frame_xml(
    shape_id: u32,
    name: Option<&str>,
    program_id: Option<&str>,
    relationship_id: &str,
    frame: Frame,
) -> String {
    let default_name = format!("OLE Object {shape_id}");
    let name = name.unwrap_or(&default_name);
    let mut xml = String::with_capacity(1024);
    let _result = write!(
        xml,
        r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="{}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></p:xfrm><a:graphic><a:graphicData uri="{OLE_GRAPHIC_DATA_URI}"><p:oleObj"#,
        escape(name),
        frame.x,
        frame.y,
        frame.cx,
        frame.cy
    );
    if let Some(program_id) = program_id {
        let _result = write!(xml, r#" progId="{}""#, escape(program_id));
    }
    let _result = write!(
        xml,
        r#" r:id="{}"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>"#,
        escape(relationship_id)
    );
    xml
}

fn insert_frame(xml: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut tree_depth = None;
    let mut insertion = None;
    loop {
        let before = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth += 1;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("OLE slide depth", MAX_XML_DEPTH));
                }
                if depth == 1 {
                    crate::presentation::embedded::validate_root(
                        &namespace,
                        element.name(),
                        root_seen,
                    )?;
                    root_seen = true;
                }
                if is_presentationml_name(&namespace, element.name(), b"spTree")
                    && tree_depth.replace(depth).is_some()
                {
                    return Err(invalid("slide has multiple shape trees"));
                }
            },
            Event::Empty(element) => {
                if is_presentationml_name(&namespace, element.name(), b"spTree") {
                    return Err(invalid("slide has an empty shape tree"));
                }
            },
            Event::End(element) => {
                if tree_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"spTree")
                {
                    insertion = Some(before);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid slide XML nesting"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "slide XML rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen {
        return Err(invalid("unterminated slide XML"));
    }
    let position = insertion.ok_or_else(|| invalid("slide has no shape tree"))?;
    let size = xml.len().checked_add(fragment.len()).ok_or_else(|| {
        limit(
            "updated slide XML bytes",
            crate::presentation::embedded::MAX_XML_BYTES,
        )
    })?;
    if size > crate::presentation::embedded::MAX_XML_BYTES {
        return Err(limit(
            "updated slide XML bytes",
            crate::presentation::embedded::MAX_XML_BYTES,
        ));
    }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn next_shape_id(xml: &[u8]) -> Result<u32> {
    let mut reader = NsReader::from_reader(xml);
    let mut maximum = 1u32;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"cNvPr" =>
            {
                if let Some(value) = unqualified_attribute_value(&element, b"id", decoder)? {
                    maximum =
                        maximum.max(value.parse().map_err(|_err| invalid("invalid shape ID"))?);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "slide XML rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    maximum
        .checked_add(1)
        .ok_or_else(|| invalid("shape ID overflow"))
}

fn next_part_index(package: &OpcPackage) -> Result<u32> {
    for index in 1..1_000_000u32 {
        let name =
            PackURI::new(format!("/ppt/embeddings/oleObject{index}.bin")).map_err(Error::Uri)?;
        if package.get_part(&name).is_err() {
            return Ok(index);
        }
    }
    Err(limit("OLE embedding part namespace", 1_000_000))
}

fn validate_program_id(value: &str) -> Result<()> {
    let mut chars = value.chars();
    if value.chars().count() > MAX_PROG_ID_CHARS
        || !chars
            .next()
            .is_some_and(|value| value.is_ascii_alphabetic())
        || !chars.all(|value| value.is_ascii_alphanumeric() || value == '.')
    {
        return Err(invalid("invalid OLE program ID"));
    }
    Ok(())
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
