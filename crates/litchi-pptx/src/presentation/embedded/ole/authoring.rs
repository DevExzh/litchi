use super::codec::OLE_GRAPHIC_DATA_URI;
use super::model::{Authored, Frame, Kind};
use super::package::load_slide;
use super::{Limits, Mode};
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::presentation::embedded::{MAX_XML_DEPTH, invalid, is_presentationml_name, limit};
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
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
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn add(
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
        return Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            payload.len(),
            MAX_PAYLOAD_BYTES,
        ));
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
    refuse_signed_change(package)?;
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
    settle_unsigned_change(package)?;
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

/// Where the frame insertion is among the top-level elements of the slide.
///
/// ```text
/// Prolog --p:sld start tag-------------------------> Rooted
/// Prolog | Rooted --top-level empty-element tag----> unchanged (accepted: skipped)
/// any    --declaration, comment, text, CDATA-------> unchanged
/// ```
///
/// The insertion is lenient about empty-element tags: one at the top level is
/// never taken for the root and is skipped whatever its name, before or after
/// the root, so an empty `p:sld` alone is not a document. The last line holds
/// in every phase: a declaration is accepted anywhere and content outside the
/// root is ignored. Refused: a top-level start tag other than `p:sld` in
/// `Prolog`; any top-level start tag in `Rooted`; a processing instruction or
/// DOCTYPE anywhere; and the end of input in `Prolog` or inside an element.
#[derive(Clone, Copy, PartialEq, Eq)]
enum RootPhase {
    /// No root start tag has been read.
    Prolog,
    /// The root start tag has been read: inside the root, or after it.
    Rooted,
}

fn insert_frame(xml: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut phase = RootPhase::Prolog;
    let mut tree_depth = None;
    let mut insertion = None;
    loop {
        let before = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth += 1;
                if depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if depth == 1 {
                    crate::presentation::embedded::validate_root(
                        &namespace,
                        element.name(),
                        phase == RootPhase::Rooted,
                    )?;
                    phase = RootPhase::Rooted;
                }
                if is_presentationml_name(&namespace, element.name(), b"spTree")
                    && tree_depth.replace(depth).is_some()
                {
                    return Err(invalid("slide has multiple shape trees"));
                }
            },
            // An empty-element tag is never the root: at the top level it is
            // skipped in either phase.
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
            // Accepted in every phase and ignored.
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
    }
    match phase {
        RootPhase::Rooted if depth == 0 => {},
        RootPhase::Prolog | RootPhase::Rooted => {
            return Err(invalid("unterminated slide XML"));
        },
    }
    let position = insertion.ok_or_else(|| invalid("slide has no shape tree"))?;
    let size = xml.len().checked_add(fragment.len()).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::XmlBytes,
            usize::MAX,
            crate::presentation::embedded::MAX_XML_BYTES,
        )
    })?;
    if size > crate::presentation::embedded::MAX_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            size,
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
        match package.get_part(&name) {
            Ok(_) => {},
            Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(index),
            Err(error) => return Err(error.into()),
        }
    }
    Err(limit(
        crate::error::CodecLimitKind::Parts,
        1_000_000,
        1_000_000,
    ))
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

#[cfg(test)]
mod envelope_tests {
    //! What the frame insertion accepts before, around and after the slide root, and the exact error reported otherwise. Each expected
    //! line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&[u8]) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(xml.as_bytes())))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The insertion of a frame before the first top-level shape tree close. The
    /// result is the patched document.
    #[test]
    fn frame_insertion_envelope() {
        let root = Root {
            open: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            body: r"<p:cSld><p:spTree><p:nvGrpSpPr/></p:spTree></p:cSld>",
            close: r"</p:sld>",
            empty: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"unterminated slide XML\")",
                "whitespace only: Invalid(\"unterminated slide XML\")",
                "declaration only: Invalid(\"unterminated slide XML\")",
                "empty-element root: Invalid(\"unterminated slide XML\")",
                "start and end root: Invalid(\"slide has no shape tree\")",
                "root then whitespace: Ok(<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>\n )",
                "root then comment: Ok(<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld><!--c-->)",
                "root then PI: Invalid(\"slide XML rejects DTDs and processing instructions\")",
                "root then text: Ok(<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>text)",
                "root then second root: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "root then foreign start: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "root then foreign empty: Ok(<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld><foreign/>)",
                "declaration after root: Ok(<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld><?xml version=\"1.0\" encoding=\"UTF-8\"?>)",
                "two declarations: Ok(<?xml version=\"1.0\" encoding=\"UTF-8\"?><?xml version=\"1.0\" encoding=\"UTF-8\"?><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "comment before declaration: Ok(<!--c--><?xml version=\"1.0\" encoding=\"UTF-8\"?><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "text before root: Ok(text<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "unterminated root: Invalid(\"unterminated slide XML\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"slide XML rejects DTDs and processing instructions\")",
                "top-level CDATA: Ok(<![CDATA[x]]><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "declaration then root: Ok(<?xml version=\"1.0\" encoding=\"UTF-8\"?><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "whitespace before declaration: Ok( <?xml version=\"1.0\" encoding=\"UTF-8\"?><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "PI before declaration: Invalid(\"slide XML rejects DTDs and processing instructions\")",
                "comment then root: Ok(<!--c--><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "PI then root: Invalid(\"slide XML rejects DTDs and processing instructions\")",
                "root then text then second root: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "empty-element root then root: Ok(<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/><p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr/><frame/></p:spTree></p:cSld></p:sld>)",
                "declaration after empty-element root: Invalid(\"unterminated slide XML\")",
            ],
            |xml| {
                outcome(super::insert_frame(xml, b"<frame/>"), |bytes| {
                    String::from_utf8_lossy(&bytes).into_owned()
                })
            },
        );
    }
}
