//! Embedded OLE object authoring for PowerPoint slides.
//!
//! This module closes the authoring gap for `p:oleObj` shapes. Embedding an
//! inert binary payload into a slide:
//!
//! - stores the payload verbatim as `/ppt/embeddings/oleObjectN.bin` with the
//!   OOXML Embedded Object (or Embedded Package) content type;
//! - adds an `rt::OLE_OBJECT` (or `rt::PACKAGE`) relationship from the slide
//!   part to the payload part;
//! - patches a `p:graphicFrame` carrying the `p:oleObj` element (with the
//!   required `p:embed` child) into the slide's shape tree, using prefix-safe
//!   XML like the master/layout and theme authors;
//! - verifies the patched slide through the read-side OLE inventory, so an
//!   authored object resolves cleanly through `Package::ole_objects`.
//!
//! Payloads are stored verbatim and are never parsed, activated, rendered,
//! or executed.

use crate::error::{OoxmlError, Result};
use crate::pptx::ole::{
    OleLoadLimits, PptxOleObjectMode, PptxOlePayloadKind, load_slide_ole_objects,
};
use litchi_core::xml::escape_xml;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::BlobPart;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fmt::Write as FmtWrite;

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// `a:graphicData/@uri` that marks a graphic frame as an OLE object.
const OLE_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";

/// Bounded ceiling for slide XML this module parses or patches; matches the
/// read-side OLE inventory limit.
const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
/// Bounded ceiling for one embedded payload.
const MAX_OLE_PAYLOAD_BYTES: usize = 32 * 1024 * 1024;
/// Bounded ceiling for XML node counts while scanning.
const MAX_SCAN_NODES: usize = 100_000;
/// Bounded ceiling for XML nesting depth while scanning.
const MAX_SCAN_DEPTH: usize = 128;
/// Bounded ceiling for an OLE program ID (`ST_ProgID` allows up to 39).
const MAX_PROG_ID_CHARS: usize = 39;
/// Bounded ceiling for authored shape names.
const MAX_NAME_CHARS: usize = 256;
/// Shape ID 1 is reserved for the group-shape root of every shape tree.
const FIRST_SHAPE_ID: u32 = 2;
/// Depth of `p:spTree` inside a slide part (root → cSld → spTree).
const SPTREE_DEPTH: usize = 3;

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

// ============================================================================
// Typed inputs
// ============================================================================

/// Position and extent of the OLE graphic frame, in EMUs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OleObjectFrame {
    /// Horizontal offset (`a:off/@x`).
    pub x: i64,
    /// Vertical offset (`a:off/@y`).
    pub y: i64,
    /// Width (`a:ext/@cx`); must be positive.
    pub cx: i64,
    /// Height (`a:ext/@cy`); must be positive.
    pub cy: i64,
}

impl OleObjectFrame {
    /// Create a frame with the given offset and extent, in EMUs.
    pub const fn new(x: i64, y: i64, cx: i64, cy: i64) -> Self {
        Self { x, y, cx, cy }
    }
}

/// Identity of an OLE object embedded by [`add_ole_object`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredOleObject {
    /// Part name of the embedded payload, e.g. `/ppt/embeddings/oleObject1.bin`.
    pub part_name: String,
    /// Relationship ID from the slide part to the payload part.
    pub relationship_id: String,
    /// Shape ID of the authored graphic frame on the slide.
    pub shape_id: u32,
}

// ============================================================================
// Authoring operation
// ============================================================================

/// Embed an inert binary payload into a slide as an OLE object shape.
///
/// The payload is stored verbatim in the next free
/// `/ppt/embeddings/oleObjectN.bin` part with the content type implied by
/// `kind`, the slide gains a matching relationship, and a `p:graphicFrame`
/// carrying `p:oleObj` (with `p:embed`) is appended to the slide's shape
/// tree with the next free shape ID. `prog_id`, when supplied, must be a
/// valid `ST_ProgID` (leading ASCII letter, then ASCII letters, digits, or
/// periods, at most 39 characters). The patched slide is verified through
/// the read-side OLE inventory before the operation returns.
pub fn add_ole_object(
    package: &mut OpcPackage,
    slide_part_name: &str,
    kind: PptxOlePayloadKind,
    prog_id: Option<&str>,
    name: Option<&str>,
    frame: OleObjectFrame,
    payload: &[u8],
) -> Result<AuthoredOleObject> {
    if let Some(prog_id) = prog_id {
        require_prog_id(prog_id)?;
    }
    if let Some(name) = name {
        require_name(name)?;
    }
    require_frame(frame)?;
    require_payload(payload)?;

    let slide_uri = PackURI::new(slide_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("slide partname: {error}")))?;
    let slide_part = package.get_part(&slide_uri)?;
    if slide_part.content_type() != ct::PML_SLIDE {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::PML_SLIDE.to_string(),
            got: slide_part.content_type().to_string(),
        });
    }
    let slide_xml = slide_part.blob().to_vec();
    let shape_id = next_shape_id(&slide_xml)?;

    let index = next_part_index(package, "/ppt/embeddings/oleObject", ".bin")?;
    let embedding_uri = PackURI::new(format!("/ppt/embeddings/oleObject{index}.bin"))
        .map_err(|error| OoxmlError::InvalidUri(format!("embedding partname: {error}")))?;

    let slide_dir = slide_part_name
        .rsplit_once('/')
        .map(|(directory, _)| format!("{directory}/"))
        .ok_or_else(|| OoxmlError::InvalidUri(format!("slide partname: {slide_part_name}")))?;
    let target = relative_target(&slide_dir, embedding_uri.as_str())?;

    let slide_part = package.get_part_mut(&slide_uri)?;
    let relationship_id = slide_part.relate_to(&target, relationship_type(kind));
    let frame_xml = graphic_frame_xml(shape_id, name, prog_id, &relationship_id, frame);
    let patched = match insert_graphic_frame(&slide_xml, frame_xml.as_bytes()) {
        Ok(patched) => patched,
        Err(error) => {
            slide_part.rels_mut().remove(&relationship_id);
            return Err(error);
        },
    };
    slide_part.set_blob(patched);
    package.add_part(Box::new(BlobPart::new(
        embedding_uri.clone(),
        content_type(kind).to_string(),
        payload.to_vec(),
    )));

    // The patched slide must inventory the authored object back through the
    // same scan the read side performs.
    verify_authored_object(
        package,
        &slide_uri,
        shape_id,
        prog_id,
        &relationship_id,
        kind,
        &embedding_uri,
    )?;
    invalidate_signatures(package)?;
    Ok(AuthoredOleObject {
        part_name: embedding_uri.to_string(),
        relationship_id,
        shape_id,
    })
}

fn relationship_type(kind: PptxOlePayloadKind) -> &'static str {
    match kind {
        PptxOlePayloadKind::OleObject => rt::OLE_OBJECT,
        PptxOlePayloadKind::Package => rt::PACKAGE,
    }
}

fn content_type(kind: PptxOlePayloadKind) -> &'static str {
    match kind {
        PptxOlePayloadKind::OleObject => ct::OFC_OLE_OBJECT,
        PptxOlePayloadKind::Package => ct::OFC_PACKAGE,
    }
}

/// Run the read-side OLE inventory over the patched slide and require the
/// authored object with exactly the authored metadata.
fn verify_authored_object(
    package: &OpcPackage,
    slide_uri: &PackURI,
    shape_id: u32,
    prog_id: Option<&str>,
    relationship_id: &str,
    kind: PptxOlePayloadKind,
    embedding_uri: &PackURI,
) -> Result<()> {
    let slide_part = package.get_part(slide_uri)?;
    let mut limits = OleLoadLimits::default();
    let objects = load_slide_ole_objects(package, 0, slide_part, &mut limits)?;
    let object = objects
        .iter()
        .find(|object| object.shape_id() == Some(shape_id))
        .ok_or_else(|| invalid("read-side OLE inventory lost the authored object"))?;
    if object.program_id() != prog_id
        || object.mode() != PptxOleObjectMode::Embedded
        || object.relationship_id() != Some(relationship_id)
        || object.payload_kind() != Some(kind)
        || object.target().and_then(|target| target.part_name()) != Some(embedding_uri)
    {
        return Err(invalid("authored OLE object did not round-trip"));
    }
    Ok(())
}

// ============================================================================
// XML generation
// ============================================================================

/// Serialize the OLE graphic frame with its own namespace declarations so it
/// can be patched into a slide with unknown prefix bindings.
fn graphic_frame_xml(
    shape_id: u32,
    name: Option<&str>,
    prog_id: Option<&str>,
    relationship_id: &str,
    frame: OleObjectFrame,
) -> String {
    let default_name;
    let name = match name {
        Some(name) => name,
        None => {
            default_name = format!("OLE Object {shape_id}");
            &default_name
        },
    };
    let mut xml = String::with_capacity(1024);
    let _ = write!(
        xml,
        "<p:graphicFrame xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\"><p:nvGraphicFramePr><p:cNvPr id=\"{shape_id}\" name=\"{}\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></p:xfrm><a:graphic><a:graphicData uri=\"{OLE_GRAPHIC_DATA_URI}\"><p:oleObj",
        escape_xml(name),
        frame.x,
        frame.y,
        frame.cx,
        frame.cy
    );
    if let Some(prog_id) = prog_id {
        let _ = write!(xml, " progId=\"{}\"", escape_xml(prog_id));
    }
    let _ = write!(
        xml,
        " r:id=\"{}\"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>",
        escape_xml(relationship_id)
    );
    xml
}

// ============================================================================
// Bounded XML scanning and patching
// ============================================================================

/// Byte span of an XML element.
#[derive(Debug, Clone, Copy)]
struct ElementSpan {
    /// Offset of the `</` that opens the closing tag.
    close_start: usize,
    /// Whether the element uses the self-closing form.
    empty: bool,
}

fn check_size(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(invalid("slide XML exceeds 32 MiB"));
    }
    Ok(())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Find the first element with `target` as local name at exactly `depth`.
fn scan_element_span(xml: &[u8], target: &str, depth: usize) -> Result<Option<ElementSpan>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<(usize, String)> = Vec::new();
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES || stack.len() >= MAX_SCAN_DEPTH {
                    return Err(invalid("slide XML resource limit exceeded"));
                }
                let local =
                    String::from_utf8_lossy(local_name(element.name().as_ref())).into_owned();
                stack.push((before, local));
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("slide XML resource limit exceeded"));
                }
                if stack.len() + 1 == depth
                    && local_name(element.name().as_ref()) == target.as_bytes()
                {
                    return Ok(Some(ElementSpan {
                        close_start: before,
                        empty: true,
                    }));
                }
            },
            Ok(Event::End(element)) => {
                let (_, local) = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element in slide XML"))?;
                if stack.len() + 1 == depth && local == target {
                    return Ok(Some(ElementSpan {
                        close_start: before,
                        empty: false,
                    }));
                }
                if local_name(element.name().as_ref()) != local.as_bytes() {
                    return Err(invalid("mismatched closing element in slide XML"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated slide XML"));
    }
    Ok(None)
}

/// Insert the graphic frame at the end of the slide's shape tree.
fn insert_graphic_frame(xml: &[u8], frame: &[u8]) -> Result<Vec<u8>> {
    let tree = scan_element_span(xml, "spTree", SPTREE_DEPTH)?
        .ok_or_else(|| invalid("slide has no shape tree"))?;
    if tree.empty {
        return Err(invalid("slide has an empty shape tree"));
    }
    let mut output = Vec::with_capacity(xml.len() + frame.len());
    output.extend_from_slice(&xml[..tree.close_start]);
    output.extend_from_slice(frame);
    output.extend_from_slice(&xml[tree.close_start..]);
    check_size(&output)?;
    Ok(output)
}

/// Allocate the next free shape ID for a slide (max existing + 1, starting
/// at [`FIRST_SHAPE_ID`]).
fn next_shape_id(xml: &[u8]) -> Result<u32> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut max_id = FIRST_SHAPE_ID - 1;
    let mut nodes = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("slide XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == b"cNvPr" {
                    for attribute in element.attributes().with_checks(true) {
                        let attribute =
                            attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        if attribute.key.as_ref() == b"id" {
                            let value = std::str::from_utf8(attribute.value.as_ref())
                                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                            let id = value
                                .parse::<u32>()
                                .map_err(|_| invalid(format!("invalid shape ID '{value}'")))?;
                            max_id = max_id.max(id);
                        }
                    }
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    max_id
        .checked_add(1)
        .ok_or_else(|| invalid("shape ID overflow"))
}

// ============================================================================
// Misc helpers and validators
// ============================================================================

/// Find the lowest free numeric suffix for a part-name pattern.
fn next_part_index(package: &OpcPackage, prefix: &str, suffix: &str) -> Result<u32> {
    let mut index = 1u32;
    loop {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
            .map_err(|error| OoxmlError::InvalidUri(format!("partname allocation: {error}")))?;
        if package.get_part(&candidate).is_err() {
            return Ok(index);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("part-name index overflow"))?;
    }
}

/// Compute the relationship target for `target` relative to `source_dir`.
///
/// Both names must be absolute pack URIs; the result uses `..` segments to
/// climb out of the source directory.
fn relative_target(source_dir: &str, target: &str) -> Result<String> {
    let source = source_dir.trim_matches('/');
    let target = target.trim_start_matches('/');
    let source_segments: Vec<&str> = source.split('/').filter(|item| !item.is_empty()).collect();
    let target_segments: Vec<&str> = target.split('/').filter(|item| !item.is_empty()).collect();
    let common = source_segments
        .iter()
        .zip(target_segments.iter())
        .take_while(|(left, right)| left == right)
        .count();
    if common == 0 && !source_segments.is_empty() {
        return Err(OoxmlError::InvalidUri(format!(
            "cannot relativize '{target}' against '/{source}/'"
        )));
    }
    let mut result = String::new();
    for _ in common..source_segments.len() {
        result.push_str("../");
    }
    result.push_str(&target_segments[common..].join("/"));
    Ok(result)
}

/// Validate an OLE program ID (`ST_ProgID`): a leading ASCII letter, then
/// ASCII letters, digits, or periods, at most 39 characters.
fn require_prog_id(prog_id: &str) -> Result<()> {
    let mut chars = prog_id.chars();
    let length = chars.clone().count();
    let valid = length <= MAX_PROG_ID_CHARS
        && chars
            .next()
            .is_some_and(|first| first.is_ascii_alphabetic())
        && chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '.');
    if valid {
        return Ok(());
    }
    Err(invalid(format!(
        "invalid OLE program ID '{prog_id}'; expected a leading letter, then letters, digits, or periods, at most {MAX_PROG_ID_CHARS} characters"
    )))
}

fn require_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(invalid("OLE object name cannot be empty"));
    }
    if name.chars().count() > MAX_NAME_CHARS {
        return Err(invalid("OLE object name exceeds 256 characters"));
    }
    Ok(())
}

fn require_frame(frame: OleObjectFrame) -> Result<()> {
    if frame.cx <= 0 || frame.cy <= 0 {
        return Err(invalid("OLE object frame extent must be positive"));
    }
    Ok(())
}

fn require_payload(payload: &[u8]) -> Result<()> {
    if payload.is_empty() {
        return Err(invalid("OLE payload cannot be empty"));
    }
    if payload.len() > MAX_OLE_PAYLOAD_BYTES {
        return Err(invalid("OLE payload exceeds 32 MiB"));
    }
    Ok(())
}

fn invalidate_signatures(package: &mut OpcPackage) -> Result<()> {
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::Other(format!("cannot invalidate package signatures: {error}"))
    })
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::ole::PptxOleObjectTarget;
    use crate::pptx::{Package, PptxOleObjectMode as Mode, PptxOlePayloadKind as Kind};
    use litchi_opc::PackageWriter;
    use litchi_opc::part::Part;
    use std::io::Cursor;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn roundtrip(package: &Package) -> Package {
        let bytes = PackageWriter::to_bytes(package.opc_package()).unwrap();
        Package::from_reader(Cursor::new(bytes)).unwrap()
    }

    fn slide_xml(extra_shapes: &str) -> String {
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{extra_shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        )
    }

    /// Register a slide part in the presentation (relationship + sldIdLst).
    fn package_with_slides(count: u32, extra_shapes: &str) -> (Package, Vec<String>) {
        let mut package = Package::new().unwrap();
        let mut names = Vec::new();
        {
            let opc = package.opc_package_mut();
            let presentation_name = opc.main_document_part().unwrap().partname().clone();
            let mut entries = String::new();
            for index in 1..=count {
                let part_name = format!("/ppt/slides/slide{index}.xml");
                let mut slide = BlobPart::new(
                    PackURI::new(&part_name).unwrap(),
                    ct::PML_SLIDE.to_string(),
                    slide_xml(extra_shapes).into_bytes(),
                );
                slide.relate_to("../slideLayouts/slideLayout1.xml", rt::SLIDE_LAYOUT);
                opc.add_part(Box::new(slide));
                let relationship_id = opc
                    .get_part_mut(&presentation_name)
                    .unwrap()
                    .relate_to(&format!("slides/slide{index}.xml"), rt::SLIDE);
                let _ = write!(
                    entries,
                    "<p:sldId id=\"{}\" r:id=\"{relationship_id}\"/>",
                    255 + index
                );
                names.push(part_name);
            }
            let presentation = opc.get_part_mut(&presentation_name).unwrap();
            let xml = String::from_utf8(presentation.blob().to_vec()).unwrap();
            let entry = format!("<p:sldIdLst>{entries}</p:sldIdLst>");
            let position = xml
                .find("<p:sldSz")
                .expect("default presentation has a slide size");
            let mut patched = xml;
            patched.insert_str(position, &entry);
            presentation.set_blob(patched.into_bytes());
        }
        (package, names)
    }

    fn sample_frame() -> OleObjectFrame {
        OleObjectFrame::new(914400, 914400, 2743200, 1828800)
    }

    #[test]
    fn authored_ole_object_roundtrips_through_inventory() {
        let payload: Vec<u8> = (0u32..4096).map(|value| (value % 251) as u8).collect();
        let (mut package, slides) = package_with_slides(1, "");
        let authored = package
            .add_ole_object(
                &slides[0],
                Kind::OleObject,
                Some("Acme.Document.1"),
                Some("Quarterly & Numbers"),
                sample_frame(),
                &payload,
            )
            .unwrap();
        assert_eq!(authored.part_name, "/ppt/embeddings/oleObject1.bin");
        assert_eq!(authored.shape_id, 2);

        let reopened = roundtrip(&package);
        let objects = reopened.ole_objects().unwrap();
        assert_eq!(objects.len(), 1);
        let object = &objects[0];
        assert_eq!(object.slide_index(), 0);
        assert_eq!(object.object_index(), 0);
        assert_eq!(object.shape_id(), Some(2));
        assert_eq!(object.shape_name(), Some("Quarterly & Numbers"));
        assert_eq!(object.program_id(), Some("Acme.Document.1"));
        assert_eq!(object.mode(), Mode::Embedded);
        assert_eq!(
            object.relationship_id(),
            Some(authored.relationship_id.as_str())
        );
        assert_eq!(object.payload_kind(), Some(Kind::OleObject));
        let PptxOleObjectTarget::Internal {
            part_name,
            content_type: target_content_type,
            relationship_type: target_relationship_type,
        } = object.target().unwrap()
        else {
            panic!("authored OLE object must have an internal target");
        };
        assert_eq!(part_name.as_str(), "/ppt/embeddings/oleObject1.bin");
        assert_eq!(target_content_type, ct::OFC_OLE_OBJECT);
        assert_eq!(target_relationship_type, rt::OLE_OBJECT);

        // The payload round-trips byte-identically.
        let part = reopened
            .opc_package()
            .get_part(&PackURI::new(&authored.part_name).unwrap())
            .unwrap();
        assert_eq!(part.blob(), payload.as_slice());
    }

    #[test]
    fn multiple_embeddings_across_slides_get_distinct_parts() {
        let (mut package, slides) = package_with_slides(2, "");
        let first = package
            .add_ole_object(
                &slides[0],
                Kind::OleObject,
                Some("Acme.Chart"),
                None,
                sample_frame(),
                b"one",
            )
            .unwrap();
        let second = package
            .add_ole_object(
                &slides[0],
                Kind::Package,
                Some("Package"),
                None,
                sample_frame(),
                b"two",
            )
            .unwrap();
        let third = package
            .add_ole_object(
                &slides[1],
                Kind::OleObject,
                None,
                None,
                sample_frame(),
                b"three",
            )
            .unwrap();
        assert_eq!(first.part_name, "/ppt/embeddings/oleObject1.bin");
        assert_eq!(second.part_name, "/ppt/embeddings/oleObject2.bin");
        assert_eq!(third.part_name, "/ppt/embeddings/oleObject3.bin");
        assert_ne!(
            first.shape_id, second.shape_id,
            "shape IDs must not collide on one slide"
        );

        let reopened = roundtrip(&package);
        let objects = reopened.ole_objects().unwrap();
        assert_eq!(objects.len(), 3);
        let by_part = |name: &str| {
            objects
                .iter()
                .find(|object| {
                    object
                        .target()
                        .and_then(|target| target.part_name())
                        .map(PackURI::as_str)
                        == Some(name)
                })
                .unwrap()
        };
        let one = by_part("/ppt/embeddings/oleObject1.bin");
        assert_eq!(one.slide_index(), 0);
        assert_eq!(one.program_id(), Some("Acme.Chart"));
        let two = by_part("/ppt/embeddings/oleObject2.bin");
        assert_eq!(two.slide_index(), 0);
        assert_eq!(two.payload_kind(), Some(Kind::Package));
        assert_eq!(two.target().unwrap().content_type(), Some(ct::OFC_PACKAGE));
        let three = by_part("/ppt/embeddings/oleObject3.bin");
        assert_eq!(three.slide_index(), 1);
        assert_eq!(three.program_id(), None);
        assert_eq!(three.shape_name(), Some("OLE Object 2"));
    }

    #[test]
    fn embedding_coexists_with_existing_slide_content() {
        let picture = "<p:pic><p:nvPicPr><p:cNvPr id=\"2\" name=\"Logo\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>";
        let (mut package, slides) = package_with_slides(1, picture);
        let authored = package
            .add_ole_object(
                &slides[0],
                Kind::OleObject,
                Some("Acme.Doc"),
                None,
                sample_frame(),
                b"data",
            )
            .unwrap();
        assert_eq!(
            authored.shape_id, 3,
            "shape ID must clear the existing picture"
        );

        let reopened = roundtrip(&package);
        let objects = reopened.ole_objects().unwrap();
        assert_eq!(objects.len(), 1);
        assert_eq!(objects[0].shape_id(), Some(3));
        // The pre-existing picture survives the patch untouched.
        let slide = reopened
            .opc_package()
            .get_part(&PackURI::new(&slides[0]).unwrap())
            .unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("<p:pic>"));
        assert!(xml.contains("name=\"Logo\""));
        // The slide still resolves and parses through the read side.
        let presentation = reopened.presentation().unwrap();
        assert_eq!(presentation.slide_count().unwrap(), 1);
        presentation.slides().unwrap()[0].text().unwrap();
    }

    #[test]
    fn invalid_ole_inputs_are_rejected() {
        let (mut package, slides) = package_with_slides(1, "");
        // Bad program IDs: leading digit, illegal character, too long, empty.
        for bad in ["1Acme", "Acme!Doc", &"A".repeat(40), ""] {
            assert!(
                package
                    .add_ole_object(
                        &slides[0],
                        Kind::OleObject,
                        Some(bad),
                        None,
                        sample_frame(),
                        b"x"
                    )
                    .is_err(),
                "program ID '{bad}' must be rejected"
            );
        }
        // A 39-character ID with periods is accepted.
        let long = format!("A{}.9", "b".repeat(36));
        assert_eq!(long.chars().count(), 39);
        assert!(
            package
                .add_ole_object(
                    &slides[0],
                    Kind::OleObject,
                    Some(&long),
                    None,
                    sample_frame(),
                    b"x"
                )
                .is_ok()
        );
        // Empty and oversize payloads are rejected.
        assert!(
            package
                .add_ole_object(&slides[0], Kind::OleObject, None, None, sample_frame(), &[])
                .is_err()
        );
        let oversize = vec![0u8; MAX_OLE_PAYLOAD_BYTES + 1];
        assert!(
            package
                .add_ole_object(
                    &slides[0],
                    Kind::OleObject,
                    None,
                    None,
                    sample_frame(),
                    &oversize
                )
                .is_err()
        );
        // Non-positive extents are rejected.
        assert!(
            package
                .add_ole_object(
                    &slides[0],
                    Kind::OleObject,
                    None,
                    None,
                    OleObjectFrame::new(0, 0, 0, 10),
                    b"x"
                )
                .is_err()
        );
        // Empty and overlong names are rejected.
        assert!(
            package
                .add_ole_object(
                    &slides[0],
                    Kind::OleObject,
                    None,
                    Some(""),
                    sample_frame(),
                    b"x"
                )
                .is_err()
        );
        assert!(
            package
                .add_ole_object(
                    &slides[0],
                    Kind::OleObject,
                    None,
                    Some(&"n".repeat(257)),
                    sample_frame(),
                    b"x"
                )
                .is_err()
        );
        // Missing slide part and non-slide part are rejected.
        assert!(
            package
                .add_ole_object(
                    "/ppt/slides/slide99.xml",
                    Kind::OleObject,
                    None,
                    None,
                    sample_frame(),
                    b"x"
                )
                .is_err()
        );
        assert!(
            package
                .add_ole_object(
                    "/ppt/presentation.xml",
                    Kind::OleObject,
                    None,
                    None,
                    sample_frame(),
                    b"x"
                )
                .is_err()
        );
        // Rejections leave the package clean: one object from the valid call.
        let objects = package.ole_objects().unwrap();
        assert_eq!(objects.len(), 1);
    }

    #[test]
    fn authored_ole_objects_serialize_deterministically() {
        let build = || {
            let (mut package, slides) = package_with_slides(1, "");
            package
                .add_ole_object(
                    &slides[0],
                    Kind::OleObject,
                    Some("Acme.Document.1"),
                    Some("Same"),
                    sample_frame(),
                    b"payload",
                )
                .unwrap();
            package
                .add_ole_object(
                    &slides[0],
                    Kind::Package,
                    Some("Package"),
                    None,
                    sample_frame(),
                    b"second",
                )
                .unwrap();
            package
        };
        let first = build();
        let second = build();
        for part_name in [
            "/ppt/slides/slide1.xml",
            "/ppt/embeddings/oleObject1.bin",
            "/ppt/embeddings/oleObject2.bin",
        ] {
            let uri = PackURI::new(part_name).unwrap();
            assert_eq!(
                first.opc_package().get_part(&uri).unwrap().blob(),
                second.opc_package().get_part(&uri).unwrap().blob(),
                "part {part_name} must serialize deterministically"
            );
        }
    }
}
