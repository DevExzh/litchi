//! Inert SpreadsheetML worksheet-control and ActiveX persistence metadata.
//!
//! ActiveX binaries are deliberately opaque. This module never instantiates a
//! control, resolves a CLSID, executes a macro, decodes MS-OFORMS/CFB data or
//! pictures, or follows an external relationship.

use crate::common::mce::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::{NsReader, XmlVersion};
use std::collections::HashSet;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SML_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const XDR_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const AX: &str = "http://schemas.microsoft.com/office/2006/activeX";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const CONTROL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
const CONTROL_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/control";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const IMAGE_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
const BINARY_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const DESCRIPTOR_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX+xml";
const BINARY_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX";
const MAX_XML: usize = 16 * 1024 * 1024;
const MAX_OUTPUT_XML: usize = 32 * 1024 * 1024;
const MAX_BINARY: usize = 64 * 1024 * 1024;
const MAX_TOTAL_BINARY: usize = 256 * 1024 * 1024;
const MAX_CONTROLS: usize = 65_535;
const MAX_SHAPE_ID: u32 = 67_098_623;
const MAX_CONTROL_NAME_CHARS: usize = 32;
const MAX_PROPERTIES: usize = 65_536;
const MAX_STRING: usize = 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 400_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Persistence {
    PropertyBag,
    Stream,
    StreamInit,
    Storage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Marker {
    pub column: i32,
    pub column_offset: i64,
    pub row: i32,
    pub row_offset: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectAnchor {
    pub from: Marker,
    pub to: Marker,
    pub move_with_cells: Option<bool>,
    pub size_with_cells: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlProperties {
    pub anchor: ObjectAnchor,
    pub locked: Option<bool>,
    pub default_size: Option<bool>,
    pub print: Option<bool>,
    pub disabled: Option<bool>,
    pub recalc_always: Option<bool>,
    pub ui_object: Option<bool>,
    pub auto_fill: Option<bool>,
    pub auto_line: Option<bool>,
    pub auto_picture: Option<bool>,
    pub macro_name: Option<String>,
    pub alternate_text: Option<String>,
    pub preview_relationship_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetControl {
    pub shape_id: u32,
    pub relationship_id: String,
    pub name: Option<String>,
    pub properties: Option<ControlProperties>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorksheetControls {
    pub controls: Vec<WorksheetControl>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveXPicture {
    pub relationship_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveXFont {
    pub persistence: Option<Persistence>,
    pub relationship_id: Option<String>,
    pub properties: Vec<ActiveXProperty>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ActiveXPropertyObject {
    Font(ActiveXFont),
    Picture(ActiveXPicture),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveXProperty {
    pub name: String,
    pub value: Option<String>,
    pub object: Option<ActiveXPropertyObject>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveXDescriptor {
    pub class_id: String,
    pub license: Option<String>,
    pub persistence: Persistence,
    pub relationship_id: Option<String>,
    pub properties: Vec<ActiveXProperty>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueActiveXBinary {
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueActiveXPreviewImage {
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoadedActiveXControl {
    pub control: WorksheetControl,
    pub descriptor_uri: PackURI,
    pub descriptor: ActiveXDescriptor,
    pub binaries: Vec<OpaqueActiveXBinary>,
    pub preview: Option<OpaqueActiveXPreviewImage>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ActiveXControlSet {
    pub controls: Vec<LoadedActiveXControl>,
}

impl WorksheetControls {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML {
            return Err(limit("worksheet XML bytes"));
        }
        let root = mce_dom(xml, true)?;
        if root.local != "worksheet" || !is_sml(&root.ns) {
            return Err(invalid("expected SpreadsheetML worksheet root"));
        }
        let mut containers = root
            .children
            .iter()
            .filter(|n| n.local == "controls" && is_sml(&n.ns));
        let Some(container) = containers.next() else {
            return Ok(Self::default());
        };
        if containers.next().is_some() {
            return Err(invalid("worksheet has multiple controls collections"));
        }
        check_attrs(container, &[])?;
        if container.children.is_empty() {
            return Err(invalid("controls requires at least one control"));
        }
        if container.children.len() > MAX_CONTROLS {
            return Err(limit("worksheet controls"));
        }
        let mut controls = Vec::with_capacity(container.children.len());
        let mut shape_ids = HashSet::new();
        let mut names = HashSet::new();
        for node in &container.children {
            if node.local != "control" || !is_sml(&node.ns) {
                return Err(invalid("unexpected child in controls"));
            }
            check_attrs(
                node,
                &[
                    ("", "shapeId"),
                    (REL, "id"),
                    (REL_STRICT, "id"),
                    ("", "name"),
                ],
            )?;
            let shape_id = req_u32(node, "", "shapeId")?;
            if !(1..=MAX_SHAPE_ID).contains(&shape_id) || !shape_ids.insert(shape_id) {
                return Err(invalid(
                    "control shapeId must be unique and within Office's supported range",
                ));
            }
            let relationship_id =
                relationship_attr(node, "id")?.ok_or_else(|| invalid("control is missing r:id"))?;
            nonempty(&relationship_id, "control relationship ID")?;
            let name = attr(node, "", "name")?;
            if let Some(name) = name.as_ref() {
                bounded(name, "control name")?;
                if name.chars().count() > MAX_CONTROL_NAME_CHARS {
                    return Err(invalid("control name exceeds Office's 32-character limit"));
                }
                if !names.insert(name.clone()) {
                    return Err(invalid("duplicate control name"));
                }
            }
            if node.children.len() > 1 {
                return Err(invalid("control permits at most one controlPr"));
            }
            let properties = node
                .children
                .first()
                .map(parse_control_properties)
                .transpose()?;
            controls.push(WorksheetControl {
                shape_id,
                relationship_id,
                name,
                properties,
            });
        }
        Ok(Self { controls })
    }

    /// Writes a minimal canonical worksheet containing only the controls collection.
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate_controls(self)?;
        let sml = if strict { SML_STRICT } else { SML };
        let rel = if strict { REL_STRICT } else { REL };
        let xdr = if strict { XDR_STRICT } else { XDR };
        let mut out = String::with_capacity(512 + self.controls.len() * 256);
        out.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"",
        );
        out.push_str(sml);
        out.push_str("\" xmlns:r=\"");
        out.push_str(rel);
        out.push_str("\" xmlns:xdr=\"");
        out.push_str(xdr);
        out.push_str("\"><controls>");
        for control in &self.controls {
            write_control(&mut out, control);
        }
        out.push_str("</controls></worksheet>");
        if out.len() > MAX_OUTPUT_XML {
            return Err(limit("canonical worksheet XML bytes"));
        }
        Ok(out.into_bytes())
    }
}

/// Replaces the direct worksheet `controls` child while preserving unrelated bytes.
///
/// An empty value removes the collection. Controls selected only through an MCE
/// `AlternateContent` branch are rejected because rewriting that branch would not
/// be byte-preserving.
pub fn replace_worksheet_controls_xml(xml: &[u8], controls: &WorksheetControls) -> Result<Vec<u8>> {
    let parsed = WorksheetControls::parse(xml)?;
    if !controls.controls.is_empty() {
        validate_controls(controls)?;
    }
    let location = worksheet_controls_span(xml)?;
    if !parsed.controls.is_empty() && location.span.is_none() {
        return Err(invalid(
            "MCE-selected controls cannot be mutated as a direct worksheet child",
        ));
    }
    let fragment = if controls.controls.is_empty() {
        Vec::new()
    } else {
        controls_fragment(controls, location.strict)?
    };
    let (start, end) = location
        .span
        .unwrap_or((location.insertion, location.insertion));
    let size = xml
        .len()
        .checked_sub(end - start)
        .and_then(|n| n.checked_add(fragment.len()))
        .ok_or_else(|| limit("updated worksheet XML bytes"))?;
    if size > MAX_OUTPUT_XML {
        return Err(limit("updated worksheet XML bytes"));
    }
    let mut out = Vec::with_capacity(size);
    out.extend_from_slice(&xml[..start]);
    out.extend_from_slice(&fragment);
    out.extend_from_slice(&xml[end..]);
    Ok(out)
}

impl ActiveXDescriptor {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML {
            return Err(limit("ActiveX descriptor XML bytes"));
        }
        let root = mce_dom(xml, false)?;
        if root.local != "ocx" || root.ns != AX {
            return Err(invalid("expected ActiveX ocx root"));
        }
        check_attrs(
            &root,
            &[
                (AX, "classid"),
                (AX, "license"),
                (AX, "persistence"),
                (REL, "id"),
                (REL_STRICT, "id"),
            ],
        )?;
        let class_id = req_attr(&root, AX, "classid")?;
        let license = attr(&root, AX, "license")?;
        let persistence = parse_persistence(&req_attr(&root, AX, "persistence")?)?;
        let relationship_id = relationship_attr(&root, "id")?;
        bounded(&class_id, "ActiveX class ID")?;
        if let Some(value) = license.as_ref() {
            bounded(value, "ActiveX license")?;
        }
        let mut count = 0usize;
        let properties = parse_properties(&root.children, 0, &mut count)?;
        let value = Self {
            class_id,
            license,
            persistence,
            relationship_id,
            properties,
        };
        validate_descriptor(&value)?;
        Ok(value)
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        validate_descriptor(self)?;
        let mut out = String::with_capacity(512);
        out.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><ax:ocx xmlns:ax=\"",
        );
        out.push_str(AX);
        out.push_str("\" xmlns:r=\"");
        out.push_str(REL);
        out.push('"');
        qattr(&mut out, "ax:classid", &self.class_id);
        if let Some(v) = self.license.as_deref() {
            qattr(&mut out, "ax:license", v);
        }
        qattr(
            &mut out,
            "ax:persistence",
            persistence_str(self.persistence),
        );
        if let Some(v) = self.relationship_id.as_deref() {
            qattr(&mut out, "r:id", v);
        }
        if self.properties.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            for property in &self.properties {
                write_property(&mut out, property);
            }
            out.push_str("</ax:ocx>");
        }
        if out.len() > MAX_OUTPUT_XML {
            return Err(limit("canonical ActiveX XML bytes"));
        }
        Ok(out.into_bytes())
    }
}

/// Loads every ActiveX control referenced by one worksheet. All payload bytes remain inert.
pub fn load_from_worksheet(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
) -> Result<ActiveXControlSet> {
    let worksheet = package.get_part(worksheet_uri)?;
    if worksheet.content_type() != WORKSHEET_CONTENT_TYPE {
        return Err(OoxmlError::InvalidContentType {
            expected: WORKSHEET_CONTENT_TYPE.into(),
            got: worksheet.content_type().into(),
        });
    }
    let parsed = WorksheetControls::parse(worksheet.blob())?;
    let referenced: HashSet<&str> = parsed
        .controls
        .iter()
        .map(|c| c.relationship_id.as_str())
        .collect();
    for rel in worksheet.rels().iter() {
        if matches!(rel.reltype(), CONTROL_REL | CONTROL_REL_STRICT)
            && !referenced.contains(rel.r_id())
        {
            return Err(relerr(
                "worksheet has an unreferenced ActiveX control relationship",
            ));
        }
    }
    let mut loaded = Vec::with_capacity(parsed.controls.len());
    let mut total_binary = 0usize;
    for control in parsed.controls {
        let preview = if let Some(id) = control
            .properties
            .as_ref()
            .and_then(|p| p.preview_relationship_id.as_deref())
        {
            let rel = worksheet
                .rels()
                .get(id)
                .ok_or_else(|| relerr("control preview relationship is missing"))?;
            if rel.is_external() || !matches!(rel.reltype(), IMAGE_REL | IMAGE_REL_STRICT) {
                return Err(relerr(
                    "control preview must be an internal image relationship",
                ));
            }
            let part = package.get_part(&rel.target_partname()?)?;
            if !part.content_type().starts_with("image/") {
                return Err(OoxmlError::InvalidContentType {
                    expected: "image/*".into(),
                    got: part.content_type().into(),
                });
            }
            if part.blob().len() > MAX_BINARY {
                return Err(limit("ActiveX preview image bytes"));
            }
            total_binary = total_binary
                .checked_add(part.blob().len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
            if total_binary > MAX_TOTAL_BINARY {
                return Err(limit("aggregate ActiveX resource bytes"));
            }
            Some(OpaqueActiveXPreviewImage {
                relationship_id: id.to_string(),
                part_uri: rel.target_partname()?,
                content_type: part.content_type().to_string(),
                bytes: part.blob().to_vec(),
            })
        } else {
            None
        };
        let rel = worksheet
            .rels()
            .get(&control.relationship_id)
            .ok_or_else(|| relerr("control relationship is missing"))?;
        if rel.is_external() || !matches!(rel.reltype(), CONTROL_REL | CONTROL_REL_STRICT) {
            return Err(relerr("control must target an internal ActiveX descriptor"));
        }
        let descriptor_uri = rel.target_partname()?;
        let part = package.get_part(&descriptor_uri)?;
        if part.content_type() != DESCRIPTOR_CONTENT_TYPE {
            return Err(OoxmlError::InvalidContentType {
                expected: DESCRIPTOR_CONTENT_TYPE.into(),
                got: part.content_type().into(),
            });
        }
        let descriptor = ActiveXDescriptor::parse(part.blob())?;
        let mut ids = HashSet::new();
        collect_binary_ids_descriptor(&descriptor, &mut ids)?;
        if part.rels().iter().count() != ids.len() {
            return Err(relerr(
                "ActiveX descriptor has unexpected or duplicate outgoing relationships",
            ));
        }
        let mut binaries = Vec::with_capacity(ids.len());
        for id in ids {
            let binary_rel = part
                .rels()
                .get(&id)
                .ok_or_else(|| relerr("ActiveX binary relationship is missing"))?;
            if binary_rel.is_external() || binary_rel.reltype() != BINARY_REL {
                return Err(relerr(
                    "ActiveX descriptor may relate only to internal ActiveX binaries",
                ));
            }
            let binary_uri = binary_rel.target_partname()?;
            let binary = package.get_part(&binary_uri)?;
            if binary.content_type() != BINARY_CONTENT_TYPE {
                return Err(OoxmlError::InvalidContentType {
                    expected: BINARY_CONTENT_TYPE.into(),
                    got: binary.content_type().into(),
                });
            }
            if binary.rels().iter().next().is_some() {
                return Err(relerr("ActiveX binary part must not have relationships"));
            }
            if binary.blob().len() > MAX_BINARY {
                return Err(limit("ActiveX binary bytes"));
            }
            total_binary = total_binary
                .checked_add(binary.blob().len())
                .ok_or_else(|| limit("aggregate ActiveX binary bytes"))?;
            if total_binary > MAX_TOTAL_BINARY {
                return Err(limit("aggregate ActiveX binary bytes"));
            }
            binaries.push(OpaqueActiveXBinary {
                relationship_id: id,
                part_uri: binary_uri,
                bytes: binary.blob().to_vec(),
            });
        }
        binaries.sort_by(|a, b| a.relationship_id.cmp(&b.relationship_id));
        loaded.push(LoadedActiveXControl {
            control,
            descriptor_uri,
            descriptor,
            binaries,
            preview,
        });
    }
    Ok(ActiveXControlSet { controls: loaded })
}

/// Stores a complete, inert ActiveX graph on a worksheet that has no controls.
pub fn store_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ActiveXControlSet,
) -> Result<()> {
    let prepared = prepare_graph(package, worksheet_uri, value, true)?;
    install_graph(package, worksheet_uri, prepared)
}

/// Atomically replaces the complete inert ActiveX graph of a worksheet.
///
/// An empty set removes the graph. An in-memory package snapshot is used only
/// for rollback; ActiveX payloads are still copied and never interpreted.
pub fn replace_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ActiveXControlSet,
) -> Result<()> {
    if value.controls.is_empty() {
        remove_from_worksheet(package, worksheet_uri)?;
        return Ok(());
    }
    validate_control_set(value)?;
    let snapshot = litchi_opc::PackageWriter::to_bytes(package)?;
    let result = (|| {
        remove_from_worksheet(package, worksheet_uri)?;
        store_on_worksheet(package, worksheet_uri, value)
    })();
    if let Err(error) = result {
        *package = OpcPackage::from_bytes(&snapshot)?;
        return Err(error);
    }
    Ok(())
}

/// Removes the complete ActiveX graph from a worksheet.
///
/// Shared descriptor, binary, or preview parts are retained while any other
/// internal relationship still targets them.
pub fn remove_from_worksheet(package: &mut OpcPackage, worksheet_uri: &PackURI) -> Result<bool> {
    let loaded = load_from_worksheet(package, worksheet_uri)?;
    if loaded.controls.is_empty() {
        return Ok(false);
    }
    let worksheet_xml = package.get_part(worksheet_uri)?.blob().to_vec();
    let updated = replace_worksheet_controls_xml(&worksheet_xml, &WorksheetControls::default())?;
    let control_ids: Vec<String> = loaded
        .controls
        .iter()
        .map(|item| item.control.relationship_id.clone())
        .collect();
    let preview_ids: Vec<String> = loaded
        .controls
        .iter()
        .filter_map(|item| item.preview.as_ref().map(|p| p.relationship_id.clone()))
        .collect();
    let remaining_ids = relationship_ids_in_xml(&updated)?;
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::Other(format!(
            "failed to clear signatures before removing ActiveX controls: {error}"
        ))
    })?;
    {
        let worksheet = package.get_part_mut(worksheet_uri)?;
        for id in &control_ids {
            worksheet.rels_mut().remove(id);
        }
        for id in &preview_ids {
            if !remaining_ids.contains(id) {
                worksheet.rels_mut().remove(id);
            }
        }
        worksheet.set_blob(updated);
    }

    let mut binary_candidates = Vec::new();
    for item in &loaded.controls {
        if !part_has_inbound_relationship(package, &item.descriptor_uri)? {
            package.remove_part(&item.descriptor_uri);
            binary_candidates.extend(item.binaries.iter().map(|b| b.part_uri.clone()));
        }
    }
    for uri in binary_candidates {
        if !part_has_inbound_relationship(package, &uri)? {
            package.remove_part(&uri);
        }
    }
    for preview in loaded
        .controls
        .iter()
        .filter_map(|item| item.preview.as_ref())
    {
        if !part_has_inbound_relationship(package, &preview.part_uri)? {
            package.remove_part(&preview.part_uri);
        }
    }
    Ok(true)
}

struct PreparedGraph {
    worksheet_xml: Vec<u8>,
    strict: bool,
    descriptors: Vec<PreparedDescriptor>,
    resources: Vec<(PackURI, String, Vec<u8>)>,
    worksheet_relationships: Vec<(String, PackURI, bool)>,
}

struct PreparedDescriptor {
    uri: PackURI,
    xml: Vec<u8>,
    relationships: Vec<(String, PackURI)>,
}

fn prepare_graph(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    value: &ActiveXControlSet,
    require_empty: bool,
) -> Result<PreparedGraph> {
    validate_control_set(value)?;
    let worksheet = package.get_part(worksheet_uri)?;
    if worksheet.content_type() != WORKSHEET_CONTENT_TYPE {
        return Err(OoxmlError::InvalidContentType {
            expected: WORKSHEET_CONTENT_TYPE.into(),
            got: worksheet.content_type().into(),
        });
    }
    let existing = WorksheetControls::parse(worksheet.blob())?;
    if require_empty
        && (!existing.controls.is_empty()
            || worksheet
                .rels()
                .iter()
                .any(|rel| matches!(rel.reltype(), CONTROL_REL | CONTROL_REL_STRICT)))
    {
        return Err(invalid("worksheet already has an ActiveX control graph"));
    }
    let controls = WorksheetControls {
        controls: value
            .controls
            .iter()
            .map(|item| item.control.clone())
            .collect(),
    };
    let worksheet_xml = replace_worksheet_controls_xml(worksheet.blob(), &controls)?;
    let strict = worksheet_controls_span(worksheet.blob())?.strict;

    let mut occupied_ids: HashSet<String> = worksheet
        .rels()
        .iter()
        .map(|r| r.r_id().to_string())
        .collect();
    let mut part_uris = HashSet::new();
    let mut descriptors = Vec::with_capacity(value.controls.len());
    let mut resources = Vec::new();
    let mut worksheet_relationships = Vec::new();
    for item in &value.controls {
        validate_rel_id(&item.control.relationship_id)?;
        if !occupied_ids.insert(item.control.relationship_id.clone()) {
            return Err(relerr("duplicate or occupied worksheet relationship ID"));
        }
        validate_part_location(&item.descriptor_uri, "/xl/activeX/", "ActiveX descriptor")?;
        reserve_new_part(package, &mut part_uris, &item.descriptor_uri)?;
        let descriptor_xml = item.descriptor.to_xml()?;
        let mut expected_ids = HashSet::new();
        collect_binary_ids_descriptor(&item.descriptor, &mut expected_ids)?;
        let actual_ids: HashSet<String> = item
            .binaries
            .iter()
            .map(|binary| binary.relationship_id.clone())
            .collect();
        if expected_ids != actual_ids || actual_ids.len() != item.binaries.len() {
            return Err(relerr(
                "descriptor relationship IDs must exactly match supplied binaries",
            ));
        }
        let mut descriptor_rels = Vec::with_capacity(item.binaries.len());
        for binary in &item.binaries {
            validate_rel_id(&binary.relationship_id)?;
            validate_part_location(&binary.part_uri, "/xl/activeX/", "ActiveX binary")?;
            if binary.bytes.len() > MAX_BINARY {
                return Err(limit("ActiveX binary bytes"));
            }
            reserve_new_part(package, &mut part_uris, &binary.part_uri)?;
            descriptor_rels.push((binary.relationship_id.clone(), binary.part_uri.clone()));
            resources.push((
                binary.part_uri.clone(),
                BINARY_CONTENT_TYPE.into(),
                binary.bytes.clone(),
            ));
        }
        worksheet_relationships.push((
            item.control.relationship_id.clone(),
            item.descriptor_uri.clone(),
            false,
        ));
        match (&item.control.properties, &item.preview) {
            (Some(properties), Some(preview)) => {
                if properties.preview_relationship_id.as_deref()
                    != Some(preview.relationship_id.as_str())
                {
                    return Err(relerr(
                        "control preview relationship ID does not match supplied preview",
                    ));
                }
                validate_rel_id(&preview.relationship_id)?;
                if !occupied_ids.insert(preview.relationship_id.clone()) {
                    return Err(relerr("duplicate or occupied worksheet relationship ID"));
                }
                validate_part_location(&preview.part_uri, "/xl/media/", "ActiveX preview")?;
                if !preview.content_type.starts_with("image/") {
                    return Err(invalid("ActiveX preview content type must be image/*"));
                }
                if preview.bytes.len() > MAX_BINARY {
                    return Err(limit("ActiveX preview image bytes"));
                }
                reserve_new_part(package, &mut part_uris, &preview.part_uri)?;
                worksheet_relationships.push((
                    preview.relationship_id.clone(),
                    preview.part_uri.clone(),
                    true,
                ));
                resources.push((
                    preview.part_uri.clone(),
                    preview.content_type.clone(),
                    preview.bytes.clone(),
                ));
            },
            (Some(properties), None) if properties.preview_relationship_id.is_some() => {
                return Err(relerr("control references a preview that was not supplied"));
            },
            (_, Some(_)) => return Err(relerr("supplied preview is not referenced by controlPr")),
            _ => {},
        }
        descriptors.push(PreparedDescriptor {
            uri: item.descriptor_uri.clone(),
            xml: descriptor_xml,
            relationships: descriptor_rels,
        });
    }
    Ok(PreparedGraph {
        worksheet_xml,
        strict,
        descriptors,
        resources,
        worksheet_relationships,
    })
}

fn install_graph(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    prepared: PreparedGraph,
) -> Result<()> {
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::Other(format!(
            "failed to clear signatures before storing ActiveX controls: {error}"
        ))
    })?;
    for (uri, content_type, bytes) in prepared.resources {
        package.try_add_part(Box::new(BlobPart::new(uri, content_type, bytes)))?;
    }
    for descriptor in prepared.descriptors {
        let mut part = BlobPart::new(
            descriptor.uri.clone(),
            DESCRIPTOR_CONTENT_TYPE.into(),
            descriptor.xml,
        );
        for (id, target) in descriptor.relationships {
            part.rels_mut().try_add_relationship(
                BINARY_REL.into(),
                target.relative_ref(descriptor.uri.base_uri()),
                id,
                TargetMode::Internal,
            )?;
        }
        package.try_add_part(Box::new(part))?;
    }
    let worksheet = package.get_part_mut(worksheet_uri)?;
    for (id, target, preview) in prepared.worksheet_relationships {
        worksheet.rels_mut().try_add_relationship(
            if preview {
                if prepared.strict {
                    IMAGE_REL_STRICT
                } else {
                    IMAGE_REL
                }
            } else if prepared.strict {
                CONTROL_REL_STRICT
            } else {
                CONTROL_REL
            }
            .into(),
            target.relative_ref(worksheet_uri.base_uri()),
            id,
            TargetMode::Internal,
        )?;
    }
    worksheet.set_blob(prepared.worksheet_xml);
    Ok(())
}

fn controls_fragment(value: &WorksheetControls, strict: bool) -> Result<Vec<u8>> {
    validate_controls(value)?;
    let rel = if strict { REL_STRICT } else { REL };
    let xdr = if strict { XDR_STRICT } else { XDR };
    let mut out = String::with_capacity(256 + value.controls.len() * 256);
    out.push_str("<controls xmlns:r=\"");
    out.push_str(rel);
    out.push_str("\" xmlns:xdr=\"");
    out.push_str(xdr);
    out.push_str("\">");
    for control in &value.controls {
        write_control(&mut out, control);
    }
    out.push_str("</controls>");
    if out.len() > MAX_OUTPUT_XML {
        return Err(limit("controls XML bytes"));
    }
    Ok(out.into_bytes())
}

struct WorksheetControlsLocation {
    strict: bool,
    span: Option<(usize, usize)>,
    insertion: usize,
}

fn worksheet_controls_span(xml: &[u8]) -> Result<WorksheetControlsLocation> {
    if xml.len() > MAX_XML {
        return Err(limit("worksheet XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut strict = false;
    let mut root = false;
    let mut controls_start = None;
    let mut controls_span = None;
    let mut insertion = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("worksheet XML offset overflow"))?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                let namespace = resolved_ns(&resolved)?;
                if depth == 0 {
                    if element.local_name().as_ref() != b"worksheet"
                        || !matches!(namespace.as_str(), SML | SML_STRICT)
                    {
                        return Err(invalid("expected SpreadsheetML worksheet root"));
                    }
                    strict = namespace == SML_STRICT;
                    root = true;
                } else if depth == 1 && namespace == if strict { SML_STRICT } else { SML } {
                    match element.local_name().as_ref() {
                        b"controls" if controls_start.replace(start).is_some() => {
                            return Err(invalid("worksheet has multiple direct controls"));
                        },
                        b"controls" => {},
                        b"webPublishItems" | b"tableParts" | b"extLst" => {
                            insertion.get_or_insert(start);
                        },
                        _ => {},
                    }
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("worksheet XML depth"));
                }
            },
            Event::Empty(element) => {
                let namespace = resolved_ns(&resolved)?;
                if depth == 1 && namespace == if strict { SML_STRICT } else { SML } {
                    if element.local_name().as_ref() == b"controls" {
                        return Err(invalid("empty controls collection is not valid"));
                    }
                    if matches!(
                        element.local_name().as_ref(),
                        b"webPublishItems" | b"tableParts" | b"extLst"
                    ) {
                        insertion.get_or_insert(start);
                    }
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected worksheet closing element"));
                }
                if depth == 2 && element.local_name().as_ref() == b"controls" {
                    let start = controls_start
                        .take()
                        .ok_or_else(|| invalid("mismatched controls closing element"))?;
                    let end = usize::try_from(reader.buffer_position())
                        .map_err(|_| invalid("worksheet XML offset overflow"))?;
                    controls_span = Some((start, end));
                }
                if depth == 1 && element.local_name().as_ref() == b"worksheet" {
                    insertion.get_or_insert(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 || controls_start.is_some() {
        return Err(invalid("invalid worksheet XML"));
    }
    Ok(WorksheetControlsLocation {
        strict,
        span: controls_span,
        insertion: insertion.ok_or_else(|| invalid("missing worksheet closing element"))?,
    })
}

fn validate_control_set(value: &ActiveXControlSet) -> Result<()> {
    if value.controls.is_empty() || value.controls.len() > MAX_CONTROLS {
        return Err(invalid("ActiveX control set requires 1..65535 controls"));
    }
    validate_controls(&WorksheetControls {
        controls: value
            .controls
            .iter()
            .map(|item| item.control.clone())
            .collect(),
    })?;
    let mut total = 0usize;
    for item in &value.controls {
        validate_descriptor(&item.descriptor)?;
        for binary in &item.binaries {
            total = total
                .checked_add(binary.bytes.len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
        }
        if let Some(preview) = item.preview.as_ref() {
            total = total
                .checked_add(preview.bytes.len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
        }
    }
    if total > MAX_TOTAL_BINARY {
        return Err(limit("aggregate ActiveX resource bytes"));
    }
    Ok(())
}

fn reserve_new_part(
    package: &OpcPackage,
    reserved: &mut HashSet<PackURI>,
    uri: &PackURI,
) -> Result<()> {
    if reserved.iter().any(|other| other.is_equivalent_to(uri)) {
        return Err(invalid("ActiveX graph contains conflicting part names"));
    }
    package.validate_new_part_name(uri)?;
    reserved.insert(uri.clone());
    Ok(())
}

fn validate_part_location(uri: &PackURI, prefix: &str, kind: &str) -> Result<()> {
    let Some(filename) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{kind} must be stored below {prefix}")));
    };
    if filename.is_empty() || filename.contains('/') {
        return Err(invalid(format!(
            "{kind} must be a direct child of {prefix}"
        )));
    }
    Ok(())
}

fn validate_rel_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    if !matches!(bytes.next(), Some(b'A'..=b'Z' | b'a'..=b'z' | b'_'))
        || !bytes.all(|b| matches!(b, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'.' | b'-'))
    {
        return Err(relerr("relationship ID must be an XML NCName"));
    }
    bounded(value, "relationship ID")
}

fn relationship_ids_in_xml(xml: &[u8]) -> Result<HashSet<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut ids = HashSet::new();
    loop {
        let resolver = reader.resolver().clone();
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes().with_checks(true) {
                    let attribute =
                        attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    let (namespace, _) = resolver.resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(Namespace(value)) if matches!(value, b"http://schemas.openxmlformats.org/officeDocument/2006/relationships" | b"http://purl.oclc.org/ooxml/officeDocument/relationships"))
                    {
                        ids.insert(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                                .into_owned(),
                        );
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(ids)
}

fn part_has_inbound_relationship(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external() && relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

#[derive(Debug, Clone)]
struct Attribute {
    ns: String,
    local: String,
    value: String,
}
#[derive(Debug, Clone)]
struct Node {
    ns: String,
    local: String,
    attrs: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

fn mce_dom(xml: &[u8], worksheet: bool) -> Result<Node> {
    let mut caps = MceCapabilities::ooxml_baseline();
    caps.understand_namespace(X14)
        .understand_namespace(XDR)
        .understand_namespace(XDR_STRICT)
        .understand_namespace(AX);
    let limits = MceLimits {
        max_input_bytes: MAX_XML,
        max_output_bytes: MAX_OUTPUT_XML,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &caps, &limits)?;
    parse_dom(processed.xml.as_ref(), worksheet)
}

fn parse_dom(xml: &[u8], worksheet: bool) -> Result<Node> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    let mut buffer = Vec::new();
    loop {
        let decoder = reader.decoder();
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        match event {
            Event::Start(e) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                let namespace = resolved_ns(&resolved)?;
                drop(resolved);
                let resolver = reader.resolver().clone();
                stack.push(make_node(&resolver, namespace, &e, decoder)?);
            },
            Event::Empty(e) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                let namespace = resolved_ns(&resolved)?;
                drop(resolved);
                let resolver = reader.resolver().clone();
                let node = make_node(&resolver, namespace, &e, decoder)?;
                append_node(&mut stack, &mut root, node)?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
                append_node(&mut stack, &mut root, node)?;
            },
            Event::Text(e) => {
                let decoded = e.decode().map_err(|x| OoxmlError::Xml(x.to_string()))?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|x| OoxmlError::Xml(x.to_string()))?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_OUTPUT_XML {
                    return Err(limit("XML text bytes"));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::CData(e) => {
                let value = e.decode().map_err(|x| OoxmlError::Xml(x.to_string()))?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_OUTPUT_XML {
                    return Err(limit("XML text bytes"));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(e) => {
                let name = e.decode().map_err(|x| OoxmlError::Xml(x.to_string()))?;
                let value = if let Some(c) = e
                    .resolve_char_ref()
                    .map_err(|x| OoxmlError::Xml(x.to_string()))?
                {
                    c.to_string()
                } else {
                    match name.as_ref() {
                        "amp" => "&",
                        "lt" => "<",
                        "gt" => ">",
                        "apos" => "'",
                        "quot" => "\"",
                        _ => return Err(invalid("custom XML entity is rejected")),
                    }
                    .into()
                };
                if let Some(parent) = stack.last_mut() {
                    parent.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    let root = root.ok_or_else(|| invalid("missing XML root"))?;
    if worksheet && root.children.len() > MAX_NODES {
        return Err(limit("worksheet nodes"));
    }
    Ok(root)
}

fn make_node(
    resolver: &NamespaceResolver,
    ns: String,
    e: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Node> {
    let local = std::str::from_utf8(e.local_name().as_ref())
        .map_err(|x| OoxmlError::Xml(x.to_string()))?
        .to_string();
    let mut attrs = Vec::new();
    for item in e.attributes().with_checks(true) {
        let item = item.map_err(|x| OoxmlError::Xml(x.to_string()))?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, _) = resolver.resolve_attribute(item.key);
        let ans = resolved_ns(&resolved)?;
        let alocal = std::str::from_utf8(item.key.local_name().as_ref())
            .map_err(|x| OoxmlError::Xml(x.to_string()))?
            .to_string();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|x| OoxmlError::Xml(x.to_string()))?
            .into_owned();
        bounded(&value, "XML attribute")?;
        if attrs
            .iter()
            .any(|a: &Attribute| a.ns == ans && a.local == alocal)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attrs.push(Attribute {
            ns: ans,
            local: alocal,
            value,
        });
    }
    Ok(Node {
        ns,
        local,
        attrs,
        children: Vec::new(),
        text: String::new(),
    })
}

fn resolved_ns(value: &ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(v)) => std::str::from_utf8(v)
            .map(str::to_string)
            .map_err(|x| OoxmlError::Xml(x.to_string())),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}
fn append_node(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn parse_control_properties(node: &Node) -> Result<ControlProperties> {
    if node.local != "controlPr" || !is_sml(&node.ns) {
        return Err(invalid("unexpected control child"));
    }
    check_attrs(
        node,
        &[
            ("", "locked"),
            ("", "defaultSize"),
            ("", "print"),
            ("", "disabled"),
            ("", "recalcAlways"),
            ("", "uiObject"),
            ("", "autoFill"),
            ("", "autoLine"),
            ("", "autoPict"),
            ("", "macro"),
            ("", "altText"),
            (REL, "id"),
            (REL_STRICT, "id"),
        ],
    )?;
    if node.children.len() != 1 {
        return Err(invalid("controlPr requires exactly one anchor"));
    }
    let anchor = parse_anchor(&node.children[0])?;
    let macro_name = attr(node, "", "macro")?;
    let alternate_text = attr(node, "", "altText")?;
    if let Some(v) = macro_name.as_ref() {
        bounded(v, "control macro name")?;
    }
    if let Some(v) = alternate_text.as_ref() {
        bounded(v, "control alternate text")?;
    }
    Ok(ControlProperties {
        anchor,
        locked: opt_bool(node, "locked")?,
        default_size: opt_bool(node, "defaultSize")?,
        print: opt_bool(node, "print")?,
        disabled: opt_bool(node, "disabled")?,
        recalc_always: opt_bool(node, "recalcAlways")?,
        ui_object: opt_bool(node, "uiObject")?,
        auto_fill: opt_bool(node, "autoFill")?,
        auto_line: opt_bool(node, "autoLine")?,
        auto_picture: opt_bool(node, "autoPict")?,
        macro_name,
        alternate_text,
        preview_relationship_id: relationship_attr(node, "id")?,
    })
}
fn parse_anchor(node: &Node) -> Result<ObjectAnchor> {
    if node.local != "anchor" || !is_sml(&node.ns) {
        return Err(invalid("controlPr requires anchor"));
    }
    check_attrs(node, &[("", "moveWithCells"), ("", "sizeWithCells")])?;
    if node.children.len() != 2
        || node.children[0].local != "from"
        || node.children[1].local != "to"
        || !is_sml(&node.children[0].ns)
        || !is_sml(&node.children[1].ns)
    {
        return Err(invalid("anchor requires from then to"));
    }
    Ok(ObjectAnchor {
        from: parse_marker(&node.children[0])?,
        to: parse_marker(&node.children[1])?,
        move_with_cells: opt_bool(node, "moveWithCells")?,
        size_with_cells: opt_bool(node, "sizeWithCells")?,
    })
}
fn parse_marker(node: &Node) -> Result<Marker> {
    check_attrs(node, &[])?;
    let expected = ["col", "colOff", "row", "rowOff"];
    if node.children.len() != expected.len() {
        return Err(invalid("anchor marker requires col, colOff, row, rowOff"));
    }
    for (child, expected) in node.children.iter().zip(expected) {
        if child.local != expected
            || !is_xdr(&child.ns)
            || !child.children.is_empty()
            || !child.attrs.is_empty()
        {
            return Err(invalid("invalid anchor marker grammar"));
        }
    }
    Ok(Marker {
        column: text_i32(&node.children[0])?,
        column_offset: text_i64(&node.children[1])?,
        row: text_i32(&node.children[2])?,
        row_offset: text_i64(&node.children[3])?,
    })
}

fn parse_properties(
    nodes: &[Node],
    depth: usize,
    count: &mut usize,
) -> Result<Vec<ActiveXProperty>> {
    if depth >= MAX_DEPTH {
        return Err(limit("ActiveX property nesting"));
    }
    let mut result = Vec::with_capacity(nodes.len());
    let mut names = HashSet::new();
    for node in nodes {
        *count = count
            .checked_add(1)
            .ok_or_else(|| limit("ActiveX properties"))?;
        if *count > MAX_PROPERTIES {
            return Err(limit("ActiveX properties"));
        }
        if node.local != "ocxPr" || node.ns != AX {
            return Err(invalid("unexpected ActiveX descriptor child"));
        }
        check_attrs(node, &[(AX, "name"), (AX, "value")])?;
        let name = req_attr(node, AX, "name")?;
        bounded(&name, "ActiveX property name")?;
        if !names.insert(name.clone()) {
            return Err(invalid("duplicate ActiveX property name"));
        }
        let value = attr(node, AX, "value")?;
        if let Some(v) = value.as_ref() {
            bounded(v, "ActiveX property value")?;
        }
        if node.children.len() > 1 {
            return Err(invalid("ActiveX property permits at most one object child"));
        }
        let object = node
            .children
            .first()
            .map(|child| parse_property_object(child, depth + 1, count))
            .transpose()?;
        if value.is_some() && object.is_some() {
            return Err(invalid(
                "ActiveX property value cannot coexist with font or picture",
            ));
        }
        result.push(ActiveXProperty {
            name,
            value,
            object,
        });
    }
    Ok(result)
}
fn parse_property_object(
    node: &Node,
    depth: usize,
    count: &mut usize,
) -> Result<ActiveXPropertyObject> {
    if node.ns != AX {
        return Err(invalid("invalid ActiveX property object namespace"));
    }
    match node.local.as_str() {
        "font" => {
            check_attrs(
                node,
                &[(AX, "persistence"), (REL, "id"), (REL_STRICT, "id")],
            )?;
            let persistence = attr(node, AX, "persistence")?
                .map(|v| parse_persistence(&v))
                .transpose()?;
            let font = ActiveXFont {
                persistence,
                relationship_id: relationship_attr(node, "id")?,
                properties: parse_properties(&node.children, depth, count)?,
            };
            validate_font(&font)?;
            Ok(ActiveXPropertyObject::Font(font))
        },
        "picture" => {
            check_attrs(node, &[(REL, "id"), (REL_STRICT, "id")])?;
            if !node.children.is_empty() {
                return Err(invalid("ActiveX picture must be empty"));
            }
            Ok(ActiveXPropertyObject::Picture(ActiveXPicture {
                relationship_id: relationship_attr(node, "id")?,
            }))
        },
        _ => Err(invalid("ActiveX property child must be font or picture")),
    }
}

fn validate_controls(value: &WorksheetControls) -> Result<()> {
    if value.controls.is_empty() || value.controls.len() > MAX_CONTROLS {
        return Err(invalid("controls requires 1..65535 controls"));
    }
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for c in &value.controls {
        if !(1..=MAX_SHAPE_ID).contains(&c.shape_id) || !ids.insert(c.shape_id) {
            return Err(invalid(
                "control shapeId must be unique and within Office's supported range",
            ));
        }
        nonempty(&c.relationship_id, "control relationship ID")?;
        bounded(&c.relationship_id, "control relationship ID")?;
        if let Some(n) = c.name.as_ref() {
            bounded(n, "control name")?;
            if n.chars().count() > MAX_CONTROL_NAME_CHARS {
                return Err(invalid("control name exceeds Office's 32-character limit"));
            }
            if !names.insert(n) {
                return Err(invalid("duplicate control name"));
            }
        }
        if let Some(p) = c.properties.as_ref() {
            validate_control_properties(p)?;
        }
    }
    Ok(())
}
fn validate_control_properties(value: &ControlProperties) -> Result<()> {
    if let Some(v) = value.macro_name.as_ref() {
        bounded(v, "control macro name")?;
    }
    if let Some(v) = value.alternate_text.as_ref() {
        bounded(v, "control alternate text")?;
    }
    if let Some(v) = value.preview_relationship_id.as_ref() {
        nonempty(v, "preview relationship ID")?;
        bounded(v, "preview relationship ID")?;
    }
    Ok(())
}
fn validate_descriptor(value: &ActiveXDescriptor) -> Result<()> {
    nonempty(&value.class_id, "ActiveX class ID")?;
    bounded(&value.class_id, "ActiveX class ID")?;
    if let Some(v) = value.license.as_ref() {
        bounded(v, "ActiveX license")?;
    }
    match value.persistence {
        Persistence::PropertyBag => {
            if value.properties.is_empty() || value.relationship_id.is_some() {
                return Err(invalid(
                    "property-bag ActiveX requires properties and forbids r:id",
                ));
            }
        },
        _ => {
            if !value.properties.is_empty()
                || value.relationship_id.as_deref().is_none_or(str::is_empty)
            {
                return Err(invalid(
                    "binary ActiveX persistence requires r:id and forbids properties",
                ));
            }
        },
    }
    let mut count = 0usize;
    validate_properties(&value.properties, 0, &mut count)
}
fn validate_properties(values: &[ActiveXProperty], depth: usize, count: &mut usize) -> Result<()> {
    if depth >= MAX_DEPTH {
        return Err(limit("ActiveX property nesting"));
    }
    let mut names = HashSet::new();
    for p in values {
        *count = count
            .checked_add(1)
            .ok_or_else(|| limit("ActiveX properties"))?;
        if *count > MAX_PROPERTIES {
            return Err(limit("ActiveX properties"));
        }
        nonempty(&p.name, "ActiveX property name")?;
        bounded(&p.name, "ActiveX property name")?;
        if !names.insert(&p.name) {
            return Err(invalid("duplicate ActiveX property name"));
        }
        if let Some(v) = p.value.as_ref() {
            bounded(v, "ActiveX property value")?;
        }
        if p.value.is_some() && p.object.is_some() {
            return Err(invalid("ActiveX property value cannot coexist with object"));
        }
        if let Some(ActiveXPropertyObject::Font(f)) = p.object.as_ref() {
            validate_font(f)?;
            validate_properties(&f.properties, depth + 1, count)?;
        }
    }
    Ok(())
}
fn validate_font(font: &ActiveXFont) -> Result<()> {
    match font.persistence {
        Some(Persistence::PropertyBag) => {
            if font.properties.is_empty() || font.relationship_id.is_some() {
                return Err(invalid(
                    "property-bag font requires properties and forbids r:id",
                ));
            }
        },
        Some(_) => {
            if !font.properties.is_empty()
                || font.relationship_id.as_deref().is_none_or(str::is_empty)
            {
                return Err(invalid(
                    "binary font persistence requires r:id and forbids properties",
                ));
            }
        },
        None => {
            if font.relationship_id.is_some() {
                return Err(invalid("font r:id requires a binary persistence mode"));
            }
        },
    }
    Ok(())
}

fn collect_binary_ids_descriptor(
    value: &ActiveXDescriptor,
    ids: &mut HashSet<String>,
) -> Result<()> {
    if let Some(id) = value.relationship_id.as_ref() {
        insert_rel_id(ids, id)?;
    }
    collect_binary_ids_properties(&value.properties, ids)
}
fn collect_binary_ids_properties(
    values: &[ActiveXProperty],
    ids: &mut HashSet<String>,
) -> Result<()> {
    for value in values {
        match value.object.as_ref() {
            Some(ActiveXPropertyObject::Font(font)) => {
                if let Some(id) = font.relationship_id.as_ref() {
                    insert_rel_id(ids, id)?;
                }
                collect_binary_ids_properties(&font.properties, ids)?;
            },
            Some(ActiveXPropertyObject::Picture(picture)) => {
                if let Some(id) = picture.relationship_id.as_ref() {
                    insert_rel_id(ids, id)?;
                }
            },
            None => {},
        }
    }
    Ok(())
}
fn insert_rel_id(ids: &mut HashSet<String>, id: &str) -> Result<()> {
    if !ids.insert(id.to_string()) {
        Err(relerr(
            "ActiveX relationship ID is referenced more than once",
        ))
    } else {
        Ok(())
    }
}

fn write_control(out: &mut String, c: &WorksheetControl) {
    out.push_str("<control");
    qattr(out, "shapeId", &c.shape_id.to_string());
    qattr(out, "r:id", &c.relationship_id);
    if let Some(v) = c.name.as_deref() {
        qattr(out, "name", v);
    }
    let Some(p) = c.properties.as_ref() else {
        out.push_str("/>");
        return;
    };
    out.push_str("><controlPr");
    bool_attr(out, "locked", p.locked);
    bool_attr(out, "defaultSize", p.default_size);
    bool_attr(out, "print", p.print);
    bool_attr(out, "disabled", p.disabled);
    bool_attr(out, "recalcAlways", p.recalc_always);
    bool_attr(out, "uiObject", p.ui_object);
    bool_attr(out, "autoFill", p.auto_fill);
    bool_attr(out, "autoLine", p.auto_line);
    bool_attr(out, "autoPict", p.auto_picture);
    if let Some(v) = p.macro_name.as_deref() {
        qattr(out, "macro", v);
    }
    if let Some(v) = p.alternate_text.as_deref() {
        qattr(out, "altText", v);
    }
    if let Some(v) = p.preview_relationship_id.as_deref() {
        qattr(out, "r:id", v);
    }
    out.push_str("><anchor");
    bool_attr(out, "moveWithCells", p.anchor.move_with_cells);
    bool_attr(out, "sizeWithCells", p.anchor.size_with_cells);
    out.push('>');
    write_marker(out, "from", &p.anchor.from);
    write_marker(out, "to", &p.anchor.to);
    out.push_str("</anchor></controlPr></control>");
}
fn write_marker(out: &mut String, name: &str, m: &Marker) {
    out.push('<');
    out.push_str(name);
    out.push_str("><xdr:col>");
    out.push_str(&m.column.to_string());
    out.push_str("</xdr:col><xdr:colOff>");
    out.push_str(&m.column_offset.to_string());
    out.push_str("</xdr:colOff><xdr:row>");
    out.push_str(&m.row.to_string());
    out.push_str("</xdr:row><xdr:rowOff>");
    out.push_str(&m.row_offset.to_string());
    out.push_str("</xdr:rowOff></");
    out.push_str(name);
    out.push('>');
}
fn write_property(out: &mut String, p: &ActiveXProperty) {
    out.push_str("<ax:ocxPr");
    qattr(out, "ax:name", &p.name);
    if let Some(v) = p.value.as_deref() {
        qattr(out, "ax:value", v);
    }
    match p.object.as_ref() {
        None => out.push_str("/>"),
        Some(o) => {
            out.push('>');
            write_object(out, o);
            out.push_str("</ax:ocxPr>");
        },
    }
}
fn write_object(out: &mut String, value: &ActiveXPropertyObject) {
    match value {
        ActiveXPropertyObject::Picture(p) => {
            out.push_str("<ax:picture");
            if let Some(v) = p.relationship_id.as_deref() {
                qattr(out, "r:id", v);
            }
            out.push_str("/>");
        },
        ActiveXPropertyObject::Font(f) => {
            out.push_str("<ax:font");
            if let Some(v) = f.persistence {
                qattr(out, "ax:persistence", persistence_str(v));
            }
            if let Some(v) = f.relationship_id.as_deref() {
                qattr(out, "r:id", v);
            }
            if f.properties.is_empty() {
                out.push_str("/>");
            } else {
                out.push('>');
                for p in &f.properties {
                    write_property(out, p);
                }
                out.push_str("</ax:font>");
            }
        },
    }
}
fn bool_attr(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(v) = value {
        qattr(out, name, if v { "1" } else { "0" });
    }
}
fn qattr(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    escape(out, value);
    out.push('"');
}
fn escape(out: &mut String, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            '\r' => out.push_str("&#xD;"),
            '\n' => out.push_str("&#xA;"),
            '\t' => out.push_str("&#x9;"),
            _ => out.push(c),
        }
    }
}

fn check_attrs(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for a in &node.attrs {
        if !allowed
            .iter()
            .any(|(ns, local)| *ns == a.ns && *local == a.local)
        {
            return Err(invalid(format!(
                "unexpected attribute {{{}}}{} on {}",
                a.ns, a.local, node.local
            )));
        }
    }
    Ok(())
}
fn attr(node: &Node, ns: &str, local: &str) -> Result<Option<String>> {
    let mut values = node.attrs.iter().filter(|a| a.ns == ns && a.local == local);
    let value = values.next().map(|a| a.value.clone());
    if values.next().is_some() {
        return Err(invalid("duplicate attribute"));
    }
    Ok(value)
}
fn req_attr(node: &Node, ns: &str, local: &str) -> Result<String> {
    attr(node, ns, local)?
        .filter(|v| !v.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing required {}", node.local, local)))
}
fn relationship_attr(node: &Node, local: &str) -> Result<Option<String>> {
    let a = attr(node, REL, local)?;
    let b = attr(node, REL_STRICT, local)?;
    if a.is_some() && b.is_some() {
        return Err(invalid("duplicate relationship attribute"));
    }
    Ok(a.or(b))
}
fn req_u32(node: &Node, ns: &str, local: &str) -> Result<u32> {
    req_attr(node, ns, local)?
        .parse()
        .map_err(|_| invalid(format!("invalid unsigned integer {local}")))
}
fn opt_bool(node: &Node, local: &str) -> Result<Option<bool>> {
    attr(node, "", local)?.map(|v| parse_bool(&v)).transpose()
}
fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid XML boolean")),
    }
}
fn text_i32(node: &Node) -> Result<i32> {
    node.text
        .trim()
        .parse()
        .map_err(|_| invalid("invalid anchor signed integer"))
}
fn text_i64(node: &Node) -> Result<i64> {
    node.text
        .trim()
        .parse()
        .map_err(|_| invalid("invalid anchor coordinate"))
}
fn parse_persistence(value: &str) -> Result<Persistence> {
    match value {
        "persistPropertyBag" => Ok(Persistence::PropertyBag),
        "persistStream" => Ok(Persistence::Stream),
        "persistStreamInit" => Ok(Persistence::StreamInit),
        "persistStorage" => Ok(Persistence::Storage),
        _ => Err(invalid("invalid ActiveX persistence")),
    }
}
fn persistence_str(value: Persistence) -> &'static str {
    match value {
        Persistence::PropertyBag => "persistPropertyBag",
        Persistence::Stream => "persistStream",
        Persistence::StreamInit => "persistStreamInit",
        Persistence::Storage => "persistStorage",
    }
}
fn is_sml(ns: &str) -> bool {
    matches!(ns, SML | SML_STRICT)
}
fn is_xdr(ns: &str) -> bool {
    matches!(ns, XDR | XDR_STRICT)
}
fn bounded(value: &str, what: &str) -> Result<()> {
    if value.len() > MAX_STRING {
        Err(limit(what))
    } else {
        Ok(())
    }
}
fn nonempty(value: &str, what: &str) -> Result<()> {
    if value.is_empty() {
        Err(invalid(format!("{what} must not be empty")))
    } else {
        Ok(())
    }
}
fn invalid(value: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(value.into())
}
fn relerr(value: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidRelationship(value.into())
}
fn limit(value: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("ActiveX resource limit exceeded: {}", value.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, PackageWriter, Part};

    const WS: &str = "/xl/worksheets/sheet1.xml";
    fn fixture(bytes: &[u8]) -> ActiveXControlSet {
        let p = OpcPackage::from_bytes(bytes).unwrap();
        load_from_worksheet(&p, &PackURI::new(WS).unwrap()).unwrap()
    }

    #[test]
    fn libreoffice_stream_init_is_opaque_and_anchored() {
        let set = fixture(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/activex_checkbox.xlsx"
        ));
        assert_eq!(set.controls.len(), 1);
        let item = &set.controls[0];
        assert_eq!(item.control.shape_id, 1025);
        assert_eq!(item.descriptor.persistence, Persistence::StreamInit);
        assert_eq!(item.binaries.len(), 1);
        assert_eq!(item.binaries[0].bytes.len(), 116);
        assert_eq!(
            item.control.properties.as_ref().unwrap().anchor.from.column,
            1
        );
        assert_eq!(item.binaries[0].bytes, item.binaries[0].bytes.clone());
    }

    #[test]
    fn libreoffice_radio_buttons_resolve_five_inert_payloads() {
        let set = fixture(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf111980_radioButtons.xlsx"
        ));
        assert_eq!(set.controls.len(), 5);
        assert!(
            set.controls
                .iter()
                .all(|c| c.descriptor.persistence == Persistence::StreamInit
                    && c.binaries.len() == 1)
        );
    }

    #[test]
    fn poi_property_bag_header_and_footer() {
        for bytes in [
            include_bytes!(
                "../../../../test-data/poi/test-data/spreadsheet/45540_form_Header.xlsx"
            )
            .as_slice(),
            include_bytes!(
                "../../../../test-data/poi/test-data/spreadsheet/45540_form_Footer.xlsx"
            )
            .as_slice(),
        ] {
            let set = fixture(bytes);
            assert_eq!(set.controls.len(), 40);
            assert!(
                set.controls
                    .iter()
                    .all(|c| c.descriptor.persistence == Persistence::PropertyBag
                        && c.binaries.is_empty()
                        && !c.descriptor.properties.is_empty())
            );
        }
    }

    #[test]
    fn strict_nested_mce_controls_roundtrip() {
        let xml = format!(
            r#"<worksheet xmlns="{SML_STRICT}" xmlns:r="{REL_STRICT}" xmlns:xdr="{XDR_STRICT}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="{X14}"><mc:AlternateContent><mc:Choice Requires="x14"><controls><mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="7" r:id="rId1" name="safe"><controlPr macro="inert"><anchor moveWithCells="true"><from><xdr:col>1</xdr:col><xdr:colOff>2</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>4</xdr:rowOff></from><to><xdr:col>5</xdr:col><xdr:colOff>6</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>8</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls></mc:Choice><mc:Fallback/></mc:AlternateContent></worksheet>"#
        );
        let controls = WorksheetControls::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            controls.controls[0]
                .properties
                .as_ref()
                .unwrap()
                .macro_name
                .as_deref(),
            Some("inert")
        );
        let canonical = controls.to_xml(true).unwrap();
        assert_eq!(WorksheetControls::parse(&canonical).unwrap(), controls);
    }

    #[test]
    fn descriptor_persistence_variants_and_nested_objects() {
        for mode in [
            Persistence::Stream,
            Persistence::StreamInit,
            Persistence::Storage,
        ] {
            let d = ActiveXDescriptor {
                class_id: "{inert}".into(),
                license: Some("not-used".into()),
                persistence: mode,
                relationship_id: Some("rId1".into()),
                properties: vec![],
            };
            assert_eq!(ActiveXDescriptor::parse(&d.to_xml().unwrap()).unwrap(), d);
        }
        let d = ActiveXDescriptor {
            class_id: "not-activated".into(),
            license: None,
            persistence: Persistence::PropertyBag,
            relationship_id: None,
            properties: vec![
                ActiveXProperty {
                    name: "Font".into(),
                    value: None,
                    object: Some(ActiveXPropertyObject::Font(ActiveXFont {
                        persistence: Some(Persistence::PropertyBag),
                        relationship_id: None,
                        properties: vec![ActiveXProperty {
                            name: "Name".into(),
                            value: Some("A&B".into()),
                            object: None,
                        }],
                    })),
                },
                ActiveXProperty {
                    name: "Picture".into(),
                    value: None,
                    object: Some(ActiveXPropertyObject::Picture(ActiveXPicture {
                        relationship_id: Some("rId2".into()),
                    })),
                },
            ],
        };
        assert_eq!(ActiveXDescriptor::parse(&d.to_xml().unwrap()).unwrap(), d);
    }

    fn package(external: bool, wrong_type: bool, outbound_binary: bool) -> OpcPackage {
        let worksheet_xml = format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><controls><control shapeId="1" r:id="rId1"/></controls></worksheet>"#).into_bytes();
        let descriptor_xml = format!(r#"<ax:ocx xmlns:ax="{AX}" xmlns:r="{REL}" ax:classid="inert" ax:persistence="persistStreamInit" r:id="rId1"/>"#).into_bytes();
        let mut worksheet = BlobPart::new(
            PackURI::new(WS).unwrap(),
            WORKSHEET_CONTENT_TYPE.into(),
            worksheet_xml,
        );
        worksheet.rels_mut().add_relationship(
            CONTROL_REL.into(),
            if external {
                "https://invalid.example/control".into()
            } else {
                "../activeX/activeX1.xml".into()
            },
            "rId1".into(),
            external,
        );
        let mut descriptor = BlobPart::new(
            PackURI::new("/xl/activeX/activeX1.xml").unwrap(),
            if wrong_type {
                "text/xml".into()
            } else {
                DESCRIPTOR_CONTENT_TYPE.into()
            },
            descriptor_xml,
        );
        descriptor.rels_mut().add_relationship(
            BINARY_REL.into(),
            "activeX1.bin".into(),
            "rId1".into(),
            false,
        );
        let mut binary = BlobPart::new(
            PackURI::new("/xl/activeX/activeX1.bin").unwrap(),
            BINARY_CONTENT_TYPE.into(),
            vec![0, 1, 2, 255],
        );
        if outbound_binary {
            binary.rels_mut().add_relationship(
                IMAGE_REL.into(),
                "../media/x.png".into(),
                "rId9".into(),
                false,
            );
        }
        let mut package = OpcPackage::new();
        package.add_part(Box::new(worksheet));
        package.add_part(Box::new(descriptor));
        package.add_part(Box::new(binary));
        package
    }

    #[test]
    fn package_validation_and_exact_opaque_roundtrip() {
        let p = package(false, false, false);
        let set = load_from_worksheet(&p, &PackURI::new(WS).unwrap()).unwrap();
        assert_eq!(set.controls[0].binaries[0].bytes, vec![0, 1, 2, 255]);
        assert!(
            load_from_worksheet(&package(true, false, false), &PackURI::new(WS).unwrap()).is_err()
        );
        assert!(
            load_from_worksheet(&package(false, true, false), &PackURI::new(WS).unwrap()).is_err()
        );
        assert!(
            load_from_worksheet(&package(false, false, true), &PackURI::new(WS).unwrap()).is_err()
        );
    }

    #[test]
    fn malformed_and_resource_matrix() {
        assert!(WorksheetControls::parse(br#"<!DOCTYPE x><worksheet/>"#).is_err());
        assert!(WorksheetControls::parse(format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><controls><control shapeId="0" r:id="x"/></controls></worksheet>"#).as_bytes()).is_err());
        assert!(
            ActiveXDescriptor::parse(
                format!(r#"<ax:ocx xmlns:ax="{AX}" ax:classid="x" ax:persistence="bad"/>"#)
                    .as_bytes()
            )
            .is_err()
        );
        assert!(ActiveXDescriptor::parse(format!(r#"<ax:ocx xmlns:ax="{AX}" ax:classid="x" ax:persistence="persistPropertyBag"><ax:ocxPr ax:name="a" ax:value="1"/><ax:ocxPr ax:name="a" ax:value="2"/></ax:ocx>"#).as_bytes()).is_err());
        let huge = "x".repeat(MAX_STRING + 1);
        assert!(
            ActiveXDescriptor {
                class_id: huge,
                license: None,
                persistence: Persistence::PropertyBag,
                relationship_id: None,
                properties: vec![ActiveXProperty {
                    name: "a".into(),
                    value: None,
                    object: None
                }]
            }
            .to_xml()
            .is_err()
        );
    }

    fn blank_package() -> OpcPackage {
        let xml = format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData><row r="1"/></sheetData><oleObjects/><tableParts count="0"/><extLst/></worksheet>"#
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(WS).unwrap(),
            WORKSHEET_CONTENT_TYPE.into(),
            xml.into_bytes(),
        )));
        package
    }

    fn binary_control(descriptor_uri: &str, binary_uri: &str) -> ActiveXControlSet {
        ActiveXControlSet {
            controls: vec![LoadedActiveXControl {
                control: WorksheetControl {
                    shape_id: 42,
                    relationship_id: "rIdControl".into(),
                    name: Some("Generated control".into()),
                    properties: Some(ControlProperties {
                        anchor: ObjectAnchor {
                            from: Marker {
                                column: 1,
                                column_offset: 2,
                                row: 3,
                                row_offset: 4,
                            },
                            to: Marker {
                                column: 5,
                                column_offset: 6,
                                row: 7,
                                row_offset: 8,
                            },
                            move_with_cells: Some(true),
                            size_with_cells: Some(false),
                        },
                        locked: Some(true),
                        default_size: None,
                        print: Some(false),
                        disabled: None,
                        recalc_always: None,
                        ui_object: None,
                        auto_fill: None,
                        auto_line: None,
                        auto_picture: None,
                        macro_name: Some("inert_callback_name".into()),
                        alternate_text: Some("generated".into()),
                        preview_relationship_id: Some("rIdPreview".into()),
                    }),
                },
                descriptor_uri: PackURI::new(descriptor_uri).unwrap(),
                descriptor: ActiveXDescriptor {
                    class_id: "{00000000-0000-0000-0000-000000000000}".into(),
                    license: None,
                    persistence: Persistence::StreamInit,
                    relationship_id: Some("rIdBinary".into()),
                    properties: Vec::new(),
                },
                binaries: vec![OpaqueActiveXBinary {
                    relationship_id: "rIdBinary".into(),
                    part_uri: PackURI::new(binary_uri).unwrap(),
                    bytes: vec![0, 1, 2, 0xff],
                }],
                preview: Some(OpaqueActiveXPreviewImage {
                    relationship_id: "rIdPreview".into(),
                    part_uri: PackURI::new("/xl/media/generated.png").unwrap(),
                    content_type: "image/png".into(),
                    bytes: vec![0x89, b'P', b'N', b'G'],
                }),
            }],
        }
    }

    #[test]
    fn generated_graph_store_reload_and_remove_preserves_unrelated_xml() {
        let mut package = blank_package();
        let worksheet_uri = PackURI::new(WS).unwrap();
        let original = package.get_part(&worksheet_uri).unwrap().blob().to_vec();
        let expected = binary_control("/xl/activeX/generated.xml", "/xl/activeX/generated.bin");
        store_on_worksheet(&mut package, &worksheet_uri, &expected).unwrap();
        let xml = std::str::from_utf8(package.get_part(&worksheet_uri).unwrap().blob()).unwrap();
        assert!(xml.contains("<sheetData><row r=\"1\"/></sheetData><oleObjects/><controls"));
        assert!(xml.contains("</controls><tableParts count=\"0\"/><extLst/>"));

        let bytes = PackageWriter::to_bytes(&package).unwrap();
        let reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            load_from_worksheet(&reopened, &worksheet_uri).unwrap(),
            expected
        );

        assert!(remove_from_worksheet(&mut package, &worksheet_uri).unwrap());
        assert_eq!(package.get_part(&worksheet_uri).unwrap().blob(), original);
        assert!(
            package
                .get_part(&PackURI::new("/xl/activeX/generated.xml").unwrap())
                .is_err()
        );
        assert!(!remove_from_worksheet(&mut package, &worksheet_uri).unwrap());
    }

    #[test]
    fn generated_replace_rolls_back_on_conflicting_part_name() {
        let mut package = blank_package();
        let worksheet_uri = PackURI::new(WS).unwrap();
        let first = binary_control("/xl/activeX/first.xml", "/xl/activeX/first.bin");
        store_on_worksheet(&mut package, &worksheet_uri, &first).unwrap();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/activeX/occupied.bin").unwrap(),
            "application/octet-stream".into(),
            vec![9],
        )));
        let before_xml = package.get_part(&worksheet_uri).unwrap().blob().to_vec();
        let before_parts = package.part_count();
        let replacement = binary_control("/xl/activeX/second.xml", "/xl/activeX/occupied.bin");
        assert!(replace_on_worksheet(&mut package, &worksheet_uri, &replacement).is_err());
        assert_eq!(package.part_count(), before_parts);
        assert_eq!(package.get_part(&worksheet_uri).unwrap().blob(), before_xml);
        assert_eq!(
            load_from_worksheet(&package, &worksheet_uri).unwrap(),
            first
        );
    }

    #[test]
    fn generated_remove_retains_shared_preview_and_rejects_limits() {
        let mut package = blank_package();
        let worksheet_uri = PackURI::new(WS).unwrap();
        let value = binary_control("/xl/activeX/a.xml", "/xl/activeX/a.bin");
        store_on_worksheet(&mut package, &worksheet_uri, &value).unwrap();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                IMAGE_REL.into(),
                "../media/generated.png".into(),
                "rIdShared".into(),
                TargetMode::Internal,
            )
            .unwrap();
        remove_from_worksheet(&mut package, &worksheet_uri).unwrap();
        assert!(
            package
                .get_part(&PackURI::new("/xl/media/generated.png").unwrap())
                .is_ok()
        );

        let mut invalid = value;
        invalid.controls[0].control.shape_id = MAX_SHAPE_ID + 1;
        assert!(store_on_worksheet(&mut blank_package(), &worksheet_uri, &invalid).is_err());
        invalid.controls[0].control.shape_id = 1;
        invalid.controls[0].control.name = Some("x".repeat(MAX_CONTROL_NAME_CHARS + 1));
        assert!(store_on_worksheet(&mut blank_package(), &worksheet_uri, &invalid).is_err());
    }

    #[test]
    fn direct_xml_mutation_refuses_mce_selected_collection() {
        let xml = format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="{X14}"><mc:AlternateContent><mc:Choice Requires="x14"><controls><control shapeId="1" r:id="rId1"/></controls></mc:Choice></mc:AlternateContent></worksheet>"#
        );
        assert!(
            replace_worksheet_controls_xml(xml.as_bytes(), &WorksheetControls::default()).is_err()
        );
    }
}
