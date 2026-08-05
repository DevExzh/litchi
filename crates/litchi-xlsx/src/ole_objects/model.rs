//! Semantic OLE-object anchors and inert embedded resources.

use crate::error::{Error, Result};
use litchi_opc::PackURI;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use std::collections::HashSet;
use std::fmt;
use std::str::FromStr;

use super::{invalid, limit};

pub(super) const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(super) const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
pub(super) const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
pub(super) const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
pub(super) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_NODES: usize = 500_000;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
pub(super) const MAX_OBJECTS: usize = 1024;
pub(super) const MAX_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_TOTAL_PAYLOAD_BYTES: usize = 256 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OleObjectConformance {
    Transitional,
    Strict,
}

impl OleObjectConformance {
    pub(super) fn sml(self) -> &'static str {
        match self {
            Self::Transitional => SML,
            Self::Strict => STRICT_SML,
        }
    }
    pub(super) fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }
    pub(super) fn xdr(self) -> &'static str {
        match self {
            Self::Transitional => XDR,
            Self::Strict => STRICT_XDR,
        }
    }
    pub(super) fn ole_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::OLE_OBJECT,
            Self::Strict => rt::STRICT_OLE_OBJECT,
        }
    }
    pub(super) fn package_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::PACKAGE,
            Self::Strict => rt::STRICT_PACKAGE,
        }
    }
    pub(super) fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::IMAGE,
            Self::Strict => rt::STRICT_IMAGE,
        }
    }
}

/// How an OLE object is represented when rendered.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OleObjectAspect {
    /// Render the embedded object's content (`DVASPECT_CONTENT`).
    Content,
    /// Render the embedded object's icon (`DVASPECT_ICON`).
    Icon,
}

impl OleObjectAspect {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Content => "DVASPECT_CONTENT",
            Self::Icon => "DVASPECT_ICON",
        }
    }
}

impl FromStr for OleObjectAspect {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        match value {
            "DVASPECT_CONTENT" => Ok(Self::Content),
            "DVASPECT_ICON" => Ok(Self::Icon),
            _ => Err(invalid(format!("invalid OLE data/view aspect '{value}'"))),
        }
    }
}

impl TryFrom<&str> for OleObjectAspect {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        value.parse()
    }
}

impl fmt::Display for OleObjectAspect {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OleObjectUpdate {
    Always,
    OnCall,
}

impl OleObjectUpdate {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Always => "OLEUPDATE_ALWAYS",
            Self::OnCall => "OLEUPDATE_ONCALL",
        }
    }
}

impl FromStr for OleObjectUpdate {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        match value {
            "OLEUPDATE_ALWAYS" => Ok(Self::Always),
            "OLEUPDATE_ONCALL" => Ok(Self::OnCall),
            _ => Err(invalid(format!("invalid OLE update mode '{value}'"))),
        }
    }
}

impl fmt::Display for OleObjectUpdate {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OleObjectRelationshipKind {
    OleObject,
    Package,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleObjectResource {
    pub part_name: String,
    pub content_type: String,
    /// Stored and returned without format sniffing, parsing, or activation.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OleObjectTarget {
    Internal(OleObjectResource),
    /// An inert OPC external target. It is never fetched or activated.
    External(String),
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct OleObjectMarker {
    pub column: u32,
    pub column_offset: i64,
    pub row: u32,
    pub row_offset: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleObjectAnchor {
    pub move_with_cells: Option<bool>,
    pub size_with_cells: Option<bool>,
    pub from: OleObjectMarker,
    pub to: OleObjectMarker,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleObjectProperties {
    pub preview_relationship_id: String,
    pub preview: Option<OleObjectResource>,
    pub default_size: Option<bool>,
    pub print: Option<bool>,
    pub disabled: Option<bool>,
    pub ui_object: Option<bool>,
    pub auto_fill: Option<bool>,
    pub auto_line: Option<bool>,
    pub auto_pict: Option<bool>,
    pub dde: Option<bool>,
    pub macro_name: Option<String>,
    pub alt_text: Option<String>,
    pub anchor: OleObjectAnchor,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleObject {
    pub program_id: Option<String>,
    pub data_or_view_aspect: Option<OleObjectAspect>,
    pub link: Option<String>,
    pub update: Option<OleObjectUpdate>,
    pub auto_load: Option<bool>,
    pub shape_id: u32,
    pub relationship_id: String,
    pub relationship_kind: OleObjectRelationshipKind,
    /// Filled by package loading and required by package storage.
    pub target: Option<OleObjectTarget>,
    pub properties: Option<OleObjectProperties>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OleObjects {
    pub objects: Vec<OleObject>,
}

pub(super) fn validate_value(value: &OleObjects, require_targets: bool) -> Result<()> {
    if value.objects.len() > MAX_OBJECTS {
        return Err(limit("object count"));
    }
    let mut shapes = HashSet::new();
    let mut ids = HashSet::new();
    let mut total = 0usize;
    for object in &value.objects {
        if !(1..=67_098_623).contains(&object.shape_id) {
            return Err(invalid("OLE shapeId is outside Office's supported range"));
        }
        if !shapes.insert(object.shape_id) {
            return Err(invalid(format!(
                "duplicate OLE shapeId {}",
                object.shape_id
            )));
        }
        validate_id(&object.relationship_id)?;
        if !ids.insert(object.relationship_id.clone()) {
            return Err(invalid(format!(
                "duplicate OLE relationship ID '{}'",
                object.relationship_id
            )));
        }
        if let Some(value) = &object.program_id {
            bounded(value)?;
            if value.len() >= 39
                || value
                    .bytes()
                    .next()
                    .is_some_and(|byte| byte.is_ascii_digit())
                || !value
                    .bytes()
                    .all(|byte| byte.is_ascii_alphanumeric() || byte == b'.')
            {
                return Err(invalid(format!("invalid Office ProgID '{value}'")));
            }
        }
        if let Some(value) = &object.link {
            bounded(value)?;
            if value.len() > 8192 {
                return Err(invalid("OLE link moniker exceeds Office's limit"));
            }
        }
        if require_targets && object.target.is_none() {
            return Err(invalid("OLE target is required for package storage"));
        }
        if let Some(target) = &object.target {
            match target {
                OleObjectTarget::External(value) => {
                    bounded(value)?;
                    if object.link.is_none() {
                        return Err(invalid("external OLE target requires a link moniker"));
                    }
                },
                OleObjectTarget::Internal(resource) => {
                    validate_resource(resource, "/xl/embeddings/")?;
                    if object.relationship_kind == OleObjectRelationshipKind::OleObject
                        && resource.content_type != ct::OFC_OLE_OBJECT
                    {
                        return Err(invalid(
                            "OLE relationship requires the OLE Object content type",
                        ));
                    }
                    add_payload(&mut total, resource.data.len())?;
                },
            }
        }
        if let Some(properties) = &object.properties {
            validate_id(&properties.preview_relationship_id)?;
            if properties.preview_relationship_id == object.relationship_id {
                return Err(invalid("payload and preview relationship IDs must differ"));
            }
            if let Some(value) = &properties.macro_name {
                bounded(value)?;
            }
            if let Some(value) = &properties.alt_text {
                bounded(value)?;
            }
            if require_targets && properties.preview.is_none() {
                return Err(invalid("object preview is required for package storage"));
            }
            if let Some(preview) = &properties.preview {
                validate_resource(preview, "/xl/media/")?;
                if !is_image_content_type(&preview.content_type) {
                    return Err(invalid("object preview has a non-image content type"));
                }
                add_payload(&mut total, preview.data.len())?;
            }
        }
    }
    Ok(())
}

pub(super) fn validate_resource(resource: &OleObjectResource, prefix: &str) -> Result<()> {
    let uri = PackURI::new(&resource.part_name).map_err(invalid)?;
    if !uri.as_str().starts_with(prefix) {
        return Err(invalid(format!("resource '{uri}' is outside {prefix}")));
    }
    if resource.content_type.is_empty()
        || resource
            .content_type
            .bytes()
            .any(|byte| byte.is_ascii_whitespace())
    {
        return Err(invalid("invalid embedded resource content type"));
    }
    if resource.data.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("individual payload bytes"));
    }
    Ok(())
}

pub(super) fn add_payload(total: &mut usize, size: usize) -> Result<()> {
    if size > MAX_PAYLOAD_BYTES {
        return Err(limit("individual payload bytes"));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("total payload bytes"))?;
    if *total > MAX_TOTAL_PAYLOAD_BYTES {
        Err(limit("total payload bytes"))
    } else {
        Ok(())
    }
}

pub(super) fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("string bytes"))
    }
}

pub(super) fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
pub(super) fn is_image_content_type(value: &str) -> bool {
    value.starts_with("image/") || matches!(value, "application/x-emf" | "application/x-wmf")
}
