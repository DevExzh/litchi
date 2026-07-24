//! PowerPoint Changes Information package parts.
//!
//! Nested edit descriptors are retained as bounded XML. They are never
//! executed and relationship-looking content inside them is never resolved.

use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

pub const CHANGES_INFORMATION_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.changesinfo+xml";
pub const CHANGES_INFORMATION_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2016/11/relationships/changesInfo";

const PC: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const PC_TEXT: &str = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const P: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const A: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const MAX_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 192;
const MAX_NODES: usize = 250_000;
const MAX_LISTS: usize = 16_384;
const MAX_CHANGES: usize = 100_000;
const MAX_EXTENSIONS: usize = 4_096;
const MAX_STRING_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChangesNamespaceDeclaration {
    /// Empty means the default namespace.
    pub prefix: String,
    pub uri: String,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChangesData {
    pub name: Option<String>,
    pub user_id: Option<String>,
    pub provider_id: Option<String>,
    pub client_id: Option<String>,
    pub email: Option<String>,
    pub date_time: Option<String>,
    pub version: Option<u32>,
    pub change_id: Option<String>,
    pub action_id: Option<i32>,
    /// Optional complete DrawingML `a:extLst` fragment.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentChangeKind {
    CustomSelection,
    AddSlide,
    DeleteSlide,
    ModifySlide,
    SlideOrder,
    ModifyMainMaster,
    ModifyNotesMaster,
    ModifyHandoutMaster,
    AddSection,
    DeleteSection,
    ModifySection,
}

impl DocumentChangeKind {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "custSel" => Ok(Self::CustomSelection),
            "addSld" => Ok(Self::AddSlide),
            "delSld" => Ok(Self::DeleteSlide),
            "modSld" => Ok(Self::ModifySlide),
            "sldOrd" => Ok(Self::SlideOrder),
            "modMainMaster" => Ok(Self::ModifyMainMaster),
            "modNotesMaster" => Ok(Self::ModifyNotesMaster),
            "modHandoutMaster" => Ok(Self::ModifyHandoutMaster),
            "addSection" => Ok(Self::AddSection),
            "delSection" => Ok(Self::DeleteSection),
            "modSection" => Ok(Self::ModifySection),
            _ => Err(invalid(format!("unknown document change bit '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentChangeDescriptor {
    pub change_kinds: Vec<DocumentChangeKind>,
    /// Complete `pc:docChg` fragment with nested commands kept inert.
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentChangesList {
    pub author: Option<ChangesData>,
    pub changes: Vec<DocumentChangeDescriptor>,
    /// Optional complete PresentationML `p:extLst` fragment.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChangesInformation {
    pub command_prefix: String,
    pub namespace_declarations: Vec<ChangesNamespaceDeclaration>,
    pub document_change_lists: Vec<DocumentChangesList>,
}

impl Default for ChangesInformation {
    fn default() -> Self {
        Self {
            command_prefix: "pc".into(),
            namespace_declarations: Vec::new(),
            document_change_lists: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChangesInformationPart {
    pub relationship_id: String,
    pub part_name: String,
    pub changes_information: ChangesInformation,
}

impl ChangesInformation {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        parse_changes_information(xml)
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        validate_model(self)?;
        let prefix = &self.command_prefix;
        let mut out = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
        open_root(&mut out, prefix, &self.namespace_declarations);
        if self.document_change_lists.is_empty() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for list in &self.document_change_lists {
                open_pc(&mut out, prefix, "docChgLst");
                out.push(b'>');
                if let Some(author) = &list.author {
                    write_changes_data(&mut out, prefix, author);
                }
                for change in &list.changes {
                    out.extend_from_slice(&change.xml);
                }
                if let Some(extension) = &list.extension_xml {
                    out.extend_from_slice(extension);
                }
                close_pc(&mut out, prefix, "docChgLst");
            }
            close_pc(&mut out, prefix, "chgInfo");
        }
        if out.len() > MAX_BYTES {
            return Err(limit("serialized Changes Information bytes"));
        }
        parse_changes_information(&out)?;
        Ok(out)
    }
}

pub fn load_changes_information(package: &OpcPackage) -> Result<Option<ChangesInformationPart>> {
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    let presentation_name = presentation.partname().to_string();
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == CHANGES_INFORMATION_RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "Changes Information relationship cannot originate at the package root",
        ));
    }
    for source in package.iter_parts() {
        if source.partname().as_str() != presentation_name.as_str()
            && source
                .rels()
                .iter()
                .any(|relationship| relationship.reltype() == CHANGES_INFORMATION_RELATIONSHIP_TYPE)
        {
            return Err(invalid(
                "Changes Information relationship has a non-Presentation source",
            ));
        }
    }
    let relationships: Vec<_> = presentation
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == CHANGES_INFORMATION_RELATIONSHIP_TYPE)
        .collect();
    if relationships.len() > 1 {
        return Err(invalid(
            "Presentation has multiple Changes Information relationships",
        ));
    }
    let Some(relationship) = relationships.first().copied() else {
        if package
            .iter_parts()
            .any(|part| part.content_type() == CHANGES_INFORMATION_CONTENT_TYPE)
        {
            return Err(invalid(
                "package contains an orphan Changes Information part",
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(invalid(
            "Changes Information relationship cannot be external",
        ));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != CHANGES_INFORMATION_CONTENT_TYPE {
        return Err(OoxmlError::InvalidContentType {
            expected: CHANGES_INFORMATION_CONTENT_TYPE.into(),
            got: part.content_type().into(),
        });
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Changes Information part cannot have outbound relationships",
        ));
    }
    if package.iter_parts().any(|candidate| {
        candidate.content_type() == CHANGES_INFORMATION_CONTENT_TYPE
            && candidate.partname() != &target
    }) {
        return Err(invalid(
            "package contains an orphan Changes Information part",
        ));
    }
    Ok(Some(ChangesInformationPart {
        relationship_id: relationship.r_id().to_string(),
        part_name: target.to_string(),
        changes_information: ChangesInformation::parse(part.blob())?,
    }))
}

pub fn store_changes_information(
    package: &mut OpcPackage,
    value: &ChangesInformationPart,
) -> Result<()> {
    if load_changes_information(package)?.is_some() {
        return Err(invalid("package already contains Changes Information"));
    }
    validate_ncname(&value.relationship_id, "relationship ID")?;
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    if presentation.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(
            "Changes Information relationship ID already exists",
        ));
    }
    let presentation_name = presentation.partname().clone();
    let part_name = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    if package
        .iter_parts()
        .any(|part| part.partname() == &part_name)
    {
        return Err(invalid(format!("part '{part_name}' already exists")));
    }
    let xml = value.changes_information.to_xml()?;
    let target = part_name.relative_ref(presentation_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(
        part_name,
        CHANGES_INFORMATION_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .add_relationship(
            CHANGES_INFORMATION_RELATIONSHIP_TYPE.into(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

#[derive(Debug)]
enum Frame {
    Root,
    List,
    Metadata,
    MetadataExtension,
    Document {
        kinds: Vec<DocumentChangeKind>,
        order: u8,
        has_moniker: bool,
    },
    MonikerList {
        count: u8,
    },
    Moniker,
    DescriptorOpaque,
    ListExtension {
        extensions: usize,
    },
    Extension {
        payloads: u8,
    },
    Payload,
    Opaque,
}

#[derive(Debug)]
enum PendingSlice {
    Metadata(usize),
    Document(usize, Vec<DocumentChangeKind>),
    ListExtension(usize),
}

fn parse_changes_information(xml: &[u8]) -> Result<ChangesInformation> {
    if xml.len() > MAX_BYTES {
        return Err(limit("Changes Information part bytes"));
    }
    let selected = process_ooxml(xml)?;
    if selected.len() > MAX_BYTES {
        return Err(limit("MCE-processed Changes Information bytes"));
    }
    let bytes = selected.as_ref();
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut command_prefix = String::new();
    let mut namespaces = Vec::new();
    let mut lists = Vec::new();
    let mut current_list: Option<DocumentChangesList> = None;
    let mut list_order = 0u8;
    let mut metadata_extension_start = None;
    let mut document_start = None;
    let mut list_extension_start = None;
    let mut pending: Option<PendingSlice> = None;
    let mut nodes = 0usize;
    let mut total_changes = 0usize;

    loop {
        let start = reader.buffer_position() as usize;
        if let Some(slice) = pending.take() {
            let list = current_list
                .as_mut()
                .ok_or_else(|| invalid("captured change XML outside docChgLst"))?;
            match slice {
                PendingSlice::Metadata(from) => {
                    list.author
                        .as_mut()
                        .ok_or_else(|| invalid("missing changes metadata"))?
                        .extension_xml = Some(bytes[from..start].to_vec());
                },
                PendingSlice::Document(from, kinds) => {
                    list.changes.push(DocumentChangeDescriptor {
                        change_kinds: kinds,
                        xml: bytes[from..start].to_vec(),
                    });
                },
                PendingSlice::ListExtension(from) => {
                    list.extension_xml = Some(bytes[from..start].to_vec());
                },
            }
        }
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("Changes Information nodes"))?;
                if nodes > MAX_NODES || stack.len() + 1 > MAX_DEPTH {
                    return Err(limit("Changes Information XML resources"));
                }
                let local = element.local_name();
                let frame = if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("Changes Information has multiple roots"));
                    }
                    expect(&namespace, PC, local.as_ref(), b"chgInfo")?;
                    command_prefix = element_prefix(&element)?;
                    namespaces = root_namespaces(&element, reader.decoder(), &command_prefix)?;
                    root_seen = true;
                    Frame::Root
                } else {
                    match stack.last_mut().expect("nonempty") {
                        Frame::Root => {
                            expect(&namespace, PC, local.as_ref(), b"docChgLst")?;
                            no_attributes(&element, reader.decoder())?;
                            if lists.len() >= MAX_LISTS || current_list.is_some() {
                                return Err(limit("document change lists"));
                            }
                            current_list = Some(DocumentChangesList::default());
                            list_order = 0;
                            Frame::List
                        },
                        Frame::List => {
                            if is(&namespace, PC, local.as_ref(), b"chgData") {
                                if list_order != 0
                                    || current_list
                                        .as_ref()
                                        .is_some_and(|list| list.author.is_some())
                                {
                                    return Err(invalid(
                                        "docChgLst chgData is duplicated or out of order",
                                    ));
                                }
                                let data = parse_changes_data(&element, reader.decoder())?;
                                current_list.as_mut().expect("active list").author = Some(data);
                                list_order = 1;
                                Frame::Metadata
                            } else if is(&namespace, PC, local.as_ref(), b"docChg") {
                                if list_order > 1 {
                                    return Err(invalid("docChg is out of order"));
                                }
                                total_changes += 1;
                                if total_changes > MAX_CHANGES || empty {
                                    return Err(limit("document changes"));
                                }
                                let kinds = parse_change_kinds(&element, reader.decoder())?;
                                document_start = Some(start);
                                list_order = 1;
                                Frame::Document {
                                    kinds,
                                    order: 0,
                                    has_moniker: false,
                                }
                            } else if is(&namespace, P, local.as_ref(), b"extLst") {
                                if list_order > 1
                                    || current_list
                                        .as_ref()
                                        .is_some_and(|list| list.extension_xml.is_some())
                                {
                                    return Err(invalid(
                                        "docChgLst extLst is duplicated or out of order",
                                    ));
                                }
                                no_attributes(&element, reader.decoder())?;
                                list_order = 2;
                                list_extension_start = Some(start);
                                Frame::ListExtension { extensions: 0 }
                            } else {
                                return Err(invalid("unexpected docChgLst child"));
                            }
                        },
                        Frame::Metadata => {
                            expect(&namespace, A, local.as_ref(), b"extLst")?;
                            no_attributes(&element, reader.decoder())?;
                            metadata_extension_start = Some(start);
                            Frame::MetadataExtension
                        },
                        Frame::Document {
                            order, has_moniker, ..
                        } => {
                            let next = if is(&namespace, PC, local.as_ref(), b"chgData") {
                                0
                            } else if is(&namespace, PC, local.as_ref(), b"docMkLst") {
                                1
                            } else if is(&namespace, PC, local.as_ref(), b"sldChg") {
                                2
                            } else if is(&namespace, PC, local.as_ref(), b"sldMasterChg") {
                                3
                            } else if is(&namespace, PC, local.as_ref(), b"cmAuthorChg") {
                                4
                            } else if is(&namespace, P, local.as_ref(), b"extLst") {
                                5
                            } else {
                                return Err(invalid("unexpected docChg child"));
                            };
                            if next < *order || (next == 0 && *order == 0 && *has_moniker) {
                                return Err(invalid(
                                    "docChg children are duplicated or out of order",
                                ));
                            }
                            *order = next;
                            match next {
                                0 => {
                                    parse_changes_data(&element, reader.decoder())?;
                                    Frame::DescriptorOpaque
                                },
                                1 => {
                                    if *has_moniker {
                                        return Err(invalid("docChg has duplicate docMkLst"));
                                    }
                                    *has_moniker = true;
                                    no_attributes(&element, reader.decoder())?;
                                    Frame::MonikerList { count: 0 }
                                },
                                2..=4 => {
                                    required_change_attribute(&element, reader.decoder())?;
                                    Frame::DescriptorOpaque
                                },
                                _ => {
                                    no_attributes(&element, reader.decoder())?;
                                    Frame::DescriptorOpaque
                                },
                            }
                        },
                        Frame::MonikerList { count } => {
                            expect(&namespace, PC, local.as_ref(), b"docMk")?;
                            no_attributes(&element, reader.decoder())?;
                            *count += 1;
                            if *count > 1 {
                                return Err(invalid("docMkLst permits exactly one docMk"));
                            }
                            Frame::Moniker
                        },
                        Frame::ListExtension { extensions } => {
                            expect(&namespace, P, local.as_ref(), b"ext")?;
                            *extensions += 1;
                            if *extensions > MAX_EXTENSIONS {
                                return Err(limit("Changes Information extensions"));
                            }
                            extension_attributes(&element, reader.decoder())?;
                            Frame::Extension { payloads: 0 }
                        },
                        Frame::Extension { payloads } => {
                            if *payloads != 0 || !other_than_p(&namespace) {
                                return Err(invalid("p:ext requires one foreign payload"));
                            }
                            *payloads = 1;
                            any_attributes(&element, reader.decoder())?;
                            Frame::Payload
                        },
                        Frame::MetadataExtension
                        | Frame::DescriptorOpaque
                        | Frame::Moniker
                        | Frame::Payload
                        | Frame::Opaque => {
                            any_attributes(&element, reader.decoder())?;
                            Frame::Opaque
                        },
                    }
                };
                if empty {
                    close_empty(&frame)?;
                    match &frame {
                        Frame::Root => root_closed = true,
                        Frame::List => {
                            lists.push(current_list.take().expect("active list"));
                        },
                        Frame::MetadataExtension => {
                            pending = Some(PendingSlice::Metadata(
                                metadata_extension_start.take().expect("metadata start"),
                            ));
                        },
                        Frame::ListExtension { .. } => {
                            pending = Some(PendingSlice::ListExtension(
                                list_extension_start.take().expect("list ext start"),
                            ));
                        },
                        _ => {},
                    }
                } else {
                    stack.push(frame);
                }
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                validate_end(&namespace, element.local_name().as_ref(), &frame)?;
                match frame {
                    Frame::Root => root_closed = true,
                    Frame::List => lists.push(current_list.take().expect("active list")),
                    Frame::MetadataExtension => {
                        pending = Some(PendingSlice::Metadata(
                            metadata_extension_start.take().expect("metadata start"),
                        ));
                    },
                    Frame::Document {
                        kinds, has_moniker, ..
                    } => {
                        if !has_moniker {
                            return Err(invalid("docChg requires docMkLst"));
                        }
                        pending = Some(PendingSlice::Document(
                            document_start.take().expect("document start"),
                            kinds,
                        ));
                    },
                    Frame::MonikerList { count } if count != 1 => {
                        return Err(invalid("docMkLst requires exactly one docMk"));
                    },
                    Frame::ListExtension { .. } => {
                        pending = Some(PendingSlice::ListExtension(
                            list_extension_start.take().expect("list ext start"),
                        ));
                    },
                    Frame::Extension { payloads } if payloads != 1 => {
                        return Err(invalid("p:ext requires one foreign payload"));
                    },
                    _ => {},
                }
            },
            Event::Text(text) => {
                if !matches!(
                    stack.last(),
                    Some(
                        Frame::MetadataExtension
                            | Frame::DescriptorOpaque
                            | Frame::Payload
                            | Frame::Opaque
                    )
                ) {
                    let decoded = text.decode().map_err(xml_error)?;
                    let unescaped = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    if !unescaped.trim().is_empty() {
                        return Err(invalid("unexpected Changes Information text"));
                    }
                }
            },
            Event::CData(text) => {
                if !matches!(
                    stack.last(),
                    Some(
                        Frame::MetadataExtension
                            | Frame::DescriptorOpaque
                            | Frame::Payload
                            | Frame::Opaque
                    )
                ) && !text.decode().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(invalid("unexpected Changes Information CDATA"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid("DTD, PI, and general references are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if let Some(slice) = pending.take() {
        let end = reader.buffer_position() as usize;
        let list = current_list
            .as_mut()
            .ok_or_else(|| invalid("captured change XML outside docChgLst"))?;
        match slice {
            PendingSlice::Metadata(from) => {
                list.author.as_mut().expect("author").extension_xml =
                    Some(bytes[from..end].to_vec())
            },
            PendingSlice::Document(from, kinds) => list.changes.push(DocumentChangeDescriptor {
                change_kinds: kinds,
                xml: bytes[from..end].to_vec(),
            }),
            PendingSlice::ListExtension(from) => {
                list.extension_xml = Some(bytes[from..end].to_vec())
            },
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() || current_list.is_some() {
        return Err(invalid("unterminated Changes Information part"));
    }
    let value = ChangesInformation {
        command_prefix,
        namespace_declarations: namespaces,
        document_change_lists: lists,
    };
    validate_model(&value)?;
    Ok(value)
}

fn parse_changes_data(element: &BytesStart<'_>, decoder: Decoder) -> Result<ChangesData> {
    let attrs = known_attributes(
        element,
        decoder,
        &[
            "name",
            "userId",
            "providerId",
            "clId",
            "email",
            "dt",
            "v",
            "id",
            "actId",
        ],
    )?;
    let get = |name: &str| {
        attrs
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.clone())
    };
    let date_time = get("dt");
    if let Some(value) = &date_time {
        validate_date_time(value)?;
    }
    let version = get("v")
        .map(|value| value.parse().map_err(|_| invalid("invalid change version")))
        .transpose()?;
    let action_id = get("actId")
        .map(|value| {
            value
                .parse()
                .map_err(|_| invalid("invalid change action ID"))
        })
        .transpose()?;
    let change_id = get("id");
    if let Some(value) = &change_id {
        validate_guid(value)?;
    }
    Ok(ChangesData {
        name: get("name"),
        user_id: get("userId"),
        provider_id: get("providerId"),
        client_id: get("clId"),
        email: get("email"),
        date_time,
        version,
        change_id,
        action_id,
        extension_xml: None,
    })
}

fn parse_change_kinds(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Vec<DocumentChangeKind>> {
    let attrs = known_attributes(element, decoder, &["chg"])?;
    let value = attrs
        .iter()
        .find(|(key, _)| key == "chg")
        .map(|(_, value)| value.as_str())
        .ok_or_else(|| invalid("docChg is missing required chg"))?;
    let kinds: Vec<_> = value
        .split_whitespace()
        .map(DocumentChangeKind::parse)
        .collect::<Result<_>>()?;
    if kinds.is_empty() {
        return Err(invalid("docChg chg list cannot be empty"));
    }
    let mut unique = HashSet::new();
    if kinds.iter().any(|kind| !unique.insert(*kind as u8)) {
        return Err(invalid("docChg contains duplicate change bits"));
    }
    Ok(kinds)
}

fn required_change_attribute(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    let attrs = known_attributes(element, decoder, &["chg"])?;
    if attrs
        .iter()
        .find(|(key, _)| key == "chg")
        .is_none_or(|(_, value)| value.split_whitespace().next().is_none())
    {
        return Err(invalid("nested change descriptor requires chg"));
    }
    Ok(())
}

fn validate_model(value: &ChangesInformation) -> Result<()> {
    if !value.command_prefix.is_empty() {
        validate_ncname(&value.command_prefix, "command prefix")?;
    }
    if value.document_change_lists.len() > MAX_LISTS {
        return Err(limit("document change lists"));
    }
    let mut prefixes = HashSet::new();
    for declaration in &value.namespace_declarations {
        if (!declaration.prefix.is_empty() && !ncname(&declaration.prefix))
            || matches!(declaration.prefix.as_str(), "xml" | "xmlns")
            || declaration.uri.is_empty()
            || declaration.prefix == value.command_prefix
            || !prefixes.insert(declaration.prefix.as_str())
        {
            return Err(invalid("invalid preserved Changes Information namespace"));
        }
        bounded(&declaration.uri)?;
    }
    let mut total = 0usize;
    for list in &value.document_change_lists {
        if let Some(author) = &list.author {
            validate_changes_data(author)?;
        }
        total = total
            .checked_add(list.changes.len())
            .ok_or_else(|| limit("document changes"))?;
        if total > MAX_CHANGES {
            return Err(limit("document changes"));
        }
        for change in &list.changes {
            if change.change_kinds.is_empty() || change.xml.len() > MAX_BYTES {
                return Err(invalid("invalid document change descriptor"));
            }
        }
        if list
            .extension_xml
            .as_ref()
            .is_some_and(|xml| xml.len() > MAX_BYTES)
        {
            return Err(limit("Changes Information extension bytes"));
        }
    }
    Ok(())
}

fn validate_changes_data(value: &ChangesData) -> Result<()> {
    for text in [
        value.name.as_deref(),
        value.user_id.as_deref(),
        value.provider_id.as_deref(),
        value.client_id.as_deref(),
        value.email.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        bounded(text)?;
    }
    if let Some(date) = &value.date_time {
        validate_date_time(date)?;
    }
    if let Some(id) = &value.change_id {
        validate_guid(id)?;
    }
    Ok(())
}

fn write_changes_data(out: &mut Vec<u8>, prefix: &str, value: &ChangesData) {
    open_pc(out, prefix, "chgData");
    for (name, text) in [
        ("name", value.name.as_deref()),
        ("userId", value.user_id.as_deref()),
        ("providerId", value.provider_id.as_deref()),
        ("clId", value.client_id.as_deref()),
        ("email", value.email.as_deref()),
        ("dt", value.date_time.as_deref()),
        ("id", value.change_id.as_deref()),
    ] {
        if let Some(text) = text {
            attr(out, name, text);
        }
    }
    if let Some(version) = value.version {
        attr(out, "v", &version.to_string());
    }
    if let Some(action_id) = value.action_id {
        attr(out, "actId", &action_id.to_string());
    }
    if let Some(extension) = &value.extension_xml {
        out.push(b'>');
        out.extend_from_slice(extension);
        close_pc(out, prefix, "chgData");
    } else {
        out.extend_from_slice(b"/>");
    }
}

fn open_root(out: &mut Vec<u8>, prefix: &str, declarations: &[ChangesNamespaceDeclaration]) {
    out.push(b'<');
    if !prefix.is_empty() {
        out.extend_from_slice(prefix.as_bytes());
        out.push(b':');
    }
    out.extend_from_slice(b"chgInfo xmlns");
    if !prefix.is_empty() {
        out.push(b':');
        out.extend_from_slice(prefix.as_bytes());
    }
    out.extend_from_slice(b"=\"");
    escape(out, PC_TEXT);
    out.push(b'"');
    for declaration in declarations {
        out.extend_from_slice(b" xmlns");
        if !declaration.prefix.is_empty() {
            out.push(b':');
            out.extend_from_slice(declaration.prefix.as_bytes());
        }
        out.extend_from_slice(b"=\"");
        escape(out, &declaration.uri);
        out.push(b'"');
    }
}

fn open_pc(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.push(b'<');
    if !prefix.is_empty() {
        out.extend_from_slice(prefix.as_bytes());
        out.push(b':');
    }
    out.extend_from_slice(local.as_bytes());
}

fn close_pc(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.extend_from_slice(b"</");
    if !prefix.is_empty() {
        out.extend_from_slice(prefix.as_bytes());
        out.push(b':');
    }
    out.extend_from_slice(local.as_bytes());
    out.push(b'>');
}

fn root_namespaces(
    element: &BytesStart<'_>,
    decoder: Decoder,
    root_prefix: &str,
) -> Result<Vec<ChangesNamespaceDeclaration>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if key != "xmlns" && !key.starts_with("xmlns:") {
            return Err(invalid(format!("unexpected chgInfo attribute '{key}'")));
        }
        let prefix = key.strip_prefix("xmlns:").unwrap_or("").to_string();
        if !seen.insert(prefix.clone()) {
            return Err(invalid("duplicate root namespace declaration"));
        }
        let uri = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded(&uri)?;
        if prefix != root_prefix {
            output.push(ChangesNamespaceDeclaration { prefix, uri });
        }
    }
    Ok(output)
}

fn known_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    known: &[&str],
) -> Result<Vec<(String, String)>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if key == "xmlns" || key.starts_with("xmlns:") {
            continue;
        }
        if key.contains(':') || !known.contains(&key) || !seen.insert(key.to_string()) {
            return Err(invalid(format!(
                "unexpected or duplicate attribute '{key}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded(&value)?;
        output.push((key.to_string(), value));
    }
    Ok(output)
}

fn no_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    known_attributes(element, decoder, &[]).map(|_| ())
}

fn extension_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    let values = known_attributes(element, decoder, &["uri"])?;
    if values
        .iter()
        .find(|(key, _)| key == "uri")
        .is_none_or(|(_, value)| value.trim().is_empty())
    {
        return Err(invalid("p:ext requires nonempty uri"));
    }
    Ok(())
}

fn any_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        bounded(&value)?;
    }
    Ok(())
}

fn close_empty(frame: &Frame) -> Result<()> {
    match frame {
        Frame::Document { .. } => Err(invalid("docChg cannot be empty")),
        Frame::MonikerList { .. } => Err(invalid("docMkLst requires docMk")),
        Frame::Extension { .. } => Err(invalid("p:ext requires one foreign payload")),
        _ => Ok(()),
    }
}

fn validate_end(namespace: &ResolveResult<'_>, local: &[u8], frame: &Frame) -> Result<()> {
    match frame {
        Frame::Root => expect(namespace, PC, local, b"chgInfo"),
        Frame::List => expect(namespace, PC, local, b"docChgLst"),
        Frame::Metadata => expect(namespace, PC, local, b"chgData"),
        Frame::MetadataExtension => expect(namespace, A, local, b"extLst"),
        Frame::Document { .. } => expect(namespace, PC, local, b"docChg"),
        Frame::MonikerList { .. } => expect(namespace, PC, local, b"docMkLst"),
        Frame::Moniker => expect(namespace, PC, local, b"docMk"),
        Frame::ListExtension { .. } => expect(namespace, P, local, b"extLst"),
        Frame::Extension { .. } => expect(namespace, P, local, b"ext"),
        Frame::DescriptorOpaque | Frame::Payload | Frame::Opaque => Ok(()),
    }
}

fn is(namespace: &ResolveResult<'_>, expected: &[u8], local: &[u8], name: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected) && local == name
}

fn expect(namespace: &ResolveResult<'_>, expected: &[u8], local: &[u8], name: &[u8]) -> Result<()> {
    if is(namespace, expected, local, name) {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected namespaced element '{}'",
            String::from_utf8_lossy(name)
        )))
    }
}

fn other_than_p(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() != P)
}

fn element_prefix(element: &BytesStart<'_>) -> Result<String> {
    let name = element.name();
    let qualified = std::str::from_utf8(name.as_ref()).map_err(xml_error)?;
    Ok(qualified
        .rsplit_once(':')
        .map(|(prefix, _)| prefix)
        .unwrap_or("")
        .to_string())
}

fn validate_guid(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    if bytes.len() != 38
        || bytes[0] != b'{'
        || bytes[37] != b'}'
        || [9, 14, 19, 24].iter().any(|index| bytes[*index] != b'-')
        || bytes[1..37]
            .iter()
            .enumerate()
            .any(|(index, byte)| ![8, 13, 18, 23].contains(&index) && !byte.is_ascii_hexdigit())
    {
        Err(invalid(format!("invalid GUID '{value}'")))
    } else {
        Ok(())
    }
}

fn validate_date_time(value: &str) -> Result<()> {
    bounded(value)?;
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML dateTime '{value}'")))
    }
}

fn validate_ncname(value: &str, label: &str) -> Result<()> {
    if ncname(value) {
        Ok(())
    } else {
        Err(invalid(format!("{label} is not an XML NCName")))
    }
}

fn ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters
        .next()
        .is_some_and(|character| character == '_' || character.is_alphabetic())
        && characters.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
}

fn require_presentation_content_type(value: &str) -> Result<()> {
    if matches!(
        value,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
            | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid("main document is not a Presentation part"))
    }
}

fn bounded(value: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        Err(limit("Changes Information string bytes"))
    } else {
        Ok(())
    }
}

fn attr(out: &mut Vec<u8>, name: &str, value: &str) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    escape(out, value);
    out.push(b'"');
}

fn escape(out: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                out.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(label: &str) -> OoxmlError {
    invalid(format!("{label} exceed configured limit"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const MAIN_CT: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            MAIN_CT.into(),
            Vec::new(),
        )));
        package
    }

    fn value() -> ChangesInformationPart {
        let xml = format!(
            r#"<pc:chgInfo xmlns:pc="{PC_TEXT}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:v="urn:vendor"><pc:docChgLst><pc:chgData name="Ada" userId="ada@example.test" providerId="AD" clId="Web-{{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}}"/><pc:docChg chg="addSld modSld"><pc:docMkLst><pc:docMk/></pc:docMkLst><p:extLst><p:ext uri="urn:nested"><v:data href="https://example.invalid/not-opened"/></p:ext></p:extLst></pc:docChg><p:extLst><p:ext uri="urn:list"><v:data xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdNeverFetched"/></p:ext></p:extLst></pc:docChgLst></pc:chgInfo>"#
        );
        ChangesInformationPart {
            relationship_id: "rIdChanges".into(),
            part_name: "/ppt/changesInfos/changesInfo1.xml".into(),
            changes_information: ChangesInformation::parse(xml.as_bytes()).unwrap(),
        }
    }

    #[test]
    fn loads_powerpoint_and_libreoffice_changes_packages() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let powerpoint = OpcPackage::open(
            root.join("test-data/poi/test-data/slideshow/ArtisticEffectSample.pptx"),
        )
        .unwrap();
        let loaded = load_changes_information(&powerpoint).unwrap().unwrap();
        assert!(!loaded.changes_information.document_change_lists.is_empty());
        assert_eq!(
            loaded.changes_information.document_change_lists[0]
                .author
                .as_ref()
                .unwrap()
                .name
                .as_deref(),
            Some("Frank Pavelski")
        );

        let libreoffice = OpcPackage::open(
            root.join("test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-sections.pptx"),
        )
        .unwrap();
        let loaded = load_changes_information(&libreoffice).unwrap().unwrap();
        assert!(
            loaded
                .changes_information
                .document_change_lists
                .iter()
                .flat_map(|list| &list.changes)
                .any(|change| change
                    .change_kinds
                    .contains(&DocumentChangeKind::AddSection))
        );
    }

    #[test]
    fn package_round_trip_keeps_nested_commands_and_extensions_inert() {
        let expected = value();
        let mut package = package();
        store_changes_information(&mut package, &expected).unwrap();
        let loaded = load_changes_information(&package).unwrap().unwrap();
        assert_eq!(loaded, expected);
        let text = String::from_utf8(loaded.changes_information.to_xml().unwrap()).unwrap();
        assert!(text.contains("rIdNeverFetched"));
        assert!(text.contains("https://example.invalid/not-opened"));
    }

    #[test]
    fn rejects_hostile_changes_grammar() {
        let wrap = |body: &str| {
            format!(
                r#"<pc:chgInfo xmlns:pc="{PC_TEXT}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{body}</pc:chgInfo>"#
            )
        };
        let cases = [
            wrap("<pc:wrong/>"),
            wrap("<pc:docChgLst bad=\"1\"/>"),
            wrap(
                "<pc:docChgLst><pc:docChg chg=\"unknown\"><pc:docMkLst><pc:docMk/></pc:docMkLst></pc:docChg></pc:docChgLst>",
            ),
            wrap(
                "<pc:docChgLst><pc:docChg chg=\"addSld\"><pc:sldChg chg=\"mod\"/></pc:docChg></pc:docChgLst>",
            ),
            wrap(
                "<pc:docChgLst><pc:docChg chg=\"addSld\"><pc:docMkLst/></pc:docChg></pc:docChgLst>",
            ),
            wrap("<pc:docChgLst><p:extLst/><pc:chgData/></pc:docChgLst>"),
            wrap("<pc:docChgLst><p:extLst><p:ext uri=\"urn:x\"/></p:extLst></pc:docChgLst>"),
            format!("<!DOCTYPE x>{}", wrap("")),
        ];
        for xml in cases {
            assert!(
                ChangesInformation::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_invalid_package_graphs_and_failed_store_is_atomic() {
        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                CHANGES_INFORMATION_RELATIONSHIP_TYPE.into(),
                "https://example.invalid/changes.xml".into(),
                "rIdChanges".into(),
                true,
            );
        assert!(load_changes_information(&external).is_err());

        let mut outbound = package();
        store_changes_information(&mut outbound, &value()).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/changesInfos/changesInfo1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship("urn:forbidden".into(), "x".into(), "rId1".into(), false);
        assert!(load_changes_information(&outbound).is_err());

        let mut bad = value();
        bad.changes_information.command_prefix = "bad prefix".into();
        let mut package = package();
        let count = package.part_count();
        assert!(store_changes_information(&mut package, &bad).is_err());
        assert_eq!(package.part_count(), count);
        assert!(load_changes_information(&package).unwrap().is_none());
    }
}
