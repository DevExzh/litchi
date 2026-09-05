//! MS-XLDM outer-storage codec: bounded XML projection and exact snapshots.

use super::model::{
    BOM, CRC_SIZE, Compression, FileEntry, FileKind, Header, MAX_DIRECTORY_BYTES, MAX_FILES,
    MAX_PARTITIONS, MAX_PATH_BYTES, MAX_STORAGE_BYTES, MAX_XML_DEPTH, MAX_XML_NODES,
    MAX_XML_TEXT_BYTES, Node, Offset, PartitionMarker, Size, Storage, XLDM_PAGE_SIZE,
    XLDM_STREAM_SIGNATURE, XmlEncoding,
};
use super::semantic::parse_backup_log;
use super::validation::{validate_allocations, validate_backup_log, validate_paths};
use crate::error::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

/// Validate and inspect the outer MS-XLDM virtual storage.
pub fn inspect(bytes: &[u8]) -> Result<Storage<'_>> {
    if bytes.len() > MAX_STORAGE_BYTES {
        return Err(limit("storage bytes"));
    }
    if bytes.len() < XLDM_PAGE_SIZE * 3 || !bytes.len().is_multiple_of(XLDM_PAGE_SIZE) {
        return Err(invalid(
            "MS-XLDM storage must contain at least three complete 4096-byte pages",
        ));
    }
    if bytes[..2] != BOM {
        return Err(invalid("MS-XLDM header byte-order mark is missing"));
    }
    let signature = utf16le(XLDM_STREAM_SIGNATURE);
    if bytes.get(2..2 + signature.len()) != Some(signature.as_slice()) {
        return Err(invalid("MS-XLDM stream storage signature is invalid"));
    }
    let header_xml_start = 2 + signature.len();
    let close = utf16le("</BackupLog>");
    let relative_end = memchr::memmem::rfind(&bytes[header_xml_start..XLDM_PAGE_SIZE], &close)
        .ok_or_else(|| invalid("MS-XLDM header BackupLog closing element is missing"))?;
    let header_xml_end = header_xml_start
        .checked_add(relative_end)
        .and_then(|value| value.checked_add(close.len()))
        .ok_or_else(|| limit("header XML range"))?;
    if bytes[header_xml_end..XLDM_PAGE_SIZE]
        .iter()
        .any(|byte| *byte != 0)
    {
        return Err(invalid("MS-XLDM header padding must be zero"));
    }
    let (header_xml, header_encoding) = decode_xml(&bytes[header_xml_start..header_xml_end], true)?;
    let header = parse_header(&parse_xml(&header_xml)?)?;
    let data_offset = checked_usize(header.data_offset.0, "data offset")?;
    let directory_offset = checked_usize(header.directory_offset.0, "directory offset")?;
    let directory_size = checked_usize(header.directory_size.0, "directory size")?;
    if data_offset != XLDM_PAGE_SIZE {
        return Err(invalid(
            "MS-XLDM data offset must equal the 4096-byte header size",
        ));
    }
    if directory_offset < XLDM_PAGE_SIZE * 2 || directory_offset % XLDM_PAGE_SIZE != 0 {
        return Err(invalid(
            "MS-XLDM directory offset must be page aligned after the files section",
        ));
    }
    if directory_size == 0 || directory_size > MAX_DIRECTORY_BYTES {
        return Err(limit("directory bytes"));
    }
    let directory_end = directory_offset
        .checked_add(directory_size)
        .ok_or_else(|| limit("directory range"))?;
    if directory_end > bytes.len() {
        return Err(invalid("MS-XLDM virtual directory extends beyond storage"));
    }
    if bytes[directory_end..].iter().any(|byte| *byte != 0) {
        return Err(invalid("MS-XLDM directory padding must be zero"));
    }
    if bytes.get(data_offset..data_offset + 2) != Some(&BOM) {
        return Err(invalid("MS-XLDM files-section byte-order mark is missing"));
    }
    if bytes.get(directory_offset..directory_offset + 2) == Some(&BOM) {
        return Err(invalid(
            "MS-XLDM virtual directory must not have a byte-order mark",
        ));
    }
    let (directory_xml, directory_encoding) =
        decode_xml(&bytes[directory_offset..directory_end], false)?;
    let mut files = parse_directory(&parse_xml(&directory_xml)?)?;
    if files.len() != header.file_count as usize {
        return Err(invalid(format!(
            "header file count {} does not match directory count {}",
            header.file_count,
            files.len()
        )));
    }
    if files.len() < 2 {
        return Err(invalid(
            "MS-XLDM storage requires partition and backup-log allocations",
        ));
    }
    validate_paths(&files)?;
    let mut order: Vec<usize> = (0..files.len()).collect();
    order.sort_by_key(|index| files[*index].offset);
    validate_allocations(bytes, data_offset, directory_offset, &mut files, &order)?;
    let first = order[0];
    let partition_bytes = payload_slice(bytes, &files[first])?;
    let (partitions_xml, partition_encoding) = decode_xml(partition_bytes, false)?;
    let partition_count = parse_partitions(&parse_xml(&partitions_xml)?)?;
    files[first].kind = FileKind::Partitions;
    let last = *order.last().unwrap_or_else(|| {
        crate::error::panic_missing_invariant("required value was checked before extraction")
    });
    files[last].kind = FileKind::BackupLog;
    let (backup_xml, backup_encoding) = decode_xml(payload_slice(bytes, &files[last])?, false)?;
    let backup_log = parse_backup_log(&parse_xml(&backup_xml)?, backup_encoding)?;
    validate_backup_log(&backup_log, &files, first, last)?;
    Ok(Storage {
        header,
        header_encoding,
        directory_encoding,
        partition_marker: PartitionMarker {
            partition_count,
            encoding: partition_encoding,
            encoded_xml: partition_bytes,
        },
        backup_log,
        files,
        bytes,
    })
}

/// Revalidate and return the original byte stream exactly.
pub fn write(storage: &Storage<'_>) -> Result<Vec<u8>> {
    inspect(storage.bytes)?;
    Ok(storage.bytes.to_vec())
}

pub(super) fn bounded_leaf(node: &Node, label: &str) -> Result<String> {
    let value = leaf_text(node)?;
    if value.len() > MAX_PATH_BYTES {
        return Err(limit(label));
    }
    Ok(value.to_owned())
}

fn parse_header(root: &Node) -> Result<Header> {
    let names = [
        "BackupRestoreSyncVersion",
        "Fault",
        "faultcode",
        "ErrorCode",
        "EncryptionFlag",
        "EncryptionKey",
        "ApplyCompression",
        "m_cbOffsetHeader",
        "DataSize",
        "Files",
        "ObjectID",
        "m_cbOffsetData",
    ];
    let values = exact_children(root, "BackupLog", &names)?;
    let backup_restore_sync_version = i32_value(values[0])?;
    if backup_restore_sync_version != 140 {
        return Err(invalid("BackupRestoreSyncVersion must equal 140"));
    }
    if bool_value(values[1])? {
        return Err(invalid("header Fault must be false"));
    }
    let fault_code = u32_value(values[2])?;
    if !bool_value(values[3])? {
        return Err(invalid("header ErrorCode must be true"));
    }
    if bool_value(values[4])? {
        return Err(invalid("header EncryptionFlag must be false"));
    }
    let encryption_key_version = i32_value(values[5])?;
    if !bool_value(values[6])? {
        return Err(invalid("header ApplyCompression must be true"));
    }
    let directory_offset = Offset(u64_value(values[7])?);
    let directory_size = Size(u64_value(values[8])?);
    let file_count = u32_value(values[9])?;
    if file_count as usize > MAX_FILES {
        return Err(limit("file count"));
    }
    let object_id = leaf_text(values[10])?.to_owned();
    if !valid_upper_guid(&object_id) {
        return Err(invalid("header ObjectID must be an uppercase UUID"));
    }
    let data_offset = Offset(u64_value(values[11])?);
    Ok(Header {
        backup_restore_sync_version,
        fault_code,
        encryption_key_version,
        compression: Compression::Xpress,
        directory_offset,
        directory_size,
        file_count,
        object_id,
        data_offset,
    })
}

fn parse_directory(root: &Node) -> Result<Vec<FileEntry>> {
    if root.name != "VirtualDirectory" || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid("expected attribute-free VirtualDirectory root"));
    }
    if root.children.len() > MAX_FILES {
        return Err(limit("file count"));
    }
    let names = [
        "Path",
        "Size",
        "m_cbOffsetHeader",
        "Delete",
        "CreatedTimestamp",
        "Access",
        "LastWriteTime",
    ];
    root.children
        .iter()
        .map(|child| {
            let values = exact_children(child, "BackupFile", &names)?;
            let path = leaf_text(values[0])?.to_owned();
            let stored_size = Size(u64_value(values[1])?);
            let offset = Offset(u64_value(values[2])?);
            let delete = bool_value(values[3])?;
            let created_timestamp = i64_value(values[4])?;
            let access_timestamp = i64_value(values[5])?;
            let last_write_timestamp = i64_value(values[6])?;
            let lower = path.to_ascii_lowercase();
            let kind = if lower.ends_with("cryptkey.bin") {
                FileKind::CryptographicKey
            } else if lower.ends_with(".xml") {
                FileKind::XmlMetadata
            } else {
                FileKind::OpaqueBinary
            };
            Ok(FileEntry {
                path,
                kind,
                offset,
                stored_size,
                crc32: 0,
                delete,
                created_timestamp,
                access_timestamp,
                last_write_timestamp,
            })
        })
        .collect()
}

fn parse_partitions(root: &Node) -> Result<usize> {
    if root.name != "Partitions" || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid("expected attribute-free Partitions root"));
    }
    if root.children.len() > MAX_PARTITIONS {
        return Err(limit("partition count"));
    }
    let names = [
        "ObjectPath",
        "Name",
        "DataSize",
        "Location",
        "DataSourceID",
        "ConnectionString",
    ];
    for partition in &root.children {
        let values = exact_children(partition, "Partition", &names)?;
        let _ = i64_value(values[2])?;
        for index in [0usize, 1, 3, 4, 5] {
            if leaf_text(values[index])?.len() > MAX_XML_TEXT_BYTES {
                return Err(limit("partition field bytes"));
            }
        }
    }
    Ok(root.children.len())
}

fn payload_slice<'a>(bytes: &'a [u8], entry: &FileEntry) -> Result<&'a [u8]> {
    let start = checked_usize(entry.offset.0, "file offset")?;
    let size = checked_usize(entry.stored_size.0, "file size")?;
    let end = start
        .checked_add(size)
        .and_then(|value| value.checked_sub(CRC_SIZE))
        .ok_or_else(|| limit("payload range"))?;
    bytes
        .get(start..end)
        .ok_or_else(|| invalid("payload range is outside storage"))
}

fn parse_xml(xml: &str) -> Result<Node> {
    if xml.len() > MAX_DIRECTORY_BYTES {
        return Err(limit("XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_XML_NODES || stack.len() >= MAX_XML_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(element)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(decoded.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_XML_TEXT_BYTES {
                    return Err(limit("XML text bytes"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_XML_TEXT_BYTES {
                    return Err(limit("XML text bytes"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::CData(_) => {
                return Err(invalid(
                    "DTDs, processing instructions, and CDATA are rejected",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(element: &BytesStart<'_>) -> Result<Node> {
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut attributes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            attributes += 1;
        }
    }
    Ok(Node {
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}
fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
pub(super) fn exact_children<'a>(
    root: &'a Node,
    root_name: &str,
    names: &[&str],
) -> Result<Vec<&'a Node>> {
    if root.name != root_name || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid(format!(
            "expected attribute-free {root_name} element"
        )));
    }
    if root.children.len() != names.len() {
        return Err(invalid(format!("{root_name} has an invalid child count")));
    }
    for (child, expected) in root.children.iter().zip(names) {
        if child.name != *expected {
            return Err(invalid(format!("expected {expected} in {root_name}")));
        }
    }
    Ok(root.children.iter().collect())
}
pub(super) fn leaf_text(node: &Node) -> Result<&str> {
    if node.attributes != 0 || !node.children.is_empty() {
        return Err(invalid(format!(
            "{} must be an attribute-free leaf",
            node.name
        )));
    }
    Ok(&node.text)
}
pub(super) fn bool_value(node: &Node) -> Result<bool> {
    match leaf_text(node)?.trim() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{} is not an XML boolean", node.name))),
    }
}
pub(super) fn u64_value(node: &Node) -> Result<u64> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not an unsigned 64-bit integer", node.name)))
}
pub(super) fn u32_value(node: &Node) -> Result<u32> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not an unsigned 32-bit integer", node.name)))
}
pub(super) fn i64_value(node: &Node) -> Result<i64> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not a signed 64-bit integer", node.name)))
}
pub(super) fn i32_value(node: &Node) -> Result<i32> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not a signed 32-bit integer", node.name)))
}

fn decode_xml(bytes: &[u8], require_utf16: bool) -> Result<(String, XmlEncoding)> {
    if bytes.is_empty() {
        return Err(invalid("empty XML allocation"));
    }
    if bytes.starts_with(&BOM) {
        return Err(invalid("unexpected XML byte-order mark"));
    }
    if require_utf16 || bytes.get(1) == Some(&0) {
        if !bytes.len().is_multiple_of(2) {
            return Err(invalid("odd-length UTF-16LE XML"));
        }
        let words: Vec<u16> = bytes
            .as_chunks::<2>()
            .0
            .iter()
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        Ok((
            String::from_utf16(&words).map_err(xml_error)?,
            XmlEncoding::Utf16Le,
        ))
    } else {
        Ok((
            std::str::from_utf8(bytes).map_err(xml_error)?.to_owned(),
            XmlEncoding::Utf8,
        ))
    }
}
pub(super) fn utf16le(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}
pub(super) fn valid_upper_guid(value: &str) -> bool {
    if value.len() != 36 {
        return false;
    }
    value.bytes().enumerate().all(|(index, byte)| {
        if matches!(index, 8 | 13 | 18 | 23) {
            byte == b'-'
        } else {
            byte.is_ascii_digit() || (b'A'..=b'F').contains(&byte)
        }
    })
}
pub(super) fn checked_usize(value: u64, name: &str) -> Result<usize> {
    usize::try_from(value).map_err(|_source| limit(name))
}
pub(super) fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for byte in bytes {
        let mut index = ((crc >> 24) ^ u32::from(*byte)) & 0xFF;
        let mut table = index << 24;
        for _ in 0..8 {
            table = if table & 0x8000_0000 != 0 {
                (table << 1) ^ 0x04C1_1DB7
            } else {
                table << 1
            };
        }
        index = table;
        crc = (crc << 8) ^ index;
    }
    crc
}
pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
pub(super) fn limit(name: &str) -> Error {
    invalid(format!("MS-XLDM {name} limit exceeded"))
}
