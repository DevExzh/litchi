use crate::ole::writer::OleWriter;
use crate::ooxml::error::{OoxmlError, Result};

/// Build an OLE compound file that wraps the given OOXML `EncryptionInfo`
/// and `EncryptedPackage` streams with the standard StrongEncryptionDataSpace
/// DataSpaces structure.
pub(crate) fn build_ole_encrypted_package(
    encryption_info: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    let dataspace_map = build_dataspace_map_stream();
    let dataspace_def = build_dataspace_definition_stream();
    let transform_primary = build_transform_primary_stream();
    let dataspace_version = build_dataspace_version_stream();

    let ds_root = "\u{0006}DataSpaces";

    let mut writer = OleWriter::new();

    writer
        .create_stream(&["EncryptionInfo"], encryption_info)
        .map_err(|e| OoxmlError::Other(format!("failed to create EncryptionInfo stream: {e}")))?;

    writer
        .create_stream(&["EncryptedPackage"], encrypted_package)
        .map_err(|e| OoxmlError::Other(format!("failed to create EncryptedPackage stream: {e}")))?;

    writer
        .create_storage(&[ds_root])
        .map_err(|e| OoxmlError::Other(format!("failed to create DataSpaces storage: {e}")))?;
    writer
        .create_storage(&[ds_root, "DataSpaceInfo"])
        .map_err(|e| OoxmlError::Other(format!("failed to create DataSpaceInfo storage: {e}")))?;
    writer
        .create_storage(&[ds_root, "TransformInfo"])
        .map_err(|e| OoxmlError::Other(format!("failed to create TransformInfo storage: {e}")))?;
    writer
        .create_storage(&[ds_root, "TransformInfo", "StrongEncryptionTransform"])
        .map_err(|e| {
            OoxmlError::Other(format!(
                "failed to create StrongEncryptionTransform storage: {e}"
            ))
        })?;

    writer
        .create_stream(&[ds_root, "DataSpaceMap"], &dataspace_map)
        .map_err(|e| OoxmlError::Other(format!("failed to create DataSpaceMap stream: {e}")))?;
    writer
        .create_stream(
            &[ds_root, "DataSpaceInfo", "StrongEncryptionDataSpace"],
            &dataspace_def,
        )
        .map_err(|e| {
            OoxmlError::Other(format!(
                "failed to create StrongEncryptionDataSpace stream: {e}"
            ))
        })?;
    writer
        .create_stream(
            &[
                ds_root,
                "TransformInfo",
                "StrongEncryptionTransform",
                "\u{0006}Primary",
            ],
            &transform_primary,
        )
        .map_err(|e| {
            OoxmlError::Other(format!(
                "failed to create StrongEncryptionTransform/Primary stream: {e}"
            ))
        })?;
    writer
        .create_stream(&[ds_root, "Version"], &dataspace_version)
        .map_err(|e| {
            OoxmlError::Other(format!("failed to create DataSpaces/Version stream: {e}"))
        })?;

    let mut cursor = std::io::Cursor::new(Vec::new());
    writer
        .write_to(&mut cursor)
        .map_err(|e| OoxmlError::Other(format!("failed to write OLE container: {e}")))?;

    Ok(cursor.into_inner())
}

fn write_unicode_lpp4(buf: &mut Vec<u8>, s: &str) {
    let mut bytes = Vec::with_capacity(s.len() * 2);
    for ch in s.encode_utf16() {
        let b = ch.to_le_bytes();
        bytes.push(b[0]);
        bytes.push(b[1]);
    }
    let len = bytes.len() as u32;
    buf.extend_from_slice(&len.to_le_bytes());
    buf.extend_from_slice(&bytes);
    if (len % 4) == 2 {
        buf.extend_from_slice(&0u16.to_le_bytes());
    }
}

fn write_utf8_lpp4_null(buf: &mut Vec<u8>) {
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
}

fn build_dataspace_map_stream() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&8u32.to_le_bytes());
    buf.extend_from_slice(&1u32.to_le_bytes());

    let entry_start = buf.len();
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&1u32.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    write_unicode_lpp4(&mut buf, "EncryptedPackage");
    write_unicode_lpp4(&mut buf, "StrongEncryptionDataSpace");
    let entry_len = (buf.len() - entry_start) as u32;
    buf[entry_start..entry_start + 4].copy_from_slice(&entry_len.to_le_bytes());

    buf
}

fn build_dataspace_definition_stream() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&8u32.to_le_bytes());
    buf.extend_from_slice(&1u32.to_le_bytes());
    write_unicode_lpp4(&mut buf, "StrongEncryptionTransform");
    buf
}

fn build_transform_primary_stream() -> Vec<u8> {
    let mut buf = Vec::new();

    let header_start = buf.len();
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&1u32.to_le_bytes());
    write_unicode_lpp4(&mut buf, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
    let header_len = (buf.len() - header_start) as u32;
    buf[header_start..header_start + 4].copy_from_slice(&header_len.to_le_bytes());

    write_unicode_lpp4(&mut buf, "Microsoft.Container.EncryptionTransform");
    buf.extend_from_slice(&1u16.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&1u16.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&1u16.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());

    buf.extend_from_slice(&0u32.to_le_bytes());
    write_utf8_lpp4_null(&mut buf);
    buf.extend_from_slice(&4u32.to_le_bytes());

    buf
}

fn build_dataspace_version_stream() -> Vec<u8> {
    let mut buf = Vec::new();
    write_unicode_lpp4(&mut buf, "Microsoft.Container.DataSpaces");
    buf.extend_from_slice(&1u16.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&1u16.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&1u16.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf
}
