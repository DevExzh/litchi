use crate::error::{OoxmlError, Result};
use litchi_cfb::OleFile;
use litchi_cfb::writer::OleWriter;
use litchi_crypto::spaces::{
    Definition, ENCRYPTION_ID, ENCRYPTION_NAME, EncryptionTransform, Header, Map, MapEntry,
    PRIMARY, Reference, ReferenceKind, STORAGE, Version, VersionInfo, inspect, write_definition,
    write_encryption_transform, write_map, write_version_info,
};

/// Build an OLE compound file that wraps the given OOXML `EncryptionInfo`
/// and `EncryptedPackage` streams with the standard StrongEncryptionDataSpace
/// DataSpaces structure.
pub(crate) fn build_ole_encrypted_package(
    encryption_info: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    let dataspace_map = write_map(&Map {
        entries: vec![MapEntry {
            references: vec![Reference {
                kind: ReferenceKind::Stream,
                component: "EncryptedPackage".to_string(),
            }],
            data_space_name: "StrongEncryptionDataSpace".to_string(),
        }],
    })
    .map_err(data_space_error)?;
    let dataspace_def = write_definition(&Definition {
        transforms: vec!["StrongEncryptionTransform".to_string()],
    })
    .map_err(data_space_error)?;
    let transform_primary = write_encryption_transform(&EncryptionTransform {
        header: Header {
            transform_id: ENCRYPTION_ID.to_string(),
            transform_name: ENCRYPTION_NAME.to_string(),
            reader: Version::V1_0,
            updater: Version::V1_0,
            writer: Version::V1_0,
        },
        encryption_name: None,
        encryption_block_size: 16,
        cipher_mode: 0,
    })
    .map_err(data_space_error)?;
    let dataspace_version =
        write_version_info(&VersionInfo::default()).map_err(data_space_error)?;

    let mut writer = OleWriter::new();

    writer
        .create_stream(&["EncryptionInfo"], encryption_info)
        .map_err(|e| OoxmlError::Other(format!("failed to create EncryptionInfo stream: {e}")))?;

    writer
        .create_stream(&["EncryptedPackage"], encrypted_package)
        .map_err(|e| OoxmlError::Other(format!("failed to create EncryptedPackage stream: {e}")))?;

    writer
        .create_storage(&[STORAGE])
        .map_err(|e| OoxmlError::Other(format!("failed to create DataSpaces storage: {e}")))?;
    writer
        .create_storage(&[STORAGE, "DataSpaceInfo"])
        .map_err(|e| OoxmlError::Other(format!("failed to create DataSpaceInfo storage: {e}")))?;
    writer
        .create_storage(&[STORAGE, "TransformInfo"])
        .map_err(|e| OoxmlError::Other(format!("failed to create TransformInfo storage: {e}")))?;
    writer
        .create_storage(&[STORAGE, "TransformInfo", "StrongEncryptionTransform"])
        .map_err(|e| {
            OoxmlError::Other(format!(
                "failed to create StrongEncryptionTransform storage: {e}"
            ))
        })?;

    writer
        .create_stream(&[STORAGE, "DataSpaceMap"], &dataspace_map)
        .map_err(|e| OoxmlError::Other(format!("failed to create DataSpaceMap stream: {e}")))?;
    writer
        .create_stream(
            &[STORAGE, "DataSpaceInfo", "StrongEncryptionDataSpace"],
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
                STORAGE,
                "TransformInfo",
                "StrongEncryptionTransform",
                PRIMARY,
            ],
            &transform_primary,
        )
        .map_err(|e| {
            OoxmlError::Other(format!(
                "failed to create StrongEncryptionTransform/Primary stream: {e}"
            ))
        })?;
    writer
        .create_stream(&[STORAGE, "Version"], &dataspace_version)
        .map_err(|e| {
            OoxmlError::Other(format!("failed to create DataSpaces/Version stream: {e}"))
        })?;

    let mut cursor = std::io::Cursor::new(Vec::new());
    writer
        .write_to(&mut cursor)
        .map_err(|e| OoxmlError::Other(format!("failed to write OLE container: {e}")))?;

    Ok(cursor.into_inner())
}

pub(crate) fn parse_ole_encrypted_package(bytes: &[u8]) -> Result<(Vec<u8>, Vec<u8>)> {
    use std::io::Cursor;

    let cursor = Cursor::new(bytes);
    let mut ole = OleFile::open(cursor).map_err(|e| {
        OoxmlError::InvalidFormat(format!("invalid OLE container for encrypted OOXML: {}", e))
    })?;
    // LibreOffice has historically emitted otherwise valid encrypted
    // packages without DataSpaces. Preserve that compatibility fallback, but
    // require a complete, exact StrongEncryption graph whenever it is present.
    if let Some(graph) = inspect(&mut ole).map_err(data_space_error)?
        && graph.map.entries.as_slice()
            != [MapEntry {
                references: vec![Reference {
                    kind: ReferenceKind::Stream,
                    component: "EncryptedPackage".to_string(),
                }],
                data_space_name: "StrongEncryptionDataSpace".to_string(),
            }]
    {
        return Err(OoxmlError::InvalidFormat(
            "encrypted OOXML has an invalid StrongEncryptionDataSpace map".into(),
        ));
    }

    let encryption_info = ole.open_stream(&["EncryptionInfo"]).map_err(|e| {
        OoxmlError::InvalidFormat(format!("failed to read EncryptionInfo stream: {}", e))
    })?;

    let encrypted_package = ole.open_stream(&["EncryptedPackage"]).map_err(|e| {
        OoxmlError::InvalidFormat(format!("failed to read EncryptedPackage stream: {}", e))
    })?;

    Ok((encryption_info, encrypted_package))
}

fn data_space_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::InvalidFormat(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_cfb::OleWriter;
    use litchi_crypto::spaces::{ENCRYPTION_ID, inspect_bytes};

    #[test]
    fn encrypted_package_wrapper_has_valid_typed_spaces() {
        let encryption_info = [3, 0, 2, 0, 0, 0, 0, 0];
        let encrypted_package = [4, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 4];
        let bytes = build_ole_encrypted_package(&encryption_info, &encrypted_package).unwrap();

        let graph = inspect_bytes(&bytes).unwrap().unwrap();
        assert!(graph.irm.is_none());
        assert_eq!(
            graph.map.entries[0].data_space_name,
            "StrongEncryptionDataSpace"
        );
        assert_eq!(graph.transforms[0].header.transform_id, ENCRYPTION_ID);
        assert_eq!(
            graph.transforms[0]
                .encryption
                .as_ref()
                .unwrap()
                .encryption_block_size,
            16
        );
        let (parsed_info, parsed_package) = parse_ole_encrypted_package(&bytes).unwrap();
        assert_eq!(parsed_info, encryption_info);
        assert_eq!(parsed_package, encrypted_package);
    }

    #[test]
    fn accepts_libreoffice_package_without_spaces() {
        let encryption_info = [3, 0, 2, 0, 0, 0, 0, 0];
        let encrypted_package = [0; 8];
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["EncryptionInfo"], &encryption_info)
            .unwrap();
        writer
            .create_stream(&["EncryptedPackage"], &encrypted_package)
            .unwrap();
        let mut bytes = std::io::Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();

        let (parsed_info, parsed_package) =
            parse_ole_encrypted_package(&bytes.into_inner()).unwrap();
        assert_eq!(parsed_info, encryption_info);
        assert_eq!(parsed_package, encrypted_package);
    }
}
