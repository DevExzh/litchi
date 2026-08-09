//! Shared mutation guards for protected CFB containers.

use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

const STORAGE_ENTRY: u8 = 1;

/// Returns whether a CFB path component marks signed, encrypted, or DRM
/// content that mutation-capable format crates must preserve unchanged.
#[must_use]
pub fn is_protected_component(name: &str) -> bool {
    [
        "_xmlsignatures",
        "_signatures",
        "DigitalSignature",
        "\u{0005}DigitalSignature",
        "\u{0006}DataSpaces",
        "\u{0006}DataSpaceInfo",
        "\u{0006}TransformInfo",
        "\u{0006}Primary",
        "\u{0009}DRMContent",
        "\u{0009}DRMViewerContent",
        "EncryptedPackage",
        "EncryptionInfo",
    ]
    .iter()
    .any(|marker| name.eq_ignore_ascii_case(marker))
}

/// Reject containers whose directory contains a signing, encryption, or DRM
/// component. Both streams and storages are inspected, including empty
/// storages that are not visible through `OleFile::list_streams`.
pub(crate) fn reject_protected_container<R: Read + Seek>(
    ole: &OleFile<R>,
    operation: &'static str,
) -> Result<(), OleError> {
    let mut pending = Vec::<Vec<String>>::new();
    pending
        .try_reserve(1)
        .map_err(|source| OleError::Allocation {
            resource: "protected-container traversal",
            source,
        })?;
    pending.push(Vec::new());

    while let Some(path) = pending.pop() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        for entry in ole.list_directory_entries(&refs)? {
            if is_protected_component(&entry.name) {
                return Err(OleError::InvalidFormat(format!(
                    "signed, encrypted, or DRM containers are not eligible for {operation}"
                )));
            }
            if entry.entry_type == STORAGE_ENTRY {
                let mut child = path.clone();
                child
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "protected-container path",
                        source,
                    })?;
                child.push(entry.name.clone());
                pending
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "protected-container traversal",
                        source,
                    })?;
                pending.push(child);
            }
        }
    }
    Ok(())
}
