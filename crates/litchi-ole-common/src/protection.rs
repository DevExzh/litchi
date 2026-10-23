//! Shared mutation guards for protected CFB containers.

use litchi_cfb::{OleError, OleFile, SharedOleFile};
use std::io::{Read, Seek};

const STORAGE_ENTRY: u8 = 1;
const ROOT_ENTRY: u8 = 5;

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

/// Returns whether a directory component carries document-signature state.
///
/// Signature components are the one protected family that a read-only probe
/// may retain.  Mutation-capable owners continue to reject every component
/// through [`is_protected_component`].
#[must_use]
pub(crate) fn is_signature_component(name: &str) -> bool {
    [
        "_xmlsignatures",
        "_signatures",
        "DigitalSignature",
        "\u{0005}DigitalSignature",
    ]
    .iter()
    .any(|marker| name.eq_ignore_ascii_case(marker))
}

/// Returns whether a protected component is encryption or DRM state.
#[must_use]
pub(crate) fn is_encryption_or_drm_component(name: &str) -> bool {
    is_protected_component(name) && !is_signature_component(name)
}

/// Reject containers whose directory contains a signing, encryption, or DRM
/// component. Both streams and storages are inspected, including empty
/// storages that are not visible through `OleFile::list_streams`.
pub(crate) fn reject_protected_container<R: Read + Seek>(
    ole: &OleFile<R>,
    operation: &'static str,
    max_entries: usize,
) -> Result<(), OleError> {
    inspect_protected_container(ole, operation, max_entries, false).map(|_| ())
}

/// Admits a bounded read-only view of a container while retaining signature
/// streams. Encryption and DRM components remain refused because the probe
/// cannot decrypt or safely expose those payloads.
pub(crate) fn admit_readonly_container<R: Read + Seek>(
    ole: &OleFile<R>,
    operation: &'static str,
    max_entries: usize,
) -> Result<bool, OleError> {
    inspect_protected_container(ole, operation, max_entries, true)
}

fn inspect_protected_container<R: Read + Seek>(
    ole: &OleFile<R>,
    operation: &'static str,
    max_entries: usize,
    allow_signatures: bool,
) -> Result<bool, OleError> {
    let mut pending = Vec::<Vec<String>>::new();
    pending
        .try_reserve(1)
        .map_err(|source| OleError::Allocation {
            resource: "protected-container traversal",
            source,
        })?;
    pending.push(Vec::new());

    // The CFB parser has already retained a validated directory index.  Use
    // its finite slot count as a hard upper bound even when a format owner
    // does not provide a tighter object/package budget.
    let max_entries = max_entries.min(ole.directory_entry_count());
    let mut discovered = 0usize;
    let mut signed = false;
    while let Some(path) = pending.pop() {
        let remaining = max_entries.saturating_sub(discovered);
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        ole.visit_directory_entry_refs(&refs, remaining, |entry| {
            discovered = discovered.saturating_add(1);
            if is_protected_component(&entry.name) {
                if allow_signatures && is_signature_component(&entry.name) {
                    signed = true;
                } else if is_encryption_or_drm_component(&entry.name) || !allow_signatures {
                    return Err(OleError::InvalidFormat(format!(
                        "signed, encrypted, or DRM containers are not eligible for {operation}"
                    )));
                }
            }
            if entry.entry_type == STORAGE_ENTRY {
                let mut child = clone_path(&path)?;
                child
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "protected-container path",
                        source,
                    })?;
                let mut name = String::new();
                name.try_reserve_exact(entry.name.len()).map_err(|source| {
                    OleError::Allocation {
                        resource: "protected-container path component",
                        source,
                    }
                })?;
                name.push_str(&entry.name);
                child.push(name);
                pending
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "protected-container traversal",
                        source,
                    })?;
                pending.push(child);
            }
            Ok(())
        })?;
    }
    Ok(signed)
}

fn clone_path(path: &[String]) -> Result<Vec<String>, OleError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(path.len())
        .map_err(|source| OleError::Allocation {
            resource: "protected-container path",
            source,
        })?;
    for component in path {
        let mut value = String::new();
        value
            .try_reserve_exact(component.len())
            .map_err(|source| OleError::Allocation {
                resource: "protected-container path component",
                source,
            })?;
        value.push_str(component);
        output.push(value);
    }
    Ok(output)
}

pub(crate) fn reject_protected_shared_container(
    ole: &SharedOleFile,
    operation: &'static str,
) -> Result<(), OleError> {
    if ole
        .directory_entries()
        .any(|entry| entry.entry_type != ROOT_ENTRY && is_protected_component(&entry.name))
    {
        let prefix = "signed, encrypted, or DRM containers are not eligible for ";
        // Both are live string slices, so their combined length necessarily
        // fits the address space even before the fallible allocation below.
        let capacity = prefix.len() + operation.len();
        let mut message = String::new();
        message
            .try_reserve_exact(capacity)
            .map_err(|source| OleError::Allocation {
                resource: "protected-container refusal",
                source,
            })?;
        message.push_str(prefix);
        message.push_str(operation);
        return Err(OleError::InvalidFormat(message));
    }
    Ok(())
}
