//! Shared mutation guards for protected CFB containers.

use litchi_cfb::{OleError, OleFile, SharedOleFile};
use std::io::{Read, Seek};

const STORAGE_ENTRY: u8 = 1;
const ROOT_ENTRY: u8 = 5;

/// Bounds used while inspecting a raw compound-file protection marker.
///
/// Protection inspection walks directory entries, including empty storages
/// which are not visible through the stream listing.  Callers must choose a
/// nonzero bound so malformed directory trees cannot make this guard perform
/// an unbounded traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum number of directory entries inspected by the guard.
    pub max_entries: usize,
}

impl Limits {
    fn validate(self) -> Result<Self, OleError> {
        if self.max_entries == 0 {
            return Err(OleError::InvalidLimit {
                resource: "protected directory entries",
                value: 0,
                maximum: usize::MAX as u64,
            });
        }
        Ok(self)
    }
}

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
    limits: Limits,
) -> Result<(), OleError> {
    inspect_protected_container(ole, operation, limits, false).map(|_| ())
}

/// Admit signatures for an inert read-only probe; encryption and DRM remain refused.
pub(crate) fn admit_readonly_container<R: Read + Seek>(
    ole: &OleFile<R>,
    operation: &'static str,
    limits: Limits,
) -> Result<bool, OleError> {
    inspect_protected_container(ole, operation, limits, true)
}

fn inspect_protected_container<R: Read + Seek>(
    ole: &OleFile<R>,
    operation: &'static str,
    limits: Limits,
    allow_signatures: bool,
) -> Result<bool, OleError> {
    let limits = limits.validate()?;
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
    let max_entries = limits.max_entries.min(ole.directory_entry_count());
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
    limits: Limits,
) -> Result<(), OleError> {
    let limits = limits.validate()?;
    let mut entries = ole.directory_entries();
    for index in 0.. {
        let Some(entry) = entries.next() else {
            break;
        };
        if index >= limits.max_entries {
            return Err(OleError::LimitExceeded {
                resource: "protected directory entries",
                observed: (index as u64).saturating_add(1),
                maximum: limits.max_entries as u64,
            });
        }
        if entry.entry_type != ROOT_ENTRY && is_protected_component(&entry.name) {
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
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "test assertions panic by design")]

    use super::*;
    use litchi_cfb::{OleFile, OleWriter, SharedOleFile};
    use litchi_core::SourceVersion;
    use std::io::Cursor;
    use std::sync::Arc;

    fn cfb_with_two_streams() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer.create_stream(&["First"], b"first").unwrap();
        writer.create_stream(&["Second"], b"second").unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn protection_limit_rejects_zero() {
        let error = Limits { max_entries: 0 }.validate().unwrap_err();
        assert!(matches!(
            error,
            OleError::InvalidLimit {
                resource: "protected directory entries",
                value: 0,
                ..
            }
        ));
    }

    #[test]
    fn raw_and_shared_guards_charge_directory_entries() {
        let bytes = cfb_with_two_streams();
        let ole = OleFile::open(
            Cursor::new(bytes.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let error =
            reject_protected_container(&ole, "test", Limits { max_entries: 1 }).unwrap_err();
        assert!(matches!(
            error,
            OleError::LimitExceeded {
                resource: "directory entries",
                maximum: 1,
                ..
            }
        ));

        let shared = SharedOleFile::open_owned_vec(
            Arc::new(bytes),
            SourceVersion::new(0x0359, 1),
            litchi_cfb::SharedOleFileLimits::default(),
        )
        .unwrap();
        let error = reject_protected_shared_container(&shared, "test", Limits { max_entries: 1 })
            .unwrap_err();
        assert!(matches!(
            error,
            OleError::LimitExceeded {
                resource: "protected directory entries",
                maximum: 1,
                ..
            }
        ));
    }
}
