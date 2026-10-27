//! Rewriting the package's metadata sidecars: `Metadata/Properties.plist`
//! and `Metadata/DocumentIdentifier`.
//!
//! Both are plain ZIP members outside the IWA object graph. A rewrite
//! replaces (or appends) only those members through the archive layer's
//! bounded reassembly, which copies every other member's records unchanged,
//! then reopens the candidate as an ordinary exact-source package.

use std::io::Cursor;

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion};

use super::{Package, PackageError, PackageResult};

/// The canonical properties sidecar.
pub(crate) const PROPERTIES: &str = "Metadata/Properties.plist";
/// The canonical document-identifier sidecar.
pub(crate) const IDENTIFIER: &str = "Metadata/DocumentIdentifier";

fn invalid(message: &str) -> PackageError {
    PackageError::InvalidFormat(message.to_owned())
}

/// The decoded properties dictionary, or `None` when the package has none.
///
/// # Errors
///
/// [`PackageError::InvalidFormat`] when the sidecar is not a plist
/// dictionary or uses an unsupported compression.
pub(crate) fn properties(package: &Package) -> PackageResult<Option<plist::Dictionary>> {
    let Some(entry) = package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == PROPERTIES)
    else {
        return Ok(None);
    };
    if entry.is_opaque() {
        return Err(invalid(
            "the Pages properties sidecar uses an unsupported compression",
        ));
    }
    match plist::Value::from_reader(Cursor::new(entry.data())) {
        Ok(plist::Value::Dictionary(dictionary)) => Ok(Some(dictionary)),
        _ => Err(invalid(
            "the Pages properties sidecar is not a property-list dictionary",
        )),
    }
}

/// Encode `dictionary` as the binary property list Pages writes.
pub(crate) fn encode(dictionary: &plist::Dictionary) -> PackageResult<Vec<u8>> {
    let mut bytes = Vec::new();
    plist::Value::Dictionary(dictionary.clone())
        .to_writer_binary(&mut bytes)
        .map_err(|_error| invalid("the Pages properties could not be encoded"))?;
    Ok(bytes)
}

/// The package with its properties replaced by `properties` and its
/// document identifier by `identifier` (each when given); a sidecar the
/// package lacks is appended.
///
/// # Errors
///
/// Archive errors when the reassembly exceeds the package's retained
/// limits; package errors when the candidate fails to reopen.
pub(crate) fn rewrite(
    package: &Package,
    properties: Option<&plist::Dictionary>,
    identifier: Option<&[u8]>,
) -> PackageResult<Package> {
    let encoded = properties.map(encode).transpose()?;
    let catalog = package.state.source.package();
    let present = |name: &str| catalog.iter().any(|entry| entry.name() == name);
    let mut edits = Vec::new();
    let mut insertions = Vec::new();
    for (name, data) in [(PROPERTIES, encoded.as_deref()), (IDENTIFIER, identifier)] {
        let Some(data) = data else { continue };
        if present(name) {
            edits.push(EntryEdit::new(name, data));
        } else {
            insertions.push(EntryInsertion::append(name, data));
        }
    }
    let limits = package.state.source.limits();
    let output = catalog.reassemble_with_changes_to_bytes(&insertions, &edits, &[], limits)?;
    let source = package.state.source.reopen(output.into())?;
    Package::from_source_catalog(source)
}
