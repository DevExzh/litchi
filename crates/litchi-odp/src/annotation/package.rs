//! Atomic ODP package reconstruction for annotation edits.

use crate::core::OwnedPackage;
use litchi_core::Result;
use litchi_odf_common::package::rebuild_package;

pub(crate) fn rebuild(source: &OwnedPackage, content: &str) -> Result<Vec<u8>> {
    rebuild_package(
        source,
        content,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}
