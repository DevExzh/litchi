//! Native text-storage conversion for structured semantic adapters.
//!
//! Protobuf text fragments are joined once into the leaf-owned storage value;
//! the semantic crate never sees the native message. Each native fragment is
//! retained as a validated archive-free run so later projections can borrow
//! the original allocation instead of reconstructing text.

use crate::protobuf::tswp::StorageArchive;
use litchi_iwa_text::storage::{MAX_RUNS, Run, Storage};

pub(super) fn from_archive(archive: StorageArchive, context: &str) -> crate::Result<Storage> {
    if archive.text.len() > MAX_RUNS {
        return Err(crate::Error::InvalidFormat(format!(
            "{context} contains {} text fragments; maximum is {MAX_RUNS}",
            archive.text.len()
        )));
    }

    let text_len = archive.text.iter().try_fold(0usize, |length, fragment| {
        length.checked_add(fragment.len()).ok_or_else(|| {
            crate::Error::InvalidFormat(format!(
                "{context} text length overflows the host address space"
            ))
        })
    })?;

    let mut text = String::new();
    text.try_reserve_exact(text_len).map_err(|_| {
        crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "structured semantic text",
            amount: text_len,
        })
    })?;

    let mut runs = Vec::new();
    runs.try_reserve_exact(archive.text.len()).map_err(|_| {
        crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "structured semantic text runs",
            amount: archive.text.len(),
        })
    })?;

    for fragment in archive.text {
        let start = text.len();
        let length = fragment.len();
        text.push_str(&fragment);
        runs.push(Run::new(start, length));
    }

    Storage::try_from_parts(text, runs).map_err(|error| {
        crate::Error::InvalidFormat(format!(
            "{context} semantic text storage is invalid: {error}"
        ))
    })
}
