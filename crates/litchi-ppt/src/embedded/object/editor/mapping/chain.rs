//! Resolves the newest visible persisted-record offset from `UserEdit` atoms.

use super::directory;
use crate::embedded::object::editor::{Result, rewrite};
use std::collections::{BTreeMap, HashSet};

const USER_EDIT: u16 = 4085;

/// Walk the `UserEdit` chain from `edit_offset`, newest first. Every
/// deviation is an error; the editor uses this.
pub(crate) fn read(document: &[u8], edit_offset: u32) -> Result<(BTreeMap<u32, u32>, u32)> {
    read_reporting(
        document,
        edit_offset,
        None,
        crate::leniency::Leniency::Strict,
    )
}

/// Walk the `UserEdit` chain from `edit_offset`, newest first, repairing
/// persist-directory defects into `report` when one is given (the read
/// path; see [`directory::merge_reporting`]).
///
/// A repeated edit offset (a cycle) is an error unless `leniency` opts into
/// [`crate::Leniency::TolerateUserEditChainDefects`] and a report is given:
/// then the walk stops at the repeated offset, keeps the mappings collected
/// so far — newest first, so the live objects are already covered — and
/// records [`crate::UserEditChainDefect::Cycle`] (finding F5.9), as
/// LibreOffice stops its walk (`svdfppt.cxx`). The chain-length bound stays
/// an error.
pub(crate) fn read_reporting(
    document: &[u8],
    mut edit_offset: u32,
    mut report: Option<&mut crate::leniency::ToleranceReport>,
    leniency: crate::leniency::Leniency,
) -> Result<(BTreeMap<u32, u32>, u32)> {
    let mut mapping = BTreeMap::new();
    let mut seen = HashSet::new();
    let mut document_id = 0;

    while edit_offset != 0 {
        if !seen.insert(edit_offset) {
            if leniency.tolerates_user_edit_chain_defects()
                && let Some(report) = report.as_deref_mut()
            {
                report.record(crate::leniency::UserEditChainDefect::Cycle, edit_offset);
                break;
            }
            return Err(crate::package::Error::Corrupted(
                "cyclic or excessive UserEdit chain".into(),
            ));
        }
        if seen.len() > 4_096 {
            return Err(crate::package::Error::Corrupted(
                "cyclic or excessive UserEdit chain".into(),
            ));
        }
        let record = rewrite::slice(document, edit_offset as usize)?;
        if rewrite::type_of(record)? != USER_EDIT || record.len() < 36 {
            return Err(crate::package::Error::Corrupted(
                "invalid UserEditAtom".into(),
            ));
        }

        let data = &record[8..];
        if document_id == 0 {
            document_id = rewrite::u32_at(data, 16)?;
        }
        let directory = rewrite::slice(document, rewrite::u32_at(data, 12)? as usize)?;
        directory::merge_reporting(&mut mapping, directory, report.as_deref_mut())?;
        edit_offset = rewrite::u32_at(data, 8)?;
    }

    if document_id == 0 {
        return Err(crate::package::Error::Corrupted(
            "missing Document persist ID".into(),
        ));
    }
    Ok((mapping, document_id))
}
