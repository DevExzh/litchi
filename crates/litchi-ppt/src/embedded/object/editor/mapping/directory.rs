//! Decodes one full or incremental `PersistDirectoryAtom`.

use crate::embedded::object::editor::{Result, rewrite};
use crate::leniency::{DefectLocation, PersistDefect, ToleranceReport};
use std::collections::{BTreeMap, HashSet};

const PERSIST_FULL: u16 = 6001;
const PERSIST_INCREMENTAL: u16 = 6002;

/// Merge one directory into `mapping`, first-wins (the chain is walked
/// newest first).
///
/// Without a report (the editor) every deviation is an error. With a report (the read path, finding F5.7), the deviations LibreOffice
/// reads past are repaired and recorded instead of rejected, as
/// `svdfppt.cxx` does: a pointer to a record other than a
/// `PersistDirectoryAtom` contributes nothing, a zero-count entry header is
/// skipped, and a persist ID repeated within the directory keeps its first
/// offset. A truncated directory stays an error.
pub(crate) fn merge_reporting(
    mapping: &mut BTreeMap<u32, u32>,
    directory: &[u8],
    mut report: Option<&mut ToleranceReport>,
) -> Result<()> {
    let kind = rewrite::record::type_of(directory)?;
    if !matches!(kind, PERSIST_FULL | PERSIST_INCREMENTAL) {
        let Some(report) = report else {
            return Err(crate::package::Error::Corrupted(
                "invalid PersistDirectoryAtom".into(),
            ));
        };
        report.record(PersistDefect::PersistDirectoryRecordType, u32::from(kind));
        return Ok(());
    }

    let mut seen = HashSet::new();
    let mut offset = 8usize;
    while offset < directory.len() {
        let info = rewrite::record::u32_at(directory, offset)?;
        offset += 4;
        let base = info & 0x000F_FFFF;
        let count = info >> 20;
        if count == 0 {
            let Some(report) = report.as_deref_mut() else {
                return Err(crate::package::Error::Corrupted("zero persist run".into()));
            };
            report.record(PersistDefect::ZeroCountPersistRun, base);
            continue;
        }
        for index in 0..count {
            let value = rewrite::record::u32_at(directory, offset)?;
            offset += 4;
            let persist_id = base + index;
            if let Some(report) = report.as_deref_mut()
                && !seen.insert(persist_id)
            {
                report.record_at(
                    PersistDefect::DuplicatePersistId,
                    persist_id,
                    DefectLocation {
                        persist_id: Some(persist_id),
                        shape_id: None,
                    },
                );
                continue;
            }
            mapping.entry(persist_id).or_insert(value);
        }
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn directory(record_type: u16, words: &[u32]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(words.len() * 4).unwrap().to_le_bytes());
        for word in words {
            bytes.extend_from_slice(&word.to_le_bytes());
        }
        bytes
    }

    fn merged(bytes: &[u8]) -> (Result<BTreeMap<u32, u32>>, Vec<(crate::Defect, u32)>) {
        let mut report = ToleranceReport::default();
        let mut mapping = BTreeMap::new();
        let result = merge_reporting(&mut mapping, bytes, Some(&mut report)).map(|()| mapping);
        let defects = report
            .defects()
            .iter()
            .map(|defect| (defect.defect, defect.context))
            .collect();
        (result, defects)
    }

    /// The editor stays strict; the read path repairs and records (F5.7).
    #[test]
    fn read_path_repairs_what_the_editor_rejects() {
        use crate::Defect;

        let wrong_type = directory(4085, &[0x0010_0001, 8]);
        let zero_run = directory(PERSIST_INCREMENTAL, &[0x0000_0009, 0x0010_0001, 8]);
        let duplicate = directory(PERSIST_INCREMENTAL, &[0x0020_0001, 8, 16, 0x0010_0002, 24]);
        for bytes in [&wrong_type, &zero_run] {
            assert!(merge_reporting(&mut BTreeMap::new(), bytes, None).is_err());
        }
        // Duplicates were always first-wins on the editor path.
        let mut strict = BTreeMap::new();
        merge_reporting(&mut strict, &duplicate, None).unwrap();
        assert_eq!(strict.get(&2), Some(&16));

        let (mapping, defects) = merged(&wrong_type);
        assert!(mapping.unwrap().is_empty());
        assert_eq!(
            defects,
            [(
                Defect::Persist(PersistDefect::PersistDirectoryRecordType),
                4085
            )]
        );

        let (mapping, defects) = merged(&zero_run);
        assert_eq!(mapping.unwrap().get(&1), Some(&8));
        assert_eq!(
            defects,
            [(Defect::Persist(PersistDefect::ZeroCountPersistRun), 9)]
        );

        let (mapping, defects) = merged(&duplicate);
        assert_eq!(mapping.unwrap().get(&2), Some(&16));
        assert_eq!(
            defects,
            [(Defect::Persist(PersistDefect::DuplicatePersistId), 2)]
        );

        // A truncated run stays an error.
        let truncated = directory(PERSIST_INCREMENTAL, &[0x0030_0001, 8]);
        assert!(merged(&truncated).0.is_err());
    }
}
