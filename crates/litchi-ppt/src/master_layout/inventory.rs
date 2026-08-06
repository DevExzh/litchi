//! Contextual discovery of master records in a parsed PPT tree.

use super::model::{Context, Entry, Inventory, Path};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Find master records in deterministic depth-first order.
pub fn inventory(root: &Record) -> Result<Inventory> {
    let mut entries = Vec::new();
    let mut stack = vec![(Path::root(), root)];
    while let Some((path, record)) = stack.pop() {
        if let Some(context) = classify(record)? {
            entries.push(Entry::new(path.clone(), context));
        }
        for (index, child) in record.children.iter().enumerate().rev() {
            stack.push((path.child(index), child));
        }
    }
    Ok(Inventory::new(entries))
}

fn classify(record: &Record) -> Result<Option<Context>> {
    match record.record_type {
        RecordType::MainMaster => Ok(Some(Context::Main)),
        RecordType::Handout => Ok(Some(Context::Handout)),
        RecordType::Notes => {
            let Some(atom) = record
                .children
                .iter()
                .find(|child| child.record_type == RecordType::NotesAtom)
            else {
                return Err(Error::Corrupted(
                    "Notes record is missing NotesAtom while building master inventory".into(),
                ));
            };
            if atom.data.len() < 4 {
                return Err(Error::Corrupted(
                    "NotesAtom is truncated while building master inventory".into(),
                ));
            }
            let slide_id = u32::from_le_bytes(atom.data[..4].try_into().unwrap_or([0; 4]));
            Ok((slide_id == 0).then_some(Context::Notes))
        },
        RecordType::Slide => {
            let Some(atom) = record
                .children
                .iter()
                .find(|child| child.record_type == RecordType::SlideAtom)
            else {
                return Ok(None);
            };
            if atom.data.len() < 4 {
                return Ok(None);
            }
            let geometry = u32::from_le_bytes(atom.data[..4].try_into().unwrap_or([0; 4]));
            Ok((geometry == 2).then_some(Context::Title))
        },
        _ => Ok(None),
    }
}
