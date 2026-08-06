//! Typed inventory and transactional snapshots for legacy PPT masters.
//!
//! Main and title masters are listed by `MasterPersistAtom` records, while
//! notes and handout masters are named by `DocumentAtom`.  This owner accepts
//! the existing parsed [`Record`](crate::Record) tree and an explicit persist
//! catalog, so it stays independent of a particular `Presentation` parser.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_presentation};
pub use model::{
    Handout, Identity, Inventory, Kind, Main, Master, Notes, Object, Objects, Persist, RecordRef,
    Scope, Title, Unknown,
};

use crate::package::Result;
use crate::presentation::Presentation;
use crate::records::Record;

impl Presentation {
    /// Enumerate the masters referenced by the live `DocumentContainer`.
    pub fn masters(&self) -> Result<Inventory<'_>> {
        let records = self.parser.find_records_ref();
        let mut persisted = Vec::new();
        for (&persist_id, &offset) in self.persist_mapping.iter() {
            if persist_id == 0 {
                continue;
            }
            let offset = usize::try_from(offset).map_err(|_| {
                crate::package::Error::Corrupted("persist offset exceeds usize".into())
            })?;
            let (resolved, _) = Record::parse(self.document_stream(), offset)?;
            if let Some(record) = records
                .iter()
                .copied()
                .find(|candidate| **candidate == resolved)
            {
                persisted.push((persist_id, record));
            }
        }
        let objects = Objects::from_records(persisted)?;
        parse_presentation(self, &objects)
    }
}
