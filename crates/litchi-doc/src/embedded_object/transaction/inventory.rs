//! Read-only inventory projection over the current editor snapshot.

use super::super::codec::corrupted;
use super::super::model::{Editor, Entry, Inventory, Metadata, Reference};
use crate::package::Result;
use litchi_ole_common::object::{Object, Objects};

impl Editor {
    /// Returns the current managed-object snapshot with inert OLE metadata.
    ///
    /// The projection only reads the editor's current candidate. It never
    /// opens, follows, renders, instantiates, or executes an embedded object,
    /// and it does not publish any state. Mutations continue to operate on a
    /// cloned candidate and therefore remain atomic if a later validation
    /// step fails.
    pub fn inventory(&self) -> Result<Inventory> {
        let references = self.objects()?;
        let mut entries = Vec::with_capacity(references.len());
        for reference in references {
            let object =
                object_for_reference(self.package.objects(), &reference).ok_or_else(|| {
                    corrupted(format!(
                        "managed embedded-object storage {:?} is missing",
                        reference.storage_name
                    ))
                })?;
            entries.push(Entry::from_parts(reference, Metadata::of(object)?));
        }
        Ok(Inventory::from_entries(entries))
    }
}

fn object_for_reference<'a>(objects: &'a Objects, reference: &Reference) -> Option<&'a Object> {
    objects.get(&reference.storage_name).or_else(|| {
        objects.iter().find(|object| {
            object
                .path()
                .last()
                .and_then(|name| name.strip_prefix('_'))
                .and_then(|value| value.parse::<u32>().ok())
                == Some(reference.storage_id)
        })
    })
}
