//! Worksheet-stream snapshots and clone-staged edits.

use super::codec::{self, Parsed};
use super::model::{Item, Reference, UnknownRecord, Watch, Watches};
use super::phonetic::Info;
use super::validation;
use crate::package::error::{Error, Result};
use crate::raw::kind;
use std::sync::Arc;

/// An immutable, parsed worksheet snapshot.
///
/// The source BIFF12 allocation is shared by clones and edits. Untouched
/// records therefore remain byte-for-byte identical, including unknown
/// records and their original order.
#[derive(Debug, Clone)]
pub struct Snapshot {
    parsed: Parsed,
    watches: Watches,
}

impl Snapshot {
    /// Return typed watches in their original source order.
    #[must_use]
    pub fn watches(&self) -> &[Watch] {
        self.watches.as_slice()
    }

    /// Return the worksheet-wide phonetic default, if present.
    #[must_use]
    pub const fn phonetic(&self) -> Option<Info> {
        self.parsed.phonetic
    }

    /// Whether the source contains a `BrtBeginCellWatches` collection.
    #[must_use]
    pub const fn has_collection(&self) -> bool {
        self.parsed.watch_block.is_some()
    }

    /// Whether the watch collection contains opaque producer records.
    #[must_use]
    pub fn has_opaque_records(&self) -> bool {
        self.parsed
            .items
            .iter()
            .any(|item| matches!(item, Item::Unknown(_)))
    }

    /// Materialize the bounded opaque records in their source order.
    ///
    /// The common case has no opaque records and returns an empty vector
    /// without copying the worksheet stream.
    pub fn opaque_records(&self) -> Result<Vec<UnknownRecord>> {
        self.parsed
            .items
            .iter()
            .filter_map(|item| match item {
                Item::Watch(_) => None,
                Item::Unknown(index) => Some(codec::opaque_record(&self.parsed, *index)),
            })
            .collect()
    }

    /// Start a detached edit against this snapshot.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            parsed: self.parsed.clone(),
            items: self.parsed.items.clone(),
            collection: self.parsed.watch_block.is_some(),
            phonetic: self.parsed.phonetic,
        }
    }

    /// Return the exact source worksheet bytes for diagnostics and patch
    /// lineage checks.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.parsed.source
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source_bytes() == other.source_bytes()
    }
}

impl Eq for Snapshot {}

/// A detached worksheet edit. It never mutates its source snapshot.
#[derive(Debug, Clone)]
pub struct Edit {
    parsed: Parsed,
    items: Vec<Item>,
    collection: bool,
    phonetic: Option<Info>,
}

impl Edit {
    /// Return the current typed watches in source-relative order.
    pub fn watches(&self) -> Watches {
        Watches::from_validated(watches_from_items(&self.items))
    }

    /// Add one watched cell at the end of the typed collection.
    pub fn add(&mut self, watch: Watch) -> Result<()> {
        let mut watches = self.watches();
        watches.items_mut().push(watch);
        validation::watches(&watches)?;
        self.items.push(Item::Watch(watch));
        Ok(())
    }

    /// Remove the first watch for a semantic cell reference.
    pub fn remove(&mut self, reference: Reference) -> bool {
        let Some(index) = self
            .items
            .iter()
            .position(|item| matches!(item, Item::Watch(watch) if watch.reference() == reference))
        else {
            return false;
        };
        self.items.remove(index);
        true
    }

    /// Remove all typed watches while retaining opaque records in place.
    pub fn clear(&mut self) -> bool {
        let old_len = self.items.len();
        self.items.retain(|item| matches!(item, Item::Unknown(_)));
        old_len != self.items.len()
    }

    /// Replace all typed watches while retaining every opaque slot and its
    /// relative position. New watches are appended after the retained slots.
    pub fn replace(&mut self, watches: Watches) -> Result<()> {
        validation::watches(&watches)?;
        let mut incoming = watches.as_slice().iter().copied();
        let mut items = Vec::with_capacity(self.items.len().max(watches.len()));
        for item in &self.items {
            match item {
                Item::Unknown(index) => items.push(Item::Unknown(*index)),
                Item::Watch(_) => {
                    if let Some(watch) = incoming.next() {
                        items.push(Item::Watch(watch));
                    }
                },
            }
        }
        items.extend(incoming.map(Item::Watch));
        self.items = items;
        Ok(())
    }

    /// Return the worksheet-wide phonetic default currently staged.
    #[must_use]
    pub const fn phonetic(&self) -> Option<Info> {
        self.phonetic
    }

    /// Set or replace the worksheet-wide phonetic default.
    pub fn set_phonetic(&mut self, value: Info) {
        self.phonetic = Some(value);
    }

    /// Remove the worksheet-wide phonetic default.
    pub fn clear_phonetic(&mut self) -> bool {
        self.phonetic.take().is_some()
    }

    /// Validate and publish a new immutable snapshot plus reversible patch.
    pub fn commit(self) -> Result<Commit> {
        let watches = Watches::new(watches_from_items(&self.items))?;
        let bytes = render(&self.parsed, &self.items, self.collection, self.phonetic)?;
        let (after, snapshot) = if bytes.as_slice() == self.parsed.source.as_ref() {
            let snapshot = Snapshot {
                parsed: self.parsed.clone(),
                watches,
            };
            (Arc::clone(&self.parsed.source), snapshot)
        } else {
            let parsed = codec::parse_stream(&bytes)?;
            let snapshot = Snapshot { parsed, watches };
            (Arc::from(bytes), snapshot)
        };
        Ok(Commit {
            snapshot,
            patch: Patch {
                before: Arc::clone(&self.parsed.source),
                after,
            },
        })
    }
}

/// A successful immutable edit result.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the committed worksheet snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-byte patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split the commit into its new snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible, source-checked byte patch for one worksheet stream.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    /// Whether this patch leaves the worksheet byte stream unchanged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.as_ref() == self.after.as_ref()
    }

    /// Return the exact before-image bytes.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Return the exact after-image bytes.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Apply this patch only to its exact source snapshot.
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>> {
        if source != self.before.as_ref() {
            return Err(Error::UnsupportedFeature(
                "cell-watch patch source snapshot does not match".to_string(),
            ));
        }
        Ok(self.after.to_vec())
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Parse one complete BIFF12 worksheet stream.
pub fn read(data: &[u8]) -> Result<Snapshot> {
    let parsed = codec::parse_stream(data)?;
    let watches = Watches::new(watches_from_items(&parsed.items))?;
    Ok(Snapshot { parsed, watches })
}

/// Apply a previously committed patch to a complete worksheet stream.
pub fn apply(data: &[u8], patch: &Patch) -> Result<Vec<u8>> {
    patch.apply(data)
}

fn watches_from_items(items: &[Item]) -> Vec<Watch> {
    items
        .iter()
        .filter_map(|item| match item {
            Item::Watch(watch) => Some(*watch),
            Item::Unknown(_) => None,
        })
        .collect()
}

fn render(
    parsed: &Parsed,
    items: &[Item],
    collection: bool,
    phonetic: Option<Info>,
) -> Result<Vec<u8>> {
    validation::watches(&Watches::new(watches_from_items(items))?)?;
    let mut output = Vec::with_capacity(parsed.source.len());
    let mut index = 0usize;
    let existing_block = parsed.watch_block;
    let mut watch_inserted = false;
    let mut phonetic_inserted = false;
    while index < parsed.records.len() {
        if index == parsed.end_sheet_index {
            if existing_block.is_none() && (collection || !items.is_empty()) {
                append_watch_block(&mut output, parsed, items)?;
                watch_inserted = true;
            }
            if parsed.phonetic_index.is_none() && !phonetic_inserted {
                if let Some(value) = phonetic {
                    output.extend_from_slice(&codec::encode_record(
                        kind::PHONETIC_INFO,
                        &codec::write_phonetic(value)?,
                    )?);
                    phonetic_inserted = true;
                }
            }
        }

        if let Some((begin, end)) = existing_block
            && index == begin
        {
            if collection {
                append_watch_block(&mut output, parsed, items)?;
            }
            index = end.saturating_add(1);
            continue;
        }

        if parsed.phonetic_index == Some(index) {
            if let Some(value) = phonetic {
                output.extend_from_slice(&codec::encode_record(
                    kind::PHONETIC_INFO,
                    &codec::write_phonetic(value)?,
                )?);
            }
            index = index.saturating_add(1);
            continue;
        }

        let span = parsed.records.get(index).ok_or_else(|| {
            validation::invalid(
                "cell-watch worksheet render",
                "record index is out of bounds",
            )
        })?;
        output.extend_from_slice(span.raw(&parsed.source));
        index = index.saturating_add(1);
    }

    if existing_block.is_none() && !items.is_empty() && !watch_inserted {
        // `end_sheet_index` is required by parse_stream, so reaching this
        // branch indicates only a future internal layout bug.
        return Err(validation::invalid(
            "cell-watch worksheet render",
            "watch collection insertion boundary was not reached",
        ));
    }
    Ok(output)
}

fn append_watch_block(output: &mut Vec<u8>, parsed: &Parsed, items: &[Item]) -> Result<()> {
    output.extend_from_slice(&codec::encode_record(kind::BEGIN_CELL_WATCHES, &[])?);
    for item in items {
        match item {
            Item::Watch(watch) => output.extend_from_slice(&codec::encode_record(
                kind::CELL_WATCH,
                &codec::write_watch(*watch)?,
            )?),
            Item::Unknown(index) => {
                let span = parsed.records.get(*index).ok_or_else(|| {
                    validation::invalid(
                        "cell-watch worksheet render",
                        "opaque record index is out of bounds",
                    )
                })?;
                output.extend_from_slice(span.raw(&parsed.source));
            },
        }
    }
    output.extend_from_slice(&codec::encode_record(kind::END_CELL_WATCHES, &[])?);
    Ok(())
}
