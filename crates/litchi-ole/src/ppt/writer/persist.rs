//! PersistPtr offset mapping for PPT files
//!
//! The PersistPtr system maps logical persist IDs to physical file offsets,
//! allowing efficient random access to slides and other objects.
//!
//! Based on Microsoft's "[MS-PPT]" specification and Apache POI's UserEditAtom.

use super::records::{RecordBuilder, record_type};
use super::spec::{
    DEFAULT_LAST_VIEW_TYPE, DEFAULT_LAST_VIEWED_SLIDE_ID, PPT_VERSION, USER_EDIT_PADWORD,
};
use std::collections::HashMap;

/// Persist reference - maps ID to offset
#[derive(Debug, Clone)]
pub struct PersistRef {
    /// Persist ID
    pub persist_id: u32,
    /// File offset
    pub offset: u32,
}

/// PersistPtr directory builder
#[derive(Debug)]
pub struct PersistPtrBuilder {
    /// Map of persist ID to offset
    refs: HashMap<u32, u32>,
    /// Next available persist ID
    next_id: u32,
}

impl PersistPtrBuilder {
    /// Create a new PersistPtr builder
    pub fn new() -> Self {
        Self {
            refs: HashMap::new(),
            next_id: 1, // IDs start at 1
        }
    }

    /// Allocate a new persist ID
    pub fn allocate_id(&mut self) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        id
    }

    /// Set the offset for a persist ID
    pub fn set_offset(&mut self, persist_id: u32, offset: u32) {
        self.refs.insert(persist_id, offset);
    }

    /// Get the offset for a persist ID
    pub fn get_offset(&self, persist_id: u32) -> Option<u32> {
        self.refs.get(&persist_id).copied()
    }

    /// Generate the raw PersistDirectoryAtom payload (groups of (info + offsets)).
    ///
    /// Format per MS-PPT and Apache POI:
    /// - For each contiguous run of persist IDs, write a 4-byte info word:
    ///   lower 20 bits = base persist ID, upper 12 bits = count of following offsets
    /// - Then write `count` u32 little-endian offsets, one per ID in the run
    pub fn generate_payload(&self) -> Vec<u8> {
        let mut payload = Vec::new();

        // Sort IDs to form contiguous runs
        let mut ids: Vec<u32> = self.refs.keys().copied().collect();
        ids.sort_unstable();

        if ids.is_empty() {
            return payload;
        }

        let mut i = 0;
        while i < ids.len() {
            let base = ids[i];
            let mut j = i + 1;
            while j < ids.len() && ids[j] == ids[j - 1] + 1 {
                j += 1;
            }
            let count = (j - i) as u32;
            // MS-PPT format: bits 20-31 = count, bits 0-19 = starting persist ID
            let info = ((count & 0x0FFF) << 20) | (base & 0x000F_FFFF);
            payload.extend_from_slice(&info.to_le_bytes());
            for id in &ids[i..j] {
                // Safe unwrap: id exists in refs
                let off = self.refs[id];
                payload.extend_from_slice(&off.to_le_bytes());
            }
            i = j;
        }

        payload
    }

    /// Generate the PersistPtrHolder (PersistDirectoryAtom) PPT record bytes.
    /// Uses incremental block (6002) matching POI's empty.ppt format.
    pub fn generate_record(&self) -> Vec<u8> {
        self.generate_incremental_record()
    }

    /// Generate a PersistPtrFullBlock (6001) record with the current payload.
    pub fn generate_full_record(&self) -> Vec<u8> {
        let payload = self.generate_payload();
        let mut builder = RecordBuilder::new(0x00, 0, record_type::PERSIST_PTR_HOLDER);
        builder.write_data(&payload);
        builder
            .build()
            .expect("build persist full directory record")
    }

    /// Generate a PersistPtrIncrementalBlock (6002) record with the current payload.
    pub fn generate_incremental_record(&self) -> Vec<u8> {
        let payload = self.generate_payload();
        let mut builder = RecordBuilder::new(0x00, 0, record_type::PERSIST_PTR_INCREMENTAL_BLOCK);
        builder.write_data(&payload);
        builder
            .build()
            .expect("build persist incremental directory record")
    }

    /// Get the highest allocated persist ID (seed).
    #[inline]
    pub fn persist_id_seed(&self) -> u32 {
        self.next_id.saturating_sub(1)
    }

    /// Get the number of persist references
    pub fn count(&self) -> usize {
        self.refs.len()
    }
}

impl Default for PersistPtrBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// UserEditAtom - contains persist directory offset (POI layout)
#[derive(Debug, Clone)]
pub struct UserEditAtom {
    /// lastViewedSlideID (usually first slide identifier 256 if present)
    pub last_viewed_slide_id: u32,
    /// pptVersion (opaque)
    pub ppt_version: u32,
    /// offset to previous UserEditAtom (0 for first edit)
    pub offset_last_edit: u32,
    /// offset to PersistDirectoryAtom
    pub offset_persist_dir: u32,
    /// docPersistRef (persist id of Document)
    pub doc_persist_id_ref: u32,
    /// maxPersistWritten (highest persist id allocated)
    pub max_persist_written: u32,
    /// lastView (1 = normal)
    pub last_view_type: u16,
    /// unused padding
    pub unused: u16,
}

impl UserEditAtom {
    /// Create a new minimal UserEditAtom with sane defaults.
    ///
    /// Uses constants from MS-PPT 2.4.16 specification.
    pub fn new_minimal(
        offset_persist_dir: u32,
        doc_persist_id_ref: u32,
        persist_id_seed: u32,
        _slide_count: u32,
    ) -> Self {
        Self {
            last_viewed_slide_id: DEFAULT_LAST_VIEWED_SLIDE_ID,
            ppt_version: PPT_VERSION,
            offset_last_edit: 0,
            offset_persist_dir,
            doc_persist_id_ref,
            max_persist_written: persist_id_seed,
            last_view_type: DEFAULT_LAST_VIEW_TYPE,
            unused: USER_EDIT_PADWORD,
        }
    }

    /// Generate a full PPT UserEditAtom record (type 4085) with proper header (28 bytes data).
    pub fn generate_record(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(28);
        data.extend_from_slice(&self.last_viewed_slide_id.to_le_bytes()); // +0
        data.extend_from_slice(&self.ppt_version.to_le_bytes()); // +4
        data.extend_from_slice(&self.offset_last_edit.to_le_bytes()); // +8
        data.extend_from_slice(&self.offset_persist_dir.to_le_bytes()); // +12
        data.extend_from_slice(&self.doc_persist_id_ref.to_le_bytes()); // +16
        data.extend_from_slice(&self.max_persist_written.to_le_bytes()); // +20
        data.extend_from_slice(&self.last_view_type.to_le_bytes()); // +24
        data.extend_from_slice(&self.unused.to_le_bytes()); // +26

        let mut b = RecordBuilder::new(0x00, 0, record_type::USER_EDIT_ATOM);
        b.write_data(&data);
        b.build().expect("build user edit atom")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_persist_ptr_builder() {
        let mut builder = PersistPtrBuilder::new();
        let id1 = builder.allocate_id();
        let id2 = builder.allocate_id();

        assert_eq!(id1, 1);
        assert_eq!(id2, 2);

        builder.set_offset(id1, 1000);
        builder.set_offset(id2, 2000);

        assert_eq!(builder.get_offset(id1), Some(1000));
        assert_eq!(builder.get_offset(id2), Some(2000));
    }

    #[test]
    fn test_persist_ptr_generate() {
        let mut builder = PersistPtrBuilder::new();
        let id = builder.allocate_id();
        builder.set_offset(id, 5000);

        let payload = builder.generate_payload();
        assert_eq!(payload.len(), 8); // 4 bytes info + 4 bytes offset
    }

    #[test]
    fn test_user_edit_atom() {
        let atom = UserEditAtom::new_minimal(1000, 1, 2, 0);
        let bytes = atom.generate_record();
        // 8 bytes header + 28 bytes data
        assert_eq!(bytes.len(), 36);
    }
}
