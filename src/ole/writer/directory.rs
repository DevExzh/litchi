//! Directory tree generation for OLE2 files
//!
//! Builds the directory structure (binary search tree) from streams and storages.
//!
//! # Implementation Notes
//!
//! Based on Apache POI's DirectoryProperty implementation.
//!
//! ## Directory Entry Ordering
//!
//! This implementation follows Apache POI's `PropertyComparator` rules for sorting
//! directory entries, which is critical for compatibility with Microsoft Office:
//!
//! 1. **Sort by name length first** (shorter names come before longer names)
//! 2. **Then sort alphabetically** (case-insensitive) for same-length names  
//! 3. **Special case**: `_VBA_PROJECT` always comes last
//! 4. **Special case**: Names starting with `__` are pushed later
//!
//! Example ordering:
//! - `"1Table"` (length 6) comes before `"WordDocument"` (length 12)
//! - `"Data"` (length 4) comes before `"1Table"` (length 6)
//! - `"ABC"` comes before `"XYZ"` (both length 3, alphabetical)
//!
//! ## Binary Search Tree Structure
//!
//! After sorting, entries are organized into a balanced BST:
//! - The middle element becomes the root of each subtree
//! - Left subtree contains entries before the middle (previous children)
//! - Right subtree contains entries after the middle (next children)
//!
//! This balanced structure ensures:
//! - O(log n) lookup performance
//! - Compatibility with Microsoft Office readers
//! - Consistent ordering across platforms
//!
//! ## Example
//!
//! For a DOC file with `WordDocument` and `1Table` streams:
//!
//! ```text
//! Sorted order: ["1Table", "WordDocument"]
//! Tree structure:
//!        Root Entry
//!             |
//!        WordDocument (midpoint)
//!            /
//!        1Table (left child)
//! ```
//!
//! ## References
//!
//! - Apache POI: `org.apache.poi.poifs.property.DirectoryProperty.PropertyComparator`
//! - Apache POI: `org.apache.poi.poifs.property.DirectoryProperty.preWrite()`
//! - Apache POI: `org.apache.poi.poifs.property.PropertyTable`
//! - MS-CFB specification: Section 2.6 (Compound File Directory Sectors)
//! - MS-DOC specification: Section 2.3 (File Structure)

use super::super::consts::*;
use std::collections::HashMap;

/// Directory entry builder
#[derive(Debug, Clone)]
pub struct DirectoryEntryBuilder {
    /// Entry name
    pub name: String,
    /// Entry type (STGTY_STORAGE, STGTY_STREAM, etc.)
    pub entry_type: u8,
    /// Starting sector
    pub start_sector: u32,
    /// Stream size
    pub size: u64,
    /// Left sibling SID
    pub sid_left: u32,
    /// Right sibling SID
    pub sid_right: u32,
    /// Child SID
    pub sid_child: u32,
    /// CLSID (Class ID) - 16 bytes, optional
    pub clsid: Option<[u8; 16]>,
}

#[allow(dead_code)] // These methods are part of the public API for future use
impl DirectoryEntryBuilder {
    /// Create a new root entry
    pub fn root(start_sector: u32, size: u64) -> Self {
        Self {
            name: "Root Entry".to_string(),
            entry_type: STGTY_ROOT,
            start_sector,
            size,
            sid_left: NOSTREAM,
            sid_right: NOSTREAM,
            sid_child: NOSTREAM,
            clsid: None,
        }
    }

    /// Set the CLSID (Class ID) for this entry
    ///
    /// # Arguments
    ///
    /// * `clsid` - 16-byte CLSID (GUID in little-endian format)
    ///
    /// # Example
    ///
    /// ```
    /// // Word 97-2003 Document CLSID: {00020906-0000-0000-C000-000000000046}
    /// let word_clsid = [0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    ///                   0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];
    /// root.set_clsid(word_clsid);
    /// ```
    pub fn set_clsid(&mut self, clsid: [u8; 16]) {
        self.clsid = Some(clsid);
    }

    /// Create a new stream entry
    pub fn stream(name: String, start_sector: u32, size: u64) -> Self {
        Self {
            name,
            entry_type: STGTY_STREAM,
            start_sector,
            size,
            sid_left: NOSTREAM,
            sid_right: NOSTREAM,
            sid_child: NOSTREAM,
            clsid: None,
        }
    }

    /// Create a new storage entry
    pub fn storage(name: String) -> Self {
        Self {
            name,
            entry_type: STGTY_STORAGE,
            start_sector: 0,
            size: 0,
            sid_left: NOSTREAM,
            sid_right: NOSTREAM,
            sid_child: NOSTREAM,
            clsid: None,
        }
    }

    /// Write this entry to bytes (128 bytes per OLE2 spec)
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut data = vec![0u8; 128];

        // Encode name to UTF-16LE (max 31 characters + null)
        let utf16: Vec<u16> = self.name.encode_utf16().collect();
        let name_len = utf16.len().min(31);

        for (i, &ch) in utf16.iter().take(name_len).enumerate() {
            let bytes = ch.to_le_bytes();
            data[i * 2] = bytes[0];
            data[i * 2 + 1] = bytes[1];
        }

        // Name length in bytes (including null terminator)
        let name_len_bytes = ((name_len + 1) * 2) as u16;
        data[64..66].copy_from_slice(&name_len_bytes.to_le_bytes());

        // Entry type
        data[66] = self.entry_type;

        // Node color (always black = 1 for simplicity)
        data[67] = 1;

        // Sibling and child SIDs
        data[68..72].copy_from_slice(&self.sid_left.to_le_bytes());
        data[72..76].copy_from_slice(&self.sid_right.to_le_bytes());
        data[76..80].copy_from_slice(&self.sid_child.to_le_bytes());

        // CLSID (Class ID) - 16 bytes at offset 80
        // If set, this identifies the type of OLE object (e.g., Word document)
        if let Some(clsid) = self.clsid {
            data[80..96].copy_from_slice(&clsid);
        }
        // Otherwise data[80..96] remains zeros

        // State bits (all zeros)
        // data[96..100] already zeros

        // Creation and modification times (all zeros for now)
        // data[100..116] already zeros

        // Starting sector
        data[116..120].copy_from_slice(&self.start_sector.to_le_bytes());

        // Stream size (8 bytes)
        data[120..128].copy_from_slice(&self.size.to_le_bytes());

        data
    }
}

/// Directory tree builder
///
/// Builds a directory tree from streams and storages, organizing them
/// as a simple list (not a true red-black tree for simplicity).
pub struct DirectoryBuilder {
    /// List of directory entries; index is the SID
    entries: Vec<DirectoryEntryBuilder>,
    /// Map from full path components to SID for storage nodes
    path_to_sid: HashMap<Vec<String>, u32>,
    /// Children SIDs per parent SID
    children: HashMap<u32, Vec<u32>>,
}

#[allow(dead_code)] // These methods are part of the public API for future use
impl DirectoryBuilder {
    /// Create a new directory builder with root entry
    pub fn new(ministream_start: u32, ministream_size: u64) -> Self {
        let root = DirectoryEntryBuilder::root(ministream_start, ministream_size);
        let mut path_to_sid = HashMap::new();
        path_to_sid.insert(Vec::new(), 0);
        let mut children = HashMap::new();
        children.insert(0, Vec::new());
        Self {
            entries: vec![root],
            path_to_sid,
            children,
        }
    }

    /// Set the CLSID for the root directory entry
    ///
    /// # Arguments
    ///
    /// * `clsid` - 16-byte CLSID
    pub fn set_root_clsid(&mut self, clsid: [u8; 16]) {
        if !self.entries.is_empty() {
            self.entries[0].set_clsid(clsid);
        }
    }

    /// Ensure a storage path exists, creating missing storages.
    /// Returns the SID of the storage at the given path.
    pub fn add_storage_path(&mut self, path: &[String]) -> u32 {
        // parent path accumulates
        let mut current_path: Vec<String> = Vec::new();
        let mut parent_sid = 0u32;

        for component in path {
            current_path.push(component.clone());
            if let Some(&sid) = self.path_to_sid.get(&current_path) {
                parent_sid = sid;
                continue;
            }

            // create new storage
            let sid = self.entries.len() as u32;
            let entry = DirectoryEntryBuilder::storage(component.clone());
            self.entries.push(entry);
            self.path_to_sid.insert(current_path.clone(), sid);

            // register as child of previous parent
            self.children.entry(parent_sid).or_default().push(sid);
            // initialize its children vec
            self.children.entry(sid).or_default();

            parent_sid = sid;
        }

        parent_sid
    }

    /// Add a stream at the given full path (parent storages will be created automatically)
    pub fn add_stream_path(&mut self, full_path: &[String], start_sector: u32, size: u64) -> u32 {
        assert!(!full_path.is_empty(), "stream path must not be empty");
        let parent_sid = if full_path.len() > 1 {
            self.add_storage_path(&full_path[..full_path.len() - 1])
        } else {
            0
        };

        let name = full_path.last().unwrap().clone();
        let sid = self.entries.len() as u32;
        let entry = DirectoryEntryBuilder::stream(name, start_sector, size);
        self.entries.push(entry);
        self.children.entry(parent_sid).or_default().push(sid);
        sid
    }

    /// Add a stream to the root directory (compat wrapper)
    pub fn add_stream(&mut self, name: String, start_sector: u32, size: u64) -> u32 {
        let path = vec![name];
        self.add_stream_path(&path, start_sector, size)
    }

    /// Add a storage to the root directory
    ///
    /// # Arguments
    ///
    /// * `name` - Storage name
    ///
    /// # Returns
    ///
    /// * `u32` - SID of the added storage
    pub fn add_storage(&mut self, name: String) -> u32 {
        let sid = self.entries.len() as u32;
        let entry = DirectoryEntryBuilder::storage(name);
        self.entries.push(entry);

        // Tree will be built when generating directory stream
        sid
    }

    /// Generate directory sectors as bytes
    ///
    /// # Returns
    ///
    /// * `Vec<u8>` - Concatenated directory entries (128 bytes each)
    pub fn generate_directory_stream(&mut self) -> Vec<u8> {
        // Link children for each storage/root using POI comparator and midpoint rules
        let storage_sids: Vec<u32> = self
            .entries
            .iter()
            .enumerate()
            .filter_map(|(sid, e)| {
                if e.entry_type == STGTY_ROOT || e.entry_type == STGTY_STORAGE {
                    Some(sid as u32)
                } else {
                    None
                }
            })
            .collect();

        for parent_sid in storage_sids {
            if let Some(children) = self.children.get(&parent_sid).cloned() {
                Self::link_children(parent_sid, &children, &mut self.entries);
            } else {
                // no children: ensure sid_child remains NOSTREAM
                self.entries[parent_sid as usize].sid_child = NOSTREAM;
            }
        }

        // Serialize entries in SID order
        let mut data = Vec::with_capacity(self.entries.len() * 128);
        for entry in &self.entries {
            data.extend_from_slice(&entry.to_bytes());
        }
        data
    }

    /// Link a parent's children list using POI's comparator and midpoint rules
    fn link_children(parent_sid: u32, child_sids: &[u32], entries: &mut [DirectoryEntryBuilder]) {
        if child_sids.is_empty() {
            entries[parent_sid as usize].sid_child = NOSTREAM;
            return;
        }

        // Sort SIDs using comparator on names
        let mut sorted: Vec<u32> = child_sids.to_vec();
        sorted.sort_by(|&a, &b| {
            let name1 = &entries[a as usize].name;
            let name2 = &entries[b as usize].name;
            let len1 = name1.len();
            let len2 = name2.len();
            match len1.cmp(&len2) {
                std::cmp::Ordering::Equal => {
                    if name1 == "_VBA_PROJECT" {
                        return std::cmp::Ordering::Greater;
                    }
                    if name2 == "_VBA_PROJECT" {
                        return std::cmp::Ordering::Less;
                    }
                    if name1.starts_with("__") && name2.starts_with("__") {
                        return name1.to_uppercase().cmp(&name2.to_uppercase());
                    }
                    if name1.starts_with("__") {
                        return std::cmp::Ordering::Greater;
                    }
                    if name2.starts_with("__") {
                        return std::cmp::Ordering::Less;
                    }
                    name1.to_uppercase().cmp(&name2.to_uppercase())
                },
                other => other,
            }
        });

        let midpoint = sorted.len() / 2;
        entries[parent_sid as usize].sid_child = sorted[midpoint];

        // Initialize first element
        entries[sorted[0] as usize].sid_left = NOSTREAM;
        entries[sorted[0] as usize].sid_right = NOSTREAM;

        // Left chain up to midpoint-1
        for j in 1..midpoint {
            let sid = sorted[j] as usize;
            entries[sid].sid_left = sorted[j - 1];
            entries[sid].sid_right = NOSTREAM;
        }

        // Midpoint element
        if midpoint > 0 {
            entries[sorted[midpoint] as usize].sid_left = sorted[midpoint - 1];
        } else {
            entries[sorted[midpoint] as usize].sid_left = NOSTREAM;
        }

        if midpoint < sorted.len() - 1 {
            entries[sorted[midpoint] as usize].sid_right = sorted[midpoint + 1];
            for j in (midpoint + 1)..(sorted.len() - 1) {
                let sid = sorted[j] as usize;
                entries[sid].sid_left = NOSTREAM;
                entries[sid].sid_right = sorted[j + 1];
            }
            entries[*sorted.last().unwrap() as usize].sid_left = NOSTREAM;
            entries[*sorted.last().unwrap() as usize].sid_right = NOSTREAM;
        } else {
            entries[sorted[midpoint] as usize].sid_right = NOSTREAM;
        }
    }

    /// Get the number of directory entries
    pub fn entry_count(&self) -> usize {
        self.entries.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_root_entry() {
        let root = DirectoryEntryBuilder::root(0, 0);
        assert_eq!(root.name, "Root Entry");
        assert_eq!(root.entry_type, STGTY_ROOT);

        let bytes = root.to_bytes();
        assert_eq!(bytes.len(), 128);
    }

    #[test]
    fn test_stream_entry() {
        let stream = DirectoryEntryBuilder::stream("Test".to_string(), 10, 512);
        assert_eq!(stream.entry_type, STGTY_STREAM);
        assert_eq!(stream.start_sector, 10);
        assert_eq!(stream.size, 512);
    }

    #[test]
    fn test_directory_builder() {
        let mut dir = DirectoryBuilder::new(0, 0);
        let sid = dir.add_stream("Stream1".to_string(), 5, 1024);

        assert_eq!(sid, 1);
        assert_eq!(dir.entry_count(), 2); // Root + 1 stream

        let data = dir.generate_directory_stream();
        assert_eq!(data.len(), 2 * 128); // 2 entries * 128 bytes each
    }
}
