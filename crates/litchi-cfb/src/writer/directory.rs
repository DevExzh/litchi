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
use super::super::file::OleError;
use crate::directory_name::{DirectoryNameData, directory_name_data as parse_directory_name};
use fixedbitset::FixedBitSet;
use smallvec::SmallVec;
use std::cmp::Ordering;
use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
enum NodeColor {
    Red = 0,
    Black = 1,
}

fn directory_name_data(name: &str) -> Result<(SmallVec<[u16; 32]>, SmallVec<[u16; 32]>), OleError> {
    let DirectoryNameData { utf16, comparison } =
        parse_directory_name(name).map_err(|error| OleError::InvalidData(error.to_string()))?;
    Ok((utf16, comparison))
}

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
    name_utf16: SmallVec<[u16; 32]>,
    comparison_name: SmallVec<[u16; 32]>,
    node_color: NodeColor,
}

#[allow(dead_code)] // These methods are part of the public API for future use
impl DirectoryEntryBuilder {
    /// Create a new root entry
    pub fn root(start_sector: u32, size: u64) -> Self {
        let name = "Root Entry".to_string();
        let (name_utf16, comparison_name) =
            directory_name_data(&name).expect("the fixed root entry name is valid");
        Self {
            name,
            entry_type: STGTY_ROOT,
            start_sector,
            size,
            sid_left: NOSTREAM,
            sid_right: NOSTREAM,
            sid_child: NOSTREAM,
            clsid: None,
            name_utf16,
            comparison_name,
            node_color: NodeColor::Black,
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
    /// ```ignore
    /// use litchi_cfb::writer::DirectoryEntryBuilder;
    ///
    /// // Word 97-2003 Document CLSID: {00020906-0000-0000-C000-000000000046}
    /// let word_clsid = [0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    ///                   0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];
    /// let mut root = DirectoryEntryBuilder::root();
    /// root.set_clsid(word_clsid);
    /// ```
    pub fn set_clsid(&mut self, clsid: [u8; 16]) {
        self.clsid = Some(clsid);
    }

    /// Create a new stream entry
    pub fn stream(name: String, start_sector: u32, size: u64) -> Result<Self, OleError> {
        let (name_utf16, comparison_name) = directory_name_data(&name)?;
        Ok(Self {
            name,
            entry_type: STGTY_STREAM,
            start_sector,
            size,
            sid_left: NOSTREAM,
            sid_right: NOSTREAM,
            sid_child: NOSTREAM,
            clsid: None,
            name_utf16,
            comparison_name,
            node_color: NodeColor::Black,
        })
    }

    /// Create a new storage entry
    pub fn storage(name: String) -> Result<Self, OleError> {
        let (name_utf16, comparison_name) = directory_name_data(&name)?;
        Ok(Self {
            name,
            entry_type: STGTY_STORAGE,
            start_sector: 0,
            size: 0,
            sid_left: NOSTREAM,
            sid_right: NOSTREAM,
            sid_child: NOSTREAM,
            clsid: None,
            name_utf16,
            comparison_name,
            node_color: NodeColor::Black,
        })
    }

    /// Write this entry to bytes (128 bytes per OLE2 spec)
    pub fn to_bytes(&self) -> Result<[u8; DIRENTRY_SIZE], OleError> {
        let (name_utf16, comparison_name) = directory_name_data(&self.name)?;
        if name_utf16 != self.name_utf16 || comparison_name != self.comparison_name {
            return Err(OleError::InvalidData(
                "CFB directory entry name was mutated after validation".to_string(),
            ));
        }

        let mut data = [0u8; DIRENTRY_SIZE];
        for (i, &ch) in self.name_utf16.iter().enumerate() {
            let bytes = ch.to_le_bytes();
            data[i * 2] = bytes[0];
            data[i * 2 + 1] = bytes[1];
        }

        // Name length in bytes (including null terminator)
        let name_len_bytes = u16::try_from((self.name_utf16.len() + 1) * 2)
            .expect("validated CFB directory name length fits in u16");
        data[64..66].copy_from_slice(&name_len_bytes.to_le_bytes());

        // Entry type
        data[66] = self.entry_type;

        data[67] = self.node_color as u8;

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

        Ok(data)
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
    pub fn add_storage_path(&mut self, path: &[String]) -> Result<u32, OleError> {
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
            let entry = DirectoryEntryBuilder::storage(component.clone())?;
            self.ensure_unique_child(parent_sid, &entry)?;
            let sid = u32::try_from(self.entries.len()).map_err(|_| {
                OleError::InvalidData("CFB directory contains too many entries".to_string())
            })?;
            self.entries.push(entry);
            self.path_to_sid.insert(current_path.clone(), sid);

            // register as child of previous parent
            self.children.entry(parent_sid).or_default().push(sid);
            // initialize its children vec
            self.children.entry(sid).or_default();

            parent_sid = sid;
        }

        Ok(parent_sid)
    }

    /// Add a stream at the given full path (parent storages will be created automatically)
    pub fn add_stream_path(
        &mut self,
        full_path: &[String],
        start_sector: u32,
        size: u64,
    ) -> Result<u32, OleError> {
        if full_path.is_empty() {
            return Err(OleError::InvalidData(
                "CFB stream path must not be empty".to_string(),
            ));
        }
        let parent_sid = if full_path.len() > 1 {
            self.add_storage_path(&full_path[..full_path.len() - 1])?
        } else {
            0
        };

        let name = full_path.last().unwrap().clone();
        let entry = DirectoryEntryBuilder::stream(name, start_sector, size)?;
        self.ensure_unique_child(parent_sid, &entry)?;
        let sid = u32::try_from(self.entries.len()).map_err(|_| {
            OleError::InvalidData("CFB directory contains too many entries".to_string())
        })?;
        self.entries.push(entry);
        self.children.entry(parent_sid).or_default().push(sid);
        Ok(sid)
    }

    /// Add a stream to the root directory (compat wrapper)
    pub fn add_stream(
        &mut self,
        name: String,
        start_sector: u32,
        size: u64,
    ) -> Result<u32, OleError> {
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
    pub fn add_storage(&mut self, name: String) -> Result<u32, OleError> {
        let entry = DirectoryEntryBuilder::storage(name)?;
        self.ensure_unique_child(0, &entry)?;
        let sid = u32::try_from(self.entries.len()).map_err(|_| {
            OleError::InvalidData("CFB directory contains too many entries".to_string())
        })?;
        self.entries.push(entry);
        self.children.entry(0).or_default().push(sid);
        self.children.entry(sid).or_default();
        Ok(sid)
    }

    /// Generate directory sectors as bytes
    ///
    /// # Returns
    ///
    /// * `Vec<u8>` - Concatenated directory entries (128 bytes each)
    pub fn generate_directory_stream(&mut self) -> Result<Vec<u8>, OleError> {
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
                Self::link_children(parent_sid, &children, &mut self.entries)?;
            } else {
                // no children: ensure sid_child remains NOSTREAM
                self.entries[parent_sid as usize].sid_child = NOSTREAM;
            }
        }

        // Serialize entries in SID order
        let mut data = Vec::with_capacity(self.entries.len() * 128);
        for entry in &self.entries {
            data.extend_from_slice(&entry.to_bytes()?);
        }
        Ok(data)
    }

    fn compare_entries(left: &DirectoryEntryBuilder, right: &DirectoryEntryBuilder) -> Ordering {
        left.name_utf16
            .len()
            .cmp(&right.name_utf16.len())
            .then_with(|| left.comparison_name.cmp(&right.comparison_name))
    }

    fn ensure_unique_child(
        &self,
        parent_sid: u32,
        candidate: &DirectoryEntryBuilder,
    ) -> Result<(), OleError> {
        if self
            .children
            .get(&parent_sid)
            .into_iter()
            .flatten()
            .any(|sid| {
                Self::compare_entries(&self.entries[*sid as usize], candidate) == Ordering::Equal
            })
        {
            return Err(OleError::InvalidData(format!(
                "duplicate CFB sibling name {:?}",
                candidate.name
            )));
        }
        Ok(())
    }

    fn color(entries: &[DirectoryEntryBuilder], sid: u32) -> NodeColor {
        if sid == NOSTREAM {
            NodeColor::Black
        } else {
            entries[sid as usize].node_color
        }
    }

    fn rotate_left(
        root: &mut u32,
        pivot: u32,
        parents: &mut [u32],
        entries: &mut [DirectoryEntryBuilder],
    ) {
        let child = entries[pivot as usize].sid_right;
        let child_left = entries[child as usize].sid_left;
        entries[pivot as usize].sid_right = child_left;
        if child_left != NOSTREAM {
            parents[child_left as usize] = pivot;
        }
        let pivot_parent = parents[pivot as usize];
        parents[child as usize] = pivot_parent;
        if pivot_parent == NOSTREAM {
            *root = child;
        } else if pivot == entries[pivot_parent as usize].sid_left {
            entries[pivot_parent as usize].sid_left = child;
        } else {
            entries[pivot_parent as usize].sid_right = child;
        }
        entries[child as usize].sid_left = pivot;
        parents[pivot as usize] = child;
    }

    fn rotate_right(
        root: &mut u32,
        pivot: u32,
        parents: &mut [u32],
        entries: &mut [DirectoryEntryBuilder],
    ) {
        let child = entries[pivot as usize].sid_left;
        let child_right = entries[child as usize].sid_right;
        entries[pivot as usize].sid_left = child_right;
        if child_right != NOSTREAM {
            parents[child_right as usize] = pivot;
        }
        let pivot_parent = parents[pivot as usize];
        parents[child as usize] = pivot_parent;
        if pivot_parent == NOSTREAM {
            *root = child;
        } else if pivot == entries[pivot_parent as usize].sid_right {
            entries[pivot_parent as usize].sid_right = child;
        } else {
            entries[pivot_parent as usize].sid_left = child;
        }
        entries[child as usize].sid_right = pivot;
        parents[pivot as usize] = child;
    }

    fn fix_after_insert(
        root: &mut u32,
        mut sid: u32,
        parents: &mut [u32],
        entries: &mut [DirectoryEntryBuilder],
    ) {
        while Self::color(entries, parents[sid as usize]) == NodeColor::Red {
            let parent = parents[sid as usize];
            let grandparent = parents[parent as usize];
            if parent == entries[grandparent as usize].sid_left {
                let uncle = entries[grandparent as usize].sid_right;
                if Self::color(entries, uncle) == NodeColor::Red {
                    entries[parent as usize].node_color = NodeColor::Black;
                    entries[uncle as usize].node_color = NodeColor::Black;
                    entries[grandparent as usize].node_color = NodeColor::Red;
                    sid = grandparent;
                } else {
                    if sid == entries[parent as usize].sid_right {
                        sid = parent;
                        Self::rotate_left(root, sid, parents, entries);
                    }
                    let parent = parents[sid as usize];
                    let grandparent = parents[parent as usize];
                    entries[parent as usize].node_color = NodeColor::Black;
                    entries[grandparent as usize].node_color = NodeColor::Red;
                    Self::rotate_right(root, grandparent, parents, entries);
                }
            } else {
                let uncle = entries[grandparent as usize].sid_left;
                if Self::color(entries, uncle) == NodeColor::Red {
                    entries[parent as usize].node_color = NodeColor::Black;
                    entries[uncle as usize].node_color = NodeColor::Black;
                    entries[grandparent as usize].node_color = NodeColor::Red;
                    sid = grandparent;
                } else {
                    if sid == entries[parent as usize].sid_left {
                        sid = parent;
                        Self::rotate_right(root, sid, parents, entries);
                    }
                    let parent = parents[sid as usize];
                    let grandparent = parents[parent as usize];
                    entries[parent as usize].node_color = NodeColor::Black;
                    entries[grandparent as usize].node_color = NodeColor::Red;
                    Self::rotate_left(root, grandparent, parents, entries);
                }
            }
        }
        entries[*root as usize].node_color = NodeColor::Black;
    }

    /// Link a parent's children as a conforming red-black sibling tree.
    fn link_children(
        parent_sid: u32,
        child_sids: &[u32],
        entries: &mut [DirectoryEntryBuilder],
    ) -> Result<(), OleError> {
        if child_sids.is_empty() {
            entries[parent_sid as usize].sid_child = NOSTREAM;
            return Ok(());
        }

        let mut parents = vec![NOSTREAM; entries.len()];
        let mut root = NOSTREAM;
        for &sid in child_sids {
            entries[sid as usize].sid_left = NOSTREAM;
            entries[sid as usize].sid_right = NOSTREAM;
            entries[sid as usize].node_color = NodeColor::Red;

            let mut parent = NOSTREAM;
            let mut cursor = root;
            let mut ordering = Ordering::Equal;
            while cursor != NOSTREAM {
                parent = cursor;
                ordering = Self::compare_entries(&entries[sid as usize], &entries[cursor as usize]);
                cursor = match ordering {
                    Ordering::Less => entries[cursor as usize].sid_left,
                    Ordering::Greater => entries[cursor as usize].sid_right,
                    Ordering::Equal => {
                        return Err(OleError::InvalidData(format!(
                            "duplicate CFB sibling name {:?}",
                            entries[sid as usize].name
                        )));
                    },
                };
            }
            parents[sid as usize] = parent;
            if parent == NOSTREAM {
                root = sid;
            } else if ordering == Ordering::Less {
                entries[parent as usize].sid_left = sid;
            } else {
                entries[parent as usize].sid_right = sid;
            }
            Self::fix_after_insert(&mut root, sid, &mut parents, entries);
        }

        entries[parent_sid as usize].sid_child = root;
        Self::validate_child_tree(root, child_sids, entries)
    }

    fn validate_child_tree(
        root: u32,
        child_sids: &[u32],
        entries: &[DirectoryEntryBuilder],
    ) -> Result<(), OleError> {
        if Self::color(entries, root) != NodeColor::Black {
            return Err(OleError::InvalidData(
                "CFB sibling-tree root must be black".to_string(),
            ));
        }

        let mut members = FixedBitSet::with_capacity(entries.len());
        for &sid in child_sids {
            if sid as usize >= entries.len() || members.contains(sid as usize) {
                return Err(OleError::InvalidData(
                    "CFB sibling list contains an invalid or duplicate SID".to_string(),
                ));
            }
            members.insert(sid as usize);
        }

        let mut visited = FixedBitSet::with_capacity(entries.len());
        let mut expected_black_depth = None;
        let mut stack = vec![(root, None, None, 0usize)];
        while let Some((sid, lower, upper, black_depth)) = stack.pop() {
            if sid == NOSTREAM {
                let leaf_depth = black_depth + 1;
                if expected_black_depth
                    .replace(leaf_depth)
                    .is_some_and(|depth| depth != leaf_depth)
                {
                    return Err(OleError::InvalidData(
                        "CFB sibling tree has inconsistent black height".to_string(),
                    ));
                }
                continue;
            }
            let index = sid as usize;
            if index >= entries.len() || !members.contains(index) || visited.contains(index) {
                return Err(OleError::InvalidData(
                    "CFB sibling tree contains an invalid, foreign, or repeated SID".to_string(),
                ));
            }
            visited.insert(index);
            let entry = &entries[index];
            if lower.is_some_and(|bound: u32| {
                Self::compare_entries(&entries[bound as usize], entry) != Ordering::Less
            }) || upper.is_some_and(|bound: u32| {
                Self::compare_entries(entry, &entries[bound as usize]) != Ordering::Less
            }) {
                return Err(OleError::InvalidData(
                    "CFB sibling tree violates directory-name ordering".to_string(),
                ));
            }
            if entry.node_color == NodeColor::Red
                && (Self::color(entries, entry.sid_left) == NodeColor::Red
                    || Self::color(entries, entry.sid_right) == NodeColor::Red)
            {
                return Err(OleError::InvalidData(
                    "CFB sibling tree contains adjacent red nodes".to_string(),
                ));
            }
            let black_depth = black_depth + usize::from(entry.node_color == NodeColor::Black);
            stack.push((entry.sid_right, Some(sid), upper, black_depth));
            stack.push((entry.sid_left, lower, Some(sid), black_depth));
        }
        if visited.count_ones(..) != child_sids.len() {
            return Err(OleError::InvalidData(
                "CFB sibling tree does not reach every child".to_string(),
            ));
        }
        Ok(())
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

        let bytes = root.to_bytes().unwrap();
        assert_eq!(bytes.len(), 128);
    }

    #[test]
    fn test_stream_entry() {
        let stream = DirectoryEntryBuilder::stream("Test".to_string(), 10, 512).unwrap();
        assert_eq!(stream.entry_type, STGTY_STREAM);
        assert_eq!(stream.start_sector, 10);
        assert_eq!(stream.size, 512);
    }

    #[test]
    fn test_directory_builder() {
        let mut dir = DirectoryBuilder::new(0, 0);
        let sid = dir.add_stream("Stream1".to_string(), 5, 1024).unwrap();

        assert_eq!(sid, 1);
        assert_eq!(dir.entry_count(), 2); // Root + 1 stream

        let data = dir.generate_directory_stream().unwrap();
        assert_eq!(data.len(), 2 * 128); // 2 entries * 128 bytes each
    }

    #[test]
    fn rejects_invalid_and_ambiguous_sibling_names() {
        assert!(DirectoryEntryBuilder::stream(String::new(), 0, 0).is_err());
        assert!(DirectoryEntryBuilder::stream("bad/name".to_string(), 0, 0).is_err());
        assert!(DirectoryEntryBuilder::stream("a".repeat(32), 0, 0).is_err());
        assert!(DirectoryEntryBuilder::stream("😀".repeat(16), 0, 0).is_err());

        let mut directory = DirectoryBuilder::new(0, 0);
        directory.add_stream("Report".to_string(), 0, 0).unwrap();
        assert!(directory.add_stream("report".to_string(), 0, 0).is_err());
    }

    #[test]
    fn serializes_maximum_utf16_name_without_truncation() {
        let name = format!("{}x", "😀".repeat(15));
        let entry = DirectoryEntryBuilder::stream(name, 0, 0).unwrap();
        let bytes = entry.to_bytes().unwrap();
        assert_eq!(u16::from_le_bytes([bytes[64], bytes[65]]), 64);
        assert_eq!(&bytes[60..64], &[b'x', 0, 0, 0]);
    }

    #[test]
    fn builds_valid_red_black_trees_for_adversarial_sibling_counts() {
        for count in 1..=64 {
            let mut directory = DirectoryBuilder::new(0, 0);
            for index in 0..count {
                directory
                    .add_stream(format!("stream-{index:03}"), index, 8)
                    .unwrap();
            }
            let bytes = directory.generate_directory_stream().unwrap();
            assert_eq!(bytes.len(), (count as usize + 1) * DIRENTRY_SIZE);
        }
    }
}
