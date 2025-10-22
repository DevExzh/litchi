// BlipStoreEntry (BSE) record parsing
//
// BSE records are part of the Blip Store Container (BStoreContainer) in Escher
// drawing records. They serve as an index to BLIP data, providing metadata
// about embedded images.
//
// References:
// - [MS-ODRAW] 2.2.32: OfficeArtBStoreContainerFileBlock
// - [MS-ODRAW] 2.2.33: OfficeArtBSE

use crate::common::binary::read_u32_le;
use crate::common::error::{Error, Result};
use crate::images::BlipType;
use std::borrow::Cow;

/// BlipStoreEntry - metadata and index for a BLIP record
///
/// BSE records are found in BStoreContainer (0xF001) and contain:
/// - Type information about the BLIP
/// - UIDs for identifying the BLIP
/// - Size information
/// - Reference counts
/// - Offset to the actual BLIP data (for delay-loaded BLIPs)
#[derive(Debug, Clone)]
pub struct BlipStoreEntry<'data> {
    /// BLIP type indicator (matches Windows BITMAPINFO values)
    pub blip_type_win32: u8,
    /// BLIP type
    pub blip_type: BlipType,
    /// Primary UID (16 bytes)
    pub uid: [u8; 16],
    /// Tag value (usually 0xFF)
    pub tag: u16,
    /// Size of BLIP data in bytes
    pub size: u32,
    /// Reference count
    pub ref_count: u32,
    /// Offset to BLIP data in stream
    pub offset: u32,
    /// Usage type (0=default, 1=texture, 2=pattern)
    pub usage: u8,
    /// Name length in bytes
    pub name_len: u8,
    /// Unused bytes count
    pub unused2: u8,
    /// Unused bytes count
    pub unused3: u8,
    /// Optional name (zero-copy borrow)
    pub name: Option<Cow<'data, str>>,
}

impl<'data> BlipStoreEntry<'data> {
    /// Convert this entry to owned data with 'static lifetime
    pub fn into_owned(self) -> BlipStoreEntry<'static> {
        BlipStoreEntry {
            blip_type_win32: self.blip_type_win32,
            blip_type: self.blip_type,
            uid: self.uid,
            tag: self.tag,
            size: self.size,
            ref_count: self.ref_count,
            offset: self.offset,
            usage: self.usage,
            name_len: self.name_len,
            unused2: self.unused2,
            unused3: self.unused3,
            name: self.name.map(|n| Cow::Owned(n.into_owned())),
        }
    }

    /// Parse a BSE record from binary data
    ///
    /// # Arguments
    /// * `data` - Record data (without the Escher record header)
    ///
    /// # Returns
    /// Parsed BSE or error
    ///
    /// # Format
    /// ```text
    /// Offset | Size | Field
    /// -------|------|------
    /// 0      | 1    | btWin32 (BLIP type for Windows)
    /// 1      | 1    | btMacOS (BLIP type for Mac)
    /// 2      | 16   | rgbUid (Primary UID)
    /// 18     | 2    | tag
    /// 20     | 4    | size (BLIP data size)
    /// 24     | 4    | cRef (reference count)
    /// 28     | 4    | foDelay (offset to BLIP)
    /// 32     | 1    | usage
    /// 33     | 1    | cbName (name length)
    /// 34     | 1    | unused2
    /// 35     | 1    | unused3
    /// 36     | N    | nameData (optional, present if cbName > 0)
    /// ```
    pub fn parse(data: &'data [u8]) -> Result<Self> {
        if data.len() < 36 {
            return Err(Error::ParseError("Insufficient data for BSE record".into()));
        }

        let blip_type_win32 = data[0];
        let _blip_type_macos = data[1]; // Usually same as Windows type

        // Parse UID
        let mut uid = [0u8; 16];
        uid.copy_from_slice(&data[2..18]);

        let tag = u16::from_le_bytes([data[18], data[19]]);
        let size = read_u32_le(data, 20)
            .map_err(|_| Error::ParseError("Invalid size field in BSE".into()))?;
        let ref_count = read_u32_le(data, 24)
            .map_err(|_| Error::ParseError("Invalid ref_count field in BSE".into()))?;
        let offset = read_u32_le(data, 28)
            .map_err(|_| Error::ParseError("Invalid offset field in BSE".into()))?;
        let usage = data[32];
        let name_len = data[33];
        let unused2 = data[34];
        let unused3 = data[35];

        // Parse optional name
        let name = if name_len > 0 {
            let name_start = 36;
            let name_end = name_start + name_len as usize;
            if name_end > data.len() {
                return Err(Error::ParseError("BSE name extends beyond data".into()));
            }

            // Name is stored as UTF-16 LE (2 bytes per character)
            let name_bytes = &data[name_start..name_end];
            if !name_len.is_multiple_of(2) {
                return Err(Error::ParseError(
                    "Invalid BSE name length (not UTF-16)".into(),
                ));
            }

            // Convert UTF-16 LE to String
            let utf16_chars: Vec<u16> = name_bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();

            match String::from_utf16(&utf16_chars) {
                Ok(s) => Some(Cow::Owned(s)),
                Err(_) => {
                    // Fallback: treat as Latin1
                    Some(Cow::Owned(
                        name_bytes
                            .iter()
                            .filter(|&&b| b != 0)
                            .map(|&b| b as char)
                            .collect(),
                    ))
                },
            }
        } else {
            None
        };

        // Map Windows BLIP type to BlipType
        let blip_type = Self::map_windows_blip_type(blip_type_win32)?;

        Ok(Self {
            blip_type_win32,
            blip_type,
            uid,
            tag,
            size,
            ref_count,
            offset,
            usage,
            name_len,
            unused2,
            unused3,
            name,
        })
    }

    /// Map Windows BLIP type value to BlipType enum
    ///
    /// Based on MS-ODRAW specification
    fn map_windows_blip_type(win32_type: u8) -> Result<BlipType> {
        match win32_type {
            0x00 => Err(Error::ParseError("Unknown BLIP type".into())),
            0x01 => Ok(BlipType::Wmf),  // BI_WMF
            0x02 => Ok(BlipType::Emf),  // BI_EMF
            0x03 => Ok(BlipType::Pict), // BI_PICT
            0x04 => Ok(BlipType::Jpeg), // BI_JPEG
            0x05 => Ok(BlipType::Png),  // BI_PNG
            0x06 => Ok(BlipType::Dib),  // BI_DIB
            0x07 => Ok(BlipType::Tiff), // BI_TIFF
            _ => Err(Error::ParseError(format!(
                "Unsupported BLIP type: 0x{:02X}",
                win32_type
            ))),
        }
    }

    /// Get the expected record type ID for this BLIP type
    pub fn expected_record_type(&self) -> u16 {
        self.blip_type as u16
    }

    /// Check if this BSE has delay-loaded BLIP data
    ///
    /// When offset is non-zero, the BLIP data is stored separately
    /// and must be loaded from the specified offset.
    pub const fn is_delay_loaded(&self) -> bool {
        self.offset != 0
    }
}

/// BlipStore - container for all BSE records in a document
///
/// The BlipStore provides centralized access to all image metadata
/// in an Office document. It corresponds to the BStoreContainer (0xF001)
/// in the Escher drawing layer.
#[derive(Debug, Clone)]
pub struct BlipStore<'data> {
    /// All BSE records in the store
    pub entries: Vec<BlipStoreEntry<'data>>,
}

impl<'data> BlipStore<'data> {
    /// Convert this store to owned data with 'static lifetime
    pub fn into_owned(self) -> BlipStore<'static> {
        BlipStore {
            entries: self.entries.into_iter().map(|e| e.into_owned()).collect(),
        }
    }

    /// Create an empty BlipStore
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Create a BlipStore with pre-allocated capacity
    pub fn with_capacity(capacity: usize) -> Self {
        Self {
            entries: Vec::with_capacity(capacity),
        }
    }

    /// Add a BSE entry to the store
    pub fn add_entry(&mut self, entry: BlipStoreEntry<'data>) {
        self.entries.push(entry);
    }

    /// Get a BSE entry by index
    pub fn get_entry(&self, index: usize) -> Option<&BlipStoreEntry<'data>> {
        self.entries.get(index)
    }

    /// Get number of entries
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Check if store is empty
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Find BSE entry by UID
    pub fn find_by_uid(&self, uid: &[u8; 16]) -> Option<&BlipStoreEntry<'data>> {
        self.entries.iter().find(|entry| &entry.uid == uid)
    }

    /// Get iterator over all entries
    pub fn iter(&self) -> impl Iterator<Item = &BlipStoreEntry<'data>> {
        self.entries.iter()
    }
}

impl<'data> Default for BlipStore<'data> {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_map_windows_blip_type() {
        assert_eq!(
            BlipStoreEntry::map_windows_blip_type(0x02).unwrap(),
            BlipType::Emf
        );
        assert_eq!(
            BlipStoreEntry::map_windows_blip_type(0x04).unwrap(),
            BlipType::Jpeg
        );
        assert_eq!(
            BlipStoreEntry::map_windows_blip_type(0x05).unwrap(),
            BlipType::Png
        );
    }

    #[test]
    fn test_bse_minimal_parse() {
        // Minimal BSE record (36 bytes, no name)
        let data = vec![
            0x02, // btWin32 = EMF
            0x02, // btMacOS = EMF
            // UID (16 bytes)
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
            0x0F, 0x10, 0xFF, 0x00, // tag
            0x00, 0x10, 0x00, 0x00, // size = 4096
            0x01, 0x00, 0x00, 0x00, // ref_count = 1
            0x00, 0x00, 0x00, 0x00, // offset = 0
            0x00, // usage
            0x00, // name_len = 0
            0x00, // unused2
            0x00, // unused3
        ];

        let bse = BlipStoreEntry::parse(&data).unwrap();
        assert_eq!(bse.blip_type, BlipType::Emf);
        assert_eq!(bse.size, 4096);
        assert_eq!(bse.ref_count, 1);
        assert!(bse.name.is_none());
        assert!(!bse.is_delay_loaded());
    }
}
