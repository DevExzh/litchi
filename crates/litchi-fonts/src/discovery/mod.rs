use font_kit::family_name::FamilyName;
use font_kit::handle::Handle;
use font_kit::properties::{Properties, Style as FontStyle, Weight};
use font_kit::source::SystemSource;
use std::sync::Arc;

use crate::model::{
    Charset, Family, FontData, FontError, FontProperties, License, Panose, Pitch, Request,
    Signature,
};

/// Resolves a typed family/style request against the host system font source.
pub struct Loader {
    source: SystemSource,
}

impl Loader {
    pub fn new() -> Self {
        Self {
            source: SystemSource::new(),
        }
    }

    pub fn load_system_font(&self, family_name: &str) -> Result<FontData, FontError> {
        self.load(&Request::regular(family_name))
    }

    /// Load the best system face for one typed family/style request.
    pub fn load(&self, request: &Request) -> Result<FontData, FontError> {
        let mut properties = Properties::new();
        properties.style = if request.style().is_italic() {
            FontStyle::Italic
        } else {
            FontStyle::Normal
        };
        properties.weight = if request.style().is_bold() {
            Weight::BOLD
        } else {
            Weight::NORMAL
        };
        let handle = self
            .source
            .select_best_match(
                &[FamilyName::Title(request.family().to_owned())],
                &properties,
            )
            .map_err(|_| FontError::NotFound(request.family().to_owned()))?;

        match handle {
            Handle::Path { path, font_index } => {
                let data = std::fs::read(&path)?;
                let properties = Self::extract_font_properties(&data, font_index)?;
                Ok(FontData {
                    name: request.family().to_owned(),
                    data,
                    index: font_index,
                    properties,
                })
            },
            Handle::Memory { bytes, font_index } => {
                let properties = Self::extract_font_properties(&bytes, font_index)?;
                Ok(FontData {
                    name: request.family().to_owned(),
                    data: into_owned_bytes(bytes),
                    index: font_index,
                    properties,
                })
            },
        }
    }

    /// Extract font properties from font data for Office embedding
    fn extract_font_properties(
        data: &[u8],
        font_index: u32,
    ) -> Result<Option<FontProperties>, FontError> {
        use allsorts::binary::read::ReadScope;
        use allsorts::error::ParseError;
        use allsorts::tables::{FontTableProvider, OpenTypeData, OpenTypeFont};

        let scope = ReadScope::new(data);
        let Some(font_file) = scope.read::<OpenTypeFont<'_>>().ok() else {
            return Ok(None);
        };
        let index =
            usize::try_from(font_index).map_err(|_| FontError::InvalidFaceIndex(font_index))?;
        if index != 0 && matches!(&font_file.data, OpenTypeData::Single(_)) {
            return Err(FontError::InvalidFaceIndex(font_index));
        }
        let provider = font_file
            .table_provider(index)
            .map_err(|error| match error {
                ParseError::BadIndex => FontError::InvalidFaceIndex(font_index),
                _ => FontError::InvalidData,
            })?;

        // Get OS/2 table raw bytes
        let Some(os2_table) = provider
            .table_data(allsorts::tag::OS_2)
            .map_err(|_| FontError::InvalidData)?
        else {
            return Ok(None);
        };
        let os2_table: &[u8] = os2_table.as_ref();
        parse_os2_properties(os2_table).map(Some)
    }
}

/// Adopt a uniquely owned font-kit allocation, copying only when it is shared.
#[inline]
fn into_owned_bytes(bytes: Arc<Vec<u8>>) -> Vec<u8> {
    Arc::unwrap_or_clone(bytes)
}

fn parse_os2_properties(os2_table: &[u8]) -> Result<FontProperties, FontError> {
    if os2_table.len() < 2 {
        return Err(FontError::InvalidData);
    }
    let version = u16::from_be_bytes([os2_table[0], os2_table[1]]);
    let minimum = match version {
        0 => 78,
        1 => 86,
        2..=4 => 96,
        5 => 100,
        _ => return Err(FontError::UnsupportedOs2Version(version)),
    };
    if os2_table.len() < minimum {
        return Err(FontError::TruncatedOs2 {
            version,
            expected: minimum,
            actual: os2_table.len(),
        });
    }

    // `OS/2.fsType` is a big-endian u16 at byte offset 8.
    let license = License::from_os2(version, u16::from_be_bytes([os2_table[8], os2_table[9]]))?;

    // PANOSE is at offset 32 (0x20) in OS/2 table (10 bytes).
    let panose_bytes: [u8; 10] = os2_table[32..42]
        .try_into()
        .map_err(|_| FontError::InvalidData)?;
    let panose = Panose::new(panose_bytes);

    // Unicode ranges at offset 42-57 (4 DWORDs).
    let unicode = [
        read_u32(os2_table, 42)?,
        read_u32(os2_table, 46)?,
        read_u32(os2_table, 50)?,
        read_u32(os2_table, 54)?,
    ];

    // Version zero does not define code-page ranges even if its table is padded.
    let code_pages = if version == 0 {
        [0, 0]
    } else {
        let mut first = read_u32(os2_table, 78)?;
        if version == 1 {
            // Bit 8 (Vietnamese) was assigned beginning with version two.
            first &= !(1 << 8);
        }
        [first, read_u32(os2_table, 82)?]
    };

    let signature = Signature::new(unicode, code_pages);
    let charset = charset_from_code_pages(code_pages);

    // PANOSE byte three is a proportion only for Latin Text (family kind two).
    let pitch = match (panose_bytes[0], panose_bytes[3]) {
        (2, 9) => Pitch::Fixed,
        (2, 2..=8) => Pitch::Variable,
        _ => Pitch::Default,
    };

    // The high byte of OS/2.sFamilyClass is the IBM family class. Fixed
    // advance fonts use Office's `modern` classification regardless of serif
    // style; class six is reserved and must not be treated as Roman.
    let family = if pitch == Pitch::Fixed {
        Family::Modern
    } else {
        match os2_table[30] {
            1..=5 | 7 => Family::Roman,
            8 => Family::Swiss,
            9 | 12 => Family::Decorative,
            10 => Family::Script,
            _ => Family::Auto,
        }
    };

    Ok(FontProperties::new(
        license, panose, charset, family, pitch, signature,
    ))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32, FontError> {
    let end = offset.checked_add(4).ok_or(FontError::InvalidData)?;
    let value: [u8; 4] = bytes
        .get(offset..end)
        .ok_or(FontError::InvalidData)?
        .try_into()
        .map_err(|_| FontError::InvalidData)?;
    Ok(u32::from_be_bytes(value))
}

fn charset_from_code_pages(code_pages: [u32; 2]) -> Option<Charset> {
    const CANDIDATES: &[(u64, Charset)] = &[
        (1 << 0, Charset::ANSI),
        (1 << 1, Charset::EAST_EUROPE),
        (1 << 2, Charset::RUSSIAN),
        (1 << 3, Charset::GREEK),
        (1 << 4, Charset::TURKISH),
        (1 << 5, Charset::HEBREW),
        (1 << 6, Charset::ARABIC),
        (1 << 7, Charset::BALTIC),
        (1 << 8, Charset::VIETNAMESE),
        (1 << 16, Charset::THAI),
        (1 << 17, Charset::SHIFT_JIS),
        (1 << 18, Charset::GB2312),
        (1 << 19, Charset::HANGEUL),
        (1 << 20, Charset::CHINESE_BIG5),
        (1 << 21, Charset::JOHAB),
        (1 << 29, Charset::MACINTOSH),
        (1 << 30, Charset::OEM),
        (1 << 31, Charset::SYMBOL),
        (1 << 32, Charset::GREEK),
        (1 << 33, Charset::RUSSIAN),
        (1 << 34, Charset::OEM),
        (1 << 35, Charset::ARABIC),
        (1 << 36, Charset::OEM),
        (1 << 37, Charset::HEBREW),
        (1 << 38, Charset::OEM),
        (1 << 39, Charset::OEM),
        (1 << 40, Charset::TURKISH),
        (1 << 41, Charset::RUSSIAN),
        (1 << 42, Charset::EAST_EUROPE),
        (1 << 43, Charset::BALTIC),
        (1 << 44, Charset::GREEK),
        (1 << 45, Charset::ARABIC),
        (1 << 46, Charset::OEM),
        (1 << 47, Charset::OEM),
    ];

    let bits = u64::from(code_pages[0]) | (u64::from(code_pages[1]) << 32);
    let mut selected = None;
    for &(mask, charset) in CANDIDATES {
        if bits & mask == 0 {
            continue;
        }
        match selected {
            None => selected = Some(charset),
            Some(current) if current == charset => {},
            Some(_) => return None,
        }
    }
    selected
}

impl Default for Loader {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Permission;

    #[test]
    fn uniquely_owned_memory_bytes_keep_their_allocation() {
        let bytes = Arc::new(vec![1, 2, 3, 4]);
        let allocation = bytes.as_slice().as_ptr();

        let owned = into_owned_bytes(bytes);

        assert_eq!(owned.as_ptr(), allocation);
        assert_eq!(owned, [1, 2, 3, 4]);
    }

    #[test]
    fn shared_memory_bytes_leave_the_other_owner_intact() {
        let bytes = Arc::new(vec![1, 2, 3, 4]);
        let retained = Arc::clone(&bytes);
        let shared_allocation = retained.as_slice().as_ptr();

        let owned = into_owned_bytes(bytes);

        assert_eq!(owned, retained.as_slice());
        assert_ne!(owned.as_ptr(), shared_allocation);
        assert_eq!(retained.as_slice().as_ptr(), shared_allocation);
    }

    #[test]
    fn parses_versioned_os2_lengths_and_ignores_undefined_fields() {
        let mut padded_version_zero = os2_table(0, 0, 0, 1);
        padded_version_zero.resize(86, 0);
        write_u32(&mut padded_version_zero, 78, 1 << 17);
        let properties = parse_os2_properties(&padded_version_zero).expect("version zero table");
        assert_eq!(properties.charset(), None);
        assert_eq!(properties.signature().code_pages(), &[0, 0]);

        let version_one = os2_table(1, 0, 1 << 8, 1);
        let properties = parse_os2_properties(&version_one).expect("version one table");
        assert_eq!(properties.charset(), None);
        assert_eq!(properties.signature().code_pages(), &[0, 0]);

        let version_two = os2_table(2, 0, 1 << 8, 1);
        let properties = parse_os2_properties(&version_two).expect("version two table");
        assert_eq!(properties.charset(), Some(Charset::VIETNAMESE));
        assert_eq!(properties.signature().code_pages(), &[1 << 8, 0]);

        let truncated = vec![0; 77];
        assert!(matches!(
            parse_os2_properties(&truncated),
            Err(FontError::TruncatedOs2 {
                version: 0,
                expected: 78,
                actual: 77,
            })
        ));

        let unsupported = [0, 6];
        assert!(matches!(
            parse_os2_properties(&unsupported),
            Err(FontError::UnsupportedOs2Version(6))
        ));
    }

    #[test]
    fn derives_a_charset_only_from_one_unambiguous_code_page() {
        let cases = [
            (17, Charset::SHIFT_JIS),
            (18, Charset::GB2312),
            (19, Charset::HANGEUL),
            (20, Charset::CHINESE_BIG5),
            (21, Charset::JOHAB),
        ];
        for (bit, expected) in cases {
            assert_eq!(charset_from_code_pages([1 << bit, 0]), Some(expected));
        }
        assert_eq!(charset_from_code_pages([0, 0]), None);
        assert_eq!(charset_from_code_pages([(1 << 0) | (1 << 18), 0]), None);
        assert_eq!(
            charset_from_code_pages([1 << 3, (1 << 0) | (1 << 12)]),
            Some(Charset::GREEK)
        );
        assert_eq!(charset_from_code_pages([1 << 0, 1 << 0]), None);
    }

    #[test]
    fn selected_collection_face_owns_its_metadata_and_license() {
        let installable = os2_table(3, 0, 1 << 0, 1);
        let restricted = os2_table(3, 0x0002, 1 << 17, 2);
        let collection = ttc(&installable, &restricted);

        let first = Loader::extract_font_properties(&collection, 0)
            .expect("first face")
            .expect("first OS/2 table");
        let second = Loader::extract_font_properties(&collection, 1)
            .expect("second face")
            .expect("second OS/2 table");

        assert_eq!(first.license().permission(), Permission::Installable);
        assert_eq!(first.charset(), Some(Charset::ANSI));
        assert_eq!(first.panose().bytes()[2], 1);
        assert_eq!(second.license().permission(), Permission::Restricted);
        assert_eq!(second.charset(), Some(Charset::SHIFT_JIS));
        assert_eq!(second.panose().bytes()[2], 2);
        assert!(matches!(
            Loader::extract_font_properties(&collection, 2),
            Err(FontError::InvalidFaceIndex(2))
        ));

        let single = sfnt(&installable);
        assert!(matches!(
            Loader::extract_font_properties(&single, 1),
            Err(FontError::InvalidFaceIndex(1))
        ));
    }

    fn os2_table(version: u16, fs_type: u16, code_pages: u32, marker: u8) -> Vec<u8> {
        let len = match version {
            0 => 78,
            1 => 86,
            2..=4 => 96,
            5 => 100,
            _ => 2,
        };
        let mut table = vec![0; len];
        write_u16(&mut table, 0, version);
        write_u16(&mut table, 8, fs_type);
        table[30] = 8;
        table[32..42].copy_from_slice(&[2, 11, marker, 4, 2, 2, 2, 2, 2, 4]);
        if version >= 1 {
            write_u32(&mut table, 78, code_pages);
        }
        table
    }

    fn ttc(first: &[u8], second: &[u8]) -> Vec<u8> {
        const HEADER: usize = 20;
        const DIRECTORY: usize = 28;
        let first_directory = HEADER;
        let second_directory = first_directory + DIRECTORY;
        let first_table = second_directory + DIRECTORY;
        let second_table = align4(first_table + first.len());
        let mut bytes = vec![0; second_table + second.len()];
        bytes[0..4].copy_from_slice(b"ttcf");
        write_u32(&mut bytes, 4, 0x0001_0000);
        write_u32(&mut bytes, 8, 2);
        write_u32(
            &mut bytes,
            12,
            u32::try_from(first_directory).expect("test offset fits u32"),
        );
        write_u32(
            &mut bytes,
            16,
            u32::try_from(second_directory).expect("test offset fits u32"),
        );
        sfnt_directory(&mut bytes, first_directory, first_table, first);
        sfnt_directory(&mut bytes, second_directory, second_table, second);
        bytes[first_table..first_table + first.len()].copy_from_slice(first);
        bytes[second_table..second_table + second.len()].copy_from_slice(second);
        bytes
    }

    fn sfnt(table: &[u8]) -> Vec<u8> {
        const DIRECTORY: usize = 28;
        let mut bytes = vec![0; DIRECTORY + table.len()];
        sfnt_directory(&mut bytes, 0, DIRECTORY, table);
        bytes[DIRECTORY..].copy_from_slice(table);
        bytes
    }

    fn sfnt_directory(bytes: &mut [u8], offset: usize, table_offset: usize, table: &[u8]) {
        write_u32(bytes, offset, 0x0001_0000);
        write_u16(bytes, offset + 4, 1);
        write_u16(bytes, offset + 6, 16);
        bytes[offset + 12..offset + 16].copy_from_slice(b"OS/2");
        write_u32(bytes, offset + 16, checksum(table));
        write_u32(
            bytes,
            offset + 20,
            u32::try_from(table_offset).expect("test table offset fits u32"),
        );
        write_u32(
            bytes,
            offset + 24,
            u32::try_from(table.len()).expect("test table length fits u32"),
        );
    }

    fn checksum(bytes: &[u8]) -> u32 {
        bytes.chunks(4).fold(0, |sum, chunk| {
            let mut word = [0; 4];
            word[..chunk.len()].copy_from_slice(chunk);
            sum.wrapping_add(u32::from_be_bytes(word))
        })
    }

    fn align4(value: usize) -> usize {
        (value + 3) & !3
    }

    fn write_u16(bytes: &mut [u8], offset: usize, value: u16) {
        bytes[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
    }

    fn write_u32(bytes: &mut [u8], offset: usize, value: u32) {
        bytes[offset..offset + 4].copy_from_slice(&value.to_be_bytes());
    }
}
