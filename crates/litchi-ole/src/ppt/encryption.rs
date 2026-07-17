//! Office Binary Document RC4 CryptoAPI handling for legacy PowerPoint files.

use super::current_user::CurrentUser;
use super::package::{PptEncryptionKind, PptError, Result};
#[cfg(feature = "imgconv")]
use crate::office_crypto::cryptoapi::CryptoApiContext;
use crate::office_crypto::cryptoapi::{self, CryptoApiError};
use std::collections::BTreeMap;

const USER_EDIT_TYPE: u16 = 4085;
const CRYPT_SESSION_TYPE: u16 = 12052;
pub(super) struct EncryptedPresentation {
    pub live_offsets: Vec<usize>,
    pub mappings: Vec<(u32, u32)>,
    #[cfg(feature = "imgconv")]
    pub crypto: CryptoApiContext,
}

#[derive(Clone, Copy)]
struct RecordHeader {
    version: u16,
    instance: u16,
    record_type: u16,
    data_len: usize,
}

pub(super) fn decrypt_powerpoint_document(
    document: &mut Vec<u8>,
    current_user_data: Option<&[u8]>,
    password: Option<&str>,
) -> Result<Option<EncryptedPresentation>> {
    let Some(current_user_data) = current_user_data else {
        return Ok(None);
    };
    let current_user = CurrentUser::parse(current_user_data)?;
    let user_edit_offset = usize::try_from(current_user.current_edit_offset()).map_err(|_| {
        PptError::MalformedEncryptionHeader(
            "current edit offset does not fit in memory".to_string(),
        )
    })?;
    let user_edit = record_header(document, user_edit_offset)?;
    if user_edit.record_type != USER_EDIT_TYPE || user_edit.version != 0 || user_edit.instance != 0
    {
        if current_user.is_encrypted() {
            return Err(PptError::MalformedEncryptionHeader(
                "CurrentUser does not reference a valid UserEditAtom".to_string(),
            ));
        }
        return Ok(None);
    }
    if !matches!(user_edit.data_len, 28 | 32) {
        return Err(PptError::MalformedEncryptionHeader(format!(
            "UserEditAtom has invalid length {}",
            user_edit.data_len
        )));
    }
    let user_data = checked_slice(
        document,
        user_edit_offset + 8,
        user_edit.data_len,
        "UserEditAtom",
    )?;
    let session_id = (user_edit.data_len == 32)
        .then(|| le_u32(user_data, 28))
        .transpose()?;
    let session_id = session_id.filter(|id| *id != u32::MAX && *id != 0);
    if session_id.is_none() {
        if current_user.is_encrypted() {
            return Err(PptError::MalformedEncryptionHeader(
                "encrypted CurrentUser has no encryption-session persist reference".to_string(),
            ));
        }
        return Ok(None);
    }
    let session_id = session_id.unwrap();
    if le_u32(user_data, 8)? != 0 {
        return Err(PptError::MalformedEncryptionHeader(
            "encrypted presentation contains more than one UserEditAtom".to_string(),
        ));
    }
    if le_u32(user_data, 16)? != 1 {
        return Err(PptError::MalformedEncryptionHeader(
            "encrypted UserEditAtom has an invalid document persist identifier".to_string(),
        ));
    }
    let directory_offset = usize::try_from(le_u32(user_data, 12)?).map_err(|_| {
        PptError::MalformedEncryptionHeader(
            "persist directory offset does not fit in memory".to_string(),
        )
    })?;
    if directory_offset >= user_edit_offset {
        return Err(PptError::MalformedEncryptionHeader(
            "persist directory does not precede UserEditAtom".to_string(),
        ));
    }
    let mappings = parse_persist_directory(document, directory_offset)?;
    let session_offset = mappings.get(&session_id).copied().ok_or_else(|| {
        PptError::MalformedEncryptionHeader(
            "encryption-session persist identifier is absent from the directory".to_string(),
        )
    })?;
    let session_offset = usize::try_from(session_offset).map_err(|_| {
        PptError::MalformedEncryptionHeader(
            "encryption-session offset does not fit in memory".to_string(),
        )
    })?;
    let session_header = record_header(document, session_offset)?;
    if session_header.record_type != CRYPT_SESSION_TYPE
        || session_header.version != 0x0f
        || session_header.instance != 0
    {
        return Err(PptError::MalformedEncryptionHeader(
            "session persist object is not a CryptSession10Container".to_string(),
        ));
    }
    let session_data = checked_slice(
        document,
        session_offset + 8,
        session_header.data_len,
        "CryptSession10Container",
    )?;
    let header = cryptoapi::parse_header(session_data).map_err(map_crypto_error)?;
    let password = password.ok_or(PptError::PasswordRequired)?;
    let crypto = cryptoapi::verify_password(&header, password)
        .map_err(map_crypto_error)?
        .ok_or(PptError::InvalidPassword)?;

    let mut decrypted = document.clone();
    let mut ranges = Vec::new();
    for (&persist_id, &raw_offset) in &mappings {
        if persist_id == session_id {
            continue;
        }
        let offset = usize::try_from(raw_offset).map_err(|_| {
            PptError::Corrupted("persist object offset does not fit in memory".to_string())
        })?;
        if offset == user_edit_offset || offset == directory_offset {
            continue;
        }
        let encrypted_header = checked_slice(document, offset, 8, "encrypted persist header")?;
        let mut clear_header = encrypted_header.to_vec();
        cryptoapi::apply_block_cipher(&crypto, persist_id, &mut clear_header)
            .map_err(map_crypto_error)?;
        let data_len = usize::try_from(u32::from_le_bytes(clear_header[4..8].try_into().unwrap()))
            .map_err(|_| {
                PptError::Corrupted("persist record length does not fit in memory".to_string())
            })?;
        let total = 8usize
            .checked_add(data_len)
            .ok_or_else(|| PptError::Corrupted("persist record length overflow".to_string()))?;
        let end = offset
            .checked_add(total)
            .ok_or_else(|| PptError::Corrupted("persist record range overflow".to_string()))?;
        if end > directory_offset || end > document.len() {
            return Err(PptError::Corrupted(format!(
                "persist object {persist_id} extends beyond the encrypted object region"
            )));
        }
        ranges.push((offset, end, persist_id));
    }
    ranges.sort_unstable_by_key(|range| range.0);
    for pair in ranges.windows(2) {
        if pair[0].1 > pair[1].0 {
            return Err(PptError::Corrupted(
                "encrypted persist object ranges overlap".to_string(),
            ));
        }
    }
    for &(start, end, persist_id) in &ranges {
        cryptoapi::apply_block_cipher(&crypto, persist_id, &mut decrypted[start..end])
            .map_err(map_crypto_error)?;
    }
    *document = decrypted;

    Ok(Some(EncryptedPresentation {
        live_offsets: ranges.into_iter().map(|range| range.0).collect(),
        mappings: mappings.into_iter().collect(),
        #[cfg(feature = "imgconv")]
        crypto,
    }))
}

#[cfg(feature = "imgconv")]
pub(super) fn decrypt_pictures(data: &mut Vec<u8>, crypto: &CryptoApiContext) -> Result<()> {
    let mut clear = data.clone();
    let mut record_offset = 0usize;
    while record_offset < clear.len() {
        decrypt_picture_segment(&mut clear, record_offset, 8, crypto)?;
        let header = record_header(&clear, record_offset)?;
        let end = record_offset
            .checked_add(8)
            .and_then(|value| value.checked_add(header.data_len))
            .ok_or_else(|| PptError::Corrupted("picture record range overflow".to_string()))?;
        if end > clear.len() {
            return Err(PptError::Corrupted(
                "encrypted picture record extends beyond Pictures stream".to_string(),
            ));
        }
        let mut offset = record_offset + 8;
        let mut record_type = header.record_type;
        let mut instance = header.instance;
        if record_type == 0xf007 {
            const PARTS: [usize; 11] = [1, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1];
            for part in PARTS {
                decrypt_picture_segment(&mut clear, offset, part, crypto)?;
                offset += part;
            }
            let name_len = usize::from(clear[offset - 3]);
            if name_len != 0 {
                decrypt_picture_segment(&mut clear, offset, name_len, crypto)?;
                offset += name_len;
            }
            if offset == end {
                record_offset = end;
                continue;
            }
            decrypt_picture_segment(&mut clear, offset, 8, crypto)?;
            let embedded = record_header(&clear, offset)?;
            record_type = embedded.record_type;
            instance = embedded.instance;
            offset += 8;
        }
        let uid_count = if matches!(
            instance,
            0x217 | 0x3d5 | 0x46b | 0x543 | 0x6e1 | 0x6e3 | 0x6e5 | 0x7a9
        ) {
            2
        } else {
            1
        };
        for _ in 0..uid_count {
            decrypt_picture_segment(&mut clear, offset, 16, crypto)?;
            offset += 16;
        }
        let metadata_len = if matches!(record_type, 0xf01a..=0xf01c) {
            34
        } else {
            1
        };
        decrypt_picture_segment(&mut clear, offset, metadata_len, crypto)?;
        offset += metadata_len;
        if offset > end {
            return Err(PptError::Corrupted(
                "encrypted picture fields exceed their record".to_string(),
            ));
        }
        decrypt_picture_segment(&mut clear, offset, end - offset, crypto)?;
        record_offset = end;
    }
    *data = clear;
    Ok(())
}

fn parse_persist_directory(data: &[u8], offset: usize) -> Result<BTreeMap<u32, u32>> {
    let header = record_header(data, offset)?;
    if !matches!(header.record_type, 6001 | 6002) || header.version != 0 {
        return Err(PptError::MalformedEncryptionHeader(
            "UserEditAtom does not reference a valid PersistDirectoryAtom".to_string(),
        ));
    }
    let payload = checked_slice(data, offset + 8, header.data_len, "PersistDirectoryAtom")?;
    if payload.len() % 4 != 0 {
        return Err(PptError::MalformedEncryptionHeader(
            "persist directory is not aligned to 4 bytes".to_string(),
        ));
    }
    let mut mappings = BTreeMap::new();
    let mut cursor = 0usize;
    while cursor < payload.len() {
        let info = le_u32(payload, cursor)?;
        cursor += 4;
        let base = info & 0x000f_ffff;
        let count = info >> 20;
        if count == 0 {
            return Err(PptError::MalformedEncryptionHeader(
                "persist directory entry has a zero count".to_string(),
            ));
        }
        for index in 0..count {
            let persist_offset = le_u32(payload, cursor)?;
            cursor += 4;
            let id = base.checked_add(index).ok_or_else(|| {
                PptError::MalformedEncryptionHeader("persist identifier overflow".to_string())
            })?;
            if mappings.insert(id, persist_offset).is_some() {
                return Err(PptError::MalformedEncryptionHeader(format!(
                    "duplicate persist identifier {id}"
                )));
            }
            if usize::try_from(persist_offset).map_or(true, |value| value >= offset) {
                return Err(PptError::MalformedEncryptionHeader(format!(
                    "persist object {id} does not precede its directory"
                )));
            }
        }
    }
    Ok(mappings)
}

fn map_crypto_error(error: CryptoApiError) -> PptError {
    match error {
        CryptoApiError::Malformed(message) => PptError::MalformedEncryptionHeader(message),
        CryptoApiError::UnsupportedVersion { major, minor } => {
            PptError::UnsupportedEncryption(PptEncryptionKind::Unknown { major, minor })
        },
        CryptoApiError::UnsupportedAlgorithm => {
            PptError::UnsupportedEncryption(PptEncryptionKind::CryptoApi)
        },
    }
}

#[cfg(feature = "imgconv")]
fn decrypt_picture_segment(
    data: &mut [u8],
    offset: usize,
    len: usize,
    crypto: &CryptoApiContext,
) -> Result<()> {
    let segment = checked_slice_mut(data, offset, len, "encrypted picture field")?;
    cryptoapi::apply_block_cipher(crypto, 0, segment).map_err(map_crypto_error)
}

fn record_header(data: &[u8], offset: usize) -> Result<RecordHeader> {
    let bytes = checked_slice(data, offset, 8, "PPT record header")?;
    let version_instance = u16::from_le_bytes(bytes[0..2].try_into().unwrap());
    Ok(RecordHeader {
        version: version_instance & 0x000f,
        instance: version_instance >> 4,
        record_type: u16::from_le_bytes(bytes[2..4].try_into().unwrap()),
        data_len: usize::try_from(u32::from_le_bytes(bytes[4..8].try_into().unwrap())).map_err(
            |_| PptError::Corrupted("PPT record length does not fit in memory".to_string()),
        )?,
    })
}

fn le_u32(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = checked_slice(data, offset, 4, "32-bit encryption field")?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

fn checked_slice<'a>(data: &'a [u8], offset: usize, len: usize, field: &str) -> Result<&'a [u8]> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| PptError::MalformedEncryptionHeader(format!("{field} range overflow")))?;
    data.get(offset..end)
        .ok_or_else(|| PptError::MalformedEncryptionHeader(format!("{field} is truncated")))
}

#[cfg(feature = "imgconv")]
fn checked_slice_mut<'a>(
    data: &'a mut [u8],
    offset: usize,
    len: usize,
    field: &str,
) -> Result<&'a mut [u8]> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| PptError::MalformedEncryptionHeader(format!("{field} range overflow")))?;
    data.get_mut(offset..end)
        .ok_or_else(|| PptError::MalformedEncryptionHeader(format!("{field} is truncated")))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::{Package, PptOpenOptions};
    use std::path::{Path, PathBuf};

    fn poi_fixture(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .unwrap()
            .parent()
            .unwrap()
            .join("3rdparty")
            .join("poi")
            .join("test-data")
            .join("slideshow")
            .join(name)
    }

    #[test]
    fn encrypted_presentation_requires_password() {
        let mut package = Package::open(poi_fixture("Password_Protected-hello.ppt")).unwrap();
        assert!(matches!(
            package.presentation(),
            Err(PptError::PasswordRequired)
        ));
    }

    #[test]
    fn encrypted_presentation_rejects_wrong_password() {
        let mut package = Package::open(poi_fixture("Password_Protected-hello.ppt")).unwrap();
        assert!(matches!(
            package.presentation_with_options(PptOpenOptions {
                password: Some("wrong")
            }),
            Err(PptError::InvalidPassword)
        ));
    }

    #[test]
    fn opens_poi_cryptoapi_key_variants() {
        for name in [
            "Password_Protected-hello.ppt",
            "Password_Protected-56-hello.ppt",
            "Password_Protected-np-hello.ppt",
        ] {
            let mut package = Package::open(poi_fixture(name)).unwrap();
            let presentation = package
                .presentation_with_options(PptOpenOptions {
                    password: Some("hello"),
                })
                .unwrap();
            assert!(presentation.slide_count() > 0, "fixture {name}");
        }
    }

    #[test]
    fn opens_poi_cryptoapi_content_oracle() {
        let mut package = Package::open(poi_fixture("cryptoapi-proc2356.ppt")).unwrap();
        let presentation = package
            .presentation_with_options(PptOpenOptions {
                password: Some("crypto"),
            })
            .unwrap();
        assert!(presentation.text().unwrap().contains("Dominic Salemno"));
    }

    #[cfg(feature = "imgconv")]
    #[test]
    fn decrypts_poi_encrypted_pictures_stream() {
        use crate::{OleFile, extractor::ImageExtractor};
        use sha1::{Digest, Sha1};
        use std::fs::File;

        let mut encrypted_ole =
            OleFile::open(File::open(poi_fixture("cryptoapi-proc2356.ppt")).unwrap()).unwrap();
        let mut document = encrypted_ole.open_stream(&["PowerPoint Document"]).unwrap();
        let current_user = encrypted_ole.open_stream(&["Current User"]).unwrap();
        let state = decrypt_powerpoint_document(&mut document, Some(&current_user), Some("crypto"))
            .unwrap()
            .unwrap();
        let mut encrypted_pictures = encrypted_ole.open_stream(&["Pictures"]).unwrap();
        decrypt_pictures(&mut encrypted_pictures, &state.crypto).unwrap();
        let images = ImageExtractor::extract_blips(&encrypted_pictures).unwrap();
        let hashes: Vec<String> = images
            .iter()
            .map(|image| {
                Sha1::digest(image.raw_data())
                    .iter()
                    .map(|byte| format!("{byte:02x}"))
                    .collect()
            })
            .collect();
        assert_eq!(
            hashes,
            [
                "9cab034caab14c247c2c591555694ff46493bd9d",
                "4ae34e47ef55d54558648a1e0fae65dd54da2e87",
                "425dc81abaf86cdab4ed94e9e623e0edbfa9bdaf",
                "f2976cb7d36304649f59ecd2644f3e69584845ed",
                "828eb1a96ee5be40ad94e3b9b582e231f6f8a31c",
                "81950cf18a9134be6418d7f2c98bc411eae7bc27",
                "08d5368a2a941409e4dd30d7b175758a11fd7913",
            ]
        );
    }
}
