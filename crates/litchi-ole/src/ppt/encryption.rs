//! Office Binary Document RC4 CryptoAPI handling for legacy PowerPoint files.

use super::current_user::CurrentUser;
use super::package::{PptEncryptionKind, PptError, Result};
use litchi_crypto::rc4 as office_rc4;
use litchi_crypto::rc4::{Context, Error, Flags};
use rand::{TryRng, rngs::SysRng};
use std::collections::BTreeMap;
use zeroize::Zeroizing;

const USER_EDIT_TYPE: u16 = 4085;
const CRYPT_SESSION_TYPE: u16 = 12052;
const CRYPTO_API_PROVIDER: &str = "Microsoft Enhanced Cryptographic Provider v1.0";

/// Password-to-open encryption profile for binary PowerPoint presentations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptEncryptionProfile {
    /// Office Binary Document CryptoAPI RC4/SHA-1 encryption.
    CryptoApiRc4 {
        /// RC4 key size in bits. Supported values are 40 through 128 in steps of eight.
        key_bits: u16,
    },
}

impl PptEncryptionProfile {
    pub(crate) fn validate(self) -> std::result::Result<(), String> {
        let Self::CryptoApiRc4 { key_bits } = self;
        if !(40..=128).contains(&key_bits) || key_bits % 8 != 0 {
            return Err(format!(
                "PPT CryptoAPI RC4 key size {key_bits} is not a byte-aligned value in 40..=128"
            ));
        }
        Ok(())
    }
}

pub(crate) fn validate_writer_password(
    profile: PptEncryptionProfile,
    password: &str,
) -> std::result::Result<(), String> {
    profile.validate()?;
    let units = password.encode_utf16().count();
    if units == 0 {
        return Err("PPT password-to-open password must not be empty".to_string());
    }
    if units > 255 {
        return Err(format!(
            "PPT password contains {units} UTF-16 code units and exceeds the 255-unit limit"
        ));
    }
    Ok(())
}

pub(crate) struct WriterEncryptionMaterial {
    pub session_record: Vec<u8>,
    pub crypto: Context,
}

pub(crate) fn prepare_writer_encryption(
    profile: PptEncryptionProfile,
    password: &str,
) -> std::result::Result<WriterEncryptionMaterial, String> {
    validate_writer_password(profile, password)?;
    let PptEncryptionProfile::CryptoApiRc4 { key_bits } = profile;
    let mut salt = Zeroizing::new([0u8; 16]);
    let mut verifier = Zeroizing::new([0u8; 16]);
    SysRng.try_fill_bytes(salt.as_mut()).map_err(|_| {
        "operating-system randomness unavailable for PPT CryptoAPI salt".to_string()
    })?;
    SysRng.try_fill_bytes(verifier.as_mut()).map_err(|_| {
        "operating-system randomness unavailable for PPT CryptoAPI verifier".to_string()
    })?;
    let (encryption_info, crypto) = office_rc4::build_header(
        password,
        usize::from(key_bits),
        Flags::CRYPTO_API | Flags::DOC_PROPERTIES,
        CRYPTO_API_PROVIDER,
        &salt,
        &verifier,
    )
    .map_err(|error| map_crypto_error(error).to_string())?;
    let data_len = u32::try_from(encryption_info.len())
        .map_err(|_| "PPT CryptoAPI EncryptionInfo is too large".to_string())?;
    let mut session_record = Vec::with_capacity(8 + encryption_info.len());
    session_record.extend_from_slice(&0x000fu16.to_le_bytes());
    session_record.extend_from_slice(&CRYPT_SESSION_TYPE.to_le_bytes());
    session_record.extend_from_slice(&data_len.to_le_bytes());
    session_record.extend_from_slice(&encryption_info);
    Ok(WriterEncryptionMaterial {
        session_record,
        crypto,
    })
}

pub(crate) fn encrypt_powerpoint_document_for_write(
    document: &mut [u8],
    directory_offset: usize,
    user_edit_offset: usize,
    session_id: u32,
    crypto: &Context,
) -> std::result::Result<(), String> {
    if directory_offset >= user_edit_offset || user_edit_offset >= document.len() {
        return Err("PPT encrypted bootstrap offsets are out of order".to_string());
    }
    let user_edit = record_header(document, user_edit_offset).map_err(|error| error.to_string())?;
    if user_edit.record_type != USER_EDIT_TYPE
        || user_edit.version != 0
        || user_edit.instance != 0
        || user_edit.data_len != 32
    {
        return Err("PPT encrypted output requires one 32-byte UserEditAtom".to_string());
    }
    let user_end = user_edit_offset
        .checked_add(8 + user_edit.data_len)
        .ok_or_else(|| "PPT UserEditAtom range overflow".to_string())?;
    if user_end != document.len() {
        return Err("PPT encrypted output contains data after its single UserEditAtom".to_string());
    }
    let user_data = document
        .get(user_edit_offset + 8..user_end)
        .ok_or_else(|| "PPT UserEditAtom is truncated".to_string())?;
    if le_u32(user_data, 8).map_err(|error| error.to_string())? != 0
        || le_u32(user_data, 12).map_err(|error| error.to_string())? != directory_offset as u32
        || le_u32(user_data, 16).map_err(|error| error.to_string())? != 1
        || le_u32(user_data, 28).map_err(|error| error.to_string())? != session_id
    {
        return Err("PPT encrypted UserEditAtom bootstrap fields are inconsistent".to_string());
    }
    let directory = record_header(document, directory_offset).map_err(|error| error.to_string())?;
    let directory_end = directory_offset
        .checked_add(8 + directory.data_len)
        .ok_or_else(|| "PPT persist directory range overflow".to_string())?;
    if directory_end != user_edit_offset {
        return Err("PPT persist directory is not immediately before UserEditAtom".to_string());
    }
    let mappings =
        parse_persist_directory(document, directory_offset).map_err(|error| error.to_string())?;
    let session_offset = usize::try_from(
        *mappings
            .get(&session_id)
            .ok_or_else(|| "PPT encryption session has no persist mapping".to_string())?,
    )
    .map_err(|_| "PPT encryption session offset does not fit in memory".to_string())?;

    let mut ranges = Vec::with_capacity(mappings.len());
    for (&persist_id, &raw_offset) in &mappings {
        let offset = usize::try_from(raw_offset)
            .map_err(|_| format!("PPT persist object {persist_id} offset is too large"))?;
        let header = record_header(document, offset).map_err(|error| error.to_string())?;
        let end = offset
            .checked_add(8 + header.data_len)
            .ok_or_else(|| format!("PPT persist object {persist_id} range overflow"))?;
        if end > directory_offset {
            return Err(format!(
                "PPT persist object {persist_id} extends into the bootstrap records"
            ));
        }
        if persist_id == session_id
            && (offset != session_offset
                || header.record_type != CRYPT_SESSION_TYPE
                || header.version != 0x0f
                || header.instance != 0)
        {
            return Err("PPT encryption session persist object is malformed".to_string());
        }
        ranges.push((offset, end, persist_id));
    }
    ranges.sort_unstable_by_key(|range| range.0);
    let mut cursor = 0usize;
    for &(start, end, persist_id) in &ranges {
        if start != cursor {
            return Err(format!(
                "PPT persist object {persist_id} leaves a gap or overlap before offset {start}"
            ));
        }
        cursor = end;
    }
    if cursor != directory_offset {
        return Err("PPT persist objects do not cover the complete encrypted region".to_string());
    }
    for &(start, end, persist_id) in &ranges {
        if persist_id != session_id {
            office_rc4::apply(crypto, persist_id, &mut document[start..end])
                .map_err(|error| map_crypto_error(error).to_string())?;
        }
    }
    Ok(())
}

pub(crate) fn encrypt_pictures_for_write(
    data: &mut [u8],
    crypto: &Context,
) -> std::result::Result<(), String> {
    let segments = clear_picture_segments(data)?;
    for (offset, len) in segments {
        office_rc4::apply(crypto, 0, &mut data[offset..offset + len])
            .map_err(|error| map_crypto_error(error).to_string())?;
    }
    Ok(())
}

fn clear_picture_segments(data: &[u8]) -> std::result::Result<Vec<(usize, usize)>, String> {
    let mut segments = Vec::new();
    let mut record_offset = 0usize;
    while record_offset < data.len() {
        let header = record_header(data, record_offset).map_err(|error| error.to_string())?;
        let end = record_offset
            .checked_add(8)
            .and_then(|value| value.checked_add(header.data_len))
            .ok_or_else(|| "PPT picture record range overflow".to_string())?;
        if end > data.len() {
            return Err("PPT picture record extends beyond the Pictures stream".to_string());
        }
        segments.push((record_offset, 8));
        let mut offset = record_offset + 8;
        let mut record_type = header.record_type;
        let mut instance = header.instance;
        if record_type == 0xf007 {
            for len in [1usize, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1] {
                checked_picture_segment(offset, len, end)?;
                segments.push((offset, len));
                offset += len;
            }
            let name_len = usize::from(data[offset - 3]);
            if name_len != 0 {
                checked_picture_segment(offset, name_len, end)?;
                segments.push((offset, name_len));
                offset += name_len;
            }
            if offset == end {
                record_offset = end;
                continue;
            }
            checked_picture_segment(offset, 8, end)?;
            segments.push((offset, 8));
            let embedded = record_header(data, offset).map_err(|error| error.to_string())?;
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
            checked_picture_segment(offset, 16, end)?;
            segments.push((offset, 16));
            offset += 16;
        }
        let metadata_len = if matches!(record_type, 0xf01a..=0xf01c) {
            34
        } else {
            1
        };
        checked_picture_segment(offset, metadata_len, end)?;
        segments.push((offset, metadata_len));
        offset += metadata_len;
        checked_picture_segment(offset, end - offset, end)?;
        if offset != end {
            segments.push((offset, end - offset));
        }
        record_offset = end;
    }
    Ok(segments)
}

fn checked_picture_segment(
    offset: usize,
    len: usize,
    record_end: usize,
) -> std::result::Result<(), String> {
    if offset.checked_add(len).is_none_or(|end| end > record_end) {
        return Err("PPT picture fields exceed their record".to_string());
    }
    Ok(())
}
pub(super) struct EncryptedPresentation {
    pub live_offsets: Vec<usize>,
    pub mappings: Vec<(u32, u32)>,
    pub crypto: Context,
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
    let header = office_rc4::parse_header(session_data).map_err(map_crypto_error)?;
    let password = password.ok_or(PptError::PasswordRequired)?;
    let crypto = office_rc4::verify(&header, password)
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
        office_rc4::apply(&crypto, persist_id, &mut clear_header).map_err(map_crypto_error)?;
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
        office_rc4::apply(&crypto, persist_id, &mut decrypted[start..end])
            .map_err(map_crypto_error)?;
    }
    *document = decrypted;

    Ok(Some(EncryptedPresentation {
        live_offsets: ranges.into_iter().map(|range| range.0).collect(),
        mappings: mappings.into_iter().collect(),
        crypto,
    }))
}

pub(super) fn decrypt_pictures(data: &mut Vec<u8>, crypto: &Context) -> Result<()> {
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

fn map_crypto_error(error: Error) -> PptError {
    match error {
        Error::Malformed(message) => PptError::MalformedEncryptionHeader(message),
        Error::UnsupportedVersion { major, minor } => {
            PptError::UnsupportedEncryption(PptEncryptionKind::Unknown { major, minor })
        },
        Error::UnsupportedAlgorithm => {
            PptError::UnsupportedEncryption(PptEncryptionKind::CryptoApi)
        },
    }
}

fn decrypt_picture_segment(
    data: &mut [u8],
    offset: usize,
    len: usize,
    crypto: &Context,
) -> Result<()> {
    let segment = checked_slice_mut(data, offset, len, "encrypted picture field")?;
    office_rc4::apply(crypto, 0, segment).map_err(map_crypto_error)
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
            .join("test-data")
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
        let images = ImageExtractor::pictures(&encrypted_pictures).unwrap();
        let hashes: Vec<String> = images
            .iter()
            .map(|image| {
                Sha1::digest(image.data().unwrap())
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
