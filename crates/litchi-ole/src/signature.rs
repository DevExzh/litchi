//! Trust-neutral XML digital signatures for legacy Office CFB containers.

use litchi_cfb::{DirectoryEntry, OleError, OleFile, OleWriter};
pub use litchi_opc::{
    CertificateTrust, EmbeddedCertificate, PackageSigner, ReferenceVerification, Sha1Policy,
    SignatureAlgorithm, SignatureVerificationPolicy, VerificationStatus,
};
use litchi_opc::{
    DetachedDigitalSignatureVerification, DetachedSignatureReference, DigitalSignatureError,
    author_detached_signature, verify_detached_signature,
};
use std::collections::HashSet;
use std::fmt;
use std::io::{Cursor, Read, Seek};

const XML_SIGNATURE_STORAGE: &str = "_xmlsignatures";
const LEGACY_SIGNATURE_STREAM: &str = "_signatures";
const MAX_CFB_STREAMS: usize = 20_000;
const MAX_CFB_BYTES: usize = 1024 * 1024 * 1024;
const MAX_SIGNATURES: usize = 64;

/// Binary Office application-specific digest rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOfficeFormat {
    Doc,
    Xls,
    Ppt,
}

/// Verification result for one stream in the `_xmlsignatures` storage.
#[derive(Debug, Clone)]
pub struct BinaryOfficeSignatureVerification {
    pub signature_stream: String,
    pub package_integrity: VerificationStatus,
    pub signature_value: VerificationStatus,
    pub certificate_trust: CertificateTrust,
    pub references: Vec<ReferenceVerification>,
    pub certificates: Vec<EmbeddedCertificate>,
    pub uses_sha1: bool,
    pub signing_time: Option<String>,
}

/// Errors raised while discovering, verifying, or authoring binary signatures.
#[derive(Debug)]
pub enum BinaryOfficeSignatureError {
    Ole(OleError),
    Xml(DigitalSignatureError),
    InvalidContainer(String),
    ResourceLimit(String),
    EncryptedDocument,
    LegacyCryptoApiUnsupported,
}

impl fmt::Display for BinaryOfficeSignatureError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ole(error) => write!(formatter, "OLE error: {error}"),
            Self::Xml(error) => write!(formatter, "digital-signature error: {error}"),
            Self::InvalidContainer(message) => {
                write!(formatter, "invalid signature container: {message}")
            },
            Self::ResourceLimit(message) => {
                write!(formatter, "signature resource limit exceeded: {message}")
            },
            Self::EncryptedDocument => write!(
                formatter,
                "encrypted binary Office documents cannot be signed or verified without their decrypted DRM storage"
            ),
            Self::LegacyCryptoApiUnsupported => write!(
                formatter,
                "legacy `_signatures` CryptoAPI MD5 signatures are recognized but unsupported"
            ),
        }
    }
}

impl std::error::Error for BinaryOfficeSignatureError {}

impl From<OleError> for BinaryOfficeSignatureError {
    fn from(error: OleError) -> Self {
        Self::Ole(error)
    }
}

impl From<DigitalSignatureError> for BinaryOfficeSignatureError {
    fn from(error: DigitalSignatureError) -> Self {
        Self::Xml(error)
    }
}

pub type Result<T> = std::result::Result<T, BinaryOfficeSignatureError>;

/// Discover and verify every binary Office XML signature without evaluating
/// certificate trust or executing embedded macro content.
pub fn verify_binary_office_signatures<R: Read + Seek>(
    ole: &mut OleFile<R>,
    format: BinaryOfficeFormat,
    policy: &SignatureVerificationPolicy,
) -> Result<Vec<BinaryOfficeSignatureVerification>> {
    let references = signed_references(ole, format)?;
    reject_encrypted(ole, format, &references)?;
    if ole.exists(&[LEGACY_SIGNATURE_STREAM]) {
        return Err(BinaryOfficeSignatureError::LegacyCryptoApiUnsupported);
    }
    if !ole.directory_exists(&[XML_SIGNATURE_STORAGE]) {
        return Ok(Vec::new());
    }
    let entries: Vec<DirectoryEntry> = ole
        .list_directory_entries(&[XML_SIGNATURE_STORAGE])?
        .into_iter()
        .cloned()
        .collect();
    if entries.is_empty() || entries.len() > MAX_SIGNATURES {
        return Err(BinaryOfficeSignatureError::ResourceLimit(format!(
            "signature count {} is outside 1..={MAX_SIGNATURES}",
            entries.len()
        )));
    }
    let mut numeric_names = HashSet::new();
    let mut streams = Vec::with_capacity(entries.len());
    for entry in entries {
        if entry.entry_type != 2
            || entry.name.is_empty()
            || entry.name.len() > 20
            || !entry.name.bytes().all(|byte| byte.is_ascii_digit())
        {
            return Err(BinaryOfficeSignatureError::InvalidContainer(format!(
                "signature entry {:?} is not a decimal stream",
                entry.name
            )));
        }
        let numeric = entry.name.parse::<u64>().map_err(|_| {
            BinaryOfficeSignatureError::InvalidContainer("signature stream number overflow".into())
        })?;
        if !numeric_names.insert(numeric) {
            return Err(BinaryOfficeSignatureError::InvalidContainer(
                "duplicate numeric signature stream name".into(),
            ));
        }
        let xml = ole.open_stream(&[XML_SIGNATURE_STORAGE, &entry.name])?;
        if xml.len() > policy.max_signature_part_bytes {
            return Err(BinaryOfficeSignatureError::ResourceLimit(format!(
                "signature stream {} is too large",
                entry.name
            )));
        }
        streams.push((numeric, entry.name, xml));
    }
    streams.sort_by_key(|entry| entry.0);
    streams
        .into_iter()
        .map(|(_, name, xml)| {
            let verification = verify_detached_signature(&xml, &references, policy)?;
            Ok(project_verification(name, verification))
        })
        .collect()
}

fn project_verification(
    signature_stream: String,
    value: DetachedDigitalSignatureVerification,
) -> BinaryOfficeSignatureVerification {
    BinaryOfficeSignatureVerification {
        signature_stream,
        package_integrity: value.package_integrity,
        signature_value: value.signature_value,
        certificate_trust: value.certificate_trust,
        references: value.references,
        certificates: value.certificates,
        uses_sha1: value.uses_sha1,
        signing_time: value.signing_time,
    }
}

/// Atomic in-memory editor for binary Office XML signatures.
pub struct BinaryOfficeSignatureEditor {
    original: Vec<u8>,
    format: BinaryOfficeFormat,
    sector_size: usize,
    root_clsid: [u8; 16],
    storages: Vec<Vec<String>>,
    streams: Vec<(Vec<String>, Vec<u8>)>,
    changed: bool,
}

impl fmt::Debug for BinaryOfficeSignatureEditor {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BinaryOfficeSignatureEditor")
            .field("format", &self.format)
            .field("stream_count", &self.streams.len())
            .field("changed", &self.changed)
            .finish()
    }
}

impl BinaryOfficeSignatureEditor {
    pub fn new(bytes: Vec<u8>, format: BinaryOfficeFormat) -> Result<Self> {
        if bytes.len() > MAX_CFB_BYTES {
            return Err(BinaryOfficeSignatureError::ResourceLimit(
                "compound file exceeds 1 GiB authoring limit".into(),
            ));
        }
        let mut ole = OleFile::open(Cursor::new(bytes.as_slice()))?;
        let sector_size = ole.sector_size();
        let root_clsid = parse_clsid(
            ole.root_entry()
                .map(|entry| entry.clsid.as_str())
                .unwrap_or_default(),
        )?;
        let storages = collect_storages(&ole)?;
        let mut streams = Vec::new();
        for path in ole.list_streams() {
            if streams.len() >= MAX_CFB_STREAMS {
                return Err(BinaryOfficeSignatureError::ResourceLimit(
                    "too many compound-file streams".into(),
                ));
            }
            let borrowed: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&borrowed)?;
            streams.push((path, data));
        }
        let references = references_from_streams(&streams, format)?;
        reject_encrypted_streams(&streams, format, &references)?;
        Ok(Self {
            original: bytes,
            format,
            sector_size,
            root_clsid,
            storages,
            streams,
            changed: false,
        })
    }

    pub fn verify(
        &self,
        policy: &SignatureVerificationPolicy,
    ) -> Result<Vec<BinaryOfficeSignatureVerification>> {
        let mut ole = OleFile::open(Cursor::new(self.finish_snapshot()?))?;
        verify_binary_office_signatures(&mut ole, self.format, policy)
    }

    pub fn add_signature(&mut self, signer: &PackageSigner) -> Result<&mut Self> {
        if self.streams.iter().any(|(path, _)| {
            path.len() == 1 && path[0].eq_ignore_ascii_case(LEGACY_SIGNATURE_STREAM)
        }) {
            return Err(BinaryOfficeSignatureError::LegacyCryptoApiUnsupported);
        }
        let references = references_from_streams(&self.streams, self.format)?;
        let xml = author_detached_signature(signer, &references)?;
        let existing: HashSet<u64> = self
            .streams
            .iter()
            .filter(|(path, _)| {
                path.len() == 2 && path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
            })
            .filter_map(|(path, _)| path[1].parse().ok())
            .collect();
        if existing.len() >= MAX_SIGNATURES {
            return Err(BinaryOfficeSignatureError::ResourceLimit(
                "too many signatures".into(),
            ));
        }
        let mut name = rand::random::<u64>();
        while existing.contains(&name) {
            name = name.wrapping_add(1);
        }
        if !self
            .storages
            .iter()
            .any(|path| path.len() == 1 && path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE))
        {
            self.storages.push(vec![XML_SIGNATURE_STORAGE.into()]);
        }
        self.streams
            .push((vec![XML_SIGNATURE_STORAGE.into(), name.to_string()], xml));
        self.changed = true;
        Ok(self)
    }

    pub fn resign(&mut self, signer: &PackageSigner) -> Result<&mut Self> {
        self.clear();
        self.add_signature(signer)
    }

    pub fn clear(&mut self) -> &mut Self {
        let stream_count = self.streams.len();
        self.streams.retain(|(path, _)| {
            !(path
                .first()
                .is_some_and(|name| name.eq_ignore_ascii_case(XML_SIGNATURE_STORAGE))
                || path.len() == 1 && path[0].eq_ignore_ascii_case(LEGACY_SIGNATURE_STREAM))
        });
        self.storages.retain(|path| {
            !path
                .first()
                .is_some_and(|name| name.eq_ignore_ascii_case(XML_SIGNATURE_STORAGE))
        });
        self.changed |= self.streams.len() != stream_count;
        self
    }

    pub fn finish(&self) -> Result<Vec<u8>> {
        if !self.changed {
            return Ok(self.original.clone());
        }
        self.finish_snapshot()
    }

    fn finish_snapshot(&self) -> Result<Vec<u8>> {
        if !self.changed {
            return Ok(self.original.clone());
        }
        let mut writer = OleWriter::with_sector_size(self.sector_size);
        if self.root_clsid != [0; 16] {
            writer.set_root_clsid(self.root_clsid);
        }
        for path in &self.storages {
            let borrowed: Vec<&str> = path.iter().map(String::as_str).collect();
            writer.create_storage(&borrowed)?;
        }
        for (path, data) in &self.streams {
            let borrowed: Vec<&str> = path.iter().map(String::as_str).collect();
            writer.create_stream(&borrowed, data)?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }
}

fn signed_references<R: Read + Seek>(
    ole: &mut OleFile<R>,
    format: BinaryOfficeFormat,
) -> Result<Vec<DetachedSignatureReference>> {
    let paths = ole.list_streams();
    if paths.len() > MAX_CFB_STREAMS {
        return Err(BinaryOfficeSignatureError::ResourceLimit(
            "too many compound-file streams".into(),
        ));
    }
    let mut streams = Vec::with_capacity(paths.len());
    let mut total = 0usize;
    for path in paths {
        let borrowed: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = ole.open_stream(&borrowed)?;
        total = total.checked_add(data.len()).ok_or_else(|| {
            BinaryOfficeSignatureError::ResourceLimit("stream byte count overflow".into())
        })?;
        if total > MAX_CFB_BYTES {
            return Err(BinaryOfficeSignatureError::ResourceLimit(
                "compound-file stream bytes exceed 1 GiB".into(),
            ));
        }
        streams.push((path, data));
    }
    references_from_streams(&streams, format)
}

fn references_from_streams(
    streams: &[(Vec<String>, Vec<u8>)],
    format: BinaryOfficeFormat,
) -> Result<Vec<DetachedSignatureReference>> {
    let mut references = Vec::new();
    for (path, original) in streams {
        if excluded(path) {
            continue;
        }
        let leaf = path.last().map(String::as_str).unwrap_or_default();
        let data = match format {
            BinaryOfficeFormat::Xls if leaf.eq_ignore_ascii_case("Workbook") => {
                filter_xls_write_access(original)?
            },
            BinaryOfficeFormat::Ppt if leaf.eq_ignore_ascii_case("Current User") => Vec::new(),
            _ => original.clone(),
        };
        references.push(DetachedSignatureReference {
            uri: encode_path(path),
            data,
        });
    }
    references.sort_by(|left, right| left.uri.cmp(&right.uri));
    Ok(references)
}

fn excluded(path: &[String]) -> bool {
    if path.iter().any(|component| {
        component.eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
            || component.eq_ignore_ascii_case("Xmlsignatures")
            || component.eq_ignore_ascii_case("MsoDataStore")
            || component == "\u{0006}DataSpaces"
            || component == "\u{0005}Bagaaqy23kudbhchAaq5u2chNd"
    }) {
        return true;
    }
    path.last().is_some_and(|leaf| {
        leaf.eq_ignore_ascii_case(LEGACY_SIGNATURE_STREAM)
            || leaf == "\u{0009}DRMContent"
            || leaf == "\u{0005}SummaryInformation"
            || leaf == "\u{0005}DocumentSummaryInformation"
    })
}

fn reject_encrypted<R: Read + Seek>(
    ole: &OleFile<R>,
    format: BinaryOfficeFormat,
    references: &[DetachedSignatureReference],
) -> Result<()> {
    if ole.directory_exists(&["\u{0006}DataSpaces"])
        || ole.exists(&["\u{0009}DRMContent"])
        || ole.exists(&["EncryptionInfo"])
        || ole.exists(&["EncryptedPackage"])
        || references_indicate_encryption(references, format)
    {
        Err(BinaryOfficeSignatureError::EncryptedDocument)
    } else {
        Ok(())
    }
}

fn reject_encrypted_streams(
    streams: &[(Vec<String>, Vec<u8>)],
    format: BinaryOfficeFormat,
    references: &[DetachedSignatureReference],
) -> Result<()> {
    if streams.iter().any(|(path, _)| {
        path.iter().any(|name| {
            name == "\u{0006}DataSpaces"
                || name == "\u{0009}DRMContent"
                || name.eq_ignore_ascii_case("EncryptionInfo")
                || name.eq_ignore_ascii_case("EncryptedPackage")
        })
    }) || references_indicate_encryption(references, format)
    {
        Err(BinaryOfficeSignatureError::EncryptedDocument)
    } else {
        Ok(())
    }
}

fn references_indicate_encryption(
    references: &[DetachedSignatureReference],
    format: BinaryOfficeFormat,
) -> bool {
    references.iter().any(|reference| match format {
        BinaryOfficeFormat::Doc if reference.uri.ends_with("/WordDocument") => reference
            .data
            .get(10..12)
            .is_some_and(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) & 0x0100 != 0),
        BinaryOfficeFormat::Xls if reference.uri.ends_with("/Workbook") => {
            contains_biff_record(&reference.data, 0x002f)
        },
        _ => false,
    })
}

fn contains_biff_record(data: &[u8], wanted: u16) -> bool {
    let mut offset = 0usize;
    while let Some(header) = data.get(offset..offset + 4) {
        let record = u16::from_le_bytes([header[0], header[1]]);
        let length = u16::from_le_bytes([header[2], header[3]]) as usize;
        let Some(next) = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(length))
        else {
            return false;
        };
        if next > data.len() {
            return false;
        }
        if record == wanted {
            return true;
        }
        offset = next;
    }
    false
}

fn filter_xls_write_access(data: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(data.len());
    let mut offset = 0usize;
    while offset < data.len() {
        let Some(header) = data.get(offset..offset + 4) else {
            if data[offset..].iter().all(|byte| *byte == 0) {
                output.extend_from_slice(&data[offset..]);
                break;
            }
            return Err(BinaryOfficeSignatureError::InvalidContainer(
                "truncated BIFF record header".into(),
            ));
        };
        let record = u16::from_le_bytes([header[0], header[1]]);
        let length = u16::from_le_bytes([header[2], header[3]]) as usize;
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(length))
            .filter(|end| *end <= data.len());
        let Some(end) = end else {
            if data[offset..].iter().all(|byte| *byte == 0) {
                output.extend_from_slice(&data[offset..]);
                break;
            }
            return Err(BinaryOfficeSignatureError::InvalidContainer(
                "truncated BIFF record data".into(),
            ));
        };
        output.extend_from_slice(header);
        if record != 0x005c {
            output.extend_from_slice(&data[offset + 4..end]);
        }
        offset = end;
    }
    Ok(output)
}

fn encode_path(path: &[String]) -> String {
    let mut output = String::new();
    for component in path {
        output.push('/');
        for byte in component.as_bytes() {
            if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~') {
                output.push(*byte as char);
            } else {
                const HEX: &[u8; 16] = b"0123456789ABCDEF";
                output.push('%');
                output.push(HEX[(byte >> 4) as usize] as char);
                output.push(HEX[(byte & 15) as usize] as char);
            }
        }
    }
    output
}

fn collect_storages<R: Read + Seek>(ole: &OleFile<R>) -> Result<Vec<Vec<String>>> {
    fn visit<R: Read + Seek>(
        ole: &OleFile<R>,
        parent: &mut Vec<String>,
        output: &mut Vec<Vec<String>>,
    ) -> Result<()> {
        let borrowed: Vec<&str> = parent.iter().map(String::as_str).collect();
        let children: Vec<(String, u8)> = ole
            .list_directory_entries(&borrowed)?
            .into_iter()
            .map(|entry| (entry.name.clone(), entry.entry_type))
            .collect();
        for (name, entry_type) in children {
            if entry_type == 1 {
                parent.push(name);
                output.push(parent.clone());
                visit(ole, parent, output)?;
                parent.pop();
            }
        }
        Ok(())
    }
    let mut output = Vec::new();
    visit(ole, &mut Vec::new(), &mut output)?;
    Ok(output)
}

fn parse_clsid(value: &str) -> Result<[u8; 16]> {
    if value.is_empty() {
        return Ok([0; 16]);
    }
    let compact = value.replace('-', "");
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(BinaryOfficeSignatureError::InvalidContainer(
            "invalid root CLSID".into(),
        ));
    }
    let mut canonical = [0u8; 16];
    for (index, slot) in canonical.iter_mut().enumerate() {
        *slot = u8::from_str_radix(&compact[index * 2..index * 2 + 2], 16).map_err(|_| {
            BinaryOfficeSignatureError::InvalidContainer("invalid root CLSID".into())
        })?;
    }
    canonical[0..4].reverse();
    canonical[4..6].reverse();
    canonical[6..8].reverse();
    Ok(canonical)
}
