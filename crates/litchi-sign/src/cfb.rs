//! XMLDSig storage for legacy DOC, PPT, and XLS compound files.
//!
//! Rewriting preserves the sector size, root CLSID, every storage path, every
//! storage CLSID exposed by `litchi-cfb`, stream order, and stream bytes.
//! `litchi-cfb::DirectoryEntry` does not currently expose storage state bits or
//! creation/modification times, and `OleWriter` has no setters for them. A
//! changed file therefore cannot preserve non-zero values in those fields.
//! Clean `Editor::finish` returns the caller's original `Vec` allocation and
//! does not render the compound file.

use crate::xml::{self, Profile, Ref};
use crate::{Coverage, Error, Limits, Policy, Result, Signer, Status, Trust};
use litchi_cfb::{DirectoryEntry, OleFile, OleWriter};
use std::collections::HashSet;
use std::fmt;
use std::io::{Cursor, Read, Seek};
use std::rc::Rc;

const XML_SIGNATURE_STORAGE: &str = "_xmlsignatures";
const LEGACY_SIGNATURE_STREAM: &str = "_signatures";

/// Application-specific binary Office digest rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Format {
    Doc,
    Ppt,
    Xls,
}

/// Verification result for one decimal stream in `_xmlsignatures`.
#[derive(Debug, Clone)]
pub struct Report {
    stream: String,
    signature: crate::Report,
}

impl Report {
    pub fn stream(&self) -> &str {
        &self.stream
    }

    pub fn details(&self) -> &crate::Report {
        &self.signature
    }

    pub fn integrity(&self) -> Status {
        self.signature.integrity()
    }

    pub fn signature(&self) -> Status {
        self.signature.signature()
    }

    pub fn coverage(&self) -> Coverage {
        self.signature.coverage()
    }

    pub fn trust(&self) -> Trust {
        self.signature.trust()
    }

    pub fn uses_sha1(&self) -> bool {
        self.signature.uses_sha1()
    }

    pub fn time(&self) -> Option<&str> {
        self.signature.time()
    }
}

/// Discover and verify every XMLDSig stream without evaluating PKI trust.
pub fn verify<R: Read + Seek>(
    ole: &mut OleFile<R>,
    format: Format,
    policy: &Policy,
) -> Result<Vec<Report>> {
    if ole.exists(&[LEGACY_SIGNATURE_STREAM]) {
        return Err(Error::Legacy);
    }
    reject_obvious_encryption(ole)?;
    if !ole.directory_exists(&[XML_SIGNATURE_STORAGE]) {
        return Ok(Vec::new());
    }
    let signatures = signature_entries(ole, policy.limits())?;
    let snapshot = Snapshot::read(ole, format, policy.limits(), false)?;
    let borrowed: Vec<_> = signatures
        .iter()
        .map(|(number, name, bytes)| (*number, name.as_str(), bytes.as_slice()))
        .collect();
    verify_entries(&snapshot, &borrowed, policy)
}

fn signature_entries<R: Read + Seek>(
    ole: &mut OleFile<R>,
    limits: &Limits,
) -> Result<Vec<(u64, String, Vec<u8>)>> {
    let entries = ole.list_directory_entries(&[XML_SIGNATURE_STORAGE])?;
    if entries.is_empty() || entries.len() > limits.max_signatures() {
        return Err(Error::Limit(format!(
            "signature count {} is outside 1..={}",
            entries.len(),
            limits.max_signatures()
        )));
    }
    let mut numeric = HashSet::new();
    let mut names = Vec::new();
    names
        .try_reserve(entries.len())
        .map_err(|_| Error::Limit("signature entry allocation failed".into()))?;
    for entry in entries {
        if entry.entry_type != 2
            || entry.name.is_empty()
            || entry.name.len() > 20
            || !entry.name.bytes().all(|byte| byte.is_ascii_digit())
        {
            return Err(Error::Container(format!(
                "signature entry {:?} is not a decimal stream",
                entry.name
            )));
        }
        let number = entry
            .name
            .parse::<u64>()
            .map_err(|_| Error::Container("signature stream number overflow".into()))?;
        if !numeric.insert(number) {
            return Err(Error::Container(
                "duplicate numeric signature stream name".into(),
            ));
        }
        names.push((number, entry.name.clone()));
    }
    names.sort_by_key(|entry| entry.0);
    let mut output = Vec::new();
    output
        .try_reserve(names.len())
        .map_err(|_| Error::Limit("signature stream allocation failed".into()))?;
    for (number, name) in names {
        let bytes = ole.open_stream(&[XML_SIGNATURE_STORAGE, &name])?;
        if bytes.len() > limits.max_signature_bytes() {
            return Err(Error::Limit(format!(
                "signature stream {name} is too large"
            )));
        }
        output.push((number, name, bytes));
    }
    Ok(output)
}

fn verify_entries(
    snapshot: &Snapshot,
    signatures: &[(u64, &str, &[u8])],
    policy: &Policy,
) -> Result<Vec<Report>> {
    let references = snapshot.references()?;
    signatures
        .iter()
        .map(|(_, stream, bytes)| {
            let signature = xml::verify(Profile::Binary, bytes, &references[..], policy)?;
            Ok(Report {
                stream: (*stream).to_string(),
                signature,
            })
        })
        .collect()
}

#[derive(Debug)]
struct Storage {
    path: Vec<String>,
    clsid: [u8; 16],
}

#[derive(Debug)]
struct Stream {
    path: Vec<String>,
    data: Vec<u8>,
}

#[derive(Debug)]
struct Snapshot {
    format: Format,
    sector_size: usize,
    root_clsid: [u8; 16],
    storages: Vec<Storage>,
    streams: Vec<Stream>,
    changed: bool,
}

impl Snapshot {
    fn read<R: Read + Seek>(
        ole: &mut OleFile<R>,
        format: Format,
        limits: &Limits,
        include_signatures: bool,
    ) -> Result<Self> {
        let sector_size = ole.sector_size();
        let root_clsid = parse_clsid(
            ole.root_entry()
                .map(|entry| entry.clsid.as_str())
                .unwrap_or_default(),
        )?;
        let (storages, paths) = collect_entries(ole, limits)?;
        if paths.len() > limits.max_cfb_streams() {
            return Err(Error::Limit("too many compound-file streams".into()));
        }
        let mut streams = Vec::with_capacity(paths.len());
        let mut total = 0_usize;
        for path in paths {
            if !include_signatures && excluded(&path) {
                continue;
            }
            let borrowed: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&borrowed)?;
            total = total
                .checked_add(data.len())
                .ok_or_else(|| Error::Limit("compound stream byte count overflow".into()))?;
            if total > limits.max_cfb_bytes() {
                return Err(Error::Limit("compound stream bytes exceed policy".into()));
            }
            streams.push(Stream { path, data });
        }
        let snapshot = Self {
            format,
            sector_size,
            root_clsid,
            storages,
            streams,
            changed: false,
        };
        snapshot.reject_encryption()?;
        Ok(snapshot)
    }

    fn references(&self) -> Result<Vec<Ref<'_>>> {
        let mut output = Vec::new();
        // Empty storages have no stream path through which their name would be
        // covered. Represent them explicitly with a trailing slash and an empty
        // digest input; non-empty storage names are already covered by child URIs.
        for storage in &self.storages {
            if excluded(&storage.path)
                || self
                    .streams
                    .iter()
                    .any(|stream| stream.path.starts_with(&storage.path))
            {
                continue;
            }
            output.push(Ref::borrowed_uri(encode_path(&storage.path, true), &[])?);
        }
        for stream in &self.streams {
            if excluded(&stream.path) {
                continue;
            }
            let leaf = stream.path.last().map(String::as_str).unwrap_or_default();
            let uri = encode_path(&stream.path, false);
            let reference = match self.format {
                Format::Xls if leaf.eq_ignore_ascii_case("Workbook") => {
                    Ref::owned(uri, filter_xls_write_access(&stream.data)?)?
                },
                Format::Ppt if leaf.eq_ignore_ascii_case("Current User") => {
                    Ref::borrowed_uri(uri, &[])?
                },
                _ => Ref::borrowed_uri(uri, &stream.data)?,
            };
            output.push(reference);
        }
        output.sort_by(|left, right| left.uri().cmp(right.uri()));
        if output.windows(2).any(|pair| pair[0].uri() == pair[1].uri()) {
            return Err(Error::Container(
                "duplicate encoded CFB reference URI".into(),
            ));
        }
        Ok(output)
    }

    fn reject_encryption(&self) -> Result<()> {
        if self.streams.iter().any(|stream| {
            stream.path.iter().any(|name| {
                name == "\u{0006}DataSpaces"
                    || name == "\u{0009}DRMContent"
                    || name.eq_ignore_ascii_case("EncryptionInfo")
                    || name.eq_ignore_ascii_case("EncryptedPackage")
            })
        }) {
            return Err(Error::Encrypted);
        }
        for stream in &self.streams {
            let leaf = stream.path.last().map(String::as_str).unwrap_or_default();
            match self.format {
                Format::Doc if leaf.eq_ignore_ascii_case("WordDocument") => {
                    if stream
                        .data
                        .get(10..12)
                        .is_some_and(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) & 0x0100 != 0)
                    {
                        return Err(Error::Encrypted);
                    }
                },
                Format::Xls if leaf.eq_ignore_ascii_case("Workbook") => {
                    if contains_biff_record(&stream.data, 0x002f) {
                        return Err(Error::Encrypted);
                    }
                },
                Format::Doc | Format::Ppt | Format::Xls => {},
            }
        }
        Ok(())
    }

    fn clear(&mut self) {
        let streams = self.streams.len();
        let storages = self.storages.len();
        self.streams.retain(|stream| !signature_path(&stream.path));
        self.storages.retain(|storage| {
            !storage
                .path
                .first()
                .is_some_and(|name| name.eq_ignore_ascii_case(XML_SIGNATURE_STORAGE))
        });
        self.changed |= streams != self.streams.len() || storages != self.storages.len();
    }

    fn add(&mut self, signer: &Signer, limits: &Limits) -> Result<String> {
        if self.streams.iter().any(|stream| {
            stream.path.len() == 1 && stream.path[0].eq_ignore_ascii_case(LEGACY_SIGNATURE_STREAM)
        }) {
            return Err(Error::Legacy);
        }
        let references = self.references()?;
        let xml = xml::author(Profile::Binary, signer, &references, limits)?;
        let name = self.next_name(limits)?;
        if !self.storages.iter().any(|storage| {
            storage.path.len() == 1 && storage.path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
        }) {
            self.storages.push(Storage {
                path: vec![XML_SIGNATURE_STORAGE.into()],
                clsid: [0; 16],
            });
        }
        self.streams.push(Stream {
            path: vec![XML_SIGNATURE_STORAGE.into(), name.clone()],
            data: xml,
        });
        self.changed = true;
        Ok(name)
    }

    fn next_name(&self, limits: &Limits) -> Result<String> {
        let existing: HashSet<u64> = self
            .streams
            .iter()
            .filter(|stream| {
                stream.path.len() == 2 && stream.path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
            })
            .filter_map(|stream| stream.path[1].parse().ok())
            .collect();
        if existing.len() >= limits.max_signatures() {
            return Err(Error::Limit("too many binary Office signatures".into()));
        }
        (1_u64..)
            .find(|candidate| !existing.contains(candidate))
            .map(|candidate| candidate.to_string())
            .ok_or_else(|| Error::Limit("no decimal signature stream name is available".into()))
    }

    fn render(&self) -> Result<Vec<u8>> {
        let mut writer = OleWriter::with_sector_size(self.sector_size)?;
        if self.root_clsid != [0; 16] {
            writer.set_root_clsid(self.root_clsid);
        }
        for storage in &self.storages {
            let path: Vec<&str> = storage.path.iter().map(String::as_str).collect();
            writer.create_storage(&path)?;
            if storage.clsid != [0; 16] {
                writer.set_storage_clsid(&path, storage.clsid)?;
            }
        }
        for stream in &self.streams {
            let path: Vec<&str> = stream.path.iter().map(String::as_str).collect();
            writer.create_stream(&path, &stream.data)?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }
}

/// Move-first transactional editor for binary Office signatures.
pub struct Editor {
    original: Vec<u8>,
    format: Format,
    limits: Limits,
    snapshot: Option<Snapshot>,
}

impl fmt::Debug for Editor {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Editor")
            .field("format", &self.format)
            .field("bytes", &self.original.len())
            .field("materialized", &self.snapshot.is_some())
            .field(
                "changed",
                &self
                    .snapshot
                    .as_ref()
                    .is_some_and(|snapshot| snapshot.changed),
            )
            .finish()
    }
}

impl Editor {
    /// Consume a compound-file buffer and defer stream materialization until a
    /// mutation actually needs it.
    pub fn open(bytes: Vec<u8>, format: Format) -> Result<Self> {
        Self::with_limits(bytes, format, Limits::standard())
    }

    pub fn with_limits(bytes: Vec<u8>, format: Format, limits: Limits) -> Result<Self> {
        if bytes.len() > limits.max_cfb_bytes() {
            return Err(Error::Limit("compound file exceeds byte policy".into()));
        }
        let ole = OleFile::open(Cursor::new(bytes.as_slice()))?;
        reject_obvious_encryption(&ole)?;
        // Validate sector size and root metadata now, while retaining ownership
        // of the caller's allocation for an exact clean finish.
        let _ = ole.sector_size();
        let _ = parse_clsid(
            ole.root_entry()
                .map(|entry| entry.clsid.as_str())
                .unwrap_or_default(),
        )?;
        Ok(Self {
            original: bytes,
            format,
            limits,
            snapshot: None,
        })
    }

    /// Verify the current editor state without changing it.
    pub fn verify(&self, policy: &Policy) -> Result<Vec<Report>> {
        if self
            .snapshot
            .as_ref()
            .is_some_and(|snapshot| snapshot.changed)
        {
            self.verify_snapshot(policy)
        } else {
            let mut ole = OleFile::open(Cursor::new(self.original.as_slice()))?;
            verify(&mut ole, self.format, policy)
        }
    }

    /// Add one signature, preserving existing signatures only when they verify.
    pub fn add(&mut self, signer: &Signer) -> Result<&mut Self> {
        self.materialize()?;
        let existing = self.verify_snapshot(&Policy::strict().with_limits(self.limits.clone()))?;
        if existing.iter().any(|report| {
            report.integrity() != Status::Valid || report.signature() != Status::Valid
        }) {
            return Err(Error::Sign(
                "cannot add while an existing signature is invalid".into(),
            ));
        }
        let limits = self.limits.clone();
        self.snapshot_mut()?.add(signer, &limits)?;
        Ok(self)
    }

    /// Atomically replace all signatures.
    ///
    /// XML authoring and reference hashing complete before the old signature
    /// graph is removed. Any error leaves the editor byte-for-byte unchanged.
    pub fn resign(&mut self, signer: &Signer) -> Result<&mut Self> {
        self.materialize()?;
        let limits = self.limits.clone();
        let xml = {
            let snapshot = self.snapshot_ref()?;
            let references = snapshot.references()?;
            xml::author(Profile::Binary, signer, &references, &limits)?
        };
        let name = "1".to_string();
        let snapshot = self.snapshot_mut()?;
        snapshot.clear();
        if !snapshot.storages.iter().any(|storage| {
            storage.path.len() == 1 && storage.path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
        }) {
            snapshot.storages.push(Storage {
                path: vec![XML_SIGNATURE_STORAGE.into()],
                clsid: [0; 16],
            });
        }
        snapshot.streams.push(Stream {
            path: vec![XML_SIGNATURE_STORAGE.into(), name],
            data: xml,
        });
        snapshot.changed = true;
        Ok(self)
    }

    /// Remove both XMLDSig and unsupported legacy CryptoAPI signature storage.
    pub fn clear(&mut self) -> Result<&mut Self> {
        self.materialize()?;
        self.snapshot_mut()?.clear();
        Ok(self)
    }

    /// Consume the editor. A clean editor returns the exact input allocation.
    pub fn finish(self) -> Result<Vec<u8>> {
        match self.snapshot {
            Some(snapshot) if snapshot.changed => snapshot.render(),
            Some(_) | None => Ok(self.original),
        }
    }

    fn materialize(&mut self) -> Result<()> {
        if self.snapshot.is_some() {
            return Ok(());
        }
        let mut ole = OleFile::open(Cursor::new(self.original.as_slice()))?;
        let snapshot = Snapshot::read(&mut ole, self.format, &self.limits, true)?;
        self.snapshot = Some(snapshot);
        Ok(())
    }

    fn snapshot_ref(&self) -> Result<&Snapshot> {
        self.snapshot
            .as_ref()
            .ok_or_else(|| Error::Container("editor snapshot is unavailable".into()))
    }

    fn snapshot_mut(&mut self) -> Result<&mut Snapshot> {
        self.snapshot
            .as_mut()
            .ok_or_else(|| Error::Container("editor snapshot is unavailable".into()))
    }

    fn verify_snapshot(&self, policy: &Policy) -> Result<Vec<Report>> {
        let snapshot = self.snapshot_ref()?;
        if snapshot.streams.iter().any(|stream| {
            stream.path.len() == 1 && stream.path[0].eq_ignore_ascii_case(LEGACY_SIGNATURE_STREAM)
        }) {
            return Err(Error::Legacy);
        }
        let has_storage = snapshot.storages.iter().any(|storage| {
            storage.path.len() == 1 && storage.path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
        });
        if snapshot.storages.iter().any(|storage| {
            storage.path.len() == 2 && storage.path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
        }) {
            return Err(Error::Container(
                "signature storage contains a nested storage".into(),
            ));
        }
        let mut signatures = Vec::new();
        let mut numeric = HashSet::new();
        for stream in &snapshot.streams {
            if stream.path.len() != 2 || !stream.path[0].eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
            {
                continue;
            }
            let name = &stream.path[1];
            if name.is_empty() || name.len() > 20 || !name.bytes().all(|byte| byte.is_ascii_digit())
            {
                return Err(Error::Container("invalid decimal signature stream".into()));
            }
            let number = name
                .parse::<u64>()
                .map_err(|_| Error::Container("signature stream number overflow".into()))?;
            if !numeric.insert(number) {
                return Err(Error::Container(
                    "duplicate numeric signature stream name".into(),
                ));
            }
            if stream.data.len() > policy.limits().max_signature_bytes() {
                return Err(Error::Limit(format!(
                    "signature stream {name} is too large"
                )));
            }
            signatures.push((number, name.as_str(), stream.data.as_slice()));
        }
        signatures.sort_by_key(|entry| entry.0);
        if signatures.is_empty() {
            return if has_storage {
                Err(Error::Limit("signature storage is empty".into()))
            } else {
                Ok(Vec::new())
            };
        }
        if !has_storage {
            return Err(Error::Container(
                "signature streams exist without their storage".into(),
            ));
        }
        if signatures.len() > policy.limits().max_signatures() {
            return Err(Error::Limit("too many binary Office signatures".into()));
        }
        verify_entries(snapshot, &signatures, policy)
    }
}

fn reject_obvious_encryption<R: Read + Seek>(ole: &OleFile<R>) -> Result<()> {
    if ole.directory_exists(&["\u{0006}DataSpaces"])
        || ole.exists(&["\u{0009}DRMContent"])
        || ole.exists(&["EncryptionInfo"])
        || ole.exists(&["EncryptedPackage"])
    {
        Err(Error::Encrypted)
    } else {
        Ok(())
    }
}

fn signature_path(path: &[String]) -> bool {
    path.first().is_some_and(|name| {
        name.eq_ignore_ascii_case(XML_SIGNATURE_STORAGE)
            || path.len() == 1 && name.eq_ignore_ascii_case(LEGACY_SIGNATURE_STREAM)
    })
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

fn contains_biff_record(data: &[u8], wanted: u16) -> bool {
    let mut offset = 0_usize;
    while let Some(header) = data.get(offset..offset.saturating_add(4)) {
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
    let mut offset = 0_usize;
    while offset < data.len() {
        let Some(header) = data.get(offset..offset.saturating_add(4)) else {
            if data[offset..].iter().all(|byte| *byte == 0) {
                output.extend_from_slice(&data[offset..]);
                break;
            }
            return Err(Error::Container("truncated BIFF record header".into()));
        };
        let record = u16::from_le_bytes([header[0], header[1]]);
        let length = u16::from_le_bytes([header[2], header[3]]) as usize;
        let Some(end) = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(length))
            .filter(|end| *end <= data.len())
        else {
            if data[offset..].iter().all(|byte| *byte == 0) {
                output.extend_from_slice(&data[offset..]);
                break;
            }
            return Err(Error::Container("truncated BIFF record data".into()));
        };
        output.extend_from_slice(header);
        if record != 0x005c {
            output.extend_from_slice(&data[offset + 4..end]);
        }
        offset = end;
    }
    Ok(output)
}

fn encode_path(path: &[String], storage: bool) -> String {
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
    if storage {
        output.push('/');
    }
    output
}

fn collect_entries<R: Read + Seek>(
    ole: &OleFile<R>,
    limits: &Limits,
) -> Result<(Vec<Storage>, Vec<Vec<String>>)> {
    enum Task<'a> {
        Directory(Vec<String>),
        Entry(Rc<[String]>, &'a DirectoryEntry),
    }

    let mut storages = Vec::new();
    let mut streams = Vec::new();
    let mut seen = 0_usize;
    let mut pending = vec![Task::Directory(Vec::new())];
    while let Some(task) = pending.pop() {
        match task {
            Task::Directory(parent) => {
                let path: Vec<&str> = parent.iter().map(String::as_str).collect();
                let children = ole.list_directory_entries(&path)?;
                seen = seen
                    .checked_add(children.len())
                    .ok_or_else(|| Error::Limit("CFB directory-entry count overflow".into()))?;
                if seen > limits.max_cfb_entries() {
                    return Err(Error::Limit("CFB directory entries exceed policy".into()));
                }
                pending
                    .try_reserve(children.len())
                    .map_err(|_| Error::Limit("CFB traversal allocation failed".into()))?;
                let parent: Rc<[String]> = parent.into();
                for child in children.into_iter().rev() {
                    pending.push(Task::Entry(Rc::clone(&parent), child));
                }
            },
            Task::Entry(parent, entry) => {
                let capacity = parent
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| Error::Limit("CFB path depth overflow".into()))?;
                let mut path = Vec::new();
                path.try_reserve(capacity)
                    .map_err(|_| Error::Limit("CFB path allocation failed".into()))?;
                path.extend(parent.iter().cloned());
                path.push(entry.name.clone());
                if path.len() > limits.max_cfb_depth() {
                    return Err(Error::Limit("CFB path depth exceeds policy".into()));
                }
                match entry.entry_type {
                    1 => {
                        storages
                            .try_reserve(1)
                            .map_err(|_| Error::Limit("CFB storage allocation failed".into()))?;
                        storages.push(Storage {
                            path: path.clone(),
                            clsid: parse_clsid(&entry.clsid)?,
                        });
                        pending.push(Task::Directory(path));
                    },
                    2 => {
                        streams
                            .try_reserve(1)
                            .map_err(|_| Error::Limit("CFB stream allocation failed".into()))?;
                        streams.push(path);
                    },
                    kind => {
                        return Err(Error::Container(format!(
                            "unexpected CFB directory entry type {kind}"
                        )));
                    },
                }
            },
        }
    }
    Ok((storages, streams))
}

fn parse_clsid(value: &str) -> Result<[u8; 16]> {
    if value.is_empty() {
        return Ok([0; 16]);
    }
    let compact = value.replace('-', "");
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::Container("invalid storage CLSID".into()));
    }
    let mut canonical = [0_u8; 16];
    for (index, slot) in canonical.iter_mut().enumerate() {
        *slot = u8::from_str_radix(&compact[index * 2..index * 2 + 2], 16)
            .map_err(|_| Error::Container("invalid storage CLSID".into()))?;
    }
    canonical[0..4].reverse();
    canonical[4..6].reverse();
    canonical[6..8].reverse();
    Ok(canonical)
}

#[cfg(test)]
mod tests {
    use super::*;
    use p256::ecdsa::SigningKey;

    fn signer() -> Signer {
        Signer::p256(SigningKey::from_bytes((&[7_u8; 32]).into()).unwrap())
            .time("2026-07-19T12:34:56Z")
            .unwrap()
    }

    fn ole(streams: &[(&[&str], &[u8])]) -> Vec<u8> {
        let mut writer = OleWriter::new();
        for (path, data) in streams {
            writer.create_stream(path, data).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn reports(bytes: &[u8]) -> Vec<Report> {
        let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
        verify(&mut file, Format::Doc, &Policy::strict()).unwrap()
    }

    #[test]
    fn signs_verifies_tampers_and_clears() {
        let original = ole(&[(&["Payload"], b"signed bytes")]);
        let mut editor = Editor::open(original, Format::Doc).unwrap();
        editor.add(&signer()).unwrap();
        let current = editor.verify(&Policy::strict()).unwrap();
        assert_eq!(current.len(), 1);
        assert_eq!(current[0].integrity(), Status::Valid);
        assert_eq!(current[0].signature(), Status::Valid);
        let signed = editor.finish().unwrap();
        let report = reports(&signed);
        assert_eq!(report.len(), 1);
        assert_eq!(report[0].integrity(), Status::Valid);
        assert_eq!(report[0].signature(), Status::Valid);

        let mut file = OleFile::open(Cursor::new(signed.as_slice())).unwrap();
        let signature_name = file
            .list_directory_entries(&[XML_SIGNATURE_STORAGE])
            .unwrap()
            .into_iter()
            .find(|entry| entry.entry_type == 2)
            .unwrap()
            .name
            .clone();
        let signature = file
            .open_stream(&[XML_SIGNATURE_STORAGE, &signature_name])
            .unwrap();
        let tampered = ole(&[
            (&["Payload"], b"tampered bytes"),
            (
                &[XML_SIGNATURE_STORAGE, signature_name.as_str()],
                signature.as_slice(),
            ),
        ]);
        let tampered_report = reports(&tampered);
        assert_eq!(tampered_report[0].integrity(), Status::Invalid);
        assert_eq!(tampered_report[0].signature(), Status::Valid);

        let mut editor = Editor::open(signed, Format::Doc).unwrap();
        editor.clear().unwrap();
        assert!(reports(&editor.finish().unwrap()).is_empty());
    }

    #[test]
    fn clean_finish_returns_the_original_allocation() {
        let original = ole(&[(&["Payload"], b"preserve")]);
        let pointer = original.as_ptr();
        let editor = Editor::open(original, Format::Doc).unwrap();
        let finished = editor.finish().unwrap();
        assert_eq!(finished.as_ptr(), pointer);
    }

    #[test]
    fn failed_resign_is_atomic() {
        let original = ole(&[(&["Payload"], b"preserve")]);
        let mut editor = Editor::open(original, Format::Doc).unwrap();
        editor.add(&signer()).unwrap();
        let signed = editor.finish().unwrap();

        let missing_time = Signer::p256(SigningKey::from_bytes((&[8_u8; 32]).into()).unwrap());
        let mut editor = Editor::open(signed.clone(), Format::Doc).unwrap();
        assert!(editor.resign(&missing_time).is_err());
        assert_eq!(editor.finish().unwrap(), signed);
    }

    #[test]
    fn default_add_rejects_partial_existing_coverage_atomically() {
        let original = ole(&[(&["Payload"], b"preserve")]);
        let mut editor = Editor::open(original, Format::Doc).unwrap();
        editor.add(&signer()).unwrap();
        let signed = editor.finish().unwrap();

        let mut editor = Editor::open(signed, Format::Doc).unwrap();
        editor.materialize().unwrap();
        let snapshot = editor.snapshot_mut().unwrap();
        snapshot.streams.push(Stream {
            path: vec!["Unsigned".into()],
            data: b"new bytes".to_vec(),
        });
        snapshot.changed = true;
        let before = snapshot.streams.len();
        assert!(editor.add(&signer()).is_err());
        assert_eq!(editor.snapshot_ref().unwrap().streams.len(), before);

        let bytes = editor.finish().unwrap();
        let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
        let reports = verify(&mut file, Format::Doc, &Policy::compatible()).unwrap();
        assert_eq!(reports.len(), 1);
        assert_eq!(reports[0].coverage(), Coverage::Partial);
    }

    #[test]
    fn cfb_directory_count_and_depth_are_bounded_before_mutation() {
        let crowded = ole(&[(&["One"], b"1"), (&["Two"], b"2"), (&["Three"], b"3")]);
        let limits = Limits::standard().cfb_entries(2).unwrap();
        let mut editor = Editor::with_limits(crowded, Format::Doc, limits).unwrap();
        assert!(matches!(editor.clear(), Err(Error::Limit(_))));

        let mut writer = OleWriter::new();
        writer.create_storage(&["A"]).unwrap();
        writer.create_storage(&["A", "B"]).unwrap();
        writer.create_storage(&["A", "B", "C"]).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let limits = Limits::standard().cfb_depth(2).unwrap();
        let mut editor = Editor::with_limits(output.into_inner(), Format::Doc, limits).unwrap();
        assert!(matches!(editor.clear(), Err(Error::Limit(_))));

        let signatures = ole(&[
            (&[XML_SIGNATURE_STORAGE, "1"], b"one"),
            (&[XML_SIGNATURE_STORAGE, "2"], b"two"),
        ]);
        let mut file = OleFile::open(Cursor::new(signatures)).unwrap();
        let policy = Policy::strict().with_limits(Limits::standard().signatures(1).unwrap());
        assert!(matches!(
            verify(&mut file, Format::Doc, &policy),
            Err(Error::Limit(_))
        ));
    }

    #[test]
    fn deterministic_names_and_storage_clsids_are_preserved() {
        let storage_clsid = [
            0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x46,
        ];
        let mut writer = OleWriter::new();
        writer.create_storage(&["Nested"]).unwrap();
        writer
            .set_storage_clsid(&["Nested"], storage_clsid)
            .unwrap();
        writer
            .create_stream(&["Nested", "Payload"], b"bytes")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();

        let mut editor = Editor::open(output.into_inner(), Format::Doc).unwrap();
        editor.add(&signer()).unwrap();
        editor.add(&signer()).unwrap();
        let signed = editor.finish().unwrap();
        let file = OleFile::open(Cursor::new(signed)).unwrap();
        let names: HashSet<String> = file
            .list_directory_entries(&[XML_SIGNATURE_STORAGE])
            .unwrap()
            .into_iter()
            .map(|entry| entry.name.clone())
            .collect();
        assert_eq!(names, HashSet::from(["1".into(), "2".into()]));
        let nested = file
            .list_directory_entries(&[])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Nested")
            .unwrap();
        assert_eq!(parse_clsid(&nested.clsid).unwrap(), storage_clsid);
    }
}
