//! Inert, bounded embedded-payload authoring for MS-XLS.

use super::super::LimitKind;
use crate::error::{Error, Result};
use litchi_cfb::{
    OleFile,
    consts::{STGTY_STORAGE, STGTY_STREAM},
};
use litchi_core::edit::RefusalKind;
use litchi_ole_common::property_set::PropertySetReader;
use litchi_ole_common::property_set::document_summary::DIGITAL_SIGNATURE;
use std::io::Cursor;
use std::sync::Arc;

/// A validated standalone CFB payload for one storage-backed embedded object.
///
/// The payload is retained as opaque bytes. Constructing or publishing it does
/// not inspect, activate, resolve, or execute any OLE server, macro, control,
/// link, or native data. The XLS package owner validates the complete bounded
/// CFB closure again when it is attached to a workbook.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedPayload {
    compound_file: Arc<[u8]>,
}

impl EmbeddedPayload {
    /// Creates an inert payload under an explicit retained-object byte limit.
    ///
    /// The payload remains opaque after this check. It is never interpreted as
    /// a file to open, a macro project to run, or a network target to resolve.
    /// `limits.max_object_size` bounds the standalone byte allocation here.
    /// Stream-count, stream-size, aggregate-size, and storage-depth limits are
    /// applied when the payload is attached to a workbook, where the common
    /// object editor has the enclosing CFB context needed to capture and
    /// validate the complete object closure.
    ///
    /// # Errors
    ///
    /// Returns an error when the input is empty, exceeds
    /// `limits.max_object_size`, or is not accepted by the bounded CFB parser.
    pub fn new(compound_file: Vec<u8>, limits: super::super::Limits) -> Result<Self> {
        if compound_file.is_empty() {
            return Err(Error::InvalidData(
                "embedded payload CFB must not be empty".into(),
            ));
        }
        if compound_file.len() as u64 > limits.max_object_size {
            return Err(
                LimitKind::ObjectBytes.exceeded(compound_file.len() as u64, limits.max_object_size)
            );
        }
        OleFile::open(
            Cursor::new(compound_file.as_slice()),
            litchi_cfb::OleFileLimits::default(),
        )?;
        Ok(Self {
            compound_file: Arc::from(compound_file),
        })
    }

    /// Borrows the exact standalone CFB bytes supplied by the caller.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.compound_file
    }

    pub(crate) fn to_vec(&self) -> Vec<u8> {
        self.compound_file.as_ref().to_vec()
    }

    /// Checks the limits that become meaningful when this standalone payload
    /// is captured as one workbook object. The common editor repeats these
    /// checks while publishing the enclosing CFB; this preflight makes the
    /// `add_embedded_payload` path apply the same per-object ceilings as a
    /// replacement path without retaining stream contents a second time.
    pub(crate) fn validate_for_publication(&self, limits: super::super::Limits) -> Result<()> {
        validate_compound_file_for_publication(self.as_bytes(), limits)
    }
}

/// Applies the same bounded/security admission to an opaque CFB supplied by
/// the legacy generic add/replace entry points. Keeping this check at the
/// payload boundary prevents those compatibility paths from bypassing the
/// embedded-payload policy merely because they accept raw bytes.
pub(crate) fn validate_compound_file_for_publication(
    bytes: &[u8],
    limits: super::super::Limits,
) -> Result<()> {
    // This is the retained standalone-object bound used by the common
    // `add_storage` ingress.  Check it before opening the CFB so directory or
    // property-set metadata cannot make a larger caller allocation reachable.
    if bytes.len() as u64 > limits.max_object_size {
        return Err(LimitKind::ObjectBytes.exceeded(bytes.len() as u64, limits.max_object_size));
    }
    let max_streams = limits.max_streams.min(limits.max_streams_per_object);
    let max_total_size = limits.max_total_size.min(limits.max_object_size);
    if max_streams == 0 {
        return Err(LimitKind::Streams.invalid_zero());
    }
    if max_total_size == 0 {
        return Err(LimitKind::AggregateBytes.invalid_zero());
    }

    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
    // A supplied payload is one selected object.  Match the common
    // `capture_object` ceiling rather than the aggregate package product, and
    // preflight it before nested capture so a wide fan-out cannot allocate
    // the common package vectors first.
    let max_storages = limits.max_storage_depth;
    let max_entries = max_storages.saturating_add(max_streams);
    let mut budget = PublicationBudget {
        streams: 0,
        total_size: 0,
        storages: 0,
        max_streams,
        max_stream_size: limits.max_stream_size,
        max_total_size,
        entries: 0,
        max_entries,
        max_storages,
    };
    inspect_publication_directory(&mut ole, &[], 0, limits.max_storage_depth, &mut budget)
}

struct PublicationBudget {
    streams: usize,
    total_size: u64,
    storages: usize,
    max_streams: usize,
    max_stream_size: u64,
    max_total_size: u64,
    entries: usize,
    max_entries: usize,
    max_storages: usize,
}

fn is_nested_protected_component(name: &str) -> bool {
    name.eq_ignore_ascii_case("encryption")
        || [
            "_xmlsignatures",
            "_signatures",
            "DigitalSignature",
            "\u{0005}DigitalSignature",
            "\u{0006}DataSpaces",
            "\u{0006}DataSpaceInfo",
            "\u{0006}TransformInfo",
            "\u{0006}Primary",
            "\u{0009}DRMContent",
            "\u{0009}DRMViewerContent",
            "EncryptedPackage",
            "EncryptionInfo",
        ]
        .iter()
        .any(|marker| name.eq_ignore_ascii_case(marker))
}

fn inspect_publication_directory<R: std::io::Read + std::io::Seek>(
    ole: &mut OleFile<R>,
    path: &[String],
    depth: usize,
    max_storage_depth: usize,
    budget: &mut PublicationBudget,
) -> Result<()> {
    let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
    ole.visit_directory_entries(&refs, budget.max_entries, |ole, sid| {
        let (entry_type, name, size) = {
            let entry = ole.directory_entry_by_sid(sid).ok_or_else(|| {
                Error::Cfb(litchi_cfb::OleError::CorruptedFile(
                    "directory visitor returned an unknown SID".into(),
                ))
            })?;
            (entry.entry_type, entry.name.clone(), entry.size)
        };
        if is_nested_protected_component(&name) {
            return Err(Error::refused(RefusalKind::UnsupportedOperation, "edit OLE objects", "embedded payload contains signed, encrypted, or DRM content; refusing publication"));
        }
        if budget.entries >= budget.max_entries {
            return Err(LimitKind::DirectoryEntries.exceeded(budget.entries.saturating_add(1) as u64, budget.max_entries as u64));
        }
        budget.entries += 1;
        match entry_type {
            STGTY_STORAGE => {
                if budget.storages >= budget.max_storages {
                    return Err(LimitKind::Storages.exceeded(budget.storages.saturating_add(1) as u64, budget.max_storages as u64));
                }
                let next_depth = depth.checked_add(1).ok_or_else(|| {
                    LimitKind::StorageDepth.exceeded(u64::MAX, max_storage_depth as u64)
                })?;
                if next_depth > max_storage_depth {
                    return Err(LimitKind::StorageDepth.exceeded(next_depth as u64, max_storage_depth as u64));
                }
                budget.storages += 1;
                let mut child = Vec::new();
                child
                    .try_reserve(path.len().saturating_add(1))
                    .map_err(|_| Error::Allocation("embedded payload storage path"))?;
                child.extend(path.iter().cloned());
                child.push(name);
                inspect_publication_directory(ole, &child, next_depth, max_storage_depth, budget)?;
            },
            STGTY_STREAM => {
                if budget.streams >= budget.max_streams {
                    return Err(LimitKind::Streams.exceeded(budget.streams.saturating_add(1) as u64, budget.max_streams as u64));
                }
                if size > budget.max_stream_size {
                    return Err(LimitKind::StreamBytes.exceeded(size, budget.max_stream_size));
                }
                let total = budget.total_size.checked_add(size).ok_or_else(|| {
                    LimitKind::AggregateBytes.exceeded(u64::MAX, budget.max_total_size)
                })?;
                if total > budget.max_total_size {
                    return Err(LimitKind::AggregateBytes.exceeded(total, budget.max_total_size));
                }
                budget.streams += 1;
                budget.total_size = total;

                // A PIDDSI can occur below any child storage.  Do not parse
                // it until its declared stream and aggregate sizes have been
                // admitted above; property-set parsing is an allocation-bearing
                // operation and must never be the first resource check.
                if name.eq_ignore_ascii_case("\u{0005}DocumentSummaryInformation") {
                    let mut child = Vec::new();
                    child
                        .try_reserve(path.len().saturating_add(1))
                        .map_err(|_| Error::Allocation("embedded payload property-set path"))?;
                    child.extend(path.iter().cloned());
                    child.push(name);
                    let stream = ole.property_set_stream(
                        &child.iter().map(String::as_str).collect::<Vec<_>>(),
                    )?;
                    if stream
                        .sections
                        .iter()
                        .any(|section| section.property(DIGITAL_SIGNATURE).is_some())
                    {
                        return Err(Error::refused(RefusalKind::UnsupportedOperation, "edit OLE objects", "embedded payload contains PIDDSI DigitalSignature; refusing a rewritten payload"));
                    }
                }
            },
            _ => {},
        }
        Ok(())
    })
    .map_err(|error| match error {
        Error::Cfb(litchi_cfb::OleError::LimitExceeded {
            resource: "directory entries",
            observed, maximum
        }) => LimitKind::DirectoryEntries.exceeded(observed, maximum),
        other => other,
    })
}

/// A constrained new XLS embedded-object request.
///
/// The request binds one inert payload to the drawing frame Excel shows it
/// in: a cell anchor and a preview picture. Adding it writes the frame
/// shape, its `Obj` (`cmo.ot=8`, `fDde=0`, `fPrstm=0`, a normative `FtCf`
/// unspecified clipboard format) and a new `MBDxxxxxxxx` storage
/// ([MS-XLS] 2.4.181 Obj, 2.5.150 FtPictFmla). The identities are
/// allocated when the object is added: the `Obj` identifier is unique
/// within its worksheet (as [MS-XLS] 2.5.143 requires) and the storage name
/// within the compound file. The request cannot describe a DDE link,
/// controls-stream object, or arbitrary Obj subrecord graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedObjectDraft {
    payload: EmbeddedPayload,
    anchor: crate::sheet_features::edit::Anchor,
    preview: Vec<u8>,
    prog_id: String,
}

impl EmbeddedObjectDraft {
    /// Creates a storage-backed embedded-object request shown at `anchor`
    /// with `preview` (PNG or JPEG) as its picture.
    #[must_use]
    pub fn new(
        payload: EmbeddedPayload,
        anchor: crate::sheet_features::edit::Anchor,
        preview: Vec<u8>,
    ) -> Self {
        Self {
            payload,
            anchor,
            preview,
            prog_id: String::new(),
        }
    }

    /// Names the object's ProgID (for example `Word.Document.8`) in its
    /// `Obj`; ASCII, at most 255 bytes.
    #[must_use]
    pub fn with_prog_id(mut self, prog_id: impl Into<String>) -> Self {
        self.prog_id = prog_id.into();
        self
    }

    /// Returns the inert payload selected for publication.
    #[must_use]
    pub fn payload(&self) -> &EmbeddedPayload {
        &self.payload
    }

    /// Returns the cell anchor of the object's frame.
    #[must_use]
    pub const fn anchor(&self) -> crate::sheet_features::edit::Anchor {
        self.anchor
    }

    /// Returns the preview picture.
    #[must_use]
    pub fn preview(&self) -> &[u8] {
        &self.preview
    }

    /// Returns the ProgID the `Obj` names; empty when it names none.
    #[must_use]
    pub fn prog_id(&self) -> &str {
        &self.prog_id
    }

    pub(crate) fn spec(&self) -> crate::sheet_features::edit::EmbeddedObject {
        crate::sheet_features::edit::EmbeddedObject {
            compound_file: self.payload.to_vec(),
            prog_id: self.prog_id.clone(),
            preview: self.preview.clone(),
            anchor: self.anchor,
        }
    }
}
