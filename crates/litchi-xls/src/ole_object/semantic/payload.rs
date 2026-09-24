//! Inert, bounded embedded-payload authoring for MS-XLS.

use super::{FtCf, FtCmo, FtPictFmla, FtPioGrbit, ObjSubrecord, OleObjectRecord};
use crate::error::{Error, Result};
use litchi_cfb::{
    OleFile,
    consts::{STGTY_STORAGE, STGTY_STREAM},
};
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
    /// Creates an inert payload from a standalone CFB compound file.
    ///
    /// This constructor checks only the CFB container header and structural
    /// admission. The attached workbook operation applies its explicit object
    /// limits and rejects protected or malformed nested containers before
    /// publication.
    ///
    /// # Errors
    ///
    /// Returns an error when the input is empty, exceeds the default bounded
    /// object size, or is not a valid CFB compound file.
    pub fn new(compound_file: Vec<u8>) -> Result<Self> {
        Self::with_limits(compound_file, super::super::Limits::default())
    }

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
    pub fn with_limits(compound_file: Vec<u8>, limits: super::super::Limits) -> Result<Self> {
        if compound_file.is_empty() {
            return Err(Error::InvalidData(
                "embedded payload CFB must not be empty".into(),
            ));
        }
        if compound_file.len() as u64 > limits.max_object_size {
            return Err(Error::InvalidData(
                "embedded payload exceeds the configured object size limit".into(),
            ));
        }
        OleFile::open(Cursor::new(compound_file.as_slice()))?;
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
        return Err(Error::InvalidData(
            "embedded payload exceeds the configured object size limit".into(),
        ));
    }
    let max_streams = limits.max_streams.min(limits.max_streams_per_object);
    let max_total_size = limits.max_total_size.min(limits.max_object_size);
    if max_streams == 0 || max_total_size == 0 {
        return Err(Error::InvalidData(
            "embedded payload publication limits must be non-zero".into(),
        ));
    }

    let mut ole = OleFile::open(Cursor::new(bytes))?;
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
            return Err(Error::UnsafeEdit(
                "embedded payload contains signed, encrypted, or DRM content; refusing publication"
                    .into(),
            ));
        }
        if budget.entries >= budget.max_entries {
            return Err(Error::InvalidData(
                "embedded payload directory entry count exceeds the configured limit".into(),
            ));
        }
        budget.entries += 1;
        match entry_type {
            STGTY_STORAGE => {
                if budget.storages >= budget.max_storages {
                    return Err(Error::InvalidData(
                        "embedded payload storage count exceeds the configured limit".into(),
                    ));
                }
                let next_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidData("embedded payload storage depth overflow".into())
                })?;
                if next_depth > max_storage_depth {
                    return Err(Error::InvalidData(
                        "embedded payload storage depth exceeds the configured limit".into(),
                    ));
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
                    return Err(Error::InvalidData(
                        "embedded payload stream count exceeds the configured limit".into(),
                    ));
                }
                if size > budget.max_stream_size {
                    return Err(Error::InvalidData(
                        "embedded payload stream exceeds the configured size limit".into(),
                    ));
                }
                let total = budget.total_size.checked_add(size).ok_or_else(|| {
                    Error::InvalidData("embedded payload stream size overflow".into())
                })?;
                if total > budget.max_total_size {
                    return Err(Error::InvalidData(
                        "embedded payload stream bytes exceed the configured total limit".into(),
                    ));
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
                        return Err(Error::UnsafeEdit(
                            "embedded payload contains PIDDSI DigitalSignature; refusing a rewritten payload".into(),
                        ));
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
            ..
        }) => Error::InvalidData(
            "embedded payload directory entry count exceeds the configured limit".into(),
        ),
        other => other,
    })
}

/// A constrained new XLS embedded-object request.
///
/// The request binds one nonzero BIFF object identity, one `MBDxxxxxxxx`
/// storage identity, and one inert payload. It cannot describe a DDE link,
/// controls-stream object, or arbitrary Obj subrecord graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedObjectDraft {
    object_id: u16,
    storage_position: u32,
    payload: EmbeddedPayload,
}

impl EmbeddedObjectDraft {
    /// Creates a storage-backed embedded-object request.
    ///
    /// `storage_position` is rendered as the eight hexadecimal digits in the
    /// normative `MBDxxxxxxxx` CFB storage name from [MS-XLS] 2.1.7.5 and
    /// 2.5.150. The generated `Obj` uses `cmo.ot=8`, `fDde=0`, and
    /// `fPrstm=0`, and a normative `FtCf` unspecified clipboard format.
    ///
    /// # Errors
    ///
    /// Returns an error when `object_id` is zero.
    pub fn new(object_id: u16, storage_position: u32, payload: EmbeddedPayload) -> Result<Self> {
        if object_id == 0 {
            return Err(Error::InvalidRecord {
                record_type: super::super::OBJ,
                message: "embedded object ID must be nonzero".into(),
            });
        }
        Ok(Self {
            object_id,
            storage_position,
            payload,
        })
    }

    /// Returns the BIFF `FtCmo.id` to be authored.
    #[must_use]
    pub const fn object_id(&self) -> u16 {
        self.object_id
    }

    /// Returns the `FtPictFmla.lPosInCtlStm` value used for the MBD name.
    #[must_use]
    pub const fn storage_position(&self) -> u32 {
        self.storage_position
    }

    /// Returns the normative embedding-storage name.
    #[must_use]
    pub fn storage_name(&self) -> String {
        format!("MBD{:08X}", self.storage_position)
    }

    /// Returns the inert payload selected for publication.
    #[must_use]
    pub fn payload(&self) -> &EmbeddedPayload {
        &self.payload
    }

    pub(crate) fn object_record(&self) -> OleObjectRecord {
        OleObjectRecord {
            subrecords: vec![
                ObjSubrecord::Common(FtCmo {
                    object_type: 8,
                    object_id: self.object_id,
                    flags: 0,
                    reserved: [0; 12],
                }),
                // MS-XLS 2.4.181 requires pictFormat for an OLE object
                // (`cmo.ot=8`).  `0xFFFF` is the normative unspecified
                // clipboard-format selector when the payload has no chosen
                // presentation format.
                ObjSubrecord::PictureFormat(
                    FtCf::new(FtCf::UNSPECIFIED).expect("normative FtCf selector"),
                ),
                ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0 }),
                ObjSubrecord::PictureFormula(FtPictFmla {
                    // MS-XLS 2.5.150/2.5.187 require an even cbFmla for this
                    // storage form. The ObjFmla is an ObjectParsedFormula
                    // (cce=5 plus four unused bytes), one five-byte PtgTbl,
                    // and the minimal three-byte PictFmlaEmbedInfo.
                    formula: vec![
                        0x05, 0x00, // cce = 5
                        0x00, 0x00, 0x00, 0x00, // ObjectParsedFormula unused
                        0x02, 0x00, 0x00, 0x00, 0x00, // PtgTbl
                        0x03, 0x00, 0x00, // PictFmlaEmbedInfo (no class name)
                    ],
                    storage_position: Some(self.storage_position),
                    // cbBufInCtlStm is absent for fPrstm=0.  The codec also
                    // retains the legacy eight-byte form when reading older
                    // source objects, but new embedding authoring emits the
                    // normative four-byte lPosInCtlStm tail.
                    control_buffer_size: None,
                }),
                ObjSubrecord::End,
            ],
            text_object: None,
        }
    }
}
