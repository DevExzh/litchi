//! Workbook and CFB transaction layer for XLS OLE objects.

use super::super::codec::{parse_subrecords, ranges, u32_at};
use super::super::semantic::{
    EmbeddedObjectDraft, EmbeddedPayload, FormControl, ObjectMetadataEdit, OleObjectRecord,
    validate_compound_file_for_publication,
};
use super::super::validation::validate_picture_formula;
use super::super::{BOUNDSHEET, CFB_STREAM, CONTINUE, EOF, Limits, OBJ, TXO, invalid};
use crate::error::{Error, Result};
use crate::protection::{
    FILESHARING_TYPE, OBJECTPROTECT_TYPE, PASSWORD_TYPE, PROT4REV_TYPE, PROT4REVPASS_TYPE,
    PROTECT_TYPE, SCENPROTECT_TYPE, WINPROTECT_TYPE, WRITEPROTECT_TYPE,
};
use litchi_cfb::OleFile;
use litchi_ole_common::object::{Editor as ObjectEditor, Target, Targets};
use litchi_ole_common::property_set::document_summary::DIGITAL_SIGNATURE;
use litchi_ole_common::property_set::{Binding, PropertySetReader};
use std::collections::{HashMap, HashSet};
use std::io::Cursor;

/// Identity facts recovered from `Obj` records that the typed OLE/control
/// grammar could not classify.  Such records remain opaque and are retained
/// byte-for-byte.  Their recoverable IDs/storage references still participate
/// in admission checks; an unparseable identity blocks operations that would
/// otherwise change object topology.
#[derive(Clone, Default)]
struct OpaqueObjectIdentity {
    ids: HashSet<u16>,
    storages: HashSet<String>,
    present: bool,
    ambiguous: bool,
}

impl OpaqueObjectIdentity {
    fn note_id(&mut self, object_id: u16) -> Result<()> {
        if object_id == 0 {
            self.ambiguous = true;
            return Ok(());
        }
        self.ids
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("opaque Obj identity index"))?;
        if !self.ids.insert(object_id) {
            self.ambiguous = true;
        }
        Ok(())
    }

    fn note_storage(&mut self, storage: String) -> Result<()> {
        self.storages
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("opaque Obj storage identity index"))?;
        if !self.storages.insert(storage) {
            self.ambiguous = true;
        }
        Ok(())
    }

    fn require_topology_proof(&self) -> Result<()> {
        if self.ambiguous {
            return Err(Error::UnsafeEdit(
                "an opaque Obj has an unproven identity; the requested object-topology edit is refused".into(),
            ));
        }
        Ok(())
    }

    fn check_new_identity(&self, object_id: u16, storage: &str) -> Result<()> {
        self.require_topology_proof()?;
        if self.ids.contains(&object_id) {
            return Err(invalid(OBJ, "duplicate workbook object ID in opaque Obj"));
        }
        if self.storages.contains(storage) {
            return Err(invalid(
                OBJ,
                "duplicate workbook storage identity in opaque Obj",
            ));
        }
        Ok(())
    }

    fn check_storage_rewrite(&self, storage: &str) -> Result<()> {
        if self.ambiguous || self.storages.contains(storage) {
            return Err(Error::UnsafeEdit(
                "an opaque Obj may alias the selected storage; payload rewrite is refused".into(),
            ));
        }
        Ok(())
    }
}

#[derive(Clone)]
pub struct Editor {
    package: ObjectEditor,
    limits: Limits,
    workbook_path: Vec<String>,
    workbook: Vec<u8>,
    sheets: Vec<Vec<OleObjectRecord>>,
    form_controls: Vec<Vec<FormControl>>,
    opaque_identity: OpaqueObjectIdentity,
    /// Number of form-control Obj records already present in each source
    /// worksheet. Newly authored controls are appended at the worksheet EOF;
    /// existing controls remain in their original byte representation.
    preserved_control_counts: Vec<usize>,
}

impl Editor {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        // Workbook metadata is XLS-owned. Read and parse it before handing
        // the original CFB bytes to the neutral object editor so the target
        // catalog can be derived solely from Obj/FtPictFmla records.
        let (workbook_path, workbook) = read_workbook(&bytes, limits)?;
        let (sheets, form_controls, opaque_identity) = parse_workbook(&workbook)?;
        let targets = targets_for_sheets(&sheets)?;
        let package = ObjectEditor::open(bytes, targets, limits)?;
        let preserved_control_counts = form_controls.iter().map(Vec::len).collect();
        Ok(Self {
            package,
            limits,
            workbook_path,
            workbook,
            sheets,
            form_controls,
            opaque_identity,
            preserved_control_counts,
        })
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn objects(&self, worksheet: usize) -> Result<&[OleObjectRecord]> {
        self.sheets
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    /// Form controls (checkboxes, list boxes, scroll bars, ...) anchored in a
    /// worksheet, in Obj record order.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn form_controls(&self, worksheet: usize) -> Result<&[FormControl]> {
        self.form_controls
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    pub(crate) fn worksheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Adds a typed worksheet form-control Obj record transactionally.
    ///
    /// The operation authors only the BIFF `Obj` metadata. It does not load,
    /// instantiate, or execute any external control/runtime content. Unknown
    /// subrecords supplied by the caller are serialized unchanged, and all
    /// controls already present in the source worksheet remain byte-identical.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_form_control(&mut self, worksheet: usize, control: FormControl) -> Result<()> {
        control.validate()?;
        let object_id = control.object_id();
        self.opaque_identity.require_topology_proof()?;
        if self.opaque_identity.ids.contains(&object_id) {
            return Err(invalid(OBJ, "duplicate workbook object ID in opaque Obj"));
        }
        if self
            .sheets
            .iter()
            .flatten()
            .any(|value| value.object_id() == object_id)
            || self
                .form_controls
                .iter()
                .flatten()
                .any(|value| value.object_id() == object_id)
        {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
        let mut candidate = self.clone();
        candidate
            .form_controls
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?
            .push(control);
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add(
        &mut self,
        worksheet: usize,
        object: OleObjectRecord,
        compound_file: Vec<u8>,
    ) -> Result<()> {
        object.validate()?;
        let storage = object
            .storage_name()
            .ok_or_else(|| invalid(OBJ, "new Obj has no MBD/LNK reference"))?;
        self.opaque_identity
            .check_new_identity(object.object_id(), &storage)?;
        validate_compound_file_for_publication(&compound_file, self.limits)?;
        if self
            .sheets
            .iter()
            .flatten()
            .any(|value| value.object_id() == object.object_id())
            || self
                .form_controls
                .iter()
                .flatten()
                .any(|value| value.object_id() == object.object_id())
        {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
        self.sheets
            .get(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        self.ensure_storage_capacity(&storage)?;
        let mut candidate = self.clone();
        candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?
            .push(object);
        let target = target_for_storage(storage)?;
        candidate.package.add_storage(target, compound_file)?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Adds a storage-backed, inert embedded payload with its complete XLS
    /// identity closure.
    ///
    /// The draft authors one picture `Obj` (`cmo.ot=8`, `fDde=0`, and
    /// `fPrstm=0`) together with the matching `MBDxxxxxxxx` storage.  The
    /// supplied CFB remains opaque; common CFB validation bounds and retains
    /// all of its streams, while no OLE server, macro, control, link, or
    /// native payload is activated.
    ///
    /// # Errors
    ///
    /// Returns an error if the worksheet or object identity is invalid, the
    /// identity is already present, or the payload cannot be admitted by the
    /// bounded common CFB editor.
    pub fn add_embedded_payload(
        &mut self,
        worksheet: usize,
        draft: EmbeddedObjectDraft,
    ) -> Result<()> {
        self.sheets
            .get(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        self.opaque_identity
            .check_new_identity(draft.object_id(), &draft.storage_name())?;
        // Check the target cardinality before parsing/capturing the supplied
        // CFB.  The neutral add_storage API receives the target only after it
        // has captured the candidate, so the XLS owner must perform this
        // host-level preflight to keep max_objects an ingress bound.
        self.ensure_storage_capacity(&draft.storage_name())?;
        draft.payload().validate_for_publication(self.limits)?;
        self.add(worksheet, draft.object_record(), draft.payload().to_vec())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn remove(&mut self, worksheet: usize, object_id: u16) -> Result<OleObjectRecord> {
        self.opaque_identity.require_topology_proof()?;
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        let index = sheet
            .iter()
            .position(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        let removed = sheet.remove(index);
        if let Some(storage) = removed.storage_name()
            && !candidate
                .sheets
                .iter()
                .flatten()
                .any(|value| value.storage_name().as_deref() == Some(&storage))
        {
            if candidate.opaque_identity.storages.contains(&storage) {
                return Err(Error::UnsafeEdit(
                    "removing the selected Obj would orphan an opaque MBD/LNK reference".into(),
                ));
            }
            let target = target_for_storage(storage)?;
            candidate.package.remove_storage(target.key())?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(removed)
    }

    /// Removes one storage-backed embedded payload and its unreferenced MBD
    /// storage.
    ///
    /// The selected `Obj` and its selected storage are removed as one
    /// identity-checked operation.  Unknown streams and unrelated object
    /// storages remain untouched.  A linked/DDE or controls-stream object is
    /// refused because it is outside this inert embedded-payload API.
    ///
    /// # Errors
    ///
    /// Returns an error if the selected object is absent or is not a valid
    /// storage-backed embedded object.
    pub fn remove_embedded_payload(&mut self, worksheet: usize, object_id: u16) -> Result<()> {
        self.embedded_storage(worksheet, object_id)?;
        self.remove(worksheet, object_id).map(|_| ())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn reorder(&mut self, worksheet: usize, ids: &[u16]) -> Result<()> {
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        if ids.len() != sheet.len() {
            return Err(invalid(
                OBJ,
                "reorder must contain every worksheet OLE object",
            ));
        }
        if sheet
            .iter()
            .map(OleObjectRecord::object_id)
            .eq(ids.iter().copied())
        {
            return Ok(());
        }
        if candidate.opaque_identity.present {
            return Err(Error::UnsafeEdit(
                "reordering typed Obj records with opaque Obj records is not source-complete"
                    .into(),
            ));
        }
        let mut remaining = sheet.clone();
        let mut reordered = Vec::with_capacity(ids.len());
        for id in ids {
            let index = remaining
                .iter()
                .position(|value| value.object_id() == *id)
                .ok_or_else(|| invalid(OBJ, "unknown or repeated OLE object ID"))?;
            reordered.push(remaining.remove(index));
        }
        *sheet = reordered;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn update_object_metadata(
        &mut self,
        worksheet: usize,
        object_id: u16,
        edit: ObjectMetadataEdit,
    ) -> Result<()> {
        let object = self
            .sheets
            .get(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?
            .iter()
            .find(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        if edit.is_empty() {
            // Preserve malformed/producer-specific source bytes exactly for a
            // semantic no-op; no source rewrite is needed.
            return Ok(());
        }
        let mut probe = object.clone();
        edit.apply(&mut probe)?;
        if probe == *object {
            return Ok(());
        }
        self.opaque_identity.require_topology_proof()?;
        let mut candidate = self.clone();
        let sheet = candidate
            .sheets
            .get_mut(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        let object = sheet
            .iter_mut()
            .find(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        edit.apply(object)?;
        if candidate.opaque_identity.ids.contains(&object.object_id()) {
            return Err(invalid(OBJ, "duplicate workbook object ID in opaque Obj"));
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn replace_storage(&mut self, storage_name: &str, compound_file: Vec<u8>) -> Result<()> {
        let storage = self
            .sheets
            .iter()
            .flatten()
            .find_map(|value| {
                value
                    .storage_name()
                    .filter(|value| value.as_str() == storage_name)
            })
            .ok_or_else(|| invalid(OBJ, "storage has no Obj reference"))?;
        let target = target_for_storage(storage)?;
        // Let an exact replacement remain an inert no-op even when the
        // retained payload contains a signature.  A changed replacement is
        // admitted below under the explicit nested security policy.
        if self
            .package
            .objects()
            .get(target.key())
            .is_some_and(|object| object.compound() == compound_file.as_slice())
        {
            return Ok(());
        }
        self.opaque_identity.check_storage_rewrite(storage_name)?;
        validate_compound_file_for_publication(&compound_file, self.limits)?;
        self.package
            .replace(target.key(), compound_file)
            .map_err(Into::into)
    }

    /// Replaces one storage-backed embedded payload while retaining its XLS
    /// object ID and `MBDxxxxxxxx` storage identity.
    ///
    /// The replacement owns the selected storage's complete inert CFB
    /// subtree.  Unselected worksheet records, object identities, and
    /// unrelated CFB streams remain retained by the common bounded editor.
    /// The operation never activates OLE, macro, control, link, or native
    /// content and refuses DDE, control-stream, and malformed object shapes.
    ///
    /// # Errors
    ///
    /// Returns an error if the worksheet/object identity is absent or is not
    /// a storage-backed embedded object, or if the replacement CFB is invalid,
    /// protected, or exceeds the configured limits.
    pub fn replace_embedded_payload(
        &mut self,
        worksheet: usize,
        object_id: u16,
        payload: EmbeddedPayload,
    ) -> Result<()> {
        let storage = self.embedded_storage(worksheet, object_id)?;
        self.replace_storage(&storage, payload.to_vec())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(Into::into)
    }

    fn commit(&mut self) -> Result<()> {
        validate_entries(&self.sheets, &self.form_controls)?;
        let workbook = rewrite_workbook(
            &self.workbook,
            &self.sheets,
            &self.form_controls,
            &self.preserved_control_counts,
        )?;
        self.package
            .put_stream(&self.workbook_path, workbook.clone())?;
        self.workbook = workbook;
        self.preserved_control_counts = self.form_controls.iter().map(Vec::len).collect();
        Ok(())
    }

    fn embedded_storage(&self, worksheet: usize, object_id: u16) -> Result<String> {
        let object = self
            .sheets
            .get(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?
            .iter()
            .find(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        object.validate()?;
        let flags = object
            .subrecords
            .iter()
            .find_map(|value| match value {
                super::super::semantic::ObjSubrecord::PictureFlags(value) => Some(*value),
                _ => None,
            })
            .ok_or_else(|| invalid(OBJ, "embedded Obj has no FtPioGrbit"))?;
        if flags.is_dde() || flags.is_control() || flags.uses_control_stream() {
            return Err(invalid(
                OBJ,
                "embedded payload API requires an MBD storage-backed Obj",
            ));
        }
        let storage = object
            .storage_name()
            .ok_or_else(|| invalid(OBJ, "embedded Obj has no MBD storage reference"))?;
        if !storage.starts_with("MBD") {
            return Err(invalid(
                OBJ,
                "embedded payload API requires an MBD storage reference",
            ));
        }
        Ok(storage)
    }

    fn ensure_storage_capacity(&self, storage: &str) -> Result<()> {
        if self
            .sheets
            .iter()
            .flatten()
            .any(|value| value.storage_name().as_deref() == Some(storage))
        {
            return Ok(());
        }
        let count = targets_for_sheets(&self.sheets)?.len();
        if count >= self.limits.max_objects {
            return Err(Error::InvalidData(format!(
                "adding storage would exceed the configured object limit of {}",
                self.limits.max_objects
            )));
        }
        Ok(())
    }
}

pub(crate) fn read_workbook(bytes: &[u8], limits: Limits) -> Result<(Vec<String>, Vec<u8>)> {
    let max_size = limits.max_stream_size.min(limits.max_total_size);
    if max_size == 0 {
        return Err(Error::InvalidData(
            "Workbook stream limits must be non-zero".into(),
        ));
    }
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    let max_entries = limits
        .max_objects
        .saturating_mul(limits.max_storage_depth)
        .saturating_add(limits.max_streams);
    let mut encrypted = false;
    let mut document_summary_size = None;
    let mut workbook_stream = None;
    let mut book_stream = None;
    ole.visit_directory_entries(&[], max_entries, |ole, sid| {
        let entry = ole.directory_entry_by_sid(sid).ok_or_else(|| {
            litchi_cfb::OleError::CorruptedFile("directory visitor returned an unknown SID".into())
        })?;
        if entry.entry_type != CFB_STREAM {
            return Ok::<(), Error>(());
        }
        if entry.name.eq_ignore_ascii_case("encryption") {
            encrypted = true;
        }
        if document_summary_size.is_none()
            && entry
                .name
                .eq_ignore_ascii_case("\u{0005}DocumentSummaryInformation")
        {
            document_summary_size = Some(entry.size);
        }
        if workbook_stream.is_none() && entry.name.eq_ignore_ascii_case("Workbook") {
            workbook_stream = Some((entry.name.clone(), entry.size));
        } else if book_stream.is_none() && entry.name.eq_ignore_ascii_case("Book") {
            book_stream = Some((entry.name.clone(), entry.size));
        }
        Ok::<(), Error>(())
    })?;
    if encrypted {
        return Err(Error::PasswordProtected);
    }
    if let Some(size) = document_summary_size {
        if size > max_size {
            return Err(Error::InvalidData(
                "DocumentSummaryInformation exceeds configured read limit".into(),
            ));
        }
    }
    reject_document_summary_signature(&mut ole)?;
    if let Some((actual_name, declared_size)) = workbook_stream.or(book_stream) {
        if declared_size > max_size {
            return Err(Error::InvalidData(format!(
                "{actual_name} stream exceeds configured read limit"
            )));
        }
        let workbook = ole.open_stream(&[actual_name.as_str()])?;
        if workbook.len() as u64 > max_size {
            return Err(Error::InvalidData(format!(
                "{actual_name} stream exceeds configured read limit"
            )));
        }
        reject_protected_workbook(&workbook)?;
        return Ok((vec![actual_name], workbook));
    }
    Err(Error::InvalidData("Workbook stream not found".into()))
}

fn reject_document_summary_signature<R: std::io::Read + std::io::Seek>(
    ole: &mut OleFile<R>,
) -> Result<()> {
    match ole.property_set(Binding::DocumentSummaryInformation) {
        Ok(stream) => {
            if stream
                .sections
                .iter()
                .any(|section| section.property(DIGITAL_SIGNATURE).is_some())
            {
                return Err(Error::UnsafeEdit(
                    "DocumentSummaryInformation contains PIDDSI DigitalSignature; refusing a rewrite".into(),
                ));
            }
        },
        Err(litchi_cfb::OleError::StreamNotFound) => {},
        Err(error) => return Err(Error::Cfb(error)),
    }
    Ok(())
}

const FILEPASS_TYPE: u16 = 0x002F;

fn reject_protected_workbook(input: &[u8]) -> Result<()> {
    for (_, _, kind, body_start, body_end) in ranges(input)? {
        let body = &input[body_start..body_end];
        if kind == FILEPASS_TYPE {
            return Err(Error::PasswordProtected);
        }
        let active = match kind {
            PROTECT_TYPE | WINPROTECT_TYPE | OBJECTPROTECT_TYPE | SCENPROTECT_TYPE
            | PROT4REV_TYPE => parse_protection_bool(kind, body)?,
            PASSWORD_TYPE | PROT4REVPASS_TYPE => {
                if body.len() != 2 {
                    return Err(Error::InvalidLength {
                        expected: 2,
                        found: body.len(),
                    });
                }
                u16::from_le_bytes([body[0], body[1]]) != 0
            },
            WRITEPROTECT_TYPE => {
                if !body.is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: body.len(),
                    });
                }
                true
            },
            FILESHARING_TYPE => parse_file_sharing_active(body)?,
            _ => false,
        };
        if active {
            return Err(Error::UnsafeEdit(format!(
                "BIFF protection record 0x{kind:04X} is active"
            )));
        }
    }
    Ok(())
}

fn parse_protection_bool(kind: u16, body: &[u8]) -> Result<bool> {
    if body.len() != 2 {
        return Err(Error::InvalidLength {
            expected: 2,
            found: body.len(),
        });
    }
    match u16::from_le_bytes([body[0], body[1]]) {
        0 => Ok(false),
        1 => Ok(true),
        value => Err(invalid(
            kind,
            format!("protection Boolean must be 0 or 1, found 0x{value:04X}"),
        )),
    }
}

fn parse_file_sharing_active(body: &[u8]) -> Result<bool> {
    if body.len() < 6 {
        return Err(Error::InvalidLength {
            expected: 6,
            found: body.len(),
        });
    }
    // `fReadOnlyRec` is a user-interface recommendation. It does not make
    // the workbook protected or authorize refusing an otherwise writable
    // edit. Still validate the Boolean so malformed FILESHARING records do
    // not become an accidental permissive path.
    parse_protection_bool(FILESHARING_TYPE, &body[..2])?;
    let password = u16::from_le_bytes([body[2], body[3]]) != 0;
    let marker = u16::from_le_bytes([body[4], body[5]]);
    if !password {
        if marker != 0 || body.len() != 6 {
            return Err(invalid(
                FILESHARING_TYPE,
                "FILESHARING without a password has an invalid username marker",
            ));
        }
    } else {
        if marker > 54 {
            return Err(invalid(
                FILESHARING_TYPE,
                "FILESHARING username exceeds 54 characters",
            ));
        }
        let flags = *body.get(6).ok_or(Error::InvalidLength {
            expected: 7,
            found: body.len(),
        })?;
        if flags & !1 != 0 {
            return Err(invalid(
                FILESHARING_TYPE,
                "FILESHARING string flags are invalid",
            ));
        }
        let expected = 7usize
            .checked_add(usize::from(marker) * if flags == 1 { 2 } else { 1 })
            .ok_or_else(|| invalid(FILESHARING_TYPE, "FILESHARING length overflow"))?;
        if body.len() != expected {
            return Err(Error::InvalidLength {
                expected,
                found: body.len(),
            });
        }
    }
    // `wResPass` is the security-bearing part of FILESHARING. A set
    // `fReadOnlyRec` with no write-reservation password remains readable and
    // writable through this inert editor.
    Ok(password)
}

fn target_for_storage(storage: String) -> Result<Target> {
    Ok(Target::new(storage.clone(), [storage])?)
}

pub(crate) fn targets_for_sheets(sheets: &[Vec<OleObjectRecord>]) -> Result<Targets> {
    let mut seen = HashSet::new();
    let mut targets = Vec::new();
    for object in sheets.iter().flatten() {
        let Some(storage) = object.storage_name() else {
            continue;
        };
        if seen.insert(storage.clone()) {
            targets.push(target_for_storage(storage)?);
        }
    }
    Ok(Targets::new(targets)?)
}

#[allow(
    clippy::type_complexity,
    reason = "type mirrors the decoded BIFF record structure"
)]
fn parse_workbook(
    input: &[u8],
) -> Result<(
    Vec<Vec<OleObjectRecord>>,
    Vec<Vec<FormControl>>,
    OpaqueObjectIdentity,
)> {
    let (_, starts) = bindings(input)?;
    let mut sheets = Vec::new();
    let mut form_controls = Vec::new();
    let mut opaque_identity = OpaqueObjectIdentity::default();
    for (index, (start, worksheet)) in starts.iter().enumerate() {
        if !worksheet {
            continue;
        }
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        let (objects, controls, sheet_opaque) = parse_sheet(&input[*start..end])?;
        merge_opaque_identity(&mut opaque_identity, sheet_opaque)?;
        sheets.push(objects);
        form_controls.push(controls);
    }
    validate_entries(&sheets, &form_controls)?;
    for object in sheets.iter().flatten() {
        if opaque_identity.ids.contains(&object.object_id()) {
            opaque_identity.ambiguous = true;
        }
    }
    for control in form_controls.iter().flatten() {
        if opaque_identity.ids.contains(&control.object_id()) {
            opaque_identity.ambiguous = true;
        }
    }
    for object in sheets.iter().flatten() {
        if let Some(storage) = object.storage_name()
            && opaque_identity.storages.contains(&storage)
        {
            opaque_identity.ambiguous = true;
        }
    }
    Ok((sheets, form_controls, opaque_identity))
}

fn parse_sheet(
    input: &[u8],
) -> Result<(Vec<OleObjectRecord>, Vec<FormControl>, OpaqueObjectIdentity)> {
    let records = ranges(input)?;
    let mut objects = Vec::new();
    let mut controls = Vec::new();
    let mut opaque_identity = OpaqueObjectIdentity::default();
    for (index, value) in records.iter().enumerate() {
        if value.2 != OBJ {
            continue;
        }
        let txo = if records.get(index + 1).is_some_and(|next| next.2 == TXO) {
            if records
                .get(index + 2)
                .is_some_and(|next| next.2 == CONTINUE)
            {
                return Err(invalid(
                    TXO,
                    "Continue-based TxO beside OLE Obj is unsupported",
                ));
            }
            let next = records[index + 1];
            Some(input[next.0..next.1].to_vec())
        } else {
            None
        };
        let body = &input[value.3..value.4];
        if let Ok(object) = OleObjectRecord::parse(body, txo.clone()) {
            objects.push(object);
        } else if let Some(control) = FormControl::parse(body, txo) {
            controls.push(control);
        } else {
            opaque_identity.present = true;
            inspect_opaque_obj(body, &mut opaque_identity)?;
        }
    }
    Ok((objects, controls, opaque_identity))
}

fn merge_opaque_identity(
    destination: &mut OpaqueObjectIdentity,
    source: OpaqueObjectIdentity,
) -> Result<()> {
    destination.present |= source.present;
    destination.ambiguous |= source.ambiguous;
    destination
        .ids
        .try_reserve(source.ids.len())
        .map_err(|_error| Error::Allocation("opaque Obj identity index"))?;
    for object_id in source.ids {
        if !destination.ids.insert(object_id) {
            destination.ambiguous = true;
        }
    }
    destination
        .storages
        .try_reserve(source.storages.len())
        .map_err(|_error| Error::Allocation("opaque Obj storage identity index"))?;
    for storage in source.storages {
        if !destination.storages.insert(storage) {
            destination.ambiguous = true;
        }
    }
    Ok(())
}

fn inspect_opaque_obj(body: &[u8], identity: &mut OpaqueObjectIdentity) -> Result<()> {
    let Ok(subrecords) = parse_subrecords(body) else {
        identity.ambiguous = true;
        return Ok(());
    };
    // An opaque record can participate in identity admission only when its
    // complete ordinary embedding prefix is proven.  In particular, do not
    // infer an MBD/LNK name from an arbitrary formula tail: lPosInCtlStm is a
    // Ctls offset when fPrstm is set, and camera/DDE/control forms have
    // different FtPictFmla ownership rules.
    let Some(super::super::semantic::ObjSubrecord::Common(common)) = subrecords.first() else {
        identity.ambiguous = true;
        return Ok(());
    };
    let common_count = subrecords
        .iter()
        .filter(|value| matches!(value, super::super::semantic::ObjSubrecord::Common(_)))
        .count();
    if common_count != 1 || common.object_type != 8 || common.object_id == 0 {
        identity.ambiguous = true;
        return Ok(());
    }
    identity.note_id(common.object_id)?;
    if !matches!(
        subrecords.last(),
        Some(super::super::semantic::ObjSubrecord::End)
    ) {
        identity.ambiguous = true;
        return Ok(());
    }

    // The storage identity is safe to recover only from the normative
    // two-byte FtCf selector. A malformed FtCf body is retained as an
    // opaque ClipboardFormat and must not be treated as an ordinary
    // embedded-object prefix.
    let picture_formats = subrecords
        .iter()
        .filter(|value| {
            matches!(
                value,
                super::super::semantic::ObjSubrecord::PictureFormat(_)
            )
        })
        .count();
    let malformed_picture_format = subrecords.iter().any(|value| {
        matches!(
            value,
            super::super::semantic::ObjSubrecord::ClipboardFormat(_)
        )
    });
    if picture_formats != 1 || malformed_picture_format {
        identity.ambiguous = true;
        return Ok(());
    }

    let flags = subrecords
        .iter()
        .filter_map(|value| match value {
            super::super::semantic::ObjSubrecord::PictureFlags(value) => Some(*value),
            _ => None,
        })
        .collect::<Vec<_>>();
    let formulas = subrecords
        .iter()
        .filter_map(|value| match value {
            super::super::semantic::ObjSubrecord::PictureFormula(value) => Some(value),
            _ => None,
        })
        .collect::<Vec<_>>();
    if flags.len() != 1 || formulas.len() != 1 {
        identity.ambiguous = true;
        return Ok(());
    }
    let flags = flags[0];
    if flags.is_dde() || flags.is_control() || flags.uses_control_stream() || flags.camera_picture()
    {
        identity.ambiguous = true;
        return Ok(());
    }
    let formula = formulas[0];
    if validate_picture_formula(formula, flags).is_err() {
        identity.ambiguous = true;
        return Ok(());
    }
    let Some(position) = formula.storage_position else {
        identity.ambiguous = true;
        return Ok(());
    };
    identity.note_storage(format!("MBD{position:08X}"))?;
    Ok(())
}

fn validate_entries(
    sheets: &[Vec<OleObjectRecord>],
    form_controls: &[Vec<FormControl>],
) -> Result<()> {
    let mut ids = HashSet::new();
    let mut count = 0usize;
    for object in sheets.iter().flatten() {
        count += 1;
        if count > 4_096 {
            return Err(invalid(OBJ, "workbook object count exceeds limit"));
        }
        object.validate()?;
        if !ids.insert(object.object_id()) {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
    }
    for control in form_controls.iter().flatten() {
        count += 1;
        if count > 4_096 {
            return Err(invalid(OBJ, "workbook object count exceeds limit"));
        }
        if !ids.insert(control.object_id()) {
            return Err(invalid(OBJ, "duplicate workbook object ID"));
        }
    }
    Ok(())
}

fn rewrite_workbook(
    input: &[u8],
    sheets: &[Vec<OleObjectRecord>],
    form_controls: &[Vec<FormControl>],
    preserved_control_counts: &[usize],
) -> Result<Vec<u8>> {
    let (refs, starts) = bindings(input)?;
    let first = starts.first().map_or(input.len(), |value| value.0);
    let mut output = input[..first].to_vec();
    let mut new_offsets = HashMap::new();
    let mut worksheet = 0usize;
    for (index, (start, is_worksheet)) in starts.iter().enumerate() {
        let end = starts.get(index + 1).map_or(input.len(), |value| value.0);
        new_offsets.insert(*start, output.len());
        if *is_worksheet {
            output.extend_from_slice(&rewrite_sheet(
                &input[*start..end],
                sheets
                    .get(worksheet)
                    .ok_or_else(|| invalid(BOUNDSHEET, "worksheet list missing"))?,
                form_controls
                    .get(worksheet)
                    .ok_or_else(|| invalid(BOUNDSHEET, "form-control list missing"))?,
                *preserved_control_counts
                    .get(worksheet)
                    .ok_or_else(|| invalid(BOUNDSHEET, "form-control baseline missing"))?,
            )?);
            worksheet += 1;
        } else {
            output.extend_from_slice(&input[*start..end]);
        }
    }
    if worksheet != sheets.len() {
        return Err(invalid(BOUNDSHEET, "worksheet count mismatch"));
    }
    for (payload, old) in refs {
        let new = *new_offsets
            .get(&old)
            .ok_or_else(|| invalid(BOUNDSHEET, "sheet target missing"))?;
        output[payload..payload + 4].copy_from_slice(
            &u32::try_from(new)
                .map_err(|_error| invalid(BOUNDSHEET, "sheet offset exceeds u32"))?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn rewrite_sheet(
    input: &[u8],
    objects: &[OleObjectRecord],
    form_controls: &[FormControl],
    preserved_control_count: usize,
) -> Result<Vec<u8>> {
    let records = ranges(input)?;
    let mut output = Vec::new();
    let mut next = 0usize;
    let mut skip_txo = false;
    for (index, value) in records.iter().enumerate() {
        if skip_txo && value.2 == TXO {
            skip_txo = false;
            continue;
        }
        if value.2 == OBJ && OleObjectRecord::parse(&input[value.3..value.4], None).is_ok() {
            if let Some(object) = objects.get(next) {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object {
                    output.extend_from_slice(txo);
                }
                next += 1;
            }
            skip_txo = records
                .get(index + 1)
                .is_some_and(|following| following.2 == TXO);
            continue;
        }
        if value.2 == EOF {
            for object in &objects[next..] {
                output.extend_from_slice(&object.to_record_bytes()?);
                if let Some(txo) = &object.text_object {
                    output.extend_from_slice(txo);
                }
            }
            next = objects.len();
            for control in form_controls
                .get(preserved_control_count..)
                .ok_or_else(|| {
                    invalid(OBJ, "form-control baseline exceeds current control count")
                })?
            {
                output.extend_from_slice(&control.to_record_bytes()?);
                if let Some(txo) = &control.text_object {
                    output.extend_from_slice(txo);
                }
            }
        }
        output.extend_from_slice(&input[value.0..value.1]);
    }
    if next != objects.len() {
        return Err(invalid(EOF, "worksheet has no EOF"));
    }
    Ok(output)
}

#[allow(
    clippy::type_complexity,
    reason = "type mirrors the decoded BIFF record structure"
)]
fn bindings(input: &[u8]) -> Result<(Vec<(usize, usize)>, Vec<(usize, bool)>)> {
    let mut refs = Vec::new();
    for (start, _, kind, body_start, body_end) in ranges(input)? {
        if kind != BOUNDSHEET {
            continue;
        }
        let body = &input[body_start..body_end];
        if body.len() < 6 {
            return Err(invalid(BOUNDSHEET, "BoundSheet is truncated"));
        }
        refs.push((
            start
                .checked_add(4)
                .ok_or_else(|| invalid(BOUNDSHEET, "record offset overflow"))?,
            u32_at(body, 0).ok_or_else(|| invalid(BOUNDSHEET, "sheet offset is truncated"))?
                as usize,
            body[5] == 0,
        ));
    }
    let mut starts = refs
        .iter()
        .map(|(_, offset, sheet)| (*offset, *sheet))
        .collect::<Vec<_>>();
    starts.sort_by_key(|value| value.0);
    if starts.windows(2).any(|value| value[0].0 >= value[1].0)
        || starts.iter().any(|value| value.0 >= input.len())
    {
        return Err(invalid(BOUNDSHEET, "invalid or duplicate sheet offsets"));
    }
    Ok((
        refs.into_iter()
            .map(|(payload, offset, _)| (payload, offset))
            .collect(),
        starts,
    ))
}

#[cfg(test)]
mod tests {
    use super::OpaqueObjectIdentity;

    #[test]
    fn duplicate_opaque_storage_identity_is_ambiguous() {
        let mut identity = OpaqueObjectIdentity::default();
        identity
            .note_storage("MBD0000002A".to_string())
            .expect("first opaque storage should be indexed");
        identity
            .note_storage("MBD0000002A".to_string())
            .expect("duplicate opaque storage should be retained as a refusal");
        assert!(identity.ambiguous);
    }
}
