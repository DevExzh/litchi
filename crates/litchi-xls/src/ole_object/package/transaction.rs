//! Workbook and CFB transaction layer for XLS OLE objects.

use super::super::LimitKind;
use super::super::codec::{parse_subrecords, ranges, u32_at};
use super::super::semantic::{
    EmbeddedObjectDraft, EmbeddedPayload, FormControl, ObjectMetadataChange, OleObjectRecord,
    validate_compound_file_for_publication,
};
use super::super::validation::validate_picture_formula;
use super::super::{BOUNDSHEET, CFB_STREAM, CONTINUE, Limits, OBJ, TXO, invalid};
use crate::error::{Error, Result};
use crate::protection::{
    FILESHARING_TYPE, OBJECTPROTECT_TYPE, PASSWORD_TYPE, PROT4REV_TYPE, PROT4REVPASS_TYPE,
    PROTECT_TYPE, SCENPROTECT_TYPE, WINPROTECT_TYPE, WRITEPROTECT_TYPE,
};
use litchi_cfb::OleFile;
use litchi_core::edit::RefusalKind;
use litchi_ole_common::object::{
    Edit as ObjectEditor, OpenOptions, Package as ObjectPackage, Target, Targets,
};
use litchi_ole_common::property_set::document_summary::DIGITAL_SIGNATURE;
use litchi_ole_common::property_set::{Binding, PropertySetReader};
use std::collections::{HashMap, HashSet};
use std::io::Cursor;

/// Identity facts recovered from `Obj` records that the typed OLE/control
/// grammar could not classify.  Such records remain opaque and are retained
/// byte-for-byte.  Their recoverable IDs/storage references still participate
/// in admission checks; an unparseable identity blocks operations that would
/// otherwise change object topology.
///
/// An `Obj` identifier is unique within its sheet ([MS-XLS] 2.5.143 FtCmo),
/// so identifiers are kept per worksheet; storages are workbook-wide.
#[derive(Clone, Default)]
struct OpaqueObjectIdentity {
    /// `(worksheet, Obj identifier)`; a single sheet's parse uses 0.
    ids: HashSet<(usize, u16)>,
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
        if !self.ids.insert((0, object_id)) {
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
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit OLE objects",
                "an opaque Obj has an unproven identity; the requested object-topology edit is refused",
            ));
        }
        Ok(())
    }

    fn check_storage_rewrite(&self, storage: &str) -> Result<()> {
        if self.ambiguous || self.storages.contains(storage) {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit OLE objects",
                "an opaque Obj may alias the selected storage; payload rewrite is refused",
            ));
        }
        Ok(())
    }
}

#[derive(Clone)]
pub(crate) struct Editor {
    package: ObjectEditor,
    limits: Limits,
    workbook_path: Vec<String>,
    workbook: Vec<u8>,
    sheets: Vec<Vec<OleObjectRecord>>,
    form_controls: Vec<Vec<FormControl>>,
    opaque_identity: OpaqueObjectIdentity,
}

impl Editor {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub(crate) fn new(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        // Workbook metadata is XLS-owned. Read and parse it before handing
        // the original CFB bytes to the neutral object editor so the target
        // catalog can be derived solely from Obj/FtPictFmla records.
        let (workbook_path, workbook) = read_workbook(&bytes, limits)?;
        let (sheets, form_controls, opaque_identity) = parse_workbook(&workbook)?;
        let targets = targets_for_sheets(&sheets)?;
        let package = ObjectPackage::open_with_options(
            bytes,
            OpenOptions::new().targets(targets).limits(limits),
        )?
        .edit()?;
        Ok(Self {
            package,
            limits,
            workbook_path,
            workbook,
            sheets,
            form_controls,
            opaque_identity,
        })
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub(crate) fn objects(&self, worksheet: usize) -> Result<&[OleObjectRecord]> {
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
    pub(crate) fn form_controls(&self, worksheet: usize) -> Result<&[FormControl]> {
        self.form_controls
            .get(worksheet)
            .map(Vec::as_slice)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))
    }

    pub(crate) fn worksheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Checks that a form control can be added to `worksheet`: it is
    /// valid, carries no text object, and the workbook's object identities
    /// are proven. The drawing-aware worksheet-feature edit then writes the
    /// control with a fresh identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the worksheet is absent, the control is invalid
    /// or has a text object, or an opaque `Obj` has an unproven identity.
    pub(crate) fn check_add_form_control(
        &self,
        worksheet: usize,
        control: &FormControl,
    ) -> Result<()> {
        self.sheets
            .get(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        control.validate()?;
        if control.text_object.is_some() {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit OLE objects",
                "a new form control's label is not written; add it without a text object",
            ));
        }
        self.opaque_identity.require_topology_proof()
    }

    /// The catalog position (`BoundSheet` order) of the worksheet at
    /// `worksheet` among the worksheets in substream order.
    pub(crate) fn sheet_position(&self, worksheet: usize) -> Result<usize> {
        let (refs, starts) = bindings(&self.workbook)?;
        let offset = starts
            .iter()
            .filter(|(_, is_worksheet)| *is_worksheet)
            .nth(worksheet)
            .map(|(offset, _)| *offset)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        refs.iter()
            .position(|(_, start)| *start == offset)
            .ok_or_else(|| invalid(BOUNDSHEET, "sheet target missing"))
    }

    /// Checks that a storage-backed embedded payload can be added to
    /// `worksheet`: the workbook's object identities are proven, one more
    /// storage target fits the limits, and the payload passes the bounded
    /// publication checks. The drawing-aware worksheet-feature edit then
    /// writes the object.
    ///
    /// # Errors
    ///
    /// Returns an error if the worksheet is absent, an opaque `Obj` has an
    /// unproven identity, the target count or a payload limit is exceeded,
    /// or the payload is protected.
    pub(crate) fn check_add_embedded_payload(
        &self,
        worksheet: usize,
        draft: &EmbeddedObjectDraft,
    ) -> Result<()> {
        self.sheets
            .get(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        self.opaque_identity.require_topology_proof()?;
        // Check the target cardinality before parsing the supplied CFB so
        // max_objects stays an ingress bound.
        let count = targets_for_sheets(&self.sheets)?.len();
        if count >= self.limits.max_objects {
            return Err(LimitKind::Objects.exceeded(
                count.saturating_add(1) as u64,
                self.limits.max_objects as u64,
            ));
        }
        draft.payload().validate_for_publication(self.limits)
    }

    /// Checks that the object `object_id` of `worksheet` can be removed and
    /// returns its typed record: the workbook's object identities are
    /// proven and no opaque `Obj` shares a storage the removal would drop.
    ///
    /// # Errors
    ///
    /// Returns an error if the object is absent, an opaque `Obj` has an
    /// unproven identity, or removing the storage would orphan an opaque
    /// reference.
    pub(crate) fn check_remove(&self, worksheet: usize, object_id: u16) -> Result<OleObjectRecord> {
        self.opaque_identity.require_topology_proof()?;
        let sheet = self
            .sheets
            .get(worksheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {worksheet}")))?;
        let object = sheet
            .iter()
            .find(|value| value.object_id() == object_id)
            .ok_or_else(|| invalid(OBJ, "OLE object ID not found"))?;
        if let Some(storage) = object.storage_name() {
            let owners = self
                .sheets
                .iter()
                .flatten()
                .filter(|value| value.storage_name().as_deref() == Some(&storage))
                .count();
            if owners == 1 && self.opaque_identity.storages.contains(&storage) {
                return Err(Error::refused(
                    RefusalKind::UnsupportedContent,
                    "edit OLE objects",
                    "removing the selected Obj would orphan an opaque MBD/LNK reference",
                ));
            }
        }
        Ok(object.clone())
    }

    /// Checks that the object is a storage-backed embedded object.
    ///
    /// # Errors
    ///
    /// Returns an error if the selected object is absent or is not a valid
    /// storage-backed embedded object.
    pub(crate) fn check_embedded(&self, worksheet: usize, object_id: u16) -> Result<()> {
        self.embedded_storage(worksheet, object_id).map(|_| ())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub(crate) fn update_object_metadata(
        &mut self,
        worksheet: usize,
        object_id: u16,
        edit: ObjectMetadataChange,
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
        if candidate
            .opaque_identity
            .ids
            .contains(&(worksheet, object.object_id()))
        {
            return Err(invalid(OBJ, "duplicate worksheet object ID in opaque Obj"));
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub(crate) fn replace_storage(
        &mut self,
        storage_name: &str,
        compound_file: Vec<u8>,
    ) -> Result<()> {
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
    pub(crate) fn replace_embedded_payload(
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
    pub(crate) fn finish(self) -> Result<Vec<u8>> {
        crate::scoped::object_editor_bytes(&self.package)
    }

    fn commit(&mut self) -> Result<()> {
        validate_entries(&self.sheets, &self.form_controls)?;
        let workbook = rewrite_workbook(&self.workbook, &self.sheets)?;
        self.package
            .put_stream(&self.workbook_path, workbook.clone())?;
        self.workbook = workbook;
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
}

pub(crate) fn read_workbook(bytes: &[u8], limits: Limits) -> Result<(Vec<String>, Vec<u8>)> {
    let max_size = limits.max_stream_size.min(limits.max_total_size);
    if max_size == 0 {
        return Err(if limits.max_stream_size == 0 {
            LimitKind::StreamBytes.invalid_zero()
        } else {
            LimitKind::AggregateBytes.invalid_zero()
        });
    }
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
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
            return Err(LimitKind::StreamBytes.exceeded(size, max_size));
        }
    }
    reject_document_summary_signature(&mut ole)?;
    if let Some((actual_name, declared_size)) = workbook_stream.or(book_stream) {
        if declared_size > max_size {
            return Err(LimitKind::StreamBytes.exceeded(declared_size, max_size));
        }
        let workbook = ole.open_stream(&[actual_name.as_str()])?;
        if workbook.len() as u64 > max_size {
            return Err(LimitKind::StreamBytes.exceeded(workbook.len() as u64, max_size));
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
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit OLE objects",
                    "DocumentSummaryInformation contains PIDDSI DigitalSignature; refusing a rewrite",
                ));
            }
        },
        Err(litchi_ole_common::Error::Cfb(litchi_cfb::OleError::StreamNotFound)) => {},
        Err(error) => return Err(error.into()),
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
            return Err(Error::refused(
                RefusalKind::Protected,
                "edit OLE objects",
                "BIFF protection record is active",
            ));
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
        merge_opaque_identity(&mut opaque_identity, sheet_opaque, sheets.len())?;
        sheets.push(objects);
        form_controls.push(controls);
    }
    validate_entries(&sheets, &form_controls)?;
    for (sheet, objects) in sheets.iter().enumerate() {
        if objects
            .iter()
            .any(|object| opaque_identity.ids.contains(&(sheet, object.object_id())))
        {
            opaque_identity.ambiguous = true;
        }
    }
    for (sheet, controls) in form_controls.iter().enumerate() {
        if controls
            .iter()
            .any(|control| opaque_identity.ids.contains(&(sheet, control.object_id())))
        {
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

/// Merge one worksheet's opaque identity (parsed as sheet 0) into the
/// workbook's, as worksheet `sheet`.
fn merge_opaque_identity(
    destination: &mut OpaqueObjectIdentity,
    source: OpaqueObjectIdentity,
    sheet: usize,
) -> Result<()> {
    destination.present |= source.present;
    destination.ambiguous |= source.ambiguous;
    destination
        .ids
        .try_reserve(source.ids.len())
        .map_err(|_error| Error::Allocation("opaque Obj identity index"))?;
    for (_, object_id) in source.ids {
        if !destination.ids.insert((sheet, object_id)) {
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
    // MS-XLS 2.4.181/2.5.149: a note has only FtCmo, FtNts and FtEnd.
    // Its GUID and shared-note flag cannot reference an MBD/LNK storage.
    // Retain it as opaque content while admitting its proven object ID.
    if common.object_type == 0x0019
        && common.object_id != 0
        && let [
            super::super::semantic::ObjSubrecord::Common(_),
            super::super::semantic::ObjSubrecord::Unknown { kind: 0x000D, data },
            super::super::semantic::ObjSubrecord::End,
        ] = subrecords.as_slice()
        && data.len() == 22
        && u16::from_le_bytes([data[16], data[17]]) <= 1
    {
        identity.note_id(common.object_id)?;
        return Ok(());
    }
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

/// Validate every typed object, bound the workbook's object count, and
/// require each `Obj` identifier to be unique within its worksheet
/// ([MS-XLS] 2.5.143 FtCmo).
fn validate_entries(
    sheets: &[Vec<OleObjectRecord>],
    form_controls: &[Vec<FormControl>],
) -> Result<()> {
    let mut count = 0usize;
    let empty = Vec::new();
    for (sheet, objects) in sheets.iter().enumerate() {
        let controls = form_controls.get(sheet).unwrap_or(&empty);
        let mut ids = HashSet::new();
        for object in objects {
            count += 1;
            if count > 4_096 {
                return Err(invalid(OBJ, "workbook object count exceeds limit"));
            }
            object.validate()?;
            if !ids.insert(object.object_id()) {
                return Err(invalid(OBJ, "duplicate worksheet object ID"));
            }
        }
        for control in controls {
            count += 1;
            if count > 4_096 {
                return Err(invalid(OBJ, "workbook object count exceeds limit"));
            }
            if !ids.insert(control.object_id()) {
                return Err(invalid(OBJ, "duplicate worksheet object ID"));
            }
        }
    }
    Ok(())
}

/// The workbook with every typed OLE `Obj` of its worksheets re-rendered in
/// place. Objects are only re-rendered here (their metadata may change);
/// adding or removing one goes through the drawing-aware worksheet-feature
/// edit, which keeps each `Obj` paired with its drawing shape.
fn rewrite_workbook(input: &[u8], sheets: &[Vec<OleObjectRecord>]) -> Result<Vec<u8>> {
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

fn rewrite_sheet(input: &[u8], objects: &[OleObjectRecord]) -> Result<Vec<u8>> {
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
            let object = objects
                .get(next)
                .ok_or_else(|| invalid(OBJ, "worksheet OLE object list is short"))?;
            output.extend_from_slice(&object.to_record_bytes()?);
            if let Some(txo) = &object.text_object {
                output.extend_from_slice(txo);
            }
            next += 1;
            skip_txo = records
                .get(index + 1)
                .is_some_and(|following| following.2 == TXO);
            continue;
        }
        output.extend_from_slice(&input[value.0..value.1]);
    }
    if next != objects.len() {
        return Err(invalid(OBJ, "worksheet OLE object list is long"));
    }
    Ok(output)
}

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

#[cfg(test)]
mod native_note_identity_tests {
    use super::*;
    fn note(shared: u16) -> Vec<u8> {
        let mut bytes = vec![0x15, 0, 18, 0, 25, 0, 1, 0, 0x11, 0x40];
        bytes.extend([0; 12]);
        bytes.extend([0x0d, 0, 22, 0]);
        bytes.extend([0xa5; 16]);
        bytes.extend(shared.to_le_bytes());
        bytes.extend([0xff; 4]); // undefined FtNts bytes are ignored
        bytes.extend([0; 4]);
        bytes
    }
    #[test]
    fn canonical_notes_contribute_ids_without_claiming_storage_ownership() {
        for shared in [0, 1] {
            let mut identities = OpaqueObjectIdentity::default();
            inspect_opaque_obj(&note(shared), &mut identities).unwrap();
            assert!(!identities.ambiguous);
            assert_eq!(identities.ids, HashSet::from([(0, 1)]));
            assert!(identities.storages.is_empty());
            identities.check_storage_rewrite("MBD00DD3649").unwrap();
            inspect_opaque_obj(&note(shared), &mut identities).unwrap();
            assert!(identities.require_topology_proof().is_err());
        }
    }
    #[test]
    fn incomplete_or_extended_notes_remain_unproven() {
        for case in 0..5 {
            let mut bytes = note(0);
            match case {
                0 => bytes[6..8].fill(0),
                1 => bytes[42..44].copy_from_slice(&2_u16.to_le_bytes()),
                2 => {
                    bytes.pop();
                },
                3 => {
                    bytes.splice(48..48, [0x34, 0x12, 1, 0, 0]);
                },
                _ => bytes[4] = 0x1e,
            }
            let mut identities = OpaqueObjectIdentity::default();
            inspect_opaque_obj(&bytes, &mut identities).unwrap();
            assert!(identities.ambiguous, "case {case}");
            assert!(identities.check_storage_rewrite("MBD00DD3649").is_err());
        }
    }
}
