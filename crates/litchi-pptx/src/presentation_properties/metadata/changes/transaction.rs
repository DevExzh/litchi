//! Detached, source-checked edits for a Changes Information part.
//!
//! The document-change commands described by [MS-PPTX] and [MS-ODRAWXML] are
//! intentionally inert here. Typed author metadata and document-change bits
//! can be edited, while the command payloads, extension lists, and any
//! relationship-looking XML remain opaque bytes owned by this capability.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::model::{Data, Info, Kind, List};
use super::package;
use crate::{Error, Result};

/// Stable fingerprint of the exact Changes Information source bytes.
pub type Revision = u64;

/// An immutable typed snapshot of one Changes Information owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) presentation_part_name: String,
    pub(crate) presentation_content_type: String,
    pub(crate) relationship_id: String,
    pub(crate) relationship_target: String,
    pub(crate) part_name: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) info: Info,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load the Changes Information owner from an OPC package.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load(package: &OpcPackage) -> Result<Option<Self>> {
        package::load_snapshot(package)
    }

    /// Alias for `Self::load` emphasizing the source-bound result.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn read(package: &OpcPackage) -> Result<Option<Self>> {
        Self::load(package)
    }

    pub(crate) fn from_wire(
        presentation_part_name: String,
        presentation_content_type: String,
        relationship_id: String,
        relationship_target: String,
        part_name: String,
        source_xml: Arc<Vec<u8>>,
        info: Info,
    ) -> Result<Self> {
        if source_xml.len() > MAX_BYTES {
            return Err(limit("Changes Information source bytes"));
        }
        Ok(Self {
            presentation_part_name,
            presentation_content_type,
            relationship_id,
            relationship_target,
            part_name,
            revision: fingerprint(source_xml.as_slice()),
            source_xml,
            info,
        })
    }

    /// Return the typed Changes Information document.
    #[inline]
    #[must_use]
    pub fn info(&self) -> &Info {
        &self.info
    }

    /// Alias for [`Self::info`] for callers treating this as a semantic view.
    #[inline]
    #[must_use]
    pub fn changes_information(&self) -> &Info {
        self.info()
    }

    /// Return the exact OPC part name owned by this snapshot.
    #[inline]
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    /// Return the `PresentationML` part that owns the relationship.
    #[inline]
    #[must_use]
    pub fn presentation_part_name(&self) -> &str {
        &self.presentation_part_name
    }

    /// Return the source relationship ID.
    #[inline]
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the authored relationship target reference.
    #[inline]
    #[must_use]
    pub fn relationship_target(&self) -> &str {
        &self.relationship_target
    }

    /// Return the source `PresentationML` content type.
    #[inline]
    #[must_use]
    pub fn presentation_content_type(&self) -> &str {
        &self.presentation_content_type
    }

    /// Return the fingerprint used for stale-source checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Borrow the exact source bytes captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Start a detached edit over the typed view.
    #[inline]
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            working: self.info.clone(),
        }
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.presentation_part_name == other.presentation_part_name
            && self.presentation_content_type == other.presentation_content_type
            && self.relationship_id == other.relationship_id
            && self.relationship_target == other.relationship_target
            && self.part_name == other.part_name
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.revision == other.revision
    }
}

/// A detached atomic edit over one Changes Information snapshot.
#[derive(Clone, Debug)]
pub struct Transaction {
    original: Snapshot,
    working: Info,
}

impl Transaction {
    /// Borrow the currently staged typed document.
    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Info {
        &self.working
    }

    /// Whether the staged typed document differs from the source semantics.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.original.info != self.working
    }

    /// Replace the complete typed document after validating the candidate.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, value: Info) -> Result<bool> {
        validate_candidate(&value)?;
        if self.working == value {
            return Ok(false);
        }
        self.working = value;
        Ok(true)
    }

    /// Apply a checked mutation to a cloned typed document.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit(&mut self, edit: impl FnOnce(&mut Info) -> Result<()>) -> Result<()> {
        let mut candidate = self.working.clone();
        edit(&mut candidate)?;
        validate_candidate(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Replace or remove one reviewer's metadata list entry.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author(&mut self, list_index: usize, value: Option<Data>) -> Result<bool> {
        let mut candidate = self.working.clone();
        let list_len = candidate.change_lists.len();
        let list = candidate
            .change_lists
            .get_mut(list_index)
            .ok_or_else(|| list_index_error(list_index, list_len))?;
        if list.author == value {
            return Ok(false);
        }
        list.author = value;
        validate_candidate(&candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Apply an atomic typed mutation to one existing reviewer's metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit_author(
        &mut self,
        list_index: usize,
        edit: impl FnOnce(&mut Data) -> Result<()>,
    ) -> Result<()> {
        let mut candidate = self.working.clone();
        let list_len = candidate.change_lists.len();
        let list = candidate
            .change_lists
            .get_mut(list_index)
            .ok_or_else(|| list_index_error(list_index, list_len))?;
        let author = list
            .author
            .as_mut()
            .ok_or_else(|| invalid("Changes Information list has no reviewer metadata"))?;
        edit(author)?;
        validate_candidate(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Set one reviewer's display name, preserving all other metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_name(&mut self, list_index: usize, value: Option<String>) -> Result<bool> {
        self.update_author(list_index, |author| author.name = value)
    }

    /// Set one reviewer's user identifier, preserving all other metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_user_id(&mut self, list_index: usize, value: Option<String>) -> Result<bool> {
        self.update_author(list_index, |author| author.user_id = value)
    }

    /// Set one reviewer's identity-provider identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_provider_id(
        &mut self,
        list_index: usize,
        value: Option<String>,
    ) -> Result<bool> {
        self.update_author(list_index, |author| author.provider_id = value)
    }

    /// Set one reviewer's client/device identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_client_id(
        &mut self,
        list_index: usize,
        value: Option<String>,
    ) -> Result<bool> {
        self.update_author(list_index, |author| author.client_id = value)
    }

    /// Set one reviewer's e-mail address.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_email(&mut self, list_index: usize, value: Option<String>) -> Result<bool> {
        self.update_author(list_index, |author| author.email = value)
    }

    /// Set one reviewer's XML date-time.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_date_time(
        &mut self,
        list_index: usize,
        value: Option<String>,
    ) -> Result<bool> {
        self.update_author(list_index, |author| author.date_time = value)
    }

    /// Set one reviewer's application version.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_version(&mut self, list_index: usize, value: Option<u32>) -> Result<bool> {
        self.update_author(list_index, |author| author.version = value)
    }

    /// Set one reviewer's stable change GUID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_change_id(
        &mut self,
        list_index: usize,
        value: Option<String>,
    ) -> Result<bool> {
        self.update_author(list_index, |author| author.change_id = value)
    }

    /// Set one reviewer's user-action identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_author_action_id(&mut self, list_index: usize, value: Option<i32>) -> Result<bool> {
        self.update_author(list_index, |author| author.action_id = value)
    }

    /// Replace one document-change bit list while preserving its opaque
    /// command children and extension XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_change_kinds(
        &mut self,
        list_index: usize,
        change_index: usize,
        kinds: Vec<Kind>,
    ) -> Result<bool> {
        validate_kinds(&kinds)?;
        let mut candidate = self.working.clone();
        let list_len = candidate.change_lists.len();
        let list = candidate
            .change_lists
            .get_mut(list_index)
            .ok_or_else(|| list_index_error(list_index, list_len))?;
        let change_len = list.changes.len();
        let descriptor = list
            .changes
            .get_mut(change_index)
            .ok_or_else(|| list_index_error(change_index, change_len))?;
        if descriptor.change_kinds == kinds {
            return Ok(false);
        }
        descriptor.change_kinds = kinds;
        validate_candidate(&candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Append one reviewer/change list.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn push_list(&mut self, value: List) -> Result<()> {
        let mut candidate = self.working.clone();
        candidate.change_lists.push(value);
        validate_candidate(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Remove one reviewer/change list, returning its typed value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_list(&mut self, index: usize) -> Result<List> {
        let mut candidate = self.working.clone();
        if index >= candidate.change_lists.len() {
            return Err(list_index_error(index, candidate.change_lists.len()));
        }
        let removed = candidate.change_lists.remove(index);
        validate_candidate(&candidate)?;
        self.working = candidate;
        Ok(removed)
    }

    fn update_author(&mut self, list_index: usize, update: impl FnOnce(&mut Data)) -> Result<bool> {
        let before = self.working.clone();
        self.edit_author(list_index, |author| {
            update(author);
            Ok(())
        })?;
        Ok(self.working != before)
    }

    /// Validate and consume the edit into a source-checked commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.original.clone(), self.original.clone());
            return Ok(Commit {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }

        let encoded = encode_info(&self.working)?;
        let source_xml = Arc::new(encoded.to_xml()?);
        let info = Info::parse(source_xml.as_slice())?;
        if info != encoded {
            return Err(invalid(
                "Changes Information serialization changed the typed document",
            ));
        }
        let snapshot = Snapshot::from_wire(
            self.original.presentation_part_name.clone(),
            self.original.presentation_content_type.clone(),
            self.original.relationship_id.clone(),
            self.original.relationship_target.clone(),
            self.original.part_name.clone(),
            source_xml,
            info,
        )?;
        let patch = Patch::new(self.original, snapshot.clone());
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful Changes Information edit and its reversible package patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether publication changes the exact owner bytes.
    #[inline]
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Alias for [`Self::changed`].
    #[inline]
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Borrow the projected post-edit snapshot.
    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    /// Consume the commit into its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible source-checked replacement of one Changes Information blob.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

const MAX_BYTES: usize = 16 * 1024 * 1024;

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before publication.
    #[inline]
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source context produced by publication.
    #[inline]
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Alias for [`Self::is_empty`].
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return the source fingerprint required for publication.
    #[inline]
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    /// Apply this patch atomically after checking the complete owner source.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        let current = package::load_snapshot(target)?
            .ok_or_else(|| invalid("Changes Information source is absent"))?;
        if !current.same_source(&self.before) {
            return Err(invalid("Changes Information source is stale"));
        }
        if self.is_empty() {
            return Ok(current);
        }
        let mut candidate = target.clone();
        let snapshot = package::replace_snapshot(&mut candidate, &self.after)?;
        *target = candidate;
        Ok(snapshot)
    }
}

fn validate_candidate(value: &Info) -> Result<()> {
    let _ = encode_info(value)?;
    Ok(())
}

fn encode_info(value: &Info) -> Result<Info> {
    let mut normalized = value.clone();
    for list in &mut normalized.change_lists {
        for descriptor in &mut list.changes {
            validate_kinds(&descriptor.change_kinds)?;
            descriptor.xml = replace_change_kinds(&descriptor.xml, &descriptor.change_kinds)?;
        }
    }
    let bytes = normalized.to_xml()?;
    let parsed = Info::parse(&bytes)?;
    if parsed != normalized {
        return Err(invalid(
            "Changes Information typed values do not round-trip through XML",
        ));
    }
    Ok(normalized)
}

fn validate_kinds(kinds: &[Kind]) -> Result<()> {
    if kinds.is_empty() {
        return Err(invalid("document change bit list cannot be empty"));
    }
    for (index, kind) in kinds.iter().enumerate() {
        if kinds[..index].contains(kind) {
            return Err(invalid("document change bit list contains duplicates"));
        }
    }
    Ok(())
}

fn replace_change_kinds(source: &[u8], kinds: &[Kind]) -> Result<Vec<u8>> {
    if source.len() > MAX_BYTES {
        return Err(limit("document change descriptor bytes"));
    }
    let value = kinds
        .iter()
        .map(|kind| kind_token(*kind))
        .collect::<Vec<_>>()
        .join(" ");
    replace_root_attribute(source, b"chg", value.as_bytes())
}

fn replace_root_attribute(source: &[u8], wanted: &[u8], value: &[u8]) -> Result<Vec<u8>> {
    let end = opening_tag_end(source)?;
    let mut cursor = 1usize;
    while cursor < end && !is_xml_space(source[cursor]) && source[cursor] != b'/' {
        cursor += 1;
    }
    let mut found = None;
    while cursor < end {
        while cursor < end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || source[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < end
            && !is_xml_space(source[cursor])
            && !matches!(source[cursor], b'=' | b'/')
        {
            cursor += 1;
        }
        let name = &source[name_start..cursor];
        while cursor < end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || source[cursor] != b'=' {
            return Err(invalid("malformed document change attribute"));
        }
        cursor += 1;
        while cursor < end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || !matches!(source[cursor], b'"' | b'\'') {
            return Err(invalid("document change attribute is not quoted"));
        }
        let quote = source[cursor];
        cursor += 1;
        let value_start = cursor;
        while cursor < end && source[cursor] != quote {
            cursor += 1;
        }
        if cursor >= end {
            return Err(invalid("unterminated document change attribute"));
        }
        if name == wanted {
            if found.is_some() {
                return Err(invalid("duplicate document change attribute"));
            }
            found = Some(value_start..cursor);
        }
        cursor += 1;
    }
    let Some(range) = found else {
        return Err(invalid("document change descriptor is missing chg"));
    };
    let mut output = source.to_vec();
    output.splice(range, value.iter().copied());
    Ok(output)
}

fn opening_tag_end(source: &[u8]) -> Result<usize> {
    if source.first() != Some(&b'<') {
        return Err(invalid(
            "document change descriptor does not start with an element",
        ));
    }
    let mut quote = None;
    for (index, byte) in source.iter().copied().enumerate().skip(1) {
        match quote {
            Some(expected) if byte == expected => quote = None,
            None if matches!(byte, b'"' | b'\'') => quote = Some(byte),
            None if byte == b'>' => return Ok(index),
            Some(_) | None => {},
        }
    }
    Err(invalid(
        "document change descriptor has no complete opening tag",
    ))
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

fn kind_token(kind: Kind) -> &'static str {
    match kind {
        Kind::CustomSelection => "custSel",
        Kind::AddSlide => "addSld",
        Kind::DeleteSlide => "delSld",
        Kind::ModifySlide => "modSld",
        Kind::SlideOrder => "sldOrd",
        Kind::ModifyMainMaster => "modMainMaster",
        Kind::ModifyNotesMaster => "modNotesMaster",
        Kind::ModifyHandoutMaster => "modHandoutMaster",
        Kind::AddSection => "addSection",
        Kind::DeleteSection => "delSection",
        Kind::ModifySection => "modSection",
    }
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100_0000_01b3);
    }
    hash
}

fn list_index_error(index: usize, len: usize) -> Error {
    Error::Invalid(format!(
        "Changes Information list index {index} is outside a list of length {len}"
    ))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(label: &str) -> Error {
    invalid(format!("{label} exceed configured limit"))
}
