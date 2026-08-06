//! Lossless, contextual values used by the master inventory.

use crate::package::{Error, Result};
use crate::records::Record;

/// The kind of master represented by an inventory entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A main master from a `MainMaster` container.
    Main,
    /// A title master from a `Slide` container referenced by the master list.
    Title,
    /// The notes master from a `Notes` container.
    Notes,
    /// The handout master from a `Handout` container.
    Handout,
}

/// A validated persist-object reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Persist {
    id: u32,
    offset: u32,
}

impl Persist {
    pub(crate) const fn new(id: u32, offset: u32) -> Self {
        Self { id, offset }
    }

    /// The persist-object identifier used by MS-PPT references.
    pub const fn id(self) -> u32 {
        self.id
    }

    /// The byte offset currently associated with this persist identifier.
    pub const fn offset(self) -> u32 {
        self.offset
    }
}

/// The stable identity of a main or title master.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Identity {
    persist: Persist,
    master_id: u32,
}

impl Identity {
    pub(crate) const fn new(persist: Persist, master_id: u32) -> Self {
        Self { persist, master_id }
    }

    /// The persist reference for this master.
    pub const fn persist(self) -> Persist {
        self.persist
    }

    /// The MS-PPT master identifier.
    pub const fn master_id(self) -> u32 {
        self.master_id
    }
}

/// A reference to a source record together with its semantic persist identity.
#[derive(Debug, Clone, Copy)]
pub struct RecordRef<'a> {
    persist: Persist,
    record: &'a Record,
}

impl<'a> RecordRef<'a> {
    pub(crate) const fn new(persist: Persist, record: &'a Record) -> Self {
        Self { persist, record }
    }

    /// The persist reference of the source record.
    pub const fn persist(self) -> Persist {
        self.persist
    }

    /// The lossless parsed record.
    pub const fn record(self) -> &'a Record {
        self.record
    }
}

/// An uninterpreted record retained at its original owner boundary.
#[derive(Debug, Clone, Copy)]
pub struct Unknown<'a> {
    scope: Scope,
    record: &'a Record,
}

impl<'a> Unknown<'a> {
    pub(crate) const fn new(scope: Scope, record: &'a Record) -> Self {
        Self { scope, record }
    }

    /// The semantic owner in which this record occurred.
    pub const fn scope(self) -> Scope {
        self.scope
    }

    /// The raw numeric record type from the record header.
    pub const fn raw_type(self) -> u16 {
        self.record.record_type_raw
    }

    /// The record version from the record header.
    pub const fn version(self) -> u16 {
        self.record.version
    }

    /// The record instance from the record header.
    pub const fn instance(self) -> u16 {
        self.record.instance
    }

    /// The original record body bytes.
    pub const fn bytes(self) -> &'a [u8] {
        self.record.data.as_slice()
    }

    /// The original parsed record reference, including nested children.
    pub const fn record(self) -> &'a Record {
        self.record
    }

    /// Reconstruct the complete PPT record bytes without normalizing payloads.
    pub fn wire(self) -> Result<Vec<u8>> {
        encode_record(self.record)
    }
}

/// The boundary at which an unknown record was observed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Scope {
    /// The document container, outside its typed master list and atom.
    Document,
    /// The master list container.
    List,
    /// A main master container.
    Main,
    /// A title master slide container.
    Title,
    /// A notes master container.
    Notes,
    /// A handout master container.
    Handout,
}

/// A main master entry.
#[derive(Debug, Clone)]
pub struct Main<'a> {
    identity: Identity,
    source: RecordRef<'a>,
    reference: &'a Record,
    unknown: Vec<Unknown<'a>>,
}

impl<'a> Main<'a> {
    pub(crate) fn new(
        identity: Identity,
        source: RecordRef<'a>,
        reference: &'a Record,
        unknown: Vec<Unknown<'a>>,
    ) -> Self {
        Self {
            identity,
            source,
            reference,
            unknown,
        }
    }

    /// Stable persist/master identity.
    pub const fn identity(&self) -> Identity {
        self.identity
    }

    /// The source `MainMaster` record.
    pub const fn record(&self) -> &'a Record {
        self.source.record()
    }

    /// The persist reference used to locate the source record.
    pub const fn persist(&self) -> Persist {
        self.source.persist()
    }

    /// The original `MasterPersistAtom` reference record.
    pub const fn reference(&self) -> &'a Record {
        self.reference
    }

    /// Unknown or intentionally uninterpreted child records.
    pub fn unknown(&self) -> &[Unknown<'a>] {
        &self.unknown
    }
}

/// A title master entry.
#[derive(Debug, Clone)]
pub struct Title<'a> {
    identity: Identity,
    source: RecordRef<'a>,
    reference: &'a Record,
    based_on: Identity,
    unknown: Vec<Unknown<'a>>,
}

impl<'a> Title<'a> {
    pub(crate) fn new(
        identity: Identity,
        source: RecordRef<'a>,
        reference: &'a Record,
        based_on: Identity,
        unknown: Vec<Unknown<'a>>,
    ) -> Self {
        Self {
            identity,
            source,
            reference,
            based_on,
            unknown,
        }
    }

    /// Stable persist/master identity.
    pub const fn identity(&self) -> Identity {
        self.identity
    }

    /// The main master referenced by this title master.
    pub const fn based_on(&self) -> Identity {
        self.based_on
    }

    /// The source title `Slide` record.
    pub const fn record(&self) -> &'a Record {
        self.source.record()
    }

    /// The persist reference used to locate the source record.
    pub const fn persist(&self) -> Persist {
        self.source.persist()
    }

    /// The original `MasterPersistAtom` reference record.
    pub const fn reference(&self) -> &'a Record {
        self.reference
    }

    /// Unknown or intentionally uninterpreted child records.
    pub fn unknown(&self) -> &[Unknown<'a>] {
        &self.unknown
    }
}

/// The notes master entry.
#[derive(Debug, Clone)]
pub struct Notes<'a> {
    source: RecordRef<'a>,
    unknown: Vec<Unknown<'a>>,
}

impl<'a> Notes<'a> {
    pub(crate) fn new(source: RecordRef<'a>, unknown: Vec<Unknown<'a>>) -> Self {
        Self { source, unknown }
    }

    /// The persist reference used to locate the notes master.
    pub const fn persist(&self) -> Persist {
        self.source.persist()
    }

    /// The source `Notes` record.
    pub const fn record(&self) -> &'a Record {
        self.source.record()
    }

    /// Unknown or intentionally uninterpreted child records.
    pub fn unknown(&self) -> &[Unknown<'a>] {
        &self.unknown
    }
}

/// The handout master entry.
#[derive(Debug, Clone)]
pub struct Handout<'a> {
    source: RecordRef<'a>,
    unknown: Vec<Unknown<'a>>,
}

impl<'a> Handout<'a> {
    pub(crate) fn new(source: RecordRef<'a>, unknown: Vec<Unknown<'a>>) -> Self {
        Self { source, unknown }
    }

    /// The persist reference used to locate the handout master.
    pub const fn persist(&self) -> Persist {
        self.source.persist()
    }

    /// The source `Handout` record.
    pub const fn record(&self) -> &'a Record {
        self.source.record()
    }

    /// Unknown or intentionally uninterpreted child records.
    pub fn unknown(&self) -> &[Unknown<'a>] {
        &self.unknown
    }
}

/// A zero-copy source of persist objects.
#[derive(Debug, Clone)]
pub struct Objects<'a> {
    entries: Vec<Object<'a>>,
}

impl<'a> Objects<'a> {
    /// Build a deterministic persist-object catalog from existing records.
    pub fn from_records<I>(records: I) -> Result<Self>
    where
        I: IntoIterator<Item = (u32, &'a Record)>,
    {
        let mut entries = Vec::new();
        for (id, record) in records {
            if id == 0 {
                return Err(Error::InvalidFormat(
                    "persist object catalog contains a null identifier".into(),
                ));
            }
            if entries.iter().any(|entry: &Object<'a>| entry.id == id) {
                return Err(Error::InvalidFormat(format!(
                    "persist object catalog contains duplicate identifier {id}"
                )));
            }
            entries.push(Object { id, record });
        }
        entries.sort_unstable_by_key(|entry| entry.id);
        Ok(Self { entries })
    }

    /// Number of records available for persist resolution.
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the catalog has no records.
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub(crate) fn resolve(&self, id: u32) -> Option<&'a Record> {
        self.entries
            .binary_search_by_key(&id, |entry| entry.id)
            .ok()
            .map(|index| self.entries[index].record)
    }
}

/// One persist-object catalog entry.
#[derive(Debug, Clone, Copy)]
pub struct Object<'a> {
    id: u32,
    record: &'a Record,
}

impl<'a> Object<'a> {
    /// The persist identifier.
    pub const fn id(self) -> u32 {
        self.id
    }

    /// The lossless source record.
    pub const fn record(self) -> &'a Record {
        self.record
    }
}

/// The complete contextual master inventory.
#[derive(Debug, Clone)]
pub struct Inventory<'a> {
    main: Vec<Main<'a>>,
    title: Vec<Title<'a>>,
    notes: Option<Notes<'a>>,
    handout: Option<Handout<'a>>,
    unknown: Vec<Unknown<'a>>,
}

impl<'a> Inventory<'a> {
    pub(crate) fn new(
        main: Vec<Main<'a>>,
        title: Vec<Title<'a>>,
        notes: Option<Notes<'a>>,
        handout: Option<Handout<'a>>,
        unknown: Vec<Unknown<'a>>,
    ) -> Self {
        Self {
            main,
            title,
            notes,
            handout,
            unknown,
        }
    }

    /// Main masters in master-list order.
    pub fn main(&self) -> &[Main<'a>] {
        &self.main
    }

    /// Title masters in master-list order.
    pub fn title(&self) -> &[Title<'a>] {
        &self.title
    }

    /// The optional notes master.
    pub fn notes(&self) -> Option<&Notes<'a>> {
        self.notes.as_ref()
    }

    /// The optional handout master.
    pub fn handout(&self) -> Option<&Handout<'a>> {
        self.handout.as_ref()
    }

    /// Unknown records retained directly under the document owner.
    pub fn unknown(&self) -> &[Unknown<'a>] {
        &self.unknown
    }

    /// Find a main or title master by its semantic identity.
    pub fn find(&self, identity: Identity) -> Option<Master<'_>> {
        self.main
            .iter()
            .find(|master| master.identity == identity)
            .map(Master::Main)
            .or_else(|| {
                self.title
                    .iter()
                    .find(|master| master.identity == identity)
                    .map(Master::Title)
            })
    }

    /// Iterate over main and title masters without flattening their typed views.
    pub fn masters(&self) -> impl Iterator<Item = Master<'_>> + '_ {
        self.main
            .iter()
            .map(Master::Main)
            .chain(self.title.iter().map(Master::Title))
    }
}

/// A borrowed main/title master view returned by inventory lookup.
#[derive(Debug, Clone, Copy)]
pub enum Master<'a> {
    /// A main master.
    Main(&'a Main<'a>),
    /// A title master.
    Title(&'a Title<'a>),
}

impl<'a> Master<'a> {
    /// The contextual kind of this entry.
    pub const fn kind(self) -> Kind {
        match self {
            Self::Main(_) => Kind::Main,
            Self::Title(_) => Kind::Title,
        }
    }

    /// The stable identity of this entry.
    pub const fn identity(self) -> Identity {
        match self {
            Self::Main(value) => value.identity,
            Self::Title(value) => value.identity,
        }
    }
}

fn encode_record(record: &Record) -> Result<Vec<u8>> {
    if record.data.len() != usize::try_from(record.data_length).unwrap_or(usize::MAX) {
        return Err(Error::Corrupted(
            "unknown record payload length does not match its header".into(),
        ));
    }
    if record.version > 0x000f || record.instance > 0x0fff {
        return Err(Error::Corrupted(
            "unknown record header has an out-of-range version or instance".into(),
        ));
    }
    let mut bytes = Vec::with_capacity(8 + record.data.len());
    let version_instance = record.version | (record.instance << 4);
    bytes.extend_from_slice(&version_instance.to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&record.data_length.to_le_bytes());
    bytes.extend_from_slice(&record.data);
    Ok(bytes)
}
