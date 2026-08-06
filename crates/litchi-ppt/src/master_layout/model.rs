//! Immutable semantic views for one legacy PowerPoint master layout.

use crate::consts::RecordType;
use crate::package::Result;
use crate::records::Record;

use super::inventory::inventory as scan;
use super::{codec, transaction, validation};

/// The contextual meaning of a master record.
///
/// The short names are intentional: the record kind supplies the context, so
/// callers do not need redundant MainMaster/NotesMaster type prefixes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Context {
    /// A MainMaster container.
    Main,
    /// A Slide container whose SlideAtom.geom is SL_MasterTitle.
    Title,
    /// A Notes container whose NotesAtom.slideIdRef is zero.
    Notes,
    /// A Handout container.
    Handout,
}

impl Context {
    pub(crate) const fn expected_record_type(self) -> RecordType {
        match self {
            Self::Main => RecordType::MainMaster,
            Self::Title => RecordType::Slide,
            Self::Notes => RecordType::Notes,
            Self::Handout => RecordType::Handout,
        }
    }

    /// Stable semantic label useful for diagnostics and logging.
    pub const fn name(self) -> &'static str {
        match self {
            Self::Main => "main",
            Self::Title => "title",
            Self::Notes => "notes",
            Self::Handout => "handout",
        }
    }
}

/// Resource limits applied while parsing, encoding, and validating a record
/// tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum nesting depth, including the root record.
    pub max_depth: usize,
    /// Maximum number of records in one tree.
    pub max_records: usize,
    /// Maximum encoded byte size of one layout snapshot.
    pub max_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_depth: 128,
            max_records: 262_144,
            max_bytes: 64 * 1024 * 1024,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<Self> {
        if self.max_depth == 0 || self.max_records == 0 || self.max_bytes < 8 {
            return Err(crate::package::Error::InvalidFormat(
                "master-layout limits must allow at least one record".into(),
            ));
        }
        Ok(self)
    }
}

/// A stable path into a record tree, expressed as child indexes from the
/// layout root.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct Path(Vec<usize>);

impl Path {
    /// Return the root path.
    pub const fn root() -> Self {
        Self(Vec::new())
    }

    /// Return a path containing one more child index.
    pub fn child(&self, index: usize) -> Self {
        let mut path = self.0.clone();
        path.push(index);
        Self(path)
    }

    /// Return the path segments without allocation.
    pub fn as_slice(&self) -> &[usize] {
        &self.0
    }

    /// Whether this path identifies the layout root.
    pub const fn is_root(&self) -> bool {
        self.0.is_empty()
    }
}

impl From<Vec<usize>> for Path {
    fn from(value: Vec<usize>) -> Self {
        Self(value)
    }
}

impl From<&[usize]> for Path {
    fn from(value: &[usize]) -> Self {
        Self(value.to_vec())
    }
}

impl From<&Path> for Path {
    fn from(value: &Path) -> Self {
        value.clone()
    }
}

impl From<usize> for Path {
    fn from(value: usize) -> Self {
        Self(vec![value])
    }
}

/// One contextual master found while scanning a presentation record tree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    path: Path,
    context: Context,
}

impl Entry {
    pub(crate) const fn new(path: Path, context: Context) -> Self {
        Self { path, context }
    }

    /// Location of this master in the supplied record tree.
    pub const fn path(&self) -> &Path {
        &self.path
    }

    /// Contextual kind of this master.
    pub const fn context(&self) -> Context {
        self.context
    }
}

/// A deterministic inventory of contextual masters.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Inventory {
    entries: Vec<Entry>,
}

impl Inventory {
    pub(crate) fn new(entries: Vec<Entry>) -> Self {
        Self { entries }
    }

    /// All masters in depth-first record order.
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }

    /// Number of discovered masters.
    pub const fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether no contextual masters were found.
    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Iterate over masters of one context.
    pub fn by_context(&self, context: Context) -> impl Iterator<Item = &Entry> {
        self.entries
            .iter()
            .filter(move |entry| entry.context == context)
    }
}

/// An immutable, validated source or committed state for one master layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) context: Context,
    pub(crate) root: Record,
    pub(crate) bytes: Vec<u8>,
    pub(crate) limits: Limits,
}

impl Snapshot {
    /// Validate and capture a record supplied by a parent presentation editor.
    pub fn from_record(context: Context, root: Record) -> Result<Self> {
        Self::from_record_with_limits(context, root, Limits::default())
    }

    /// Validate and capture a record under explicit resource limits.
    pub fn from_record_with_limits(context: Context, root: Record, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let bytes = codec::encode(&root, limits)?;
        let reparsed = codec::parse(&bytes, limits)?;
        validation::validate(context, &reparsed, limits)?;
        Ok(Self {
            context,
            root: reparsed,
            bytes,
            limits,
        })
    }

    /// Parse one complete PPT record and capture it as an immutable snapshot.
    pub fn parse(context: Context, bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(context, bytes, Limits::default())
    }

    /// Parse one complete PPT record under explicit resource limits.
    pub fn parse_with_limits(context: Context, bytes: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let root = codec::parse(bytes, limits)?;
        validation::validate(context, &root, limits)?;
        let encoded = codec::encode(&root, limits)?;
        if encoded != bytes {
            return Err(crate::package::Error::Corrupted(
                "master-layout record is not losslessly representable".into(),
            ));
        }
        Ok(Self {
            context,
            root,
            bytes: bytes.to_vec(),
            limits,
        })
    }

    /// Context of this layout.
    pub const fn context(&self) -> Context {
        self.context
    }

    /// Borrow the validated record tree.
    pub const fn record(&self) -> &Record {
        &self.root
    }

    /// Borrow the exact encoded record bytes represented by this snapshot.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Resource limits attached to this snapshot.
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Stable content revision for optimistic parent-owner integration.
    #[must_use]
    pub fn revision(&self) -> transaction::Revision {
        transaction::Revision::from_bytes(&self.bytes)
    }

    /// Start an isolated transactional edit from this immutable state.
    pub fn edit(&self) -> transaction::Transaction {
        transaction::Transaction::open(self.clone())
    }

    /// Borrow this snapshot through the transaction's immutable view.
    #[must_use]
    pub fn view(&self) -> transaction::View<'_> {
        transaction::View::new(self.context, &self.root, self.revision())
    }

    /// Inventory contextual masters in this snapshot's record tree.
    pub fn inventory(&self) -> Result<Inventory> {
        scan(&self.root)
    }
}
