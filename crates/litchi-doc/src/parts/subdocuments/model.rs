//! Typed master-document subdocument metadata.

use crate::parts::mail_merge::Fnpi;

/// The type and identifier that address one `SttbFnm` entry.
///
/// The identifier is allocated independently within each [`Kind`]. The
/// reserved 0x0FFF identifier is rejected by [`FileNameKey::try_new`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FileNameKey {
    kind: Kind,
    identifier: u16,
}

impl FileNameKey {
    /// Creates a checked file-name key.
    pub fn try_new(kind: Kind, identifier: u16) -> Result<Self, FileNameKeyError> {
        if identifier == 0x0FFF {
            return Err(FileNameKeyError::ReservedIdentifier);
        }
        if identifier > 0x0FFF {
            return Err(FileNameKeyError::IdentifierOutOfRange(identifier));
        }
        Ok(Self { kind, identifier })
    }

    /// The file-name kind encoded in `FNPI.fnpt`.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// The identifier encoded in `FNPI.fnpd`.
    #[must_use]
    pub const fn identifier(self) -> u16 {
        self.identifier
    }

    pub(crate) fn fnpi(self) -> Fnpi {
        Fnpi::from_raw(u16::from(self.kind.raw_type()) | (self.identifier << 4))
    }

    pub(crate) fn from_fnpi(fnpi: Fnpi) -> Self {
        Self {
            kind: Kind::from_raw(fnpi.file_type()),
            identifier: fnpi.identifier(),
        }
    }
}

/// A file-name key could not be represented by the MS-DOC `FNPI` fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileNameKeyError {
    /// `fnpd = 0x0FFF` is reserved and MUST NOT be used.
    ReservedIdentifier,
    /// `fnpd` occupies only 12 bits.
    IdentifierOutOfRange(u16),
}

impl std::fmt::Display for FileNameKeyError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ReservedIdentifier => formatter.write_str("FNPI.fnpd 0x0FFF is reserved"),
            Self::IdentifierOutOfRange(identifier) => {
                write!(formatter, "FNPI.fnpd {identifier:#06X} exceeds 12 bits")
            },
        }
    }
}

impl std::error::Error for FileNameKeyError {}

/// The kind of an externally referenced file (`FNPI.fnpt`, MS-DOC 2.9.93).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A mail merge data source file.
    MailMergeDataSource,
    /// A subdocument of a master document.
    Subdocument,
}

impl Kind {
    pub(crate) const fn raw_type(self) -> u8 {
        match self {
            Self::MailMergeDataSource => 0x3,
            Self::Subdocument => 0x5,
        }
    }

    pub(crate) fn from_raw(raw: u8) -> Self {
        match raw {
            0x5 => Self::Subdocument,
            _ => Self::MailMergeDataSource,
        }
    }
}

/// The editable `FNIF` path and file-system metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct FileNameMetadata {
    /// UTF-16 code-unit offset of the document-relative path, or `None` for
    /// the `0xFF` sentinel.
    pub relative_path_offset: Option<usize>,
    /// Whether the path is valid on FAT file systems.
    pub valid_on_fat: bool,
    /// Whether the path is valid on NTFS file systems.
    pub valid_on_ntfs: bool,
    /// Whether the path uses an external file I/O protocol rather than a
    /// native file-system path.
    pub is_non_file_system_path: bool,
}

/// One external file referenced by the document: an `SttbFnm` string plus its
/// appended `FNIF` metadata (MS-DOC 2.9.288 and 2.9.92).
///
/// The path is stored verbatim and is never opened, resolved, or followed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Name {
    pub(crate) fnpi: Fnpi,
    pub(crate) raw_fnfb: u8,
    pub(crate) fnif_unused: [u8; 4],
    /// The full path of the referenced file, including name and extension.
    pub(crate) path: String,
    /// `FNIF.ichRelative`: the character offset into `path` at which the
    /// document-relative path segment starts, or `None` when the file name
    /// carries no such segment.
    pub(crate) relative_path_offset: Option<usize>,
    /// Whether the path is valid on FAT file systems (`FNFB.fFAT`).
    pub(crate) valid_on_fat: bool,
    /// Whether the path is valid on NTFS file systems (`FNFB.fNTFS`).
    pub(crate) valid_on_ntfs: bool,
    /// Whether the path is not a native file system path and requires an
    /// external file I/O protocol (`FNFB.fNonFileSys`).
    pub(crate) is_non_file_system_path: bool,
}

impl Name {
    /// The type and identifier of this file name (`FNPI`, MS-DOC 2.9.93).
    #[must_use]
    pub fn fnpi(&self) -> Fnpi {
        self.fnpi
    }

    /// The typed key used by `SttbFnm` and `PlcfWKB` references.
    #[must_use]
    pub fn key(&self) -> FileNameKey {
        FileNameKey::from_fnpi(self.fnpi)
    }

    /// The full file path exactly as stored in `SttbFnm`.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// The `FNIF.ichRelative` UTF-16 code-unit offset, when present.
    #[must_use]
    pub fn relative_path_offset(&self) -> Option<usize> {
        self.relative_path_offset
    }

    /// The raw `FNIF.fnfb` byte, including undefined bits retained from the
    /// source. The typed booleans are available through the other accessors.
    #[must_use]
    pub fn file_system_flags(&self) -> u8 {
        self.raw_fnfb
    }

    /// The four undefined `FNIF.unused` bytes retained verbatim.
    #[must_use]
    pub fn fnif_unused(&self) -> [u8; 4] {
        self.fnif_unused
    }

    /// The kind of the referenced file.
    #[must_use]
    pub fn kind(&self) -> Kind {
        match self.fnpi.file_type() {
            0x5 => Kind::Subdocument,
            _ => Kind::MailMergeDataSource,
        }
    }

    /// The path segment relative to the folder containing the document, when
    /// the file name carries one. Never resolved against the file system.
    #[must_use]
    pub fn relative_path(&self) -> Option<&str> {
        self.relative_path_offset
            .and_then(|offset| utf16_suffix(&self.path, offset))
    }

    pub(crate) fn metadata(&self) -> FileNameMetadata {
        FileNameMetadata {
            relative_path_offset: self.relative_path_offset,
            valid_on_fat: self.valid_on_fat,
            valid_on_ntfs: self.valid_on_ntfs,
            is_non_file_system_path: self.is_non_file_system_path,
        }
    }
}

/// Convert the UTF-16 code-unit offset carried by `FNIF.ichRelative` into a
/// UTF-8 boundary without changing the source offset. An offset inside a
/// surrogate pair is malformed semantic state and is reported as absent.
fn utf16_suffix(path: &str, offset: usize) -> Option<&str> {
    let mut units = 0usize;
    for (byte, character) in path.char_indices() {
        if units == offset {
            return path.get(byte..);
        }
        units = units.checked_add(character.len_utf16())?;
        if units > offset {
            return None;
        }
    }
    (units == offset).then_some("")
}

/// One subdocument of a master document (`WKB`, MS-DOC 2.9.346).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Reference {
    /// Character position in the main document where the subdocument begins.
    pub(crate) start: u32,
    /// The outline level of the subdocument (`WKB.lvl`).
    pub(crate) outline_level: u16,
    /// The type and identifier of the subdocument file name (`WKB.fnpi`).
    pub(crate) file_name: Fnpi,
    pub(crate) file_name_index: usize,
    pub(crate) raw_flags: u16,
    pub(crate) raw_wkb: [u8; 12],
}

impl Reference {
    /// The main-document character position at which the subdocument starts.
    #[must_use]
    pub const fn start(&self) -> u32 {
        self.start
    }

    /// The mandated WKB outline level.
    #[must_use]
    pub const fn outline_level(&self) -> u16 {
        self.outline_level
    }

    /// The typed key of the referenced `SttbFnm` entry.
    #[must_use]
    pub fn file_name_key(&self) -> FileNameKey {
        FileNameKey::from_fnpi(self.file_name)
    }

    /// The `FNPI` stored in this reference.
    #[must_use]
    pub const fn file_name(&self) -> Fnpi {
        self.file_name
    }
}

/// The master-document subdocument directory and the referenced-file name
/// table, addressed by `fcPlcfWkb` and `fcSttbFnm`.
///
/// All data is inert: paths are exposed verbatim and never opened, resolved,
/// or followed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Collection {
    pub(crate) referenced_files: Vec<Name>,
    pub(crate) subdocuments: Vec<Reference>,
}

impl Collection {
    pub(crate) fn from_parts(referenced_files: Vec<Name>, subdocuments: Vec<Reference>) -> Self {
        Self {
            referenced_files,
            subdocuments,
        }
    }

    /// All externally referenced files in `SttbFnm` table order.
    #[must_use]
    pub fn referenced_files(&self) -> &[Name] {
        &self.referenced_files
    }

    /// The subdocuments in start-CP order (empty unless this is a master
    /// document).
    #[must_use]
    pub fn subdocuments(&self) -> &[Reference] {
        &self.subdocuments
    }

    /// Resolve an `FNPI` reference to its `SttbFnm` entry.
    #[must_use]
    pub fn file_name(&self, fnpi: Fnpi) -> Option<&Name> {
        self.referenced_files.iter().find(|file| file.fnpi == fnpi)
    }

    /// The referenced file of a subdocument. Always resolves: entries are
    /// validated against the `SttbFnm` during parsing.
    #[must_use]
    pub fn file_name_of(&self, reference: &Reference) -> &Name {
        &self.referenced_files[reference.file_name_index]
    }
}
