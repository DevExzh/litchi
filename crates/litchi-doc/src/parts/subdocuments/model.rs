//! Typed master-document subdocument metadata.

use crate::parts::mail_merge::Fnpi;

/// The kind of an externally referenced file (`FNPI.fnpt`, MS-DOC 2.9.93).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    /// A mail merge data source file.
    MailMergeDataSource,
    /// A subdocument of a master document.
    Subdocument,
}

/// One external file referenced by the document: an `SttbFnm` string plus its
/// appended `FNIF` metadata (MS-DOC 2.9.288 and 2.9.92).
///
/// The path is stored verbatim and is never opened, resolved, or followed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Name {
    pub(crate) fnpi: Fnpi,
    /// The full path of the referenced file, including name and extension.
    pub path: String,
    /// `FNIF.ichRelative`: the character offset into `path` at which the
    /// document-relative path segment starts, or `None` when the file name
    /// carries no such segment.
    pub relative_path_offset: Option<usize>,
    /// Whether the path is valid on FAT file systems (`FNFB.fFAT`).
    pub valid_on_fat: bool,
    /// Whether the path is valid on NTFS file systems (`FNFB.fNTFS`).
    pub valid_on_ntfs: bool,
    /// Whether the path is not a native file system path and requires an
    /// external file I/O protocol (`FNFB.fNonFileSys`).
    pub is_non_file_system_path: bool,
}

impl Name {
    /// The type and identifier of this file name (`FNPI`, MS-DOC 2.9.93).
    pub fn fnpi(&self) -> Fnpi {
        self.fnpi
    }

    /// The kind of the referenced file.
    pub fn kind(&self) -> Kind {
        match self.fnpi.file_type() {
            0x5 => Kind::Subdocument,
            _ => Kind::MailMergeDataSource,
        }
    }

    /// The path segment relative to the folder containing the document, when
    /// the file name carries one. Never resolved against the file system.
    pub fn relative_path(&self) -> Option<&str> {
        self.relative_path_offset
            .and_then(|offset| self.path.get(offset..))
    }
}

/// One subdocument of a master document (`WKB`, MS-DOC 2.9.346).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Reference {
    /// Character position in the main document where the subdocument begins.
    pub start: u32,
    /// The outline level of the subdocument (`WKB.lvl`).
    pub outline_level: u16,
    /// The type and identifier of the subdocument file name (`WKB.fnpi`).
    pub file_name: Fnpi,
    pub(crate) file_name_index: usize,
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
    pub fn referenced_files(&self) -> &[Name] {
        &self.referenced_files
    }

    /// The subdocuments in start-CP order (empty unless this is a master
    /// document).
    pub fn subdocuments(&self) -> &[Reference] {
        &self.subdocuments
    }

    /// Resolve an `FNPI` reference to its `SttbFnm` entry.
    pub fn file_name(&self, fnpi: Fnpi) -> Option<&Name> {
        self.referenced_files.iter().find(|file| file.fnpi == fnpi)
    }

    /// The referenced file of a subdocument. Always resolves: entries are
    /// validated against the `SttbFnm` during parsing.
    pub fn file_name_of(&self, reference: &Reference) -> &Name {
        &self.referenced_files[reference.file_name_index]
    }
}
