//! Fresh Numbers workbooks (ADR 0034 §4, IWORK_API_REQUESTS.md request 1).
//!
//! [`Package::new`] returns the minimal native-shaped workbook: one sheet
//! (`Sheet 1`) holding one empty 10 × 5 table (`Table 1`). Every artifact it
//! returns is byte-identical and carries the fixed [`Identity::BLANK`].
//!
//! [`Builder`] authors a workbook with a caller-chosen [`Identity`] and
//! optional sheet and table names. The identity is a required input:
//! builders never consult ambient randomness, so a fixed identity yields
//! byte-identical output ([`Identity::generate`] is the explicit opt-in to
//! a fresh native-like identity).
//!
//! ```
//! use litchi_numbers::{Builder, Identity, Package};
//!
//! # fn main() -> Result<(), litchi_numbers::create::Error> {
//! let blank = Package::new();
//! assert_eq!(blank.sheets().len(), 1);
//!
//! let identity = Identity::from_bytes([1; 16], [2; 16], [3; 16])?;
//! let package = Builder::new()
//!     .identity(identity)
//!     .sheet_name("Budget")
//!     .table_name("Q1")
//!     .build()?;
//! assert_eq!(package.sheets()[0].name(), "Budget");
//! # Ok(())
//! # }
//! ```
//!
//! # Template provenance
//!
//! The blank artifact is the output of the migration host's from-scratch
//! Numbers writer (`litchi_iwa::numbers::Builder::new()`,
//! file format 14.4.1, stored members), with its identity rewritten to
//! [`Identity::BLANK`] by this module and its build-history producer
//! string set to `Litchi`. Record 0309 then gave its table's formula owner
//! the native table extents (the spanning dependencies' total and body
//! ranges, as a table added by `table::lifecycle` carries them), so formulas
//! can be authored on a created workbook; the change was published through
//! the structural working set and the asset regenerated from it. It is
//! 6 KiB and holds no preview images.

use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_core::SnappyStream;
use thiserror::Error as ThisError;

use crate::package::{Package, Profile};
use crate::{SheetSelector, TableSelector};

/// The blank workbook every [`Package::new`] returns.
const BLANK: &[u8] = include_bytes!("create/blank.numbers");

const DOCUMENT_IDENTIFIER: &str = "Metadata/DocumentIdentifier";
const PROPERTIES: &str = "Metadata/Properties.plist";
const PACKAGE_METADATA: &str = "Index/Metadata.iwa";

/// The three independent UUIDs that identify one Numbers document: the
/// public document UUID (also its share and stable UUID), the saved-version
/// UUID and the private bookkeeping UUID.
///
/// Distinct identities keep documents derived from one source from being
/// treated as revisions of each other by Numbers or iCloud.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Identity {
    document: [u8; 16],
    version: [u8; 16],
    private: [u8; 16],
}

impl fmt::Debug for Identity {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Identity")
            .field("document", &Uuid(self.document))
            .field("version", &Uuid(self.version))
            .field("private", &Uuid(self.private))
            .finish()
    }
}

impl Identity {
    /// The fixed identity of the [`Package::new`] artifact.
    pub const BLANK: Self = Self {
        document: [
            0x5c, 0x0f, 0x6a, 0x31, 0x4f, 0x5e, 0x4d, 0x2b, 0x9c, 0x1a, 0x6e, 0x3b, 0x2d, 0x7f,
            0x8a, 0x01,
        ],
        version: [
            0x5c, 0x0f, 0x6a, 0x31, 0x4f, 0x5e, 0x4d, 0x2b, 0x9c, 0x1a, 0x6e, 0x3b, 0x2d, 0x7f,
            0x8a, 0x02,
        ],
        private: [
            0x5c, 0x0f, 0x6a, 0x31, 0x4f, 0x5e, 0x4d, 0x2b, 0x9c, 0x1a, 0x6e, 0x3b, 0x2d, 0x7f,
            0x8a, 0x03,
        ],
    };

    /// Three fresh, mutually distinct random (RFC 4122 version 4) UUIDs:
    /// the explicit opt-in to a native-like fresh identity.
    #[must_use]
    pub fn generate() -> Self {
        let document = litchi_core::id::generate_guid_bytes();
        let mut version = litchi_core::id::generate_guid_bytes();
        while version == document {
            version = litchi_core::id::generate_guid_bytes();
        }
        let mut private = litchi_core::id::generate_guid_bytes();
        while private == document || private == version {
            private = litchi_core::id::generate_guid_bytes();
        }
        Self {
            document,
            version,
            private,
        }
    }

    /// A fixed identity for deterministic pipelines and goldens.
    ///
    /// # Errors
    ///
    /// [`Error::IdentityNotDistinct`] when two of the UUIDs are equal.
    pub fn from_bytes(document: [u8; 16], version: [u8; 16], private: [u8; 16]) -> Result<Self> {
        if document == version || document == private || version == private {
            return Err(Error::IdentityNotDistinct);
        }
        Ok(Self {
            document,
            version,
            private,
        })
    }

    /// The public document UUID.
    #[must_use]
    pub const fn document(&self) -> [u8; 16] {
        self.document
    }

    /// The saved-version UUID.
    #[must_use]
    pub const fn version(&self) -> [u8; 16] {
        self.version
    }

    /// The private bookkeeping UUID.
    #[must_use]
    pub const fn private(&self) -> [u8; 16] {
        self.private
    }
}

/// Uppercase hyphenated rendering of a UUID, as Numbers stores it.
struct Uuid([u8; 16]);

impl fmt::Display for Uuid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        for (index, byte) in self.0.iter().enumerate() {
            if matches!(index, 4 | 6 | 8 | 10) {
                formatter.write_str("-")?;
            }
            write!(formatter, "{byte:02X}")?;
        }
        Ok(())
    }
}

impl fmt::Debug for Uuid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(self, formatter)
    }
}

/// The failure of a fresh-workbook build. Messages carry no authored
/// content.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// [`Builder::build`] was called without [`Builder::identity`].
    #[error(
        "a Numbers builder needs an explicit identity (Identity::generate or Identity::from_bytes)"
    )]
    MissingIdentity,
    /// Two UUIDs of an identity are equal.
    #[error("the three UUIDs of a Numbers identity must be distinct")]
    IdentityNotDistinct,
    /// A requested sheet or table name was refused by the names transaction.
    #[error(transparent)]
    Name(#[from] crate::names::Error),
    /// The template could not be rewritten or reopened.
    #[error(transparent)]
    Package(#[from] crate::Error),
}

/// The result of a fresh-workbook build.
pub type Result<T> = std::result::Result<T, Error>;

/// Option-bearing authoring of a fresh Numbers workbook.
///
/// The output is a [`Package`]; output bytes come from
/// [`Package::write_to`] or [`Package::save`].
#[derive(Debug, Clone, Default)]
#[must_use]
pub struct Builder {
    identity: Option<Identity>,
    sheet_name: Option<String>,
    table_name: Option<String>,
}

impl Builder {
    /// A builder with no identity and the default names.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the document identity (required).
    pub const fn identity(mut self, identity: Identity) -> Self {
        self.identity = Some(identity);
        self
    }

    /// Name the workbook's sheet (default `Sheet 1`).
    pub fn sheet_name(mut self, name: impl Into<String>) -> Self {
        self.sheet_name = Some(name.into());
        self
    }

    /// Name the sheet's table (default `Table 1`).
    pub fn table_name(mut self, name: impl Into<String>) -> Self {
        self.table_name = Some(name.into());
        self
    }

    /// Build the workbook.
    ///
    /// # Errors
    ///
    /// [`Error::MissingIdentity`] without an identity; [`Error::Name`] when
    /// a requested name is refused (empty, or otherwise invalid for
    /// Numbers); [`Error::Package`] when the template cannot be rewritten.
    pub fn build(self) -> Result<Package> {
        let identity = self.identity.ok_or(Error::MissingIdentity)?;
        let package = if identity == Identity::BLANK {
            Package::new()
        } else {
            let bytes = with_identity(BLANK, identity)?;
            Package::from_shared_profile(Arc::from(bytes), Profile::default())?
        };
        if self.sheet_name.is_none() && self.table_name.is_none() {
            return Ok(package);
        }
        let mut edit = package.edit_names();
        if let Some(name) = self.sheet_name.as_deref() {
            edit = edit.rename_sheet(SheetSelector::index(0), name)?;
        }
        if let Some(name) = self.table_name.as_deref() {
            edit = edit.rename_table(SheetSelector::index(0), TableSelector::index(0), name)?;
        }
        Ok(edit.commit()?.into_package())
    }
}

impl Package {
    /// The minimal Numbers workbook: one sheet (`Sheet 1`) holding one empty
    /// 10 × 5 table (`Table 1`), with the fixed [`Identity::BLANK`]. Every
    /// call returns the same bytes; use [`Builder`] for a distinct identity.
    #[must_use]
    pub fn new() -> Self {
        // The template is a compile-time constant the crate's own tests
        // open; its failure would be a build defect, not a runtime input.
        #[allow(
            clippy::expect_used,
            reason = "the embedded template is validated by the crate's tests"
        )]
        Self::from_shared_profile(Arc::from(BLANK), Profile::default())
            .expect("the embedded blank Numbers workbook opens")
    }
}

impl Default for Package {
    fn default() -> Self {
        Self::new()
    }
}

/// Rewrite the three identity carriers of `source`: the document
/// identifier, the properties plist and the package-metadata version UUID.
pub(crate) fn with_identity(source: &[u8], identity: Identity) -> crate::Result<Vec<u8>> {
    let profile = Profile::default();
    let limits = profile.archive();
    let catalog = SourceCatalog::from_bytes(source, limits)?;
    let package = catalog.package();
    let member = |name: &str| {
        package
            .iter()
            .find(|entry| entry.name() == name && !entry.is_opaque())
            .map(|entry| entry.data().to_vec())
            .ok_or_else(|| {
                crate::Error::InvalidFormat(format!("template member {name} is missing"))
            })
    };
    let document = Uuid(identity.document).to_string();
    let version = Uuid(identity.version).to_string();
    let private = Uuid(identity.private).to_string();

    let identifier = document.clone().into_bytes();

    let plist = member(PROPERTIES)?;
    let mut value = plist::Value::from_reader(Cursor::new(plist))
        .map_err(|_| crate::Error::InvalidFormat("template properties are malformed".to_owned()))?;
    let dictionary = value.as_dictionary_mut().ok_or_else(|| {
        crate::Error::InvalidFormat("template properties are not a dictionary".to_owned())
    })?;
    let old_version = dictionary
        .get("versionUUID")
        .and_then(plist::Value::as_string)
        .map(str::to_owned)
        .ok_or_else(|| {
            crate::Error::InvalidFormat("template properties lack a version".to_owned())
        })?;
    for key in ["documentUUID", "stableDocumentUUID", "shareUUID"] {
        dictionary.insert(key.to_owned(), plist::Value::String(document.clone()));
    }
    dictionary.insert("privateUUID".to_owned(), plist::Value::String(private));
    dictionary.insert(
        "versionUUID".to_owned(),
        plist::Value::String(version.clone()),
    );
    dictionary.insert(
        "revision".to_owned(),
        plist::Value::String(format!("0::{version}")),
    );
    let mut properties = Vec::new();
    value.to_writer_binary(&mut properties).map_err(|_| {
        crate::Error::InvalidFormat("template properties cannot be written".to_owned())
    })?;

    let metadata = member(PACKAGE_METADATA)?;
    let stream = SnappyStream::decompress(&metadata, limits.snappy_limits()?)
        .map_err(|error| crate::Error::Archive(litchi_iwa_archive::Error::Iwa(error)))?;
    let mut decoded = stream.into_bytes();
    let needle = old_version.as_bytes();
    let found: Vec<usize> = decoded
        .windows(needle.len())
        .enumerate()
        .filter_map(|(offset, window)| (window == needle).then_some(offset))
        .collect();
    let [offset] = found.as_slice() else {
        return Err(crate::Error::InvalidFormat(
            "template package metadata does not carry its version exactly once".to_owned(),
        ));
    };
    decoded[*offset..*offset + needle.len()].copy_from_slice(version.as_bytes());
    let metadata = SnappyStream::compress(&decoded, Default::default())
        .map_err(|error| crate::Error::Archive(litchi_iwa_archive::Error::Iwa(error)))?;

    let edits = [
        EntryEdit::new(DOCUMENT_IDENTIFIER, &identifier),
        EntryEdit::new(PROPERTIES, &properties),
        EntryEdit::new(PACKAGE_METADATA, &metadata),
    ];
    Ok(package.reassemble_with_deletions_to_bytes(&edits, &[], limits)?)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn new_is_a_one_sheet_one_table_workbook() {
        let package = Package::new();
        let sheets = package.sheets();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name(), "Sheet 1");
        let table = sheets[0].tables().next().unwrap();
        assert_eq!(table.name(), "Table 1");
        assert_eq!((table.row_count(), table.column_count()), (10, 5));
        let mut first = Vec::new();
        package.write_to(&mut first).unwrap();
        let mut second = Vec::new();
        Package::new().write_to(&mut second).unwrap();
        assert_eq!(first, second);
        assert_eq!(first, BLANK);
    }

    #[test]
    fn the_template_carries_the_blank_identity() {
        let rewritten = with_identity(BLANK, Identity::BLANK).unwrap();
        assert_eq!(rewritten, BLANK);
    }

    #[test]
    fn a_build_without_identity_is_refused() {
        assert!(matches!(
            Builder::new().build(),
            Err(Error::MissingIdentity)
        ));
        assert!(matches!(
            Identity::from_bytes([1; 16], [1; 16], [2; 16]),
            Err(Error::IdentityNotDistinct)
        ));
    }

    #[test]
    fn a_fixed_identity_is_deterministic_and_distinct_from_the_blank() {
        let identity = Identity::from_bytes([7; 16], [8; 16], [9; 16]).unwrap();
        let bytes = |package: &Package| {
            let mut out = Vec::new();
            package.write_to(&mut out).unwrap();
            out
        };
        let first = Builder::new().identity(identity).build().unwrap();
        let second = Builder::new().identity(identity).build().unwrap();
        assert_eq!(bytes(&first), bytes(&second));
        assert_ne!(bytes(&first), BLANK);
        let generated = Builder::new()
            .identity(Identity::generate())
            .build()
            .unwrap();
        assert_ne!(bytes(&generated), bytes(&first));
    }

    #[test]
    fn names_are_applied_through_the_names_transaction() {
        let package = Builder::new()
            .identity(Identity::from_bytes([4; 16], [5; 16], [6; 16]).unwrap())
            .sheet_name("Budget")
            .table_name("Q1")
            .build()
            .unwrap();
        assert_eq!(package.sheets()[0].name(), "Budget");
        assert_eq!(package.sheets()[0].tables().next().unwrap().name(), "Q1");
        assert!(matches!(
            Builder::new()
                .identity(Identity::BLANK)
                .sheet_name("")
                .build(),
            Err(Error::Name(_))
        ));
    }
}
