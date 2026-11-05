//! Exact-source transactions for the Pages document properties
//! (IWORK_API_REQUESTS.md request 14; ADR 0034 §7).
//!
//! The title, author, keywords and description live in
//! `Metadata/Properties.plist` under the `Title`, `Author`, `Keywords` and
//! `Comments` keys, the keys [`Package::metadata`] reads. A commit rewrites
//! that one ZIP member (every other member keeps its records), reopens the
//! candidate and verifies that the read-back metadata equals the request.
//!
//! Pages itself does not show these keys in its interface; they are the
//! properties this library reads and other tools may set. Setting a field
//! to `None` removes its key and the legacy fallbacks the reader honours
//! (`kDocumentTitleKey`, `kDocumentAuthorKey`, `kSFWPAuthorPropertyKey`), so
//! the read-back value is absent.

use std::fmt;
use std::sync::Arc;

use thiserror::Error as ThisError;

pub use litchi_core::edit::Diagnostics;

use super::{Package, PackageError, sidecars};

/// The largest encoded properties sidecar a commit publishes.
const MAX_PROPERTIES_BYTES: usize = litchi_iwa_detect::MAX_PROPERTIES_BYTES;

/// A finite resource governed while document properties are rewritten.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Encoded bytes of the rewritten properties sidecar.
    PropertiesBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PropertiesBytes => "properties bytes",
        })
    }
}

/// An error raised while staging or publishing document properties.
///
/// Diagnostics never carry the property values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The existing properties sidecar cannot be read or rewritten.
    #[error("the Pages properties sidecar cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Pages document-properties {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// The reopened candidate did not reproduce the requested properties.
    #[error("the edited Pages document properties failed verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Pages document-properties patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// A value holds a NUL character, which a property list cannot carry.
    #[error("a Pages document property cannot contain a NUL character")]
    InvalidValue,
    /// A keyword holds `;` or `,`, or surrounding whitespace, which the
    /// one `Keywords` string cannot keep apart from its neighbours
    /// (`litchi_core::metadata::keywords::join`).
    #[error("a Pages keyword cannot be stored in the separated Keywords string")]
    InvalidKeyword,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::InvalidValue { .. } | Self::InvalidKeyword { .. } => RefusalKind::InvalidValue,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// A staged change of one property: keep, set, or remove.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
enum Change {
    #[default]
    Keep,
    Set(Option<String>),
}

/// The four editable properties and their plist keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Field {
    Title,
    Author,
    Keywords,
    Description,
}

impl Field {
    const ALL: [Self; 4] = [Self::Title, Self::Author, Self::Keywords, Self::Description];

    /// The key this library writes.
    const fn key(self) -> &'static str {
        match self {
            Self::Title => "Title",
            Self::Author => "Author",
            Self::Keywords => "Keywords",
            Self::Description => "Comments",
        }
    }

    /// Legacy keys the reader falls back to; removed on every write so the
    /// written key is authoritative.
    const fn fallbacks(self) -> &'static [&'static str] {
        match self {
            Self::Title => &["kDocumentTitleKey"],
            Self::Author => &["kDocumentAuthorKey", "kSFWPAuthorPropertyKey"],
            Self::Keywords | Self::Description => &[],
        }
    }

    /// The text slot of a text field; `None` for the keyword list.
    fn slot(self, metadata: &mut litchi_core::Metadata) -> Option<&mut Option<String>> {
        match self {
            Self::Title => Some(&mut metadata.title),
            Self::Author => Some(&mut metadata.author),
            Self::Keywords => None,
            Self::Description => Some(&mut metadata.description),
        }
    }
}

/// A staged document-properties change against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    changes: [Change; 4],
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field(
                "staged",
                &self
                    .changes
                    .iter()
                    .filter(|change| **change != Change::Keep)
                    .count(),
            )
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn stage(&mut self, field: Field, value: Option<&str>) -> Result<&mut Self, Error> {
        if value.is_some_and(|value| value.contains('\0')) {
            return Err(Error::Refused(Refusal::InvalidValue));
        }
        self.changes[field as usize] = Change::Set(value.map(str::to_owned));
        Ok(self)
    }

    /// Stage the title; `None` removes it.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidValue`] for a value holding NUL.
    pub fn set_title(&mut self, value: Option<&str>) -> Result<&mut Self, Error> {
        self.stage(Field::Title, value)
    }

    /// Stage the author; `None` removes it.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidValue`] for a value holding NUL.
    pub fn set_author(&mut self, value: Option<&str>) -> Result<&mut Self, Error> {
        self.stage(Field::Author, value)
    }

    /// Stage the keyword list; an empty list removes the `Keywords` key.
    ///
    /// The key holds one string: the list is joined with `, ` and read
    /// back split at `;` and `,` (`litchi_core::metadata::keywords`). The
    /// list the package already holds is kept as it is, so its source
    /// separators survive.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidValue`] for a keyword holding NUL, and
    /// [`Refusal::InvalidKeyword`] for one the joined string cannot keep.
    pub fn set_keywords(&mut self, keywords: &[String]) -> Result<&mut Self, Error> {
        if keywords.iter().any(|keyword| keyword.contains('\0')) {
            return Err(Error::Refused(Refusal::InvalidValue));
        }
        if keywords == self.source.metadata().keywords {
            self.changes[Field::Keywords as usize] = Change::Keep;
            return Ok(self);
        }
        let value = if keywords.is_empty() {
            None
        } else {
            Some(
                litchi_core::metadata::keywords::join(keywords)
                    .map_err(|_error| Error::Refused(Refusal::InvalidKeyword))?,
            )
        };
        self.changes[Field::Keywords as usize] = Change::Set(value);
        Ok(self)
    }

    /// Stage the description (the `Comments` key); `None` removes it.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidValue`] for a value holding NUL.
    pub fn set_description(&mut self, value: Option<&str>) -> Result<&mut Self, Error> {
        self.stage(Field::Description, value)
    }

    /// Validate and publish the staged properties.
    ///
    /// An edit whose values equal the current ones is an exact no-op: the
    /// commit returns the source package and a no-op patch.
    ///
    /// # Errors
    ///
    /// A typed error, leaving the source untouched, when the sidecar cannot
    /// be decoded, the encoded properties exceed their ceiling, or the
    /// reopened candidate does not read back the request.
    pub fn commit(self) -> Result<Commit, Error> {
        let current = self.source.metadata();
        let mut expected = current.clone();
        for field in Field::ALL {
            if let Change::Set(value) = &self.changes[field as usize] {
                match field.slot(&mut expected) {
                    Some(slot) => slot.clone_from(value),
                    None => {
                        expected.keywords = value
                            .as_deref()
                            .map(litchi_core::metadata::keywords::split)
                            .unwrap_or_default();
                    },
                }
            }
        }
        let changed = expected != current;
        let source_bytes = self.source.state.source.shared_source();
        if !changed {
            return Ok(Commit::new(
                self.source.clone(),
                Patch {
                    source: Arc::clone(&source_bytes),
                    target: source_bytes,
                },
                Diagnostics::UNCHANGED,
            ));
        }
        let mut dictionary = sidecars::properties(self.source)
            .map_err(|_error| Error::InvalidSource)?
            .unwrap_or_default();
        for field in Field::ALL {
            let Change::Set(value) = &self.changes[field as usize] else {
                continue;
            };
            for key in field.fallbacks() {
                dictionary.remove(key);
            }
            match value {
                Some(value) => {
                    dictionary.insert(field.key().to_owned(), plist::Value::String(value.clone()));
                },
                None => {
                    dictionary.remove(field.key());
                },
            }
        }
        let mut encoded = Vec::new();
        plist::Value::Dictionary(dictionary.clone())
            .to_writer_binary(&mut encoded)
            .map_err(|_error| Error::InvalidSource)?;
        if encoded.len() > MAX_PROPERTIES_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::PropertiesBytes,
                observed: u64::try_from(encoded.len()).unwrap_or(u64::MAX),
                maximum: u64::try_from(MAX_PROPERTIES_BYTES).unwrap_or(u64::MAX),
            });
        }
        let package = sidecars::rewrite(self.source, Some(&dictionary), None)
            .map_err(|_error| Error::Verification)?;
        if package.metadata() != expected {
            return Err(Error::Verification);
        }
        let target = package.state.source.shared_source();
        Ok(Commit::new(
            package,
            Patch {
                source: source_bytes,
                target,
            },
            Diagnostics::CHANGED,
        ))
    }
}

/// An exact-source, reversible document-properties patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("is_noop", &self.is_noop())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        Arc::ptr_eq(&self.source, &self.target) || self.source == self.target
    }

    /// The patch from the target artifact back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
        }
    }
}

/// The verified result of one document-properties transaction.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

impl Package {
    /// Stage a document-properties transaction.
    #[must_use]
    pub fn edit_metadata(&self) -> Edit<'_> {
        Edit {
            source: self,
            changes: Default::default(),
        }
    }

    /// Apply a document-properties patch to the exact package it was made
    /// from and return the replay commit triad.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] unless this package's bytes equal the
    /// patch's source; [`Error::Verification`] when the target fails to
    /// reopen.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit, Error> {
        if self.state.source.source_bytes() != patch.source.as_ref() {
            return Err(Error::PatchConflict);
        }
        let changed = !patch.is_noop();
        let package = if changed {
            let limits = self.state.source.limits();
            Self::from_bytes_with_archive_limits(&patch.target, limits)
                .map_err(|_error: PackageError| Error::Verification)?
        } else {
            self.clone()
        };
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(changed),
        ))
    }
}
