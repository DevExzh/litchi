//! Ergonomic Formula document facade and source-bound transactions.

use litchi_core::{Error, Result};
use std::io::Read;
use std::path::Path;
use std::sync::Arc;

use crate::codec;
use crate::model::Element;

/// An immutable-or-transactionally-edited `OpenDocument` Formula package.
#[derive(Clone)]
pub struct Formula {
    package: crate::package::Package,
    root: Arc<Element>,
    limits: codec::Limits,
}

impl Formula {
    /// Create a standard `.odf` package from validated `MathML`.
    ///
    /// # Errors
    ///
    /// Returns an error when the `MathML` fails validation or the package
    /// cannot be created.
    pub fn create(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_with_limits(mathml, codec::Limits::default())
    }

    /// Create a standard `.odf` package using caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the `MathML` fails validation, a limit is
    /// exceeded, or the package cannot be created.
    pub fn create_with_limits(mathml: impl AsRef<str>, limits: codec::Limits) -> Result<Self> {
        Self::create_with_flavor(mathml.as_ref(), false, limits)
    }

    /// Create a Formula template `.otf` package from validated `MathML`.
    ///
    /// # Errors
    ///
    /// Returns an error when the `MathML` fails validation or the package
    /// cannot be created.
    pub fn create_template(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_template_with_limits(mathml, codec::Limits::default())
    }

    /// Create a Formula template `.otf` package using caller-selected limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the `MathML` fails validation, a limit is
    /// exceeded, or the package cannot be created.
    pub fn create_template_with_limits(
        mathml: impl AsRef<str>,
        limits: codec::Limits,
    ) -> Result<Self> {
        Self::create_with_flavor(mathml.as_ref(), true, limits)
    }

    fn create_with_flavor(mathml: &str, template: bool, limits: codec::Limits) -> Result<Self> {
        let root = codec::parse_with_limits(mathml, limits)?;
        let package = crate::package::Package::create(mathml.as_bytes(), template)?;
        Self::from_validated_parts(package, root, limits)
    }

    /// Open a Formula package from a path.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be opened or its contents are
    /// not a valid Formula package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, codec::Limits::default())
    }

    /// Open a Formula package from a path with caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be opened, a limit is exceeded,
    /// or its contents are not a valid Formula package.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: codec::Limits) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_limits(file, limits)
    }

    /// Read a Formula package from a stream.
    ///
    /// # Errors
    ///
    /// Returns an error when reading fails or the data is not a valid Formula
    /// package.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limits(reader, codec::Limits::default())
    }

    /// Read a Formula package from a stream under caller-selected limits.
    ///
    /// # Errors
    ///
    /// Returns an error when reading fails, the package byte ceiling is
    /// exceeded, or the data is not a valid Formula package.
    pub fn from_reader_with_limits(reader: impl Read, limits: codec::Limits) -> Result<Self> {
        Self::from_bytes_with_limits(read_all(reader, limits.package_bytes())?, limits)
    }

    /// Read and validate a Formula package from owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not a valid Formula package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, codec::Limits::default())
    }

    /// Read and validate owned Formula package bytes under selected limits.
    ///
    /// # Errors
    ///
    /// Returns an error when a resource limit is exceeded or the bytes are
    /// not a valid Formula package.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: codec::Limits) -> Result<Self> {
        check_package_bytes(bytes.len(), limits)?;
        let package = crate::package::Package::from_bytes(bytes)?;
        Self::from_package(package, limits)
    }

    fn from_package(package: crate::package::Package, limits: codec::Limits) -> Result<Self> {
        check_package_bytes(package.as_bytes().len(), limits)?;
        let root = codec::parse_with_limits(&package.content_xml()?, limits)?;
        Self::from_validated_parts(package, root, limits)
    }

    fn from_validated_parts(
        package: crate::package::Package,
        root: Element,
        limits: codec::Limits,
    ) -> Result<Self> {
        check_package_bytes(package.as_bytes().len(), limits)?;
        Ok(Self {
            package,
            root: Arc::new(root),
            limits,
        })
    }

    /// Return the complete `MathML` `math` root.
    #[must_use]
    pub fn root(&self) -> &Element {
        &self.root
    }

    /// Return the finite limits retained for reparsing edited candidates.
    #[must_use]
    pub const fn limits(&self) -> codec::Limits {
        self.limits
    }

    /// Start a source-bound whole-root transaction.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            replacement: None,
        }
    }

    /// Replace the `MathML` root and atomically rebuild the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the serialized root fails re-validation or the
    /// rebuilt package cannot be assembled.
    pub fn set_root(&mut self, root: Element) -> Result<()> {
        let mut edit = self.edit();
        edit.set_root(root)?;
        *self = edit.commit()?.into_formula();
        Ok(())
    }

    /// Apply a reversible patch only when this is its byte-exact source.
    ///
    /// # Errors
    ///
    /// Returns an error when this package differs from the patch source.
    pub fn apply_patch(&self, patch: &Patch) -> Result<Self> {
        patch.apply(self)
    }

    /// Return every `MathML` annotation in document order.
    #[must_use]
    pub fn annotations(&self) -> Vec<&Element> {
        let mut annotations = Vec::new();
        self.root.collect_annotations(&mut annotations);
        annotations
    }

    /// Return the first `StarMath` annotation source, when present.
    #[must_use]
    pub fn starmath_source(&self) -> Option<String> {
        self.annotations().into_iter().find_map(|annotation| {
            annotation
                .attribute(None, "encoding")
                .is_some_and(|encoding| encoding.eq_ignore_ascii_case("StarMath 5.0"))
                .then(|| annotation.all_text())
        })
    }

    /// Return the exact package MIME type.
    #[must_use]
    pub fn mimetype(&self) -> &'static str {
        self.package.mimetype()
    }

    /// Whether this package uses the Formula template MIME type.
    #[must_use]
    pub const fn is_template(&self) -> bool {
        self.package.is_template()
    }

    /// Read the exact package `content.xml` as UTF-8.
    ///
    /// # Errors
    ///
    /// Returns an error when `content.xml` is missing or not valid UTF-8.
    pub fn content_xml(&self) -> Result<String> {
        self.package.content_xml()
    }

    /// List package members in archive order.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive member list cannot be read.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Return the exact original package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the exact package bytes.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Consume the facade and return the exact package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Save the exact package bytes without rebuilding the archive.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be written.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }
}

/// A source-bound whole-root edit.
pub struct Edit<'source> {
    source: &'source Formula,
    replacement: Option<Arc<Element>>,
}

impl Edit<'_> {
    /// Stage a complete checked `MathML` root replacement.
    ///
    /// # Errors
    ///
    /// Returns an error when the root cannot be serialized and reparsed under
    /// the source snapshot's retained limits.
    pub fn set_root(&mut self, root: Element) -> Result<()> {
        let xml = codec::serialize(&root);
        drop(root);
        let validated = Arc::new(codec::parse_with_limits(&xml, self.source.limits)?);
        self.replacement = (validated.as_ref() != self.source.root()).then_some(validated);
        Ok(())
    }

    /// Atomically rebuild, reopen, and semantically verify the candidate.
    ///
    /// # Errors
    ///
    /// Returns an error when package rebuilding, bounded reopen, or semantic
    /// readback fails. The source snapshot is never mutated.
    pub fn commit(self) -> Result<Commit> {
        let Some(replacement) = self.replacement else {
            return Ok(Commit::unchanged(self.source.clone()));
        };
        let xml = codec::serialize(&replacement);
        let package = self.source.package.replace_content(xml.as_bytes())?;
        let snapshot = Formula::from_package(package, self.source.limits)?;
        if snapshot.root() != replacement.as_ref() {
            return Err(Error::InvalidFormat(
                "Formula root edit failed semantic readback".to_string(),
            ));
        }
        let change = RootChange {
            before: Arc::clone(&self.source.root),
            after: replacement,
        };
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                change: Some(change),
            },
            diagnostics: Diagnostics {
                changed: true,
                candidate_reopened: true,
            },
        })
    }
}

/// The semantic before/after roots carried by a reversible patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RootChange {
    before: Arc<Element>,
    after: Arc<Element>,
}

impl RootChange {
    /// The source snapshot's root.
    #[must_use]
    pub fn before(&self) -> &Element {
        &self.before
    }

    /// The committed replacement root.
    #[must_use]
    pub fn after(&self) -> &Element {
        &self.after
    }
}

/// Deterministic diagnostics for one root commit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    candidate_reopened: bool,
}

impl Diagnostics {
    /// Whether the transaction changed package bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Whether a changed candidate was fully reopened before publication.
    #[must_use]
    pub const fn candidate_reopened(self) -> bool {
        self.candidate_reopened
    }
}

/// A committed immutable Formula snapshot and its reversible patch.
pub struct Commit {
    snapshot: Formula,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    fn unchanged(snapshot: Formula) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                change: None,
            },
            snapshot,
            diagnostics: Diagnostics {
                changed: false,
                candidate_reopened: false,
            },
        }
    }

    /// Whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.diagnostics.changed
    }

    /// The published immutable Formula snapshot.
    #[must_use]
    pub const fn formula(&self) -> &Formula {
        &self.snapshot
    }

    /// The source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Validation and publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume this commit into the published snapshot.
    #[must_use]
    pub fn into_formula(self) -> Formula {
        self.snapshot
    }
}

/// A byte-exact source-checked reversible Formula root patch.
#[derive(Clone)]
pub struct Patch {
    source: Formula,
    target: Formula,
    change: Option<RootChange>,
}

impl Patch {
    /// Whether this patch authorizes the supplied exact source package.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Formula) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Apply this patch only to its byte-exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source is stale or unrelated.
    pub fn apply(&self, source: &Formula) -> Result<Formula> {
        if !self.is_applicable_to(source) {
            return Err(Error::InvalidFormat(
                "Formula patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.target.clone())
    }

    /// The semantic root change, or `None` for an exact no-op.
    #[must_use]
    pub const fn change(&self) -> Option<&RootChange> {
        self.change.as_ref()
    }

    /// Return the patch that restores the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            change: self.change.as_ref().map(|change| RootChange {
                before: Arc::clone(&change.after),
                after: Arc::clone(&change.before),
            }),
        }
    }
}

fn check_package_bytes(actual: usize, limits: codec::Limits) -> Result<()> {
    if actual > limits.package_bytes() {
        return Err(Error::InvalidFormat(format!(
            "Formula package has {actual} bytes, exceeding the {} byte limit",
            limits.package_bytes()
        )));
    }
    Ok(())
}

fn read_all(reader: impl Read, maximum: usize) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    let read_limit = u64::try_from(maximum).unwrap_or(u64::MAX).saturating_add(1);
    reader.take(read_limit).read_to_end(&mut bytes)?;
    if bytes.len() > maximum {
        return Err(Error::InvalidFormat(format!(
            "Formula package stream exceeds the {maximum} byte limit"
        )));
    }
    Ok(bytes)
}
