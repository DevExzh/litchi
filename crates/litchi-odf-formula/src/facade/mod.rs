//! Ergonomic Formula document facade and source-bound transactions.

use litchi_core::{Error, Result};
use std::io::Read;
use std::path::Path;
use std::sync::Arc;

use crate::codec;
use crate::model::{Content, Element};

/// Maximum UTF-8 bytes accepted by the validated opaque `StarMath` boundary.
pub const MAX_STARMATH_SOURCE_BYTES: usize = 1024 * 1024;
/// Maximum semantic operations retained by one commit or durable patch.
pub const MAX_SEMANTIC_OPERATIONS: usize = 4_096;
/// Maximum recent commit records retained by an in-memory Formula snapshot.
pub const MAX_COMMIT_HISTORY: usize = 64;
/// Maximum approximate semantic evidence retained across recent commits.
pub const MAX_COMMIT_HISTORY_BYTES: usize = 8 * 1024 * 1024;

const HISTORY_MAGIC: &[u8] = b"LITCHI-ODF-HISTORY\0\x01";
const PATCH_MAGIC: &[u8] = b"LITCHI-ODF-PATCH\0\x01";

/// Validated, inert `StarMath` source.
///
/// `StarMath` is a `LibreOffice` language rather than an `ODF` or `MathML`
/// grammar.
/// This type deliberately validates only a bounded XML-safe opaque payload and
/// a recognized syntax version; it never claims to parse or evaluate source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OpaqueStarMath {
    source: String,
    version: StarMathVersion,
}

impl OpaqueStarMath {
    /// Validate source at the explicit opaque-language boundary.
    ///
    /// # Errors
    ///
    /// Returns an error when source exceeds the fixed byte ceiling or contains
    /// a character forbidden by XML 1.0.
    pub fn new(version: StarMathVersion, source: impl Into<String>) -> Result<Self> {
        let source_text = source.into();
        if source_text.len() > MAX_STARMATH_SOURCE_BYTES {
            return Err(invalid(format!(
                "StarMath source exceeds {MAX_STARMATH_SOURCE_BYTES} bytes"
            )));
        }
        if !source_text.chars().all(is_xml_character) {
            return Err(invalid(
                "StarMath source contains a character forbidden by XML 1.0",
            ));
        }
        Ok(Self {
            source: source_text,
            version,
        })
    }

    /// Inert source text.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Declared syntax version.
    #[must_use]
    pub const fn version(&self) -> StarMathVersion {
        self.version
    }
}

/// `LibreOffice` `StarMath` syntax version carried by a `MathML` annotation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum StarMathVersion {
    /// Legacy `StarMath 5.0` syntax.
    V5,
    /// Current `StarMath 6` syntax recognized by `LibreOffice`'s checked-in
    /// `MathML` importer.
    V6,
}

impl StarMathVersion {
    /// Exact `MathML` `encoding` attribute spelling.
    #[must_use]
    pub const fn encoding(self) -> &'static str {
        match self {
            Self::V5 => "StarMath 5.0",
            Self::V6 => "StarMath 6",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            "StarMath 5.0" => Some(Self::V5),
            "StarMath 6" => Some(Self::V6),
            _ => None,
        }
    }
}

/// Inert, versioned `StarMath` source carried by a `MathML` annotation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StarMathAnnotation<'element> {
    element: &'element Element,
    opaque: OpaqueStarMath,
}

/// One bounded semantic commit retained with a published snapshot.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CommitRecord {
    bytes: usize,
    changes: Arc<[SemanticChange]>,
    source: Revision,
    target: Revision,
}

impl CommitRecord {
    /// Approximate retained semantic evidence bytes.
    #[must_use]
    pub const fn retained_bytes(&self) -> usize {
        self.bytes
    }

    /// Ordered durable operations published by the commit.
    #[must_use]
    pub fn changes(&self) -> &[SemanticChange] {
        &self.changes
    }

    /// Exact source package revision.
    #[must_use]
    pub const fn source_revision(&self) -> Revision {
        self.source
    }

    /// Exact target package revision.
    #[must_use]
    pub const fn target_revision(&self) -> Revision {
        self.target
    }
}

impl<'element> StarMathAnnotation<'element> {
    /// Underlying `MathML` `annotation` element.
    #[must_use]
    pub const fn element(&self) -> &'element Element {
        self.element
    }

    /// Checked opaque `StarMath` source. The crate never evaluates it.
    #[must_use]
    pub const fn opaque(&self) -> &OpaqueStarMath {
        &self.opaque
    }
}

/// An immutable-or-transactionally-edited `OpenDocument` Formula package.
#[derive(Clone)]
pub struct Formula {
    history: Arc<[CommitRecord]>,
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
        validate_starmath_boundaries(&root)?;
        Ok(Self {
            history: Arc::from(Vec::new().into_boxed_slice()),
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

    /// Stable fingerprint of the exact package snapshot.
    #[must_use]
    pub fn revision(&self) -> Revision {
        Revision::from_bytes(self.as_bytes())
    }

    /// Recent semantic commits coupled to this in-memory snapshot.
    ///
    /// The oldest record is evicted after [`MAX_COMMIT_HISTORY`] entries.
    #[must_use]
    pub fn commit_history(&self) -> &[CommitRecord] {
        &self.history
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
            changes: Vec::new(),
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

    /// Return the first validated opaque `StarMath` annotation, when present.
    #[must_use]
    pub fn starmath(&self) -> Option<StarMathAnnotation<'_>> {
        self.starmath_annotations().into_iter().next()
    }

    /// Return recognized, versioned `StarMath` annotations in document order.
    #[must_use]
    pub fn starmath_annotations(&self) -> Vec<StarMathAnnotation<'_>> {
        self.annotations()
            .into_iter()
            .filter_map(|element| {
                let version = StarMathVersion::parse(element.attribute(None, "encoding")?)?;
                let opaque = OpaqueStarMath::new(version, element.all_text()).ok()?;
                Some(StarMathAnnotation { element, opaque })
            })
            .collect()
    }

    fn with_commit_record(mut self, source: &Self, changes: &[SemanticChange]) -> Result<Self> {
        if changes.is_empty() {
            self.history = Arc::clone(&source.history);
            return Ok(self);
        }
        if changes.len() > MAX_SEMANTIC_OPERATIONS {
            return Err(invalid(format!(
                "Formula commit exceeds {MAX_SEMANTIC_OPERATIONS} semantic operations"
            )));
        }
        let source_revision = source.revision();
        let target_revision = self.revision();
        if self.history.last().is_some_and(|record| {
            record.source == source_revision && record.target == target_revision
        }) {
            return Ok(self);
        }
        let retained_bytes = changes.iter().try_fold(0_usize, |total, change| {
            total
                .checked_add(change.retained_bytes())
                .ok_or_else(|| invalid("Formula commit history byte count overflow"))
        })?;
        if retained_bytes > MAX_COMMIT_HISTORY_BYTES {
            return Err(invalid(format!(
                "Formula commit history exceeds {MAX_COMMIT_HISTORY_BYTES} retained bytes"
            )));
        }
        let mut history = source.history.to_vec();
        let mut history_bytes = history.iter().try_fold(0_usize, |total, record| {
            total
                .checked_add(record.bytes)
                .ok_or_else(|| invalid("Formula history byte count overflow"))
        })?;
        while history.len() >= MAX_COMMIT_HISTORY
            || history_bytes.saturating_add(retained_bytes) > MAX_COMMIT_HISTORY_BYTES
        {
            let removed = history.remove(0);
            history_bytes = history_bytes.saturating_sub(removed.bytes);
        }
        history.push(CommitRecord {
            bytes: retained_bytes,
            changes: Arc::from(changes.to_vec().into_boxed_slice()),
            source: source_revision,
            target: target_revision,
        });
        self.history = Arc::from(history.into_boxed_slice());
        Ok(self)
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
    changes: Vec<SemanticChange>,
}

impl Edit<'_> {
    /// Borrow the current transaction-local root.
    #[must_use]
    pub fn root(&self) -> &Element {
        self.replacement
            .as_deref()
            .unwrap_or_else(|| self.source.root())
    }

    /// Semantic operations staged in call order.
    #[must_use]
    pub fn changes(&self) -> &[SemanticChange] {
        &self.changes
    }

    /// Plan transfer of another edit created from the same source snapshot.
    ///
    /// This method never mutates either edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the edits do not share byte-exact source and
    /// resource limits.
    pub fn plan_join(&self, other: &Self) -> Result<DependencyTransfer> {
        self.require_same_source(other)?;
        Ok(DependencyTransfer::plan(&other.changes, &self.changes))
    }

    /// Atomically join a non-conflicting independent edit.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale source, a dependency conflict, an
    /// operation bound, or an invalid joined `MathML` candidate. This edit is
    /// unchanged on failure.
    pub fn join(&mut self, other: &Self) -> Result<()> {
        let transfer = self.plan_join(other)?;
        if let Some(conflict) = transfer.conflicts().first() {
            return Err(invalid(format!(
                "Formula sub-edit conflict {:?} at path {:?}",
                conflict.kind(),
                conflict.path().indices()
            )));
        }
        let operation_count = self
            .changes
            .len()
            .checked_add(transfer.operations().len())
            .ok_or_else(|| invalid("Formula semantic operation count overflow"))?;
        if operation_count > MAX_SEMANTIC_OPERATIONS {
            return Err(invalid(format!(
                "Formula join exceeds {MAX_SEMANTIC_OPERATIONS} semantic operations"
            )));
        }
        let mut candidate = self.root().clone();
        for change in transfer.operations() {
            candidate = apply_semantic_change(&candidate, change)?;
        }
        let validated = validate_candidate(&candidate, self.source.limits)?;
        let mut joined_changes = self.changes.clone();
        joined_changes.extend_from_slice(transfer.operations());
        self.replacement = (validated.as_ref() != self.source.root()).then_some(validated);
        self.changes = if self.replacement.is_some() {
            joined_changes
        } else {
            Vec::new()
        };
        Ok(())
    }

    /// Insert one element child at a stable child-element path.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent parent/index or an invalid candidate.
    pub fn insert_child(&mut self, parent: &NodePath, index: usize, child: Element) -> Result<()> {
        let after_root = transform_at(self.root(), parent.indices(), |element| {
            element.insert_child(index, child.clone())
        })?;
        let path = parent.child(index);
        let validated = validate_candidate(&after_root, self.source.limits)?;
        self.stage(
            validated,
            SemanticChange::new(ChangeKind::Insert, path, None, Some(child)),
        );
        Ok(())
    }

    /// Remove and return one element selected by path.
    ///
    /// # Errors
    ///
    /// Returns an error for the root path, an absent element, or an invalid
    /// candidate content model.
    pub fn remove(&mut self, path: &NodePath) -> Result<Element> {
        let (parent_path, index) = path.parent_and_index()?;
        let before = element_at(self.root(), path.indices())?.clone();
        let after_root = transform_at(self.root(), parent_path, |parent| {
            parent
                .remove_child(index)
                .ok_or_else(|| invalid("MathML removal path is out of range"))?;
            Ok(())
        })?;
        let validated = validate_candidate(&after_root, self.source.limits)?;
        self.stage(
            validated,
            SemanticChange::new(ChangeKind::Remove, path.clone(), Some(before.clone()), None),
        );
        Ok(before)
    }

    /// Replace one element selected by path and return its previous value.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent path or an invalid candidate.
    pub fn replace(&mut self, path: &NodePath, replacement: Element) -> Result<Element> {
        let before = element_at(self.root(), path.indices())?.clone();
        if before == replacement {
            return Ok(before);
        }
        if path.is_root() {
            let validated = validate_candidate(&replacement, self.source.limits)?;
            self.stage(
                validated,
                SemanticChange::new(
                    ChangeKind::Replace,
                    path.clone(),
                    Some(before.clone()),
                    Some(replacement),
                ),
            );
            return Ok(before);
        }
        let (parent_path, index) = path.parent_and_index()?;
        let after_root = transform_at(self.root(), parent_path, |parent| {
            parent
                .replace_child(index, replacement.clone())
                .ok_or_else(|| invalid("MathML replacement path is out of range"))?;
            Ok(())
        })?;
        let validated = validate_candidate(&after_root, self.source.limits)?;
        self.stage(
            validated,
            SemanticChange::new(
                ChangeKind::Replace,
                path.clone(),
                Some(before.clone()),
                Some(replacement),
            ),
        );
        Ok(before)
    }

    /// Set or replace one attribute on the selected element.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent path, invalid name, or invalid value
    /// domain.
    pub fn set_attribute(
        &mut self,
        path: &NodePath,
        namespace_uri: Option<&str>,
        local_name: &str,
        value: &str,
    ) -> Result<()> {
        self.mutate_element(path, ChangeKind::SetAttribute, |element| {
            element.set_attribute(namespace_uri, local_name, value)
        })
    }

    /// Remove one attribute from the selected element.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent path or invalid resulting content.
    pub fn remove_attribute(
        &mut self,
        path: &NodePath,
        namespace_uri: Option<&str>,
        local_name: &str,
    ) -> Result<bool> {
        let existed = element_at(self.root(), path.indices())?
            .attribute(namespace_uri, local_name)
            .is_some();
        if existed {
            self.mutate_element(path, ChangeKind::RemoveAttribute, |element| {
                element.remove_attribute(namespace_uri, local_name);
                Ok(())
            })?;
        }
        Ok(existed)
    }

    /// Replace all content of a token or annotation with one text run.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent path or a content model that does not
    /// permit direct character data.
    pub fn set_text(&mut self, path: &NodePath, text: &str) -> Result<()> {
        self.mutate_element(path, ChangeKind::SetText, |element| {
            element.clear_content();
            element.push_text(text);
            Ok(())
        })
    }

    /// Replace the first recognized annotation with validated opaque source.
    ///
    /// # Errors
    ///
    /// Returns an error when the formula has no recognized annotation or the
    /// resulting tree violates retained limits.
    pub fn set_starmath(&mut self, starmath: &OpaqueStarMath) -> Result<()> {
        let path = find_starmath_path(self.root())
            .ok_or_else(|| invalid("Formula has no recognized StarMath annotation"))?;
        self.mutate_element(&path, ChangeKind::SetText, |annotation| {
            annotation.set_attribute(None, "encoding", starmath.version().encoding())?;
            annotation.clear_content();
            annotation.push_text(starmath.source());
            Ok(())
        })
    }

    /// Stage a complete checked `MathML` root replacement.
    ///
    /// # Errors
    ///
    /// Returns an error when the root cannot be serialized and reparsed under
    /// the source snapshot's retained limits.
    pub fn set_root(&mut self, root: Element) -> Result<()> {
        let path = NodePath::root();
        let _previous = self.replace(&path, root)?;
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
        if self.changes.len() > MAX_SEMANTIC_OPERATIONS {
            return Err(invalid(format!(
                "Formula commit exceeds {MAX_SEMANTIC_OPERATIONS} semantic operations"
            )));
        }
        let xml = codec::serialize(&replacement);
        let package = self.source.package.replace_content(xml.as_bytes())?;
        let reopened = Formula::from_package(package, self.source.limits)?;
        if reopened.root() != replacement.as_ref() {
            return Err(Error::InvalidFormat(
                "Formula root edit failed semantic readback".to_string(),
            ));
        }
        let change = RootChange {
            before: Arc::clone(&self.source.root),
            after: replacement,
        };
        let snapshot = reopened.with_commit_record(self.source, &self.changes)?;
        let record = snapshot.history.last().cloned();
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                change: Some(change),
                changes: self.changes,
            },
            diagnostics: Diagnostics {
                changed: true,
                candidate_reopened: true,
            },
            record,
        })
    }

    fn mutate_element(
        &mut self,
        path: &NodePath,
        kind: ChangeKind,
        mutation: impl FnOnce(&mut Element) -> Result<()>,
    ) -> Result<()> {
        let before = element_at(self.root(), path.indices())?.clone();
        let after_root = transform_at(self.root(), path.indices(), mutation)?;
        let validated = validate_candidate(&after_root, self.source.limits)?;
        let after = element_at(&validated, path.indices())?.clone();
        if before != after {
            self.stage(
                validated,
                SemanticChange::new(kind, path.clone(), Some(before), Some(after)),
            );
        }
        Ok(())
    }

    fn require_same_source(&self, other: &Self) -> Result<()> {
        if self.source.as_bytes() == other.source.as_bytes()
            && self.source.limits == other.source.limits
        {
            Ok(())
        } else {
            Err(invalid("Formula sub-edits do not share an exact source"))
        }
    }

    fn stage(&mut self, validated: Arc<Element>, change: SemanticChange) {
        self.replacement = (validated.as_ref() != self.source.root()).then_some(validated);
        if self.replacement.is_some() {
            self.changes.push(change);
        } else {
            self.changes.clear();
        }
    }
}

/// Stable element-child path from the `MathML` root.
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct NodePath(Vec<usize>);

impl NodePath {
    /// The root element path.
    #[must_use]
    pub const fn root() -> Self {
        Self(Vec::new())
    }

    /// Build a checked path from child-element indices.
    #[must_use]
    pub fn new(indices: impl Into<Vec<usize>>) -> Self {
        Self(indices.into())
    }

    /// Return a path to one direct child.
    #[must_use]
    pub fn child(&self, index: usize) -> Self {
        let mut indices = self.0.clone();
        indices.push(index);
        Self(indices)
    }

    /// Child-element indices in root-to-leaf order.
    #[must_use]
    pub fn indices(&self) -> &[usize] {
        &self.0
    }

    /// Whether this path selects the root.
    #[must_use]
    pub fn is_root(&self) -> bool {
        self.0.is_empty()
    }

    fn parent_and_index(&self) -> Result<(&[usize], usize)> {
        let (index, parent) = self
            .0
            .split_last()
            .ok_or_else(|| invalid("the MathML root has no parent"))?;
        Ok((parent, *index))
    }
}

/// Category of one granular `MathML` semantic operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ChangeKind {
    Insert,
    Remove,
    Replace,
    SetAttribute,
    RemoveAttribute,
    SetText,
}

/// Why an independent semantic operation cannot be transferred.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum DependencyConflictKind {
    ConcurrentInsert,
    RemovedDependency,
    ReplacedDependency,
    SameTarget,
}

/// One conflict found without mutating either independent edit.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DependencyConflict {
    dependency_index: usize,
    incoming_index: usize,
    kind: DependencyConflictKind,
    path: NodePath,
}

impl DependencyConflict {
    /// Existing/dependency operation index.
    #[must_use]
    pub const fn dependency_index(&self) -> usize {
        self.dependency_index
    }

    /// Incoming operation index.
    #[must_use]
    pub const fn incoming_index(&self) -> usize {
        self.incoming_index
    }

    /// Conflict category.
    #[must_use]
    pub const fn kind(&self) -> DependencyConflictKind {
        self.kind
    }

    /// Incoming path at the point where transfer stopped.
    #[must_use]
    pub const fn path(&self) -> &NodePath {
        &self.path
    }
}

/// Non-mutating dependency-transfer result for independent operations.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DependencyTransfer {
    conflicts: Vec<DependencyConflict>,
    operations: Vec<SemanticChange>,
}

impl DependencyTransfer {
    /// Plan path transfer without changing either operation sequence.
    #[must_use]
    pub fn plan(incoming: &[SemanticChange], dependencies: &[SemanticChange]) -> Self {
        let mut operations = Vec::with_capacity(incoming.len());
        let mut conflicts = Vec::new();
        for (incoming_index, change) in incoming.iter().enumerate() {
            match transfer_change(change, incoming_index, dependencies) {
                Ok(transferred) => operations.push(transferred),
                Err(conflict) => conflicts.push(conflict),
            }
        }
        Self {
            conflicts,
            operations,
        }
    }

    /// Conflicts discovered during planning.
    #[must_use]
    pub fn conflicts(&self) -> &[DependencyConflict] {
        &self.conflicts
    }

    /// Successfully transferred operations, in incoming order.
    #[must_use]
    pub fn operations(&self) -> &[SemanticChange] {
        &self.operations
    }

    /// Whether every incoming operation transferred.
    #[must_use]
    pub fn is_complete(&self) -> bool {
        self.conflicts.is_empty()
    }
}

/// One reversible, path-addressed `MathML` semantic operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SemanticChange {
    kind: ChangeKind,
    path: NodePath,
    before: Option<Element>,
    after: Option<Element>,
}

impl SemanticChange {
    fn new(
        kind: ChangeKind,
        path: NodePath,
        before: Option<Element>,
        after: Option<Element>,
    ) -> Self {
        Self {
            kind,
            path,
            before,
            after,
        }
    }

    /// Operation category.
    #[must_use]
    pub const fn kind(&self) -> ChangeKind {
        self.kind
    }

    /// Stable element path affected by the operation.
    #[must_use]
    pub const fn path(&self) -> &NodePath {
        &self.path
    }

    /// Element state before the operation, when present.
    #[must_use]
    pub const fn before(&self) -> Option<&Element> {
        self.before.as_ref()
    }

    /// Element state after the operation, when present.
    #[must_use]
    pub const fn after(&self) -> Option<&Element> {
        self.after.as_ref()
    }

    fn inverse(&self) -> Self {
        let kind = match self.kind {
            ChangeKind::Insert => ChangeKind::Remove,
            ChangeKind::Remove => ChangeKind::Insert,
            other @ (ChangeKind::Replace
            | ChangeKind::SetAttribute
            | ChangeKind::RemoveAttribute
            | ChangeKind::SetText) => other,
        };
        Self {
            kind,
            path: self.path.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    fn retained_bytes(&self) -> usize {
        let path_bytes = self.path.indices().len().saturating_mul(size_of::<usize>());
        [self.before.as_ref(), self.after.as_ref()]
            .into_iter()
            .flatten()
            .fold(path_bytes, |total, element| {
                total.saturating_add(element_retained_bytes(element))
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
    record: Option<CommitRecord>,
}

impl Commit {
    fn unchanged(snapshot: Formula) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                change: None,
                changes: Vec::new(),
            },
            snapshot,
            diagnostics: Diagnostics {
                changed: false,
                candidate_reopened: false,
            },
            record: None,
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

    /// Commit-coupled semantic record, absent for an exact no-op.
    #[must_use]
    pub const fn record(&self) -> Option<&CommitRecord> {
        self.record.as_ref()
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
    changes: Vec<SemanticChange>,
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
        self.target
            .clone()
            .with_commit_record(source, &self.changes)
    }

    /// The semantic root change, or `None` for an exact no-op.
    #[must_use]
    pub const fn change(&self) -> Option<&RootChange> {
        self.change.as_ref()
    }

    /// Granular semantic operations in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[SemanticChange] {
        &self.changes
    }

    /// Stable fingerprint of the exact expected source package.
    #[must_use]
    pub fn source_revision(&self) -> Revision {
        Revision::from_bytes(self.source.as_bytes())
    }

    /// Stable fingerprint of the exact target package.
    #[must_use]
    pub fn target_revision(&self) -> Revision {
        Revision::from_bytes(self.target.as_bytes())
    }

    /// Plan a three-way semantic join without publishing package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error unless both patches use `base` as their byte-exact
    /// source or the non-conflicting candidate fails validation.
    pub fn plan_three_way(base: &Formula, left: &Self, right: &Self) -> Result<ThreeWayPlan> {
        ThreeWayPlan::new(base, left, right)
    }

    /// Serialize this patch into a durable, bounded binary envelope.
    ///
    /// # Errors
    ///
    /// Returns an error if a length cannot be represented by the wire format.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = WireWriter::new();
        writer.bytes.extend_from_slice(PATCH_MAGIC);
        writer.write_blob(self.source.as_bytes())?;
        writer.write_blob(self.target.as_bytes())?;
        writer.write_usize(self.changes.len())?;
        for change in &self.changes {
            writer.write_change(change)?;
        }
        Ok(writer.bytes)
    }

    /// Reopen a durable patch and verify its package and semantic evidence.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, oversized, or inconsistent data.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let mut reader = WireReader::new(bytes);
        reader.expect_magic(PATCH_MAGIC)?;
        let source = Formula::from_bytes(reader.read_blob()?.to_vec())?;
        let target = Formula::from_bytes(reader.read_blob()?.to_vec())?;
        let count = reader.read_usize()?;
        if count > source.limits().nodes().min(MAX_SEMANTIC_OPERATIONS) {
            return Err(invalid("Formula patch has too many semantic changes"));
        }
        let mut changes = Vec::with_capacity(count);
        for _index in 0..count {
            changes.push(reader.read_change(source.limits())?);
        }
        reader.finish()?;
        let change = (source.root() != target.root()).then(|| RootChange {
            before: Arc::clone(&source.root),
            after: Arc::clone(&target.root),
        });
        let patch = Self {
            source,
            target,
            change,
            changes,
        };
        patch.verify_changes()?;
        Ok(patch)
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
            changes: self
                .changes
                .iter()
                .rev()
                .map(SemanticChange::inverse)
                .collect(),
        }
    }

    fn verify_changes(&self) -> Result<()> {
        if self.change.is_none() && !self.changes.is_empty() {
            return Err(invalid("no-op Formula patch contains semantic changes"));
        }
        let mut candidate = self.source.root().clone();
        for change in &self.changes {
            if let Some(before) = change.before()
                && before.namespace_uri() == Some(crate::model::NAMESPACE)
            {
                codec::validate_subtree(before)?;
            }
            if let Some(after) = change.after()
                && after.namespace_uri() == Some(crate::model::NAMESPACE)
            {
                codec::validate_subtree(after)?;
            }
            candidate = apply_semantic_change(&candidate, change)?;
        }
        codec::validate(&candidate)?;
        if &candidate == self.target.root() {
            Ok(())
        } else {
            Err(invalid(
                "Formula patch semantic history does not produce its target",
            ))
        }
    }
}

/// Non-mutating three-way plan over two patches from one exact base.
#[derive(Clone)]
pub struct ThreeWayPlan {
    base: Formula,
    candidate: Option<Arc<Element>>,
    changes: Vec<SemanticChange>,
    conflicts: Vec<DependencyConflict>,
    left: Revision,
    right: Revision,
}

impl ThreeWayPlan {
    fn new(base: &Formula, left: &Patch, right: &Patch) -> Result<Self> {
        if !left.is_applicable_to(base) || !right.is_applicable_to(base) {
            return Err(invalid(
                "Formula three-way patches do not share the supplied exact base",
            ));
        }
        let transfer = DependencyTransfer::plan(right.changes(), left.changes());
        let operation_count = left
            .changes()
            .len()
            .checked_add(transfer.operations().len())
            .ok_or_else(|| invalid("Formula three-way operation count overflow"))?;
        if operation_count > MAX_SEMANTIC_OPERATIONS {
            return Err(invalid(format!(
                "Formula three-way plan exceeds {MAX_SEMANTIC_OPERATIONS} operations"
            )));
        }
        let mut changes = left.changes().to_vec();
        changes.extend_from_slice(transfer.operations());
        let candidate = if transfer.is_complete() {
            let mut root = base.root().clone();
            for change in &changes {
                root = apply_semantic_change(&root, change)?;
            }
            Some(validate_candidate(&root, base.limits())?)
        } else {
            None
        };
        Ok(Self {
            base: base.clone(),
            candidate,
            changes,
            conflicts: transfer.conflicts,
            left: left.target_revision(),
            right: right.target_revision(),
        })
    }

    /// Validated joined root, absent when conflicts remain.
    #[must_use]
    pub fn candidate_root(&self) -> Option<&Element> {
        self.candidate.as_deref()
    }

    /// Planned operations in publication order.
    #[must_use]
    pub fn changes(&self) -> &[SemanticChange] {
        &self.changes
    }

    /// Conflicts found without changing any package.
    #[must_use]
    pub fn conflicts(&self) -> &[DependencyConflict] {
        &self.conflicts
    }

    /// Left target revision used for this plan.
    #[must_use]
    pub const fn left_revision(&self) -> Revision {
        self.left
    }

    /// Whether this plan can be published.
    #[must_use]
    pub fn is_publishable(&self) -> bool {
        self.candidate.is_some() && self.conflicts.is_empty()
    }

    /// Atomically publish a conflict-free plan and fully reopen the package.
    ///
    /// # Errors
    ///
    /// Returns an error when conflicts exist or publication/readback fails.
    pub fn publish(&self) -> Result<Commit> {
        let candidate = self
            .candidate
            .as_ref()
            .ok_or_else(|| invalid("Formula three-way plan has unresolved conflicts"))?;
        let edit = Edit {
            source: &self.base,
            replacement: Some(Arc::clone(candidate)),
            changes: self.changes.clone(),
        };
        edit.commit()
    }

    /// Right target revision used for this plan.
    #[must_use]
    pub const fn right_revision(&self) -> Revision {
        self.right
    }
}

/// Stable FNV-1a fingerprint of exact package bytes.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325_u64 ^ bytes.len() as u64;
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Compact revision value for optimistic source checks.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// Ordered, durable chain of source-checked Formula patches.
#[derive(Clone, Default)]
pub struct History {
    patches: Vec<Patch>,
}

impl History {
    /// Create an empty history.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            patches: Vec::new(),
        }
    }

    /// Append a patch whose source is the current history tip.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch does not continue the chain.
    pub fn push(&mut self, patch: Patch) -> Result<()> {
        if let Some(previous) = self.patches.last()
            && previous.target.as_bytes() != patch.source.as_bytes()
        {
            return Err(invalid("Formula history patch does not continue its tip"));
        }
        self.patches.push(patch);
        Ok(())
    }

    /// Ordered patches.
    #[must_use]
    pub fn patches(&self) -> &[Patch] {
        &self.patches
    }

    /// Apply every entry to an exact history origin.
    ///
    /// # Errors
    ///
    /// Returns an error when the origin or any link is stale.
    pub fn apply(&self, origin: &Formula) -> Result<Formula> {
        let mut snapshot = origin.clone();
        for patch in &self.patches {
            snapshot = patch.apply(&snapshot)?;
        }
        Ok(snapshot)
    }

    /// Serialize all entries into a durable sidecar envelope.
    ///
    /// # Errors
    ///
    /// Returns an error if a size cannot be represented.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = WireWriter::new();
        writer.bytes.extend_from_slice(HISTORY_MAGIC);
        writer.write_usize(self.patches.len())?;
        for patch in &self.patches {
            writer.write_blob(&patch.to_bytes()?)?;
        }
        Ok(writer.bytes)
    }

    /// Parse and source-check a durable history sidecar.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, oversized, or disconnected history.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let mut reader = WireReader::new(bytes);
        reader.expect_magic(HISTORY_MAGIC)?;
        let count = reader.read_usize()?;
        if count > codec::HARD_MAX_NODES {
            return Err(invalid("Formula history contains too many patches"));
        }
        let mut history = Self::new();
        for _index in 0..count {
            history.push(Patch::from_bytes(reader.read_blob()?)?)?;
        }
        reader.finish()?;
        Ok(history)
    }
}

struct WireReader<'bytes> {
    bytes: &'bytes [u8],
    position: usize,
}

impl<'bytes> WireReader<'bytes> {
    const fn new(bytes: &'bytes [u8]) -> Self {
        Self { bytes, position: 0 }
    }

    fn expect_magic(&mut self, expected: &[u8]) -> Result<()> {
        if self.take(expected.len())? == expected {
            Ok(())
        } else {
            Err(invalid("Formula durable envelope has an invalid signature"))
        }
    }

    fn finish(self) -> Result<()> {
        if self.position == self.bytes.len() {
            Ok(())
        } else {
            Err(invalid("Formula durable envelope has trailing bytes"))
        }
    }

    fn read_blob(&mut self) -> Result<&'bytes [u8]> {
        let length = self.read_usize()?;
        self.take(length)
    }

    fn read_change(&mut self, limits: codec::Limits) -> Result<SemanticChange> {
        let kind = match self.read_u8()? {
            0 => ChangeKind::Insert,
            1 => ChangeKind::Remove,
            2 => ChangeKind::Replace,
            3 => ChangeKind::SetAttribute,
            4 => ChangeKind::RemoveAttribute,
            5 => ChangeKind::SetText,
            _ => return Err(invalid("Formula patch has an unknown change kind")),
        };
        let path_length = self.read_usize()?;
        if path_length > limits.depth() {
            return Err(invalid("Formula patch path exceeds its depth limit"));
        }
        let mut indices = Vec::with_capacity(path_length);
        for _index in 0..path_length {
            indices.push(self.read_usize()?);
        }
        let before = self.read_element(limits)?;
        let after = self.read_element(limits)?;
        Ok(SemanticChange::new(
            kind,
            NodePath::new(indices),
            before,
            after,
        ))
    }

    fn read_element(&mut self, limits: codec::Limits) -> Result<Option<Element>> {
        match self.read_u8()? {
            0 => Ok(None),
            1 => {
                let bytes = self.read_blob()?;
                let xml = std::str::from_utf8(bytes).map_err(|_error| {
                    invalid("Formula patch element is not encoded as UTF-8 XML")
                })?;
                codec::parse_fragment_with_limits(xml, limits).map(Some)
            },
            _ => Err(invalid("Formula patch has an invalid optional-element tag")),
        }
    }

    fn read_u8(&mut self) -> Result<u8> {
        self.take(1)?
            .first()
            .copied()
            .ok_or_else(|| invalid("Formula durable envelope is truncated"))
    }

    fn read_usize(&mut self) -> Result<usize> {
        let raw: [u8; 8] = self
            .take(8)?
            .try_into()
            .map_err(|_error| invalid("Formula durable integer is truncated"))?;
        usize::try_from(u64::from_le_bytes(raw))
            .map_err(|_error| invalid("Formula durable integer exceeds this platform"))
    }

    fn take(&mut self, length: usize) -> Result<&'bytes [u8]> {
        let end = self
            .position
            .checked_add(length)
            .ok_or_else(|| invalid("Formula durable envelope offset overflow"))?;
        let value = self
            .bytes
            .get(self.position..end)
            .ok_or_else(|| invalid("Formula durable envelope is truncated"))?;
        self.position = end;
        Ok(value)
    }
}

struct WireWriter {
    bytes: Vec<u8>,
}

impl WireWriter {
    const fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    fn write_blob(&mut self, value: &[u8]) -> Result<()> {
        self.write_usize(value.len())?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn write_change(&mut self, change: &SemanticChange) -> Result<()> {
        self.bytes.push(match change.kind {
            ChangeKind::Insert => 0,
            ChangeKind::Remove => 1,
            ChangeKind::Replace => 2,
            ChangeKind::SetAttribute => 3,
            ChangeKind::RemoveAttribute => 4,
            ChangeKind::SetText => 5,
        });
        self.write_usize(change.path.indices().len())?;
        for index in change.path.indices() {
            self.write_usize(*index)?;
        }
        self.write_element(change.before())?;
        self.write_element(change.after())?;
        Ok(())
    }

    fn write_element(&mut self, value: Option<&Element>) -> Result<()> {
        if let Some(element) = value {
            self.bytes.push(1);
            self.write_blob(codec::serialize(element).as_bytes())?;
        } else {
            self.bytes.push(0);
        }
        Ok(())
    }

    fn write_usize(&mut self, value: usize) -> Result<()> {
        let encoded = u64::try_from(value)
            .map_err(|_error| invalid("Formula durable length exceeds the wire format"))?;
        self.bytes.extend_from_slice(&encoded.to_le_bytes());
        Ok(())
    }
}

fn element_at<'element>(root: &'element Element, path: &[usize]) -> Result<&'element Element> {
    let mut current = root;
    for index in path {
        current = current
            .children()
            .nth(*index)
            .ok_or_else(|| invalid("MathML element path is out of range"))?;
    }
    Ok(current)
}

fn element_retained_bytes(root: &Element) -> usize {
    let mut retained = 0_usize;
    let mut pending = vec![root];
    while let Some(element) = pending.pop() {
        retained = retained
            .saturating_add(element.local_name().len())
            .saturating_add(element.namespace_uri().map_or(0, str::len));
        for attribute in element.attributes() {
            retained = retained
                .saturating_add(attribute.local_name().len())
                .saturating_add(attribute.namespace_uri().map_or(0, str::len))
                .saturating_add(attribute.value().len());
        }
        for content in element.content() {
            match content {
                Content::Text(text) => retained = retained.saturating_add(text.len()),
                Content::Element(child) => pending.push(child),
            }
        }
    }
    retained
}

fn apply_semantic_change(root: &Element, change: &SemanticChange) -> Result<Element> {
    match change.kind {
        ChangeKind::Insert => {
            let (parent_path, index) = change.path.parent_and_index()?;
            let after = change
                .after
                .clone()
                .ok_or_else(|| invalid("Formula insert history has no inserted element"))?;
            transform_at(root, parent_path, |parent| {
                parent.insert_child(index, after)
            })
        },
        ChangeKind::Remove => {
            let before = change
                .before
                .as_ref()
                .ok_or_else(|| invalid("Formula remove history has no prior element"))?;
            if element_at(root, change.path.indices())? != before {
                return Err(invalid(
                    "Formula remove history prior element does not match",
                ));
            }
            let (parent_path, index) = change.path.parent_and_index()?;
            transform_at(root, parent_path, |parent| {
                parent
                    .remove_child(index)
                    .ok_or_else(|| invalid("Formula remove history path is out of range"))?;
                Ok(())
            })
        },
        ChangeKind::Replace
        | ChangeKind::SetAttribute
        | ChangeKind::RemoveAttribute
        | ChangeKind::SetText => {
            let before = change
                .before
                .as_ref()
                .ok_or_else(|| invalid("Formula replacement history has no prior element"))?;
            if element_at(root, change.path.indices())? != before {
                return Err(invalid(
                    "Formula replacement history prior element does not match",
                ));
            }
            let after = change
                .after
                .clone()
                .ok_or_else(|| invalid("Formula replacement history has no target element"))?;
            transform_at(root, change.path.indices(), |element| {
                *element = after;
                Ok(())
            })
        },
    }
}

fn transfer_change(
    change: &SemanticChange,
    incoming_index: usize,
    dependencies: &[SemanticChange],
) -> std::result::Result<SemanticChange, DependencyConflict> {
    let mut transferred = change.clone();
    for (dependency_index, dependency) in dependencies.iter().enumerate() {
        if let Err(kind) = transfer_path(&mut transferred.path.0, transferred.kind, dependency) {
            return Err(DependencyConflict {
                dependency_index,
                incoming_index,
                kind,
                path: transferred.path,
            });
        }
    }
    Ok(transferred)
}

fn transfer_path(
    path: &mut [usize],
    incoming_kind: ChangeKind,
    dependency: &SemanticChange,
) -> std::result::Result<(), DependencyConflictKind> {
    if incoming_kind != ChangeKind::Insert
        && path.len() < dependency.path.indices().len()
        && dependency.path.indices().starts_with(path)
    {
        return Err(DependencyConflictKind::ReplacedDependency);
    }
    match dependency.kind {
        ChangeKind::Insert => {
            let Some((inserted_index, parent)) = dependency.path.0.split_last() else {
                return Err(DependencyConflictKind::ReplacedDependency);
            };
            if incoming_kind == ChangeKind::Insert && path == dependency.path.indices() {
                return Err(DependencyConflictKind::ConcurrentInsert);
            }
            adjust_after_insert(path, parent, *inserted_index);
            Ok(())
        },
        ChangeKind::Remove => {
            let Some((removed_index, parent)) = dependency.path.0.split_last() else {
                return Err(DependencyConflictKind::RemovedDependency);
            };
            let removed_path = dependency.path.indices();
            let targets_removed_descendant = path.starts_with(removed_path)
                && (incoming_kind != ChangeKind::Insert || path.len() > removed_path.len());
            if targets_removed_descendant {
                return Err(DependencyConflictKind::RemovedDependency);
            }
            adjust_after_remove(path, parent, *removed_index);
            Ok(())
        },
        ChangeKind::Replace => {
            if path == dependency.path.indices() || path.starts_with(dependency.path.indices()) {
                Err(DependencyConflictKind::ReplacedDependency)
            } else {
                Ok(())
            }
        },
        ChangeKind::SetAttribute | ChangeKind::RemoveAttribute | ChangeKind::SetText => {
            if path == dependency.path.indices() {
                Err(DependencyConflictKind::SameTarget)
            } else {
                Ok(())
            }
        },
    }
}

fn adjust_after_insert(path: &mut [usize], parent: &[usize], inserted_index: usize) {
    if path.len() > parent.len() && path.starts_with(parent) && path[parent.len()] >= inserted_index
    {
        path[parent.len()] = path[parent.len()].saturating_add(1);
    }
}

fn adjust_after_remove(path: &mut [usize], parent: &[usize], removed_index: usize) {
    if path.len() > parent.len() && path.starts_with(parent) && path[parent.len()] > removed_index {
        path[parent.len()] = path[parent.len()].saturating_sub(1);
    }
}

fn find_starmath_path(root: &Element) -> Option<NodePath> {
    let mut pending = vec![(root, NodePath::root())];
    while let Some((element, path)) = pending.pop() {
        if element.namespace_uri() == Some(crate::model::NAMESPACE)
            && element.local_name() == "annotation"
            && element
                .attribute(None, "encoding")
                .and_then(StarMathVersion::parse)
                .is_some()
        {
            return Some(path);
        }
        let children: Vec<_> = element.children().collect();
        for (index, child) in children.into_iter().enumerate().rev() {
            pending.push((child, path.child(index)));
        }
    }
    None
}

fn transform_at(
    root: &Element,
    path: &[usize],
    mutation: impl FnOnce(&mut Element) -> Result<()>,
) -> Result<Element> {
    let mut current = root.clone();
    let mut ancestors = Vec::with_capacity(path.len());
    for index in path {
        let child = current
            .children()
            .nth(*index)
            .cloned()
            .ok_or_else(|| invalid("MathML element path is out of range"))?;
        ancestors.push((current, *index));
        current = child;
    }
    mutation(&mut current)?;
    for (mut parent, index) in ancestors.into_iter().rev() {
        parent
            .replace_child(index, current)
            .ok_or_else(|| invalid("MathML element path changed during mutation"))?;
        current = parent;
    }
    Ok(current)
}

fn validate_candidate(root: &Element, limits: codec::Limits) -> Result<Arc<Element>> {
    let xml = codec::serialize(root);
    let parsed = codec::parse_with_limits(&xml, limits)?;
    validate_starmath_boundaries(&parsed)?;
    Ok(Arc::new(parsed))
}

fn validate_starmath_boundaries(root: &Element) -> Result<()> {
    let mut pending = vec![root];
    while let Some(element) = pending.pop() {
        if element.namespace_uri() == Some(crate::model::NAMESPACE)
            && element.local_name() == "annotation"
            && let Some(version) = element
                .attribute(None, "encoding")
                .and_then(StarMathVersion::parse)
        {
            let _opaque = OpaqueStarMath::new(version, element.all_text())?;
        }
        pending.extend(element.children());
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn is_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&character)
        || ('\u{E000}'..='\u{FFFD}').contains(&character)
        || ('\u{10000}'..='\u{10FFFF}').contains(&character)
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
