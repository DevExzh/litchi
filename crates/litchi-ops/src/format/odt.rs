//! Handle for OpenDocument Text (`.odt`) inputs.
//!
//! The handle retains the leaf read owner detection built — the eager
//! `litchi_odt::Document` of the byte and reader routes, with its immutable
//! `litchi_odt::Package`, or the filesystem-backed
//! `litchi_odt::SourceBackedDocument` of the path route — and reads through
//! the leaf body index (`litchi_odt::body`). Edits stage typed body edits on
//! the leaf `Package::edit()` transaction; a filesystem-backed handle
//! materializes its package once, on its first edit or publication.

pub(crate) mod children;
mod nav;
mod node;
mod parts;
mod props;
mod spec;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use props::Props;
pub use spec::Kind;
pub(crate) use spec::{ADDABLE, REMOVABLE, SETTABLE, SPEC, WRITE_GAPS};

use std::fmt;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;
use litchi_odf_common::leniency::ToleranceReport;
use litchi_odt::body::{Body, NodeId, Tree};

use crate::{Error, OpenOptions, Result};

/// An opened ODT document. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Odt(Arc<OdtState>);

struct OdtState {
    owner: OdtOwner,
    /// The caller-selected ODF ingress policy retained across publication
    /// and raw package reads.
    odf_options: litchi_odf_common::OpenOptions,
    /// The package's format: ODT, or an ODF text-family format read through
    /// the same content model (OTH web template, ODM master document,
    /// record 0282).
    format: FileFormat,
    /// The immutable leaf package: set for eager owners, materialized on
    /// first use for a filesystem-backed owner.
    package: OnceLock<litchi_odt::Package>,
    /// The body index of a filesystem-backed owner (an eager document keeps
    /// its own).
    tree: OnceLock<Tree>,
    /// Child identities by parent and unified kind, built on the first
    /// ordinal lookup past [`DIRECT_SCAN`] children.
    kinds: OnceLock<KindIndex>,
    /// The part index of record 0328 (list styles, master-page regions,
    /// tables of contents, shapes, objects and columns), its body fields and
    /// its equations, built on the first read of a part.
    parts: OnceLock<parts::Index>,
    origin: Option<Arc<Path>>,
    /// The flat single-XML document (`.fodt`) this handle reads, whose
    /// bytes it publishes; the package is its read projection (record
    /// 0292).
    flat: Option<litchi_odt::flat::Document>,
}

/// Ordinal lookups below this many same-kind children scan the parent;
/// deeper ones read the per-handle [`KindIndex`].
const DIRECT_SCAN: usize = 32;

/// The children of every body node by unified kind, in document order.
struct KindIndex {
    children: std::collections::HashMap<(NodeId, u8), Vec<NodeId>>,
}

impl KindIndex {
    fn build(body: Body<'_>) -> Self {
        let mut children: std::collections::HashMap<(NodeId, u8), Vec<NodeId>> =
            std::collections::HashMap::new();
        for node in body.nodes() {
            if let (Some(parent), Some(kind)) = (node.parent(), Kind::of(node.kind())) {
                children
                    .entry((parent.id(), kind as u8))
                    .or_default()
                    .push(node.id());
            }
        }
        Self { children }
    }
}

/// The ODT read owner, in the state today's route builds.
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to that open route"
)]
pub(crate) enum OdtOwner {
    /// Parsed from the package detection prepared (byte and reader routes,
    /// the path fallback on other targets, and every edit result).
    Eager(litchi_odt::Document),
    /// Filesystem-backed, with deferred package payloads (path route).
    #[cfg(any(unix, windows))]
    Source(litchi_odt::SourceBackedDocument),
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Odt>();
    assert!(size_of::<Odt>() == size_of::<usize>());
};

/// Typed ODF limit failures stay `Source::Odf`; every other failure keeps
/// its common-core payload as `Source::Core`, exactly as before W3.
pub(crate) fn odf_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => core_error(error),
        other => Error::from(other).with_format(FileFormat::Odt),
    }
}

pub(crate) fn core_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Odt)
}

impl Odt {
    /// Adopts the package detection prepared (`DetectedFormat::Odt`) without
    /// a second ZIP index. The ODF limits were applied when the package was
    /// prepared.
    pub(crate) fn from_prepared(prepared: PreparedPackage, options: &OpenOptions) -> Result<Self> {
        let format = prepared.format();
        let odf_options = options.odf_open_options();
        let package = if format == FileFormat::Odt {
            litchi_odt::Package::from_prepared_package_with_options(prepared, odf_options)
        } else {
            litchi_odt::Package::from_prepared_text_package_with_options(prepared, odf_options)
        }
        .map_err(|error| odf_error(error).with_format(format))?;
        Self::from_package_with_options(package, None, format, odf_options)
    }

    /// Adopts an immutable leaf package (a commit result or a fresh one).
    pub(crate) fn from_package(
        package: litchi_odt::Package,
        origin: Option<Arc<Path>>,
        format: FileFormat,
    ) -> Result<Self> {
        Self::from_package_with_options(
            package,
            origin,
            format,
            litchi_odf_common::OpenOptions::new(),
        )
    }

    fn from_package_with_options(
        package: litchi_odt::Package,
        origin: Option<Arc<Path>>,
        format: FileFormat,
        odf_options: litchi_odf_common::OpenOptions,
    ) -> Result<Self> {
        let document = package
            .document()
            .map_err(|error| odf_error(error).with_format(format))?;
        Ok(Self(Arc::new(OdtState {
            owner: OdtOwner::Eager(document),
            odf_options,
            format,
            package: OnceLock::from(package),
            tree: OnceLock::new(),
            kinds: OnceLock::new(),
            parts: OnceLock::new(),
            origin,
            flat: None,
        })))
    }

    /// Opens a flat single-XML text document (`.fodt`) detection
    /// recognized. Reads go through the leaf's packaged projection of the
    /// flat parts; publication writes the flat bytes (record 0292).
    pub(crate) fn from_flat(bytes: Vec<u8>, options: &OpenOptions) -> Result<Self> {
        let flat_options = options.odf_open_options();
        let flat = litchi_odt::flat::Document::from_bytes_with_options(&bytes, flat_options)
            .map_err(odf_error)?;
        Self::from_flat_document(flat, None)
    }

    /// Adopts a flat document (an open or a committed flat edit): its
    /// packaged projection serves reads, its bytes publication.
    pub(crate) fn from_flat_document(
        flat: litchi_odt::flat::Document,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        let package = flat.to_package().map_err(odf_error)?;
        let document = package.document().map_err(odf_error)?;
        Ok(Self(Arc::new(OdtState {
            owner: OdtOwner::Eager(document),
            odf_options: flat.options(),
            format: FileFormat::Odt,
            package: OnceLock::from(package),
            tree: OnceLock::new(),
            kinds: OnceLock::new(),
            parts: OnceLock::new(),
            origin,
            flat: Some(flat),
        })))
    }

    /// The handle over a committed package: a packaged document adopts it;
    /// a flat document republishes the edit into its flat bytes (record
    /// 0292).
    pub(crate) fn adopt(&self, package: litchi_odt::Package) -> Result<Self> {
        match self.flat() {
            Some(flat) => {
                let republished = flat.republish(&package).map_err(odf_error)?;
                // The edited package holds exactly the parts the republished
                // flat document splits into, so it stays the projection.
                let document = package.document().map_err(odf_error)?;
                Ok(Self(Arc::new(OdtState {
                    owner: OdtOwner::Eager(document),
                    odf_options: self.0.odf_options,
                    format: FileFormat::Odt,
                    package: OnceLock::from(package),
                    tree: OnceLock::new(),
                    kinds: OnceLock::new(),
                    parts: OnceLock::new(),
                    origin: self.origin().cloned(),
                    flat: Some(republished),
                })))
            },
            None => Self::from_package_with_options(
                package,
                self.origin().cloned(),
                self.format(),
                self.0.odf_options,
            ),
        }
    }

    /// The flat document this handle reads, when it reads one.
    pub(crate) fn flat(&self) -> Option<&litchi_odt::flat::Document> {
        self.0.flat.as_ref()
    }

    /// Whether the handle reads a flat single-XML document.
    pub(crate) fn is_flat(&self) -> bool {
        self.0.flat.is_some()
    }

    /// Opens the source-pinned candidate the document path probe returned
    /// (`DocumentSourcePathDetection::Odt`). The candidate's own guards
    /// refuse a source that changed since detection.
    ///
    /// ODT-versus-OOXML polyglot arbitration stays with detection: the caller
    /// passes a candidate only once ODT owns the input.
    #[cfg(any(unix, windows))]
    pub(crate) fn from_source_candidate(
        candidate: crate::detect::detected::OdtSourcePathCandidate,
        options: &OpenOptions,
    ) -> Result<Self> {
        let document = candidate
            .into_document(options.odf_open_options())
            .map_err(core_error)?;
        Ok(Self::from_owner(
            OdtOwner::Source(document),
            options.odf_open_options(),
        ))
    }

    fn from_owner(owner: OdtOwner, odf_options: litchi_odf_common::OpenOptions) -> Self {
        Self(Arc::new(OdtState {
            owner,
            odf_options,
            format: FileFormat::Odt,
            package: OnceLock::new(),
            tree: OnceLock::new(),
            kinds: OnceLock::new(),
            parts: OnceLock::new(),
            origin: None,
            flat: None,
        }))
    }

    /// The same document, saving to `origin`. Called on a handle the open
    /// route just built, which no other handle shares yet.
    pub(crate) fn with_origin(self, origin: Option<Arc<Path>>) -> Self {
        match Arc::try_unwrap(self.0) {
            Ok(mut state) => {
                state.origin = origin;
                Self(Arc::new(state))
            },
            Err(shared) => Self(shared),
        }
    }

    /// The package's format: [`FileFormat::Odt`], [`FileFormat::Oth`] or
    /// [`FileFormat::Odm`].
    #[must_use]
    pub fn format(&self) -> FileFormat {
        self.0.format
    }

    /// The ODF ingress policy retained by this handle.
    pub(crate) fn odf_options(&self) -> litchi_odf_common::OpenOptions {
        self.0.odf_options
    }

    /// Re-tags an error the shared ODT machinery reported as ODT with this
    /// handle's format (OTH, ODM).
    pub(crate) fn retag(&self, error: Error) -> Error {
        if self.0.format != FileFormat::Odt && error.format() == Some(FileFormat::Odt) {
            error.with_format(self.0.format)
        } else {
            error
        }
    }

    /// The `File` variant of this handle's format.
    pub(crate) fn into_file(self) -> crate::File {
        match self.0.format {
            FileFormat::Oth => crate::File::Oth(self),
            FileFormat::Odm => crate::File::Odm(self),
            _ => crate::File::Odt(self),
        }
    }

    /// The path the document was opened from, if any.
    pub(crate) fn origin(&self) -> Option<&Arc<Path>> {
        self.0.origin.as_ref()
    }

    /// The document's plain text.
    ///
    /// # Errors
    ///
    /// Returns `corrupt_file` when the body cannot be parsed, and `io_error`
    /// when a filesystem-backed source changed.
    pub fn text(&self) -> Result<String> {
        match &self.0.owner {
            OdtOwner::Eager(document) => document.text().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => self.0.owner.finish(document.text()),
        }
    }

    /// The number of body paragraphs.
    ///
    /// # Errors
    ///
    /// As for [`Self::text`].
    pub fn paragraph_count(&self) -> Result<usize> {
        match &self.0.owner {
            OdtOwner::Eager(document) => document.paragraph_count().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => self.0.owner.finish(document.paragraph_count()),
        }
    }

    /// Document properties from `meta.xml`. Always `Some` for ODT.
    ///
    /// # Errors
    ///
    /// As for [`Self::text`].
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        let metadata = match &self.0.owner {
            OdtOwner::Eager(document) => document.metadata().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => self.0.owner.finish(document.metadata()),
        }?;
        Ok(Some(metadata))
    }

    /// Everything the reader repaired while opening the package and parsing
    /// the body, which happens at open.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform across
    /// formats.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(match &self.0.owner {
            OdtOwner::Eager(document) => document.tolerance_report().clone(),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => document.tolerance_report().clone(),
        })
    }

    /// The retained leaf owner, for the crate's family views.
    pub(crate) fn owner(&self) -> &OdtOwner {
        &self.0.owner
    }

    /// The structural body index, built once per handle.
    pub(crate) fn body(&self) -> Result<Body<'_>> {
        match &self.0.owner {
            OdtOwner::Eager(document) => document.body().map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => {
                let content = document.content_xml().map_err(core_error)?;
                let styles = document.styles_xml().map_err(core_error)?;
                let tree = match self.0.tree.get() {
                    Some(tree) => tree,
                    None => {
                        let tree = Tree::build(content, styles).map_err(core_error)?;
                        self.0.tree.get_or_init(|| tree)
                    },
                };
                Ok(tree.view(content, styles))
            },
        }
    }

    /// The part index of record 0328, built on the first part read.
    pub(crate) fn parts(&self) -> Result<&parts::Index> {
        if let Some(index) = self.0.parts.get() {
            return Ok(index);
        }
        let index = parts::Index::build(self)?;
        Ok(self.0.parts.get_or_init(|| index))
    }

    /// The bytes of one package member (an embedded object's part), or
    /// `None` when the package has no such member.
    pub(crate) fn member(&self, path: &str) -> Result<Option<Vec<u8>>> {
        match &self.0.owner {
            OdtOwner::Eager(document) => document.member_data(path).map_err(core_error),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => self.0.owner.finish(document.member_data(path)),
        }
    }

    /// The `ordinal`th child of `kind` under a body-backed `parent`, with
    /// its ordinal.
    pub(crate) fn nth_child(
        &self,
        parent: &Node<'_>,
        kind: Kind,
        ordinal: usize,
    ) -> Result<Option<(NodeId, usize)>> {
        let Some(leaf) = parent.leaf() else {
            return Ok(None);
        };
        if ordinal < DIRECT_SCAN {
            return Ok(leaf
                .children()
                .filter(|child| Kind::of(child.kind()) == Some(kind))
                .nth(ordinal)
                .map(|child| (child.id(), ordinal)));
        }
        let index = match self.0.kinds.get() {
            Some(index) => index,
            None => {
                let index = KindIndex::build(self.body()?);
                self.0.kinds.get_or_init(|| index)
            },
        };
        Ok(index
            .children
            .get(&(leaf.id(), kind as u8))
            .and_then(|children| children.get(ordinal))
            .map(|id| (*id, ordinal)))
    }

    /// The last child of `kind` under a body-backed `parent`, with its
    /// ordinal.
    pub(crate) fn last_child(
        &self,
        parent: &Node<'_>,
        kind: Kind,
    ) -> Result<Option<(NodeId, usize)>> {
        let Some(leaf) = parent.leaf() else {
            return Ok(None);
        };
        Ok(leaf
            .children()
            .filter(|child| Kind::of(child.kind()) == Some(kind))
            .enumerate()
            .last()
            .map(|(ordinal, child)| (child.id(), ordinal)))
    }

    /// The immutable leaf package edits start from.
    pub(crate) fn package(&self) -> Result<&litchi_odt::Package> {
        if let Some(package) = self.0.package.get() {
            return Ok(package);
        }
        match &self.0.owner {
            OdtOwner::Eager(_) => Err(Error::new(
                crate::Code::InternalError,
                "an eager ODT handle lost its package",
            )
            .with_format(FileFormat::Odt)),
            #[cfg(any(unix, windows))]
            OdtOwner::Source(document) => {
                let package = self.0.owner.finish(document.to_package())?;
                Ok(self.0.package.get_or_init(|| package))
            },
        }
    }

    /// Fails with `io_error` when a filesystem-backed source changed.
    pub(crate) fn check_source(&self) -> Result<()> {
        self.0.owner.check_source()
    }

    /// Whether two handles retain the same immutable package bytes.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        if let (Some(left), Some(right)) = (self.flat(), other.flat()) {
            return Arc::ptr_eq(&self.0, &other.0) || left.as_bytes() == right.as_bytes();
        }
        Arc::ptr_eq(&self.0, &other.0)
            || match (self.0.package.get(), other.0.package.get()) {
                (Some(left), Some(right)) => left.as_bytes() == right.as_bytes(),
                _ => false,
            }
    }

    /// Publish the document's package bytes into `sink`.
    pub(crate) fn write_to(&self, mut sink: impl std::io::Write) -> Result<()> {
        if let Some(flat) = self.flat() {
            return sink
                .write_all(flat.as_bytes())
                .map_err(|error| core_error(litchi_core::Error::from(error)));
        }
        self.package()?.write_to(sink).map_err(odf_error)
    }

    /// Atomically save the document's package to `path`: a sibling
    /// temporary file renamed over the destination, or published only when
    /// the destination is still absent under [`crate::Overwrite::Refuse`].
    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        let bytes = match self.flat() {
            Some(flat) => flat.as_bytes(),
            None => self.package()?.as_bytes(),
        };
        let overwrite = match options.overwrite {
            crate::Overwrite::Replace => litchi_core::atomic::Overwrite::Replace,
            crate::Overwrite::Refuse => litchi_core::atomic::Overwrite::Refuse,
        };
        litchi_core::atomic::save(
            path,
            options.durability,
            overwrite,
            |writer| {
                use std::io::Write as _;
                writer.write_all(bytes).map_err(|error| {
                    litchi_odf_common::Error::from(litchi_core::atomic::Error::from(error))
                })
            },
            || Ok(()),
        )
        .map_err(odf_error)?;
        self.check_source()?;
        Ok(u64::try_from(bytes.len()).unwrap_or(u64::MAX))
    }

    /// A fresh empty text document.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        write::create(options)
    }

    /// Admit a leaf edit before a unified transaction starts: the source is
    /// current and the package is available.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()?;
        self.package().map(drop)
    }

    /// Whether two handles publish the same package bytes.
    pub(crate) fn same_content(&self, other: &Self) -> Result<bool> {
        if self.is_flat() || other.is_flat() {
            return Ok(self.same_snapshot(other));
        }
        Ok(self.same_snapshot(other) || self.package()?.as_bytes() == other.package()?.as_bytes())
    }

    /// Set properties with the default unified limits.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: crate::TargetArg<'p>,
        props: impl crate::IntoProps<'v>,
    ) -> Result<write::Commit> {
        write::set_with_limits(
            self,
            target,
            props,
            &crate::transaction::SemanticLimits::default(),
        )
        .map_err(|error| self.retag(error))
    }

    /// Set properties with the caller's transaction limits.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: crate::TargetArg<'p>,
        props: impl crate::IntoProps<'v>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::set_with_limits(self, target, props, limits).map_err(|error| self.retag(error))
    }

    /// Add one element with the default unified limits.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        write::add_with_limits(
            self,
            parent,
            node,
            &crate::transaction::SemanticLimits::default(),
        )
        .map_err(|error| self.retag(error))
    }

    /// Add one element with the caller's transaction limits.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        write::add_with_limits(self, parent, node, limits).map_err(|error| self.retag(error))
    }

    /// Remove targets with the default unified limits.
    pub(crate) fn remove<'p>(&self, target: crate::TargetArg<'p>) -> Result<write::Commit> {
        write::remove_with_limits(self, target, &crate::transaction::SemanticLimits::default())
            .map_err(|error| self.retag(error))
    }

    /// Remove targets with the caller's transaction limits.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: crate::TargetArg<'p>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::remove_with_limits(self, target, limits).map_err(|error| self.retag(error))
    }

    /// Move one element with the default unified limits.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
    ) -> Result<write::Commit> {
        write::move_to_with_limits(
            self,
            path,
            placement,
            &crate::transaction::SemanticLimits::default(),
        )
        .map_err(|error| self.retag(error))
    }

    /// Move one element with the caller's transaction limits.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::move_to_with_limits(self, path, placement, limits).map_err(|error| self.retag(error))
    }

    /// Exchange two elements with the default unified limits.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<write::Commit> {
        write::swap_with_limits(
            self,
            left,
            right,
            &crate::transaction::SemanticLimits::default(),
        )
        .map_err(|error| self.retag(error))
    }

    /// Exchange two elements with the caller's transaction limits.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::swap_with_limits(self, left, right, limits).map_err(|error| self.retag(error))
    }
}

impl OdtOwner {
    /// Binds a read's result to the source revision captured at open: on a
    /// filesystem-backed owner a final source-revision failure wins over the
    /// result the leaf produced after its own check. This closes the TOCTOU
    /// window between the leaf's payload read and the handle's return (the
    /// facade's `finish_source_odt_result`).
    pub(crate) fn finish<T>(&self, result: litchi_core::Result<T>) -> Result<T> {
        match self {
            Self::Eager(_) => result.map_err(core_error),
            #[cfg(any(unix, windows))]
            Self::Source(document) => match document.source_version() {
                Err(error) => Err(core_error(error)),
                Ok(_) => result.map_err(core_error),
            },
        }
    }

    /// Fails with `io_error` when a filesystem-backed source changed since
    /// open; always succeeds for the eager owner. This is the ODT half of
    /// the facade's `check_source_freshness_after_markdown`.
    pub(crate) fn check_source(&self) -> Result<()> {
        match self {
            Self::Eager(_) => Ok(()),
            #[cfg(any(unix, windows))]
            Self::Source(document) => document.source_version().map(drop).map_err(core_error),
        }
    }
}

impl fmt::Debug for Odt {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let state = match &self.0.owner {
            OdtOwner::Eager(_) => "eager",
            #[cfg(any(unix, windows))]
            OdtOwner::Source(_) => "source",
        };
        formatter
            .debug_struct("Odt")
            .field("owner", &state)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests;
