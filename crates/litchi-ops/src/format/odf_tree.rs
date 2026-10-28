//! Handles for the OpenDocument object formats without a text, sheet or
//! slide body (record 0282): drawings (`.odg`), charts (`.odc`), images
//! (`.odi`), databases (`.odb`) and formulas (`.odf`).
//!
//! Each handle retains its leaf's immutable `Package` and an object tree
//! built once at open from the leaf's read-only projection: pages and
//! shapes, chart axes and series, the image frame, database tables, columns
//! and queries, and MathML elements. Paths, queries and node views run over
//! that tree through one generic adapter (`Nav`), parameterized by a
//! per-format `Flavor` (its element kinds and spec table). Writes stage
//! the leaf's own `Edit` operations, address their targets through the
//! leaf positions the tree recorded, and adopt the committed package as a
//! new handle with a freshly built tree.

/// Declares a field-less kind enum with its vocabulary mapping.
macro_rules! kinds {
    (
        $(#[$meta:meta])*
        $name:ident: $vocabulary:expr, root $root:ident {
            $($kind:ident => $canonical:literal,)+
        }
    ) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        pub(crate) enum $name {
            $(
                #[doc = concat!("The `", $canonical, "` element.")]
                $kind,
            )+
        }

        impl crate::path::navigate::ElementKind for $name {
            const VOCABULARY: &'static crate::path::vocabulary::Vocabulary = $vocabulary;
            const ROOT: Self = Self::$root;

            fn from_name(name: &str) -> Option<Self> {
                match name {
                    $($canonical => Some(Self::$kind),)+
                    _ => None,
                }
            }

            fn name(self) -> &'static str {
                match self {
                    $(Self::$kind => $canonical,)+
                }
            }

            fn ordinal(self) -> u8 {
                self as u8
            }
        }
    };
}

pub(crate) mod children;
pub(crate) mod coalesce;
mod nav;
mod node;
mod tree;
pub(crate) mod views;
pub(crate) mod write;

#[cfg(feature = "odf-formula")]
pub(crate) mod formula;
#[cfg(feature = "odb")]
pub(crate) mod odb;
#[cfg(feature = "odc")]
pub(crate) mod odc;
#[cfg(feature = "odg")]
pub(crate) mod odg;
#[cfg(feature = "odi")]
pub(crate) mod odi;

pub(crate) use nav::{Flavor, Nav};
pub use node::Node;
pub(crate) use tree::{Tree, TreeBuilder};

use std::fmt;
use std::path::Path;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;

use crate::{Error, OpenOptions, Result};

/// An opened OpenDocument drawing, chart, image, database or formula. Cheap
/// to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct OdfTree(Arc<State>);

struct State {
    owner: Owner,
    /// The caller-selected ODF ingress policy retained across publication
    /// and raw package reads.
    odf_options: litchi_odf_common::OpenOptions,
    /// Built at open; after a write, on the first read (record 0288).
    tree: std::sync::OnceLock<Tree>,
    origin: Option<Arc<Path>>,
    /// The flat single-XML bytes (`.fodg`) a flat handle publishes; its
    /// owner is the leaf's packaged read projection (record 0292).
    flat: Option<FlatSource>,
}

struct FlatSource {
    bytes: Box<[u8]>,
    limits: litchi_odf_common::flat::Limits,
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<OdfTree>();
    assert!(size_of::<OdfTree>() == size_of::<usize>());
};

/// The leaf owner of one object format.
#[derive(Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to the open route"
)]
pub(crate) enum Owner {
    #[cfg(feature = "odg")]
    Odg(litchi_odg::Package),
    #[cfg(feature = "odc")]
    Odc(litchi_odc::Package),
    #[cfg(feature = "odi")]
    Odi(litchi_odi::Package),
    #[cfg(feature = "odb")]
    Odb(litchi_odb::Package),
    #[cfg(feature = "odf-formula")]
    Formula(litchi_odf_formula::Package),
}

impl Owner {
    /// The format the owner reads.
    pub(crate) const fn format(&self) -> FileFormat {
        match self {
            #[cfg(feature = "odg")]
            Self::Odg(_) => FileFormat::Odg,
            #[cfg(feature = "odc")]
            Self::Odc(_) => FileFormat::Odc,
            #[cfg(feature = "odi")]
            Self::Odi(_) => FileFormat::Odi,
            #[cfg(feature = "odb")]
            Self::Odb(_) => FileFormat::Odb,
            #[cfg(feature = "odf-formula")]
            Self::Formula(_) => FileFormat::Odf,
        }
    }

    /// The exact package bytes.
    pub(crate) fn as_bytes(&self) -> &[u8] {
        match self {
            #[cfg(feature = "odg")]
            Self::Odg(package) => package.as_bytes(),
            #[cfg(feature = "odc")]
            Self::Odc(package) => package.as_bytes(),
            #[cfg(feature = "odi")]
            Self::Odi(package) => package.as_bytes(),
            #[cfg(feature = "odb")]
            Self::Odb(package) => package.as_bytes(),
            #[cfg(feature = "odf-formula")]
            Self::Formula(package) => package.as_bytes(),
        }
    }

    /// Whether building this owner's tree can fail: a database tree parses
    /// its catalog, the other trees project the model the leaf parsed.
    const fn tree_can_fail(&self) -> bool {
        match self {
            #[cfg(feature = "odg")]
            Self::Odg(_) => false,
            #[cfg(feature = "odc")]
            Self::Odc(_) => false,
            #[cfg(feature = "odi")]
            Self::Odi(_) => false,
            #[cfg(feature = "odb")]
            Self::Odb(_) => true,
            #[cfg(feature = "odf-formula")]
            Self::Formula(_) => false,
        }
    }

    /// Build the object tree of this owner.
    fn tree(&self) -> Result<Tree> {
        let format = self.format();
        match self {
            #[cfg(feature = "odg")]
            Self::Odg(package) => odg::build(package),
            #[cfg(feature = "odc")]
            Self::Odc(package) => odc::build(package),
            #[cfg(feature = "odi")]
            Self::Odi(package) => odi::build(package),
            #[cfg(feature = "odb")]
            Self::Odb(package) => odb::build(package),
            #[cfg(feature = "odf-formula")]
            Self::Formula(package) => formula::build(package),
        }
        .map_err(|error| error.with_format(format))
    }
}

/// A leaf ingress or publication failure: typed limits stay
/// `Source::Odf`, everything else keeps its common-core payload.
pub(crate) fn odf_error(error: litchi_odf_common::Error, format: FileFormat) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => core_error(error, format),
        other => Error::from(other).with_format(format),
    }
}

/// A common-core leaf failure of `format`.
pub(crate) fn core_error(error: litchi_core::Error, format: FileFormat) -> Error {
    Error::from(error).with_format(format)
}

/// Map an ODF edit error without projecting its typed package payload into
/// the common-core error.  The older object adapters still return a core
/// error for some staging paths, so those retain the caller-argument
/// classification supplied by the shared helper.
#[cfg(any(
    all(test, feature = "odb"),
    feature = "odc",
    feature = "odi",
    feature = "odf-formula"
))]
pub(crate) fn odf_staging_error(error: litchi_odf_common::Error, format: FileFormat) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => write::staging_error(error, format),
        other => odf_error(other, format),
    }
}

impl OdfTree {
    /// Adopt the package detection prepared, without a second ZIP index.
    pub(crate) fn from_prepared(prepared: PreparedPackage, options: &OpenOptions) -> Result<Self> {
        let format = prepared.format();
        let error = |error| odf_error(error, format);
        let owner = match format {
            #[cfg(feature = "odg")]
            FileFormat::Odg => Owner::Odg(
                litchi_odg::Package::from_prepared_package_with_options(
                    prepared,
                    options.odf_open_options(),
                )
                .map_err(error)?,
            ),
            #[cfg(feature = "odc")]
            FileFormat::Odc => Owner::Odc(
                litchi_odc::Package::from_prepared_package_with_options(
                    prepared,
                    options.odf_open_options(),
                )
                .map_err(error)?,
            ),
            #[cfg(feature = "odi")]
            FileFormat::Odi => Owner::Odi(
                litchi_odi::Package::from_prepared_package_with_options(
                    prepared,
                    options.odf_open_options(),
                )
                .map_err(error)?,
            ),
            #[cfg(feature = "odb")]
            FileFormat::Odb => Owner::Odb(
                litchi_odb::Package::from_prepared_package_with_options(
                    prepared,
                    options.odf_open_options(),
                )
                .map_err(error)?,
            ),
            #[cfg(feature = "odf-formula")]
            FileFormat::Odf => Owner::Formula(
                litchi_odf_formula::Package::from_prepared_package_with_options(
                    prepared,
                    options.odf_open_options(),
                )
                .map_err(error)?,
            ),
            other => {
                let _ = (prepared, error);
                return Err(Error::unsupported_format(Some(other)));
            },
        };
        Self::from_owner(owner, None, options.odf_open_options())
    }

    /// Adopt a leaf owner (an open, a commit result or a fresh package).
    /// A tree whose build can fail is built here, so the open refuses a
    /// database whose catalog it cannot read; the others are built on the
    /// first read, as after a commit (record 0377a).
    pub(crate) fn from_owner(
        owner: Owner,
        origin: Option<Arc<Path>>,
        odf_options: litchi_odf_common::OpenOptions,
    ) -> Result<Self> {
        let tree = if owner.tree_can_fail() {
            std::sync::OnceLock::from(owner.tree()?)
        } else {
            std::sync::OnceLock::new()
        };
        Ok(Self(Arc::new(State {
            owner,
            odf_options,
            tree,
            origin,
            flat: None,
        })))
    }

    /// Open a flat single-XML drawing (`.fodg`) detection recognized: the
    /// leaf reads the flat parts as a package projection; publication
    /// writes the flat bytes unchanged (record 0292).
    #[cfg(feature = "odg")]
    pub(crate) fn from_flat(bytes: Vec<u8>, options: &OpenOptions) -> Result<Self> {
        let format = FileFormat::Odg;
        let flat_limits = litchi_odf_common::flat::Limits::from_options(options.odf_open_options());
        let parts = litchi_odf_common::flat::Parts::split(&bytes, &flat_limits)
            .map_err(|error| odf_error(error, format))?;
        let package = parts
            .to_bytes(&flat_limits)
            .map_err(|error| odf_error(error, format))
            .and_then(|zip| {
                // The input limit applies to the original flat source. The
                // generated ZIP is an internal projection; semantic policy
                // still applies to every parsed drawing part and object.
                litchi_odg::Package::from_bytes_with_options(&zip, flat_limits.projection_options())
                    .map_err(|error| odf_error(error, format))
            })?;
        let owner = Owner::Odg(package);
        let tree = owner.tree()?;
        Ok(Self(Arc::new(State {
            owner,
            odf_options: options.odf_open_options(),
            tree: std::sync::OnceLock::from(tree),
            origin: None,
            flat: Some(FlatSource {
                bytes: bytes.into_boxed_slice(),
                limits: flat_limits,
            }),
        })))
    }

    /// The handle over a committed leaf owner: a packaged file adopts it;
    /// a flat drawing republishes the edit of its packaged projection into
    /// its flat bytes, element by element (record 0292).
    pub(crate) fn republish(&self, owner: Owner) -> Result<Self> {
        let Some(flat) = self.0.flat.as_ref() else {
            return Self::from_commit(owner, self.origin().cloned(), self.0.odf_options);
        };
        let format = self.format();
        let parts = litchi_odf_common::flat::Parts::split(&flat.bytes, &flat.limits)
            .map_err(|error| odf_error(error, format))?;
        let bytes = parts
            .merge_package_projection(&flat.bytes, owner.as_bytes(), &flat.limits)
            .map_err(|error| odf_error(error, format))?;
        // The edited package holds exactly the parts the republished flat
        // document splits into, so it stays the projection; its tree is
        // built on the first read.
        Ok(Self(Arc::new(State {
            owner,
            odf_options: self.0.odf_options,
            tree: std::sync::OnceLock::new(),
            origin: self.origin().cloned(),
            flat: Some(FlatSource {
                bytes: bytes.into_boxed_slice(),
                limits: flat.limits,
            }),
        })))
    }

    /// The bytes this handle publishes: the flat document, or the package.
    fn published(&self) -> &[u8] {
        self.0
            .flat
            .as_ref()
            .map_or_else(|| self.0.owner.as_bytes(), |flat| flat.bytes.as_ref())
    }

    /// Adopt a package a leaf commit published and verified. The object
    /// tree is built on the first read, so a write (or a batch of writes)
    /// does not pay for a projection nobody reads. The drawing, chart, image
    /// and formula trees build infallibly. A database tree parses its
    /// catalog, which can fail, but not on a committed package: the ODB
    /// leaf's commit parses the new catalog with the same default limits
    /// before it publishes, and a metadata commit leaves the catalog's
    /// `content.xml` as it was (records 0291 and 0292 made the same change independently; before, a database tree was
    /// built here and a single `set` on a 1,000-query database parsed the
    /// catalog twice).
    pub(crate) fn from_commit(
        owner: Owner,
        origin: Option<Arc<Path>>,
        odf_options: litchi_odf_common::OpenOptions,
    ) -> Result<Self> {
        Ok(Self(Arc::new(State {
            owner,
            odf_options,
            tree: std::sync::OnceLock::new(),
            origin,
            flat: None,
        })))
    }

    /// The same file, saving to `origin`. Called on a handle the open route
    /// just built.
    pub(crate) fn with_origin(self, origin: Option<Arc<Path>>) -> Self {
        match Arc::try_unwrap(self.0) {
            Ok(mut state) => {
                state.origin = origin;
                Self(Arc::new(state))
            },
            Err(shared) => Self(shared),
        }
    }

    /// The path the file was opened from, if any.
    pub(crate) fn origin(&self) -> Option<&Arc<Path>> {
        self.0.origin.as_ref()
    }

    /// The package's format.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        self.0.owner.format()
    }

    /// The ODF ingress policy retained by this handle.
    pub(crate) fn odf_options(&self) -> litchi_odf_common::OpenOptions {
        self.0.odf_options
    }

    /// The retained leaf owner.
    pub(crate) fn owner(&self) -> &Owner {
        &self.0.owner
    }

    /// The object tree: a database's built at open, any other on first read.
    pub(crate) fn tree(&self) -> &Tree {
        self.0.tree.get_or_init(|| {
            // Only builders that cannot fail (`from_owner`), or cannot fail
            // on a committed package (`from_commit`), are deferred; a
            // failure would still leave the root rather than panic.
            self.0.owner.tree().unwrap_or_else(|_| Tree::root_only())
        })
    }

    /// The `File` variant of this handle's format.
    pub(crate) fn into_file(self) -> crate::File {
        match self.format() {
            FileFormat::Odg => crate::File::Odg(self),
            FileFormat::Odc => crate::File::Odc(self),
            FileFormat::Odi => crate::File::Odi(self),
            FileFormat::Odb => crate::File::Odb(self),
            _ => crate::File::Odf(self),
        }
    }

    /// The file's plain text: every text-bearing object, one per line, in
    /// document order.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform.
    pub fn text(&self) -> Result<String> {
        let mut text = String::new();
        for item in self.tree().items() {
            if let Some(value) = item.text() {
                if value.is_empty() {
                    continue;
                }
                if !text.is_empty() {
                    text.push('\n');
                }
                text.push_str(value);
            }
        }
        Ok(text)
    }

    /// Document properties from `meta.xml`, when the package has them.
    ///
    /// # Errors
    ///
    /// A malformed formula `meta.xml` (the other leaves parse it at open).
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(match self.owner() {
            #[cfg(feature = "odg")]
            Owner::Odg(package) => package.metadata().cloned(),
            #[cfg(feature = "odc")]
            Owner::Odc(package) => package.chart().metadata().cloned(),
            #[cfg(feature = "odi")]
            Owner::Odi(package) => package.image().metadata().cloned(),
            #[cfg(feature = "odb")]
            Owner::Odb(package) => package.database().metadata().cloned(),
            #[cfg(feature = "odf-formula")]
            Owner::Formula(package) => package
                .metadata()
                .map_err(|error| odf_error(error, FileFormat::Odf))?,
        })
    }

    /// Everything the reader repaired while opening the package: the leaf's
    /// report where it keeps one, else a clean report.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform.
    pub fn tolerance_report(&self) -> Result<litchi_odf_common::leniency::ToleranceReport> {
        Ok(match self.owner() {
            #[cfg(feature = "odg")]
            Owner::Odg(package) => package.tolerance_report().clone(),
            #[cfg(feature = "odc")]
            Owner::Odc(package) => package.tolerance_report().clone(),
            #[cfg(feature = "odi")]
            Owner::Odi(package) => package.tolerance_report().clone(),
            #[cfg(feature = "odb")]
            Owner::Odb(package) => package.tolerance_report().clone(),
            #[cfg(feature = "odf-formula")]
            Owner::Formula(_) => litchi_odf_common::leniency::ToleranceReport::default(),
        })
    }

    /// Whether two handles retain the same package bytes.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0) || self.published() == other.published()
    }

    /// Whether two handles publish the same package bytes.
    pub(crate) fn same_content(&self, other: &Self) -> bool {
        self.same_snapshot(other)
    }

    /// The handle's source is in memory: always current.
    #[allow(
        clippy::unnecessary_wraps,
        clippy::unused_self,
        reason = "the uniform handle surface returns a freshness result"
    )]
    pub(crate) fn check_source(&self) -> Result<()> {
        Ok(())
    }

    /// Admit a leaf edit before a unified transaction starts.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()
    }

    /// Publish the package bytes into `sink`.
    pub(crate) fn write_to(&self, mut sink: impl std::io::Write) -> Result<()> {
        sink.write_all(self.published()).map_err(|error| {
            Error::from(litchi_core::Error::from(error)).with_format(self.format())
        })
    }

    /// Atomically save the package to `path`.
    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        let bytes = self.published();
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
        .map_err(|error| odf_error(error, self.format()))?;
        Ok(u64::try_from(bytes.len()).unwrap_or(u64::MAX))
    }

    /// A fresh minimal package of `format`.
    pub(crate) fn create(format: FileFormat, options: &crate::CreateOptions) -> Result<Self> {
        write::create(format, options)
    }

    /// Set properties with the default unified limits.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: crate::TargetArg<'p>,
        props: impl crate::IntoProps<'v>,
    ) -> Result<write::Commit> {
        self.set_with_limits(
            target,
            props,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Set properties with the caller's transaction limits.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: crate::TargetArg<'p>,
        props: impl crate::IntoProps<'v>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::set(self, target, props, limits).map_err(|error| self.tag(error))
    }

    /// Add one element with the default unified limits.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        self.add_with_limits(parent, node, &crate::transaction::SemanticLimits::default())
    }

    /// Add one element with the caller's transaction limits.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        write::add(self, parent, node, limits).map_err(|error| self.tag(error))
    }

    /// Remove targets with the default unified limits.
    pub(crate) fn remove<'p>(&self, target: crate::TargetArg<'p>) -> Result<write::Commit> {
        self.remove_with_limits(target, &crate::transaction::SemanticLimits::default())
    }

    /// Remove targets with the caller's transaction limits.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: crate::TargetArg<'p>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::remove(self, target, limits).map_err(|error| self.tag(error))
    }

    /// Move an element (ODG pages and shapes; other formats refuse).
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
    ) -> Result<write::Commit> {
        self.move_to_with_limits(
            path,
            placement,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Move an element with the caller's transaction limits.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::move_to(self, path, placement, limits).map_err(|error| self.tag(error))
    }

    /// Exchange two elements (ODG pages and shapes; other formats refuse).
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<write::Commit> {
        self.swap_with_limits(left, right, &crate::transaction::SemanticLimits::default())
    }

    /// Exchange two elements with the caller's transaction limits.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::swap(self, left, right, limits).map_err(|error| self.tag(error))
    }

    /// Tag an error with this handle's format.
    pub(crate) fn tag(&self, error: Error) -> Error {
        if error.format().is_none() {
            error.with_format(self.format())
        } else {
            error
        }
    }
}

impl fmt::Debug for OdfTree {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("OdfTree")
            .field("format", &self.format())
            .field("items", &self.tree().len())
            .finish_non_exhaustive()
    }
}

/// The `/metadata` path of the object formats: set-only (record 0288).
#[cfg(any(
    feature = "odg",
    feature = "odc",
    feature = "odb",
    feature = "odf-formula"
))]
pub(crate) const METADATA_PATH_GAP: crate::node::spec::GapSpec = crate::node::spec::GapSpec {
    subject: crate::node::spec::Subject::Operation {
        verb: "get",
        element: "metadata",
        context: "path-read",
    },
    code: crate::Code::UnsupportedCapability,
    follow_up: "0288 ODF object metadata node",
    reason: "`set(\"/metadata\", ..)` writes meta.xml through the ODF metadata family; the path is set-only, and reads go through File::metadata()",
};

/// One element of a flavor's spec table.
#[allow(
    clippy::too_many_arguments,
    reason = "one argument per spec table column keeps the tables readable"
)]
pub(crate) const fn element<K: Copy>(
    kind: K,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [K],
    children: &'static [K],
    paths: &'static [&'static str],
    props: &'static [crate::node::spec::PropSpec],
) -> crate::node::spec::ElementSpec<K> {
    crate::node::spec::ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
        identities: &[],
        paths,
        props,
        gap: None,
    }
}

/// A readable property that `set` also writes.
pub(crate) const fn writable(
    name: &'static str,
    ty: crate::value::ValueType,
) -> crate::node::spec::PropSpec {
    let mut spec = crate::node::spec::PropSpec::new(name, ty);
    spec.access.set = true;
    spec
}

/// A readable property that `set` writes and `add` accepts.
#[cfg(any(
    feature = "odg",
    feature = "odc",
    feature = "odb",
    feature = "odf-formula"
))]
pub(crate) const fn addable(
    name: &'static str,
    ty: crate::value::ValueType,
) -> crate::node::spec::PropSpec {
    let mut spec = writable(name, ty);
    spec.access.add = true;
    spec
}

/// Runs `$body` with `$flavor` bound to the flavor type of `$format`.
macro_rules! with_flavor {
    ($format:expr, $flavor:ident => $body:expr) => {
        match $format {
            #[cfg(feature = "odg")]
            FileFormat::Odg => {
                type $flavor = odg::Odg;
                $body
            },
            #[cfg(feature = "odc")]
            FileFormat::Odc => {
                type $flavor = odc::Odc;
                $body
            },
            #[cfg(feature = "odi")]
            FileFormat::Odi => {
                type $flavor = odi::Odi;
                $body
            },
            #[cfg(feature = "odb")]
            FileFormat::Odb => {
                type $flavor = odb::Odb;
                $body
            },
            #[cfg(feature = "odf-formula")]
            FileFormat::Odf => {
                type $flavor = formula::Formula;
                $body
            },
            other => Err(Error::unsupported_format(Some(other))),
        }
    };
}

/// `File::get` over an object tree.
pub(crate) fn get<'f>(
    file: &'f OdfTree,
    path: &crate::path::PathArg<'_>,
) -> Result<crate::node::Node<'f>> {
    with_flavor!(file.format(), F => {
        crate::path::resolve::get(&Nav::<F>::new(file), path).map(crate::node::Node::Object)
    })
}

/// The object node at `path`.
pub(crate) fn object<'f>(file: &'f OdfTree, path: &crate::path::PathArg<'_>) -> Result<Node<'f>> {
    with_flavor!(file.format(), F => crate::path::resolve::get(&Nav::<F>::new(file), path))
}

/// `File::find` over an object tree.
pub(crate) fn find<'f>(
    file: &'f OdfTree,
    path: &crate::path::PathArg<'_>,
) -> Result<Option<crate::node::Node<'f>>> {
    with_flavor!(file.format(), F => {
        crate::path::resolve::find(&Nav::<F>::new(file), path)
            .map(|node| node.map(crate::node::Node::Object))
    })
}

/// The query walk of one object format.
pub(crate) enum Walks<'q, 'f> {
    #[cfg(feature = "odg")]
    Odg(crate::query::walk::Walk<'q, 'f, Nav<'f, odg::Odg>>),
    #[cfg(feature = "odc")]
    Odc(crate::query::walk::Walk<'q, 'f, Nav<'f, odc::Odc>>),
    #[cfg(feature = "odi")]
    Odi(crate::query::walk::Walk<'q, 'f, Nav<'f, odi::Odi>>),
    #[cfg(feature = "odb")]
    Odb(crate::query::walk::Walk<'q, 'f, Nav<'f, odb::Odb>>),
    #[cfg(feature = "odf-formula")]
    Formula(crate::query::walk::Walk<'q, 'f, Nav<'f, formula::Formula>>),
}

/// Applies `$body` to the walk of any object format.
macro_rules! each_walk {
    ($walks:expr, $walk:ident => $body:expr) => {
        match $walks {
            #[cfg(feature = "odg")]
            Walks::Odg($walk) => $body,
            #[cfg(feature = "odc")]
            Walks::Odc($walk) => $body,
            #[cfg(feature = "odi")]
            Walks::Odi($walk) => $body,
            #[cfg(feature = "odb")]
            Walks::Odb($walk) => $body,
            #[cfg(feature = "odf-formula")]
            Walks::Formula($walk) => $body,
        }
    };
}

impl<'q, 'f> Walks<'q, 'f> {
    pub(crate) fn limit(self, n: usize) -> Self {
        match self {
            #[cfg(feature = "odg")]
            Self::Odg(walk) => Self::Odg(walk.limit(n)),
            #[cfg(feature = "odc")]
            Self::Odc(walk) => Self::Odc(walk.limit(n)),
            #[cfg(feature = "odi")]
            Self::Odi(walk) => Self::Odi(walk.limit(n)),
            #[cfg(feature = "odb")]
            Self::Odb(walk) => Self::Odb(walk.limit(n)),
            #[cfg(feature = "odf-formula")]
            Self::Formula(walk) => Self::Formula(walk.limit(n)),
        }
    }

    pub(crate) fn plan(&self) -> &crate::query::Plan<'q> {
        each_walk!(self, walk => walk.plan())
    }

    pub(crate) fn next(&mut self) -> Option<Result<crate::node::Node<'f>>> {
        each_walk!(self, walk => walk.next().map(|node| node.map(crate::node::Node::Object)))
    }
}

/// `File::query` over an object tree.
pub(crate) fn query<'f, 'q>(
    file: &'f OdfTree,
    selector: impl crate::query::IntoQuery<'q>,
) -> Result<Walks<'q, 'f>> {
    Ok(match file.format() {
        #[cfg(feature = "odg")]
        FileFormat::Odg => Walks::Odg(crate::query::run(Nav::new(file), selector)?),
        #[cfg(feature = "odc")]
        FileFormat::Odc => Walks::Odc(crate::query::run(Nav::new(file), selector)?),
        #[cfg(feature = "odi")]
        FileFormat::Odi => Walks::Odi(crate::query::run(Nav::new(file), selector)?),
        #[cfg(feature = "odb")]
        FileFormat::Odb => Walks::Odb(crate::query::run(Nav::new(file), selector)?),
        #[cfg(feature = "odf-formula")]
        FileFormat::Odf => Walks::Formula(crate::query::run(Nav::new(file), selector)?),
        other => {
            let _ = selector;
            return Err(Error::unsupported_format(Some(other)));
        },
    })
}

/// Compile `query` for an object format.
pub(crate) fn compile<'q>(
    query: &crate::query::Query<'q>,
    format: FileFormat,
) -> Result<crate::query::Plan<'q>> {
    with_flavor!(format, F => crate::query::compile::compile::<Nav<'static, F>>(query))
}

/// The write registry of one object format: the elements `set`, `add` and
/// `remove` accept, and its operation gaps.
pub(crate) struct Writes {
    pub(crate) settable: &'static [&'static str],
    pub(crate) addable: &'static [&'static str],
    pub(crate) removable: &'static [&'static str],
    pub(crate) gaps: &'static [crate::node::spec::GapSpec],
}

/// The write registry of `adapter`, when it is an object format.
pub(crate) const fn writes(adapter: crate::read::Adapter) -> Option<&'static Writes> {
    match adapter {
        #[cfg(feature = "odg")]
        crate::read::Adapter::Odg => Some(&odg::WRITES),
        #[cfg(feature = "odc")]
        crate::read::Adapter::Odc => Some(&odc::WRITES),
        #[cfg(feature = "odi")]
        crate::read::Adapter::Odi => Some(&odi::WRITES),
        #[cfg(feature = "odb")]
        crate::read::Adapter::Odb => Some(&odb::WRITES),
        #[cfg(feature = "odf-formula")]
        crate::read::Adapter::OdfFormula => Some(&formula::WRITES),
        #[allow(
            unreachable_patterns,
            reason = "the other adapters write through their own tables"
        )]
        _ => None,
    }
}

#[cfg(all(test, feature = "odb"))]
mod tests {
    use litchi_core::detection::FileFormat;
    use litchi_core::{Resource, SourceVersion};

    use crate::error::{Code, Source};
    use crate::{File, NewNode, props};

    fn tree(file: &File) -> &super::OdfTree {
        match file {
            File::Odb(tree) => tree,
            _ => panic!("not a database"),
        }
    }

    /// Record 0291: a database handle committed by `set` builds its tree on
    /// the first read, and that tree is the one a fresh open builds. (`add`
    /// reads the new tree to report the added node's path.)
    #[test]
    fn committed_database_trees_are_built_on_first_read() {
        let mut file = crate::create(FileFormat::Odb).expect("create");
        for index in 0..3 {
            file = file
                .add(
                    "/",
                    NewNode::new("query")
                        .prop("name", format!("q{index}").as_str())
                        .prop("command", format!("SELECT {index}").as_str()),
                )
                .expect("add query")
                .file;
        }
        let file = file
            .set("/query[2]", props! { "command" => "SELECT 20" })
            .expect("set");
        assert!(tree(&file).0.tree.get().is_none(), "set built the tree");
        let lazy = format!("{:?}", tree(&file).tree());
        let reopened = crate::open_bytes(file.to_bytes().expect("bytes")).expect("reopen");
        assert_eq!(lazy, format!("{:?}", tree(&reopened).tree()));
        assert!(lazy.contains("SELECT 20"));
    }

    /// Record 0377a: an opened chart builds its tree on the first read, and
    /// that tree is the one the leaf owner builds; an opened database builds
    /// (and so validates) its tree at open.
    #[cfg(feature = "odc")]
    #[test]
    fn opened_trees_that_cannot_fail_are_built_on_first_read() {
        let chart = crate::create(FileFormat::Odc).expect("create chart");
        let opened = crate::open_bytes(chart.to_bytes().expect("bytes")).expect("open chart");
        let File::Odc(handle) = &opened else {
            panic!("not a chart");
        };
        assert!(handle.0.tree.get().is_none(), "open built the chart tree");
        assert_eq!(
            format!("{:?}", handle.tree()),
            format!("{:?}", handle.0.owner.tree().expect("chart tree"))
        );
        let database = crate::create(FileFormat::Odb).expect("create database");
        let opened = crate::open_bytes(database.to_bytes().expect("bytes")).expect("open database");
        assert!(
            tree(&opened).0.tree.get().is_some(),
            "open deferred the catalog"
        );

        // The database leaf admits its content projection without decoding
        // the catalog, but the unified ODB handle must still reject a
        // catalog it cannot build at open.  Deferring every object tree would
        // move this refusal to the first read and change the public contract.
        let content = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:forms><d:component d:name="form"/></d:forms></o:database></o:body></o:document-content>"#;
        let malformed = content.replace("<d:forms>", "").replace("</d:forms>", "");
        let malformed = litchi_odb::Builder::new()
            .content_xml(malformed)
            .build()
            .expect("leaf package");
        let malformed_bytes = malformed.as_bytes().to_vec();
        litchi_odb::Package::from_bytes(&malformed_bytes).expect("leaf admission");
        assert!(
            crate::open_bytes(malformed_bytes).is_err(),
            "catalog was deferred"
        );
    }

    #[test]
    fn staging_refusals_keep_their_kind_through_the_odf_adapter() {
        let error = litchi_odf_common::Error::Core(litchi_core::Error::Refused(
            litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "set_shape",
                "the target does not accept this shape",
            ),
        ));
        let mapped = super::odf_staging_error(error, FileFormat::Odb);
        assert_eq!(mapped.code(), Code::UnsupportedType);
        assert!(matches!(mapped.leaf(), Source::Core(_)));
    }

    #[test]
    fn staging_preserves_source_and_limit_classifications() {
        let source_changed = litchi_core::Error::SourceChanged {
            expected: SourceVersion::new(0x10, 1),
            observed: SourceVersion::new(0x10, 2),
        };
        let mapped = super::odf_staging_error(
            litchi_odf_common::Error::Core(source_changed),
            FileFormat::Odb,
        );
        assert_eq!(mapped.code(), Code::IoError);
        assert!(matches!(mapped.leaf(), Source::Core(_)));

        let resource_limit = litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::InputBytes,
            observed: 11,
            limit: 10,
            scope: std::sync::Arc::from("ODF object bytes"),
        });
        let mapped = super::odf_staging_error(
            litchi_odf_common::Error::Core(resource_limit),
            FileFormat::Odb,
        );
        assert_eq!(mapped.code(), Code::LimitExceeded);
        assert!(matches!(mapped.leaf(), Source::Core(_)));
    }
}
