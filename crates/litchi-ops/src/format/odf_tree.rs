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
    tree: Tree,
    origin: Option<Arc<Path>>,
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

impl OdfTree {
    /// Adopt the package detection prepared, without a second ZIP index.
    pub(crate) fn from_prepared(prepared: PreparedPackage, _options: &OpenOptions) -> Result<Self> {
        let format = prepared.format();
        let error = |error| odf_error(error, format);
        let owner = match format {
            #[cfg(feature = "odg")]
            FileFormat::Odg => {
                Owner::Odg(litchi_odg::Package::from_prepared_package(prepared).map_err(error)?)
            },
            #[cfg(feature = "odc")]
            FileFormat::Odc => {
                Owner::Odc(litchi_odc::Package::from_prepared_package(prepared).map_err(error)?)
            },
            #[cfg(feature = "odi")]
            FileFormat::Odi => {
                Owner::Odi(litchi_odi::Package::from_prepared_package(prepared).map_err(error)?)
            },
            #[cfg(feature = "odb")]
            FileFormat::Odb => {
                Owner::Odb(litchi_odb::Package::from_prepared_package(prepared).map_err(error)?)
            },
            #[cfg(feature = "odf-formula")]
            FileFormat::Odf => Owner::Formula(
                litchi_odf_formula::Package::from_prepared_package(prepared).map_err(error)?,
            ),
            other => {
                let _ = (prepared, error);
                return Err(Error::unsupported_format(Some(other)));
            },
        };
        Self::from_owner(owner, None)
    }

    /// Adopt a leaf owner (an open, a commit result or a fresh package).
    pub(crate) fn from_owner(owner: Owner, origin: Option<Arc<Path>>) -> Result<Self> {
        let tree = owner.tree()?;
        Ok(Self(Arc::new(State {
            owner,
            tree,
            origin,
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

    /// The retained leaf owner.
    pub(crate) fn owner(&self) -> &Owner {
        &self.0.owner
    }

    /// The object tree built at open.
    pub(crate) fn tree(&self) -> &Tree {
        &self.0.tree
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
    /// None today; the `Result` keeps the read surface uniform.
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
            Owner::Formula(_) => None,
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
        Arc::ptr_eq(&self.0, &other.0) || self.0.owner.as_bytes() == other.0.owner.as_bytes()
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
        sink.write_all(self.0.owner.as_bytes()).map_err(|error| {
            Error::from(litchi_core::Error::from(error)).with_format(self.format())
        })
    }

    /// Atomically save the package to `path`.
    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        let bytes = self.0.owner.as_bytes();
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

    /// Moving elements is not offered by these leaves.
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

    /// Moving elements is not offered by these leaves.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        let _ = (path, placement, limits);
        Err(write::refusal(
            crate::Code::UnsupportedCapability,
            "moving elements is not supported for this OpenDocument format",
        )
        .with_format(self.format()))
    }

    /// Exchanging elements is not offered by these leaves.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<write::Commit> {
        self.swap_with_limits(left, right, &crate::transaction::SemanticLimits::default())
    }

    /// Exchanging elements is not offered by these leaves.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        let _ = (left, right, limits);
        Err(write::refusal(
            crate::Code::UnsupportedCapability,
            "swapping elements is not supported for this OpenDocument format",
        )
        .with_format(self.format()))
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
#[cfg(any(feature = "odg", feature = "odb"))]
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
