//! Handles for the Apple iWork formats (record 0296, owner decision D1):
//! Pages documents, Keynote presentations and Numbers workbooks.
//!
//! Each handle retains its leaf's immutable `Package` and an object tree
//! built from the leaf's read-only projection on the first read: the body,
//! its paragraphs and the sections of a Pages document; the slides, their
//! title and body placeholders and notes of a Keynote presentation; the
//! sheets, tables and stored cells of a Numbers workbook. Keynote slides and
//! their children and Numbers table cells are lazily built parts of the
//! tree, built as reads reach them (record 0305). Paths, queries and node
//! views run over that tree through one generic adapter (`Nav`),
//! parameterized by a per-format `Flavor`, in the shape of the ODF
//! object-tree adapter (record 0282). Writes stage the leaf's own focused
//! transactions (section text, slide text, slide skip/order/deletion, table
//! cells), address their targets through the leaf positions the tree
//! recorded, and adopt the committed package as a new handle whose tree is
//! built on its first read.
//!
//! Package directories (the app-authored bundle form) open for every format:
//! the leaf converts the bundle to its single-file form (records 0299 and
//! 0306), and `save()` writes the bundle form back in place (records 0301
//! and 0306).

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
mod error;
mod nav;
mod node;
mod part;
mod tree;
pub(crate) mod views;
pub(crate) mod write;

#[cfg(feature = "keynote")]
pub(crate) mod keynote;
#[cfg(feature = "numbers")]
pub mod numbers;
#[cfg(feature = "pages")]
pub(crate) mod pages;

pub use error::LeafError;
pub(crate) use nav::{Flavor, Nav};
pub use node::Node;
pub(crate) use part::{Kids, Part};
pub(crate) use tree::{At, Tree, TreeBuilder};

use std::fmt;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;

use crate::{Error, OpenOptions, Result};

/// An opened Pages document, Keynote presentation or Numbers workbook.
/// Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Iwork(Arc<State>);

struct State {
    owner: Owner,
    /// Built on the first read: the open route decodes only what the leaf
    /// open decodes.
    tree: OnceLock<Built>,
    /// Fixed per-container chart entries used by the canonical chart paths.
    /// Each entry is built only when its slide/sheet is selected; the normal
    /// tree remains lazy and unrelated container inventories are untouched.
    chart: OnceLock<Result<ChartCache>>,
    /// The complete Pages tree, whose section headers and footers are built
    /// the first time a read reaches them (record 0299). Every item of
    /// `tree` keeps its index here, so nodes taken from either tree stay
    /// valid.
    #[cfg(feature = "pages")]
    full: OnceLock<Built>,
    /// The global index of each Pages section's first paragraph, then the
    /// paragraph count: what a coalesced `/body/p[N]` set needs, without
    /// building the tree (record 0302).
    #[cfg(feature = "pages")]
    paragraph_starts: OnceLock<Box<[usize]>>,
    origin: Option<Arc<Path>>,
}

/// The object tree, or the error building it raised (a lazily decoded
/// Keynote slide can fail) beside a root-only tree; and its lazily built
/// parts, the `k`-th at [`At`] part `k` (record 0305).
struct Built {
    tree: Tree,
    error: Option<Error>,
    parts: Box<[Part]>,
}

/// The stable per-container namespace for canonical chart nodes.  The entry
/// boxes are fixed when the owner is opened, so a node's [`At`] remains valid
/// after another slide or sheet is selected and built.
struct ChartCache {
    entries: Box<[OnceLock<Built>]>,
}

impl ChartCache {
    fn new(owner: &Owner) -> Result<Self> {
        let count = match owner {
            #[cfg(feature = "pages")]
            Owner::Pages(_) => 1,
            #[cfg(feature = "keynote")]
            Owner::Keynote(package) => package
                .slides()
                .map_err(|error| leaf_error(LeafError::Keynote(error), FileFormat::Keynote))?
                .len(),
            #[cfg(feature = "numbers")]
            Owner::Numbers(package) => package.sheets().len(),
        };
        let mut entries = Vec::new();
        entries.try_reserve_exact(count).map_err(|_| {
            Error::new(
                crate::Code::InternalError,
                "could not allocate the canonical chart-container cache",
            )
        })?;
        entries.extend((0..count).map(|_| OnceLock::new()));
        Ok(Self {
            entries: entries.into_boxed_slice(),
        })
    }

    fn entry(&self, slot: usize) -> Option<&OnceLock<Built>> {
        self.entries.get(slot)
    }
}

impl Built {
    fn new(built: Result<(Tree, Box<[Part]>)>) -> Self {
        match built {
            Ok((tree, parts)) => Self {
                tree,
                error: None,
                parts,
            },
            Err(error) => Self {
                tree: Tree::root_only(),
                error: Some(error),
                parts: Box::new([]),
            },
        }
    }
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Iwork>();
    assert!(size_of::<Iwork>() == size_of::<usize>());
};

/// The leaf owner of one iWork format.
#[derive(Clone)]
pub(crate) enum Owner {
    #[cfg(feature = "pages")]
    Pages(litchi_pages::Package),
    #[cfg(feature = "keynote")]
    Keynote(litchi_keynote::Package),
    #[cfg(feature = "numbers")]
    Numbers(litchi_numbers::Package),
}

impl Owner {
    /// The format the owner reads.
    pub(crate) const fn format(&self) -> FileFormat {
        match self {
            #[cfg(feature = "pages")]
            Self::Pages(_) => FileFormat::Pages,
            #[cfg(feature = "keynote")]
            Self::Keynote(_) => FileFormat::Keynote,
            #[cfg(feature = "numbers")]
            Self::Numbers(_) => FileFormat::Numbers,
        }
    }

    /// Build the eager object tree of this owner and its lazily built
    /// parts.
    fn tree(&self) -> Result<(Tree, Box<[Part]>)> {
        match self {
            #[cfg(feature = "pages")]
            Self::Pages(package) => Ok(pages::build(package)),
            #[cfg(feature = "keynote")]
            Self::Keynote(package) => keynote::build(package),
            #[cfg(feature = "numbers")]
            Self::Numbers(package) => numbers::build(package),
        }
    }

    /// Stream the exact package bytes into `sink`.
    fn write_to<W: std::io::Write + ?Sized>(&self, sink: &mut W) -> Result<()> {
        let format = self.format();
        match self {
            #[cfg(feature = "pages")]
            Self::Pages(package) => package.write_to(sink).map_err(LeafError::PagesWrite),
            #[cfg(feature = "keynote")]
            Self::Keynote(package) => package.write_to(sink).map_err(LeafError::KeynoteWrite),
            #[cfg(feature = "numbers")]
            Self::Numbers(package) => package.write_to(sink).map_err(LeafError::NumbersWrite),
        }
        .map_err(|error| leaf_error(error, format))
    }

    /// Save the package as an app-style package directory at `path`.
    fn save_directory(&self, path: &Path) -> Result<()> {
        let format = self.format();
        match self {
            #[cfg(feature = "pages")]
            Self::Pages(package) => package.save_directory(path).map_err(LeafError::PagesSave),
            #[cfg(feature = "keynote")]
            Self::Keynote(package) => package.save_directory(path).map_err(LeafError::KeynoteSave),
            #[cfg(feature = "numbers")]
            Self::Numbers(package) => package.save_directory(path).map_err(LeafError::NumbersSave),
        }
        .map_err(|error| leaf_error(error, format))
    }

    /// The exact package bytes, for snapshot comparisons and size charges.
    fn bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

/// A leaf failure of `format`.
pub(crate) fn leaf_error(error: LeafError, format: FileFormat) -> Error {
    Error::from(error).with_format(format)
}

/// What the iWork arm of detection made of a ZIP input.
pub(crate) enum Detected {
    /// An iWork package of an enabled format, opened.
    Opened(Iwork),
    /// An iWork package whose cargo feature is off in this build.
    Disabled(FileFormat),
    /// Not an iWork package: the input, given back.
    Declined(Vec<u8>),
}

impl Iwork {
    /// Classify and open an in-memory ZIP package: the last ZIP candidate,
    /// after OOXML and ODF declined it. Detection prepares the container
    /// once and the leaf owner adopts it (no second ZIP pass).
    ///
    /// A caller's leaf options ([`OpenOptions::pages_options`],
    /// [`OpenOptions::keynote_options`], [`OpenOptions::numbers_options`])
    /// apply to the detected format: semantic limits are passed to the
    /// prepared-source ingress; non-default source limits make the leaf parse
    /// the input again under them (detection prepares under the hard
    /// ceilings), which costs a second container pass on that opt-in route
    /// only.
    pub(crate) fn detect_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Result<Detected> {
        let shared: Arc<[u8]> = Arc::from(bytes);
        let prepared = match litchi_iwa_detect::PreparedSource::from_shared_bytes(
            Arc::clone(&shared),
            litchi_iwa_detect::Limits::default(),
        ) {
            Ok(Some(prepared)) => prepared,
            // A container the iWork detector refuses is not claimed: the
            // input is reported as unrecognized, as before iWork joined.
            Ok(None) | Err(_) => return Ok(Detected::Declined(shared.to_vec())),
        };
        let owner = match prepared.format() {
            #[cfg(feature = "pages")]
            litchi_iwa_detect::Format::Pages => {
                let leaf = options.pages_options.unwrap_or_default();
                if leaf.source_limits_value() == litchi_pages::SourceLimits::default() {
                    drop(shared);
                    litchi_pages::Package::from_prepared_source(
                        prepared,
                        leaf.semantic_limits_value(),
                    )
                } else {
                    drop(prepared);
                    litchi_pages::Package::from_shared_bytes_with_options(shared, leaf)
                }
                .map(Owner::Pages)
                .map_err(|error| leaf_error(LeafError::Pages(error), FileFormat::Pages))?
            },
            #[cfg(feature = "keynote")]
            litchi_iwa_detect::Format::Keynote => {
                let leaf = options.keynote_options.unwrap_or_default();
                if leaf.source_limits_value() == litchi_keynote::SourceLimits::default() {
                    drop(shared);
                    litchi_keynote::Package::from_prepared_source(
                        prepared,
                        leaf.semantic_limits_value(),
                    )
                } else {
                    drop(prepared);
                    litchi_keynote::Package::from_shared_bytes_with_options(shared, leaf)
                }
                .map(Owner::Keynote)
                .map_err(|error| leaf_error(LeafError::Keynote(error), FileFormat::Keynote))?
            },
            #[cfg(feature = "numbers")]
            litchi_iwa_detect::Format::Numbers => {
                let leaf = options.numbers_options.unwrap_or_default();
                if leaf.source_limits_value() == litchi_numbers::SourceLimits::default() {
                    drop(shared);
                    litchi_numbers::Package::from_prepared_source(
                        prepared,
                        leaf.semantic_limits_value(),
                    )
                } else {
                    drop(prepared);
                    litchi_numbers::Package::from_shared_bytes_with_options(shared, leaf)
                }
                .map(Owner::Numbers)
                .map_err(|error| leaf_error(LeafError::Numbers(error), FileFormat::Numbers))?
            },
            #[allow(
                unreachable_patterns,
                reason = "the format arms depend on the enabled features"
            )]
            other => return Ok(Detected::Disabled(iwork_format(other))),
        };
        Ok(Detected::Opened(Self::from_owner(owner, None)))
    }

    /// Adopt a leaf owner (an open or a commit result). The tree is built
    /// on the first read.
    pub(crate) fn from_owner(owner: Owner, origin: Option<Arc<Path>>) -> Self {
        Self(Arc::new(State {
            owner,
            tree: OnceLock::new(),
            chart: OnceLock::new(),
            #[cfg(feature = "pages")]
            full: OnceLock::new(),
            #[cfg(feature = "pages")]
            paragraph_starts: OnceLock::new(),
            origin,
        }))
    }

    /// The Pages paragraph starts, computed by `build` on the first call.
    #[cfg(feature = "pages")]
    pub(crate) fn paragraph_starts(&self, build: impl FnOnce() -> Box<[usize]>) -> &[usize] {
        self.0.paragraph_starts.get_or_init(build)
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

    /// The retained Pages package, when this is a Pages document.
    #[cfg(feature = "pages")]
    #[must_use]
    pub fn pages(&self) -> Option<&litchi_pages::Package> {
        match self.owner() {
            Owner::Pages(package) => Some(package),
            #[allow(unreachable_patterns, reason = "the owner arms depend on features")]
            _ => None,
        }
    }

    /// The retained Keynote package, when this is a Keynote presentation.
    #[cfg(feature = "keynote")]
    #[must_use]
    pub fn keynote(&self) -> Option<&litchi_keynote::Package> {
        match self.owner() {
            Owner::Keynote(package) => Some(package),
            #[allow(unreachable_patterns, reason = "the owner arms depend on features")]
            _ => None,
        }
    }

    /// The retained Numbers package, when this is a Numbers workbook.
    #[cfg(feature = "numbers")]
    #[must_use]
    pub fn numbers(&self) -> Option<&litchi_numbers::Package> {
        match self.owner() {
            Owner::Numbers(package) => Some(package),
            #[allow(unreachable_patterns, reason = "the owner arms depend on features")]
            _ => None,
        }
    }

    fn built(&self) -> &Built {
        #[cfg(feature = "pages")]
        if let Some(full) = self.0.full.get() {
            return full;
        }
        self.0.tree.get_or_init(|| Built::new(self.0.owner.tree()))
    }

    fn chart_cache(&self) -> Result<&ChartCache> {
        match self.0.chart.get_or_init(|| ChartCache::new(self.owner())) {
            Ok(cache) => Ok(cache),
            Err(error) => Err(error.clone()),
        }
    }

    fn chart_part(slot: usize) -> Result<u32> {
        let slot = u32::try_from(slot).map_err(|_| {
            Error::new(
                crate::Code::LimitExceeded,
                "canonical chart container index exceeds the bounded namespace",
            )
        })?;
        if slot > tree::CHART_PART_BASE - tree::CHART_PART_MIN {
            return Err(Error::new(
                crate::Code::LimitExceeded,
                "canonical chart container count exceeds the bounded namespace",
            ));
        }
        Ok(tree::CHART_PART_BASE - slot)
    }

    fn checked_chart_built(&self, selector: ChartPath) -> Result<(u32, &Built)> {
        let slot = selector.container();
        let part = Self::chart_part(slot)?;
        let cache = self.chart_cache()?;
        let entry = cache.entry(slot).ok_or_else(|| {
            Error::new(
                crate::Code::PathNotFound,
                "requested chart container was not found",
            )
            .with_format(self.format())
        })?;
        let built = entry.get_or_init(|| {
            let built = match self.owner() {
                #[cfg(feature = "pages")]
                Owner::Pages(package) => pages::build_charts(package),
                #[cfg(feature = "keynote")]
                Owner::Keynote(package) => keynote::build_charts(package, slot),
                #[cfg(feature = "numbers")]
                Owner::Numbers(package) => numbers::build_charts(package, slot),
            }
            .map(|tree| (tree, Vec::<Part>::new().into_boxed_slice()));
            Built::new(built)
        });
        match &built.error {
            Some(error) => Err(error.clone()),
            None => Ok((part, built)),
        }
    }

    fn checked_chart_tree(&self, selector: ChartPath) -> Result<(u32, &Tree)> {
        let (part, built) = self.checked_chart_built(selector)?;
        Ok((part, &built.tree))
    }

    pub(crate) fn chart_item(&self, at: At) -> &tree::Item {
        let slot = at
            .chart_slot()
            .and_then(|slot| usize::try_from(slot).ok())
            .expect("chart At must carry a bounded cache slot");
        let cache = self.chart_cache().expect("chart cache must initialize");
        let built = cache
            .entry(slot)
            .and_then(|entry| entry.get())
            .and_then(|built| built.error.is_none().then_some(built))
            .expect("chart At must refer to an initialized chart entry");
        built.tree.item(at.index)
    }

    /// Complete the tree (views and whole-tree reads).
    pub(crate) fn expand_all(&self) {
        match &self.0.owner {
            #[cfg(feature = "pages")]
            Owner::Pages(package) => {
                self.0
                    .full
                    .get_or_init(|| Built::new(Ok(pages::build_full(package))));
            },
            #[allow(
                unreachable_patterns,
                reason = "the Keynote and Numbers trees build parts on demand"
            )]
            _ => self.expand_parts(),
        }
    }

    /// The object tree, built on the first call.
    ///
    /// # Errors
    ///
    /// The leaf error when the tree could not be built (a Keynote slide that
    /// fails to decode).
    pub(crate) fn checked_tree(&self) -> Result<&Tree> {
        let built = self.built();
        match &built.error {
            Some(error) => Err(error.clone()),
            None => Ok(&built.tree),
        }
    }

    /// The eager object tree; root-only when building it failed (every
    /// read entry point checks [`Self::checked_tree`] first). Items of the
    /// lazily built parts are read through [`Self::item`].
    pub(crate) fn tree(&self) -> &Tree {
        &self.built().tree
    }

    /// The `File` variant of this handle's format.
    pub(crate) fn into_file(self) -> crate::File {
        match self.owner() {
            #[cfg(feature = "pages")]
            Owner::Pages(_) => crate::File::Pages(self),
            #[cfg(feature = "keynote")]
            Owner::Keynote(_) => crate::File::Keynote(self),
            #[cfg(feature = "numbers")]
            Owner::Numbers(_) => crate::File::Numbers(self),
        }
    }

    /// The file's plain text, as the leaf renders it.
    ///
    /// # Errors
    ///
    /// The leaf's typed error when a lazily decoded part cannot be read.
    pub fn text(&self) -> Result<String> {
        let format = self.format();
        match self.owner() {
            #[cfg(feature = "pages")]
            Owner::Pages(package) => package.text().map_err(LeafError::Pages),
            #[cfg(feature = "keynote")]
            Owner::Keynote(package) => package.text().map_err(LeafError::Keynote),
            #[cfg(feature = "numbers")]
            Owner::Numbers(package) => package
                .workbook_ref()
                .plain_text()
                .map_err(LeafError::NumbersText),
        }
        .map_err(|error| leaf_error(error, format))
    }

    /// The document properties the leaf reads from the package metadata.
    ///
    /// # Errors
    ///
    /// The Keynote leaf's typed error when its properties cannot be read.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        match self.owner() {
            #[cfg(feature = "pages")]
            Owner::Pages(package) => Ok(Some(package.metadata())),
            #[cfg(feature = "keynote")]
            Owner::Keynote(package) => package
                .metadata()
                .map_err(|error| leaf_error(LeafError::Keynote(error), FileFormat::Keynote)),
            #[cfg(feature = "numbers")]
            Owner::Numbers(package) => Ok(Some(
                package
                    .workbook_ref()
                    .metadata()
                    .cloned()
                    .unwrap_or_default(),
            )),
        }
    }

    /// Whether two handles retain the same package.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    /// Whether two handles publish the same package bytes.
    pub(crate) fn same_content(&self, other: &Self) -> Result<bool> {
        if self.same_snapshot(other) {
            return Ok(true);
        }
        Ok(self.0.owner.bytes()? == other.0.owner.bytes()?)
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
        self.0.owner.write_to(&mut sink)
    }

    /// Atomically save the package to `path`: the leaf's exact bytes,
    /// staged in a sibling temporary file.
    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        // A package saved over a package directory (its origin when it was
        // opened from one) is written back as a directory (records 0301 and
        // 0306).
        if std::fs::symlink_metadata(path).is_ok_and(|metadata| metadata.is_dir()) {
            if options.overwrite == crate::Overwrite::Refuse {
                return Err(Error::new(
                    crate::Code::FileExists,
                    "the destination package directory exists",
                ));
            }
            self.0.owner.save_directory(path)?;
            let mut sink = std::io::sink();
            let mut written = 0_u64;
            self.0.owner.write_to(&mut Counter {
                inner: &mut sink,
                written: &mut written,
            })?;
            return Ok(written);
        }
        let overwrite = match options.overwrite {
            crate::Overwrite::Replace => litchi_core::atomic::Overwrite::Replace,
            crate::Overwrite::Refuse => litchi_core::atomic::Overwrite::Refuse,
        };
        let mut written = 0_u64;
        litchi_core::atomic::save(
            path,
            options.durability,
            overwrite,
            |writer| {
                let mut counter = Counter {
                    inner: writer,
                    written: &mut written,
                };
                self.0.owner.write_to(&mut counter)
            },
            || Ok::<(), Error>(()),
        )?;
        Ok(written)
    }

    /// A fresh package of `format`: the Pages (record 0299), Numbers (record
    /// 0297) and Keynote (record 0298) leaves' deterministic `Package::new()`.
    pub(crate) fn create(format: FileFormat, options: &crate::CreateOptions) -> Result<Self> {
        #[cfg(feature = "numbers")]
        if format == FileFormat::Numbers {
            return numbers::create(options)
                .map(|package| Self::from_owner(Owner::Numbers(package), None))
                .map_err(|error| error.with_format(format));
        }
        let _ = options;
        #[cfg(feature = "pages")]
        if format == FileFormat::Pages {
            let package = litchi_pages::Package::new()
                .map_err(|error| leaf_error(LeafError::Pages(error), FileFormat::Pages))?;
            return Ok(Self::from_owner(Owner::Pages(package), None));
        }
        #[cfg(feature = "keynote")]
        if format == FileFormat::Keynote {
            let package = litchi_keynote::Package::new();
            return Ok(Self::from_owner(Owner::Keynote(package), None));
        }
        Err(Error::new(
            crate::Code::UnsupportedCapability,
            "this build cannot create the format from scratch",
        )
        .with_format(format)
        .with_suggestion("open an existing package and edit it"))
    }

    /// Open an app-authored package directory on the single-pass route:
    /// capture and convert the bundle once (`litchi_iwa_archive::bundle`,
    /// under the hard physical ceilings detection also uses), then classify
    /// and open the converted bytes like any in-memory package, so the
    /// container is indexed once and the leaf adopts it. `Ok(None)` when the
    /// route does not apply: a caller's leaf source limits are set (the leaf
    /// then captures under them, [`Self::open_directory`]), a legacy
    /// application marker is present, the capture fails, or detection
    /// declines; the caller then classifies the
    /// directory and lets the leaf report the precise error.
    pub(crate) fn capture_directory(path: &Path, options: &OpenOptions) -> Result<Option<Self>> {
        // A pre-2013 application marker can contradict the index: the
        // classifier arbitrates those bundles.
        if leaf_source_limits_set(options)
            || ["index.xml", "index.apxl", "index.numbers"]
                .iter()
                .any(|marker| std::fs::symlink_metadata(path.join(marker)).is_ok())
        {
            return Ok(None);
        }
        let limits = litchi_iwa_archive::Limits::default();
        let Ok(bytes) = litchi_iwa_archive::bundle::capture(path, limits.into(), limits) else {
            return Ok(None);
        };
        if let Some(limit) = options.max_input_bytes {
            crate::open::check_input_len(u64::try_from(bytes.len()).unwrap_or(u64::MAX), limit)?;
        }
        Ok(match Self::detect_bytes(bytes, options)? {
            Detected::Opened(handle) => Some(handle.with_origin(Some(Arc::from(path)))),
            Detected::Disabled(_) | Detected::Declined(_) => None,
        })
    }

    /// Open an app-authored package directory of `format`: the leaf
    /// captures the bundle once, under the caller's leaf options, and
    /// converts it to its single-file form (records 0299 and 0306). The
    /// directory is the handle's origin, so `save()` writes the bundle back
    /// in place (records 0301 and 0306); `save_as` to a file path writes the
    /// single-file package. [`OpenOptions::max_input_bytes`] bounds the
    /// converted package.
    pub(crate) fn open_directory(
        path: &Path,
        format: FileFormat,
        options: &OpenOptions,
    ) -> Result<Self> {
        let _ = options;
        let owner = match format {
            #[cfg(feature = "pages")]
            FileFormat::Pages => Owner::Pages(
                litchi_pages::Package::open_with_options(
                    path,
                    options.pages_options.unwrap_or_default(),
                )
                .map_err(|error| leaf_error(LeafError::Pages(error), format))?,
            ),
            #[cfg(feature = "keynote")]
            FileFormat::Keynote => Owner::Keynote(
                litchi_keynote::Package::open_with_options(
                    path,
                    options.keynote_options.unwrap_or_default(),
                )
                .map_err(|error| leaf_error(LeafError::Keynote(error), format))?,
            ),
            #[cfg(feature = "numbers")]
            FileFormat::Numbers => Owner::Numbers(
                litchi_numbers::Package::open_with_options(
                    path,
                    options.numbers_options.unwrap_or_default(),
                )
                .map_err(|error| leaf_error(LeafError::Numbers(error), format))?,
            ),
            _ => return Err(Error::unsupported_format(Some(format))),
        };
        let handle = Self::from_owner(owner, Some(Arc::from(path)));
        if let Some(limit) = options.max_input_bytes {
            let mut length = 0_u64;
            handle.0.owner.write_to(&mut Counter {
                inner: &mut std::io::sink(),
                written: &mut length,
            })?;
            crate::open::check_input_len(length, limit)?;
        }
        Ok(handle)
    }

    /// Tag an error with this handle's format.
    pub(crate) fn tag(&self, error: Error) -> Error {
        if error.format().is_none() {
            error.with_format(self.format())
        } else {
            error
        }
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

    /// Add one element with the default unified limits.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        self.add_with_limits(parent, node, &crate::transaction::SemanticLimits::default())
    }

    /// Remove targets with the default unified limits.
    pub(crate) fn remove<'p>(&self, target: crate::TargetArg<'p>) -> Result<write::Commit> {
        self.remove_with_limits(target, &crate::transaction::SemanticLimits::default())
    }

    /// Move an element with the default unified limits.
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

    /// Exchange two elements with the default unified limits.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<write::Commit> {
        self.swap_with_limits(left, right, &crate::transaction::SemanticLimits::default())
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

    /// Add one element with the caller's transaction limits.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        write::add(self, parent, node, limits).map_err(|error| self.tag(error))
    }

    /// Remove targets with the caller's transaction limits.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: crate::TargetArg<'p>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::remove(self, target, limits).map_err(|error| self.tag(error))
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

    /// Exchange two elements with the caller's transaction limits.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::Commit> {
        write::swap(self, left, right, limits).map_err(|error| self.tag(error))
    }
}

/// Counts the bytes an atomic save stages.
struct Counter<'w, W> {
    inner: &'w mut W,
    written: &'w mut u64,
}

impl<W: std::io::Write> std::io::Write for Counter<'_, W> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(buffer)?;
        *self.written += u64::try_from(written).unwrap_or(u64::MAX);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

/// Whether a caller selected leaf source limits other than the defaults for
/// any iWork format: the routes that prepare under the hard ceilings then
/// hand over to the leaf.
fn leaf_source_limits_set(options: &OpenOptions) -> bool {
    let _ = options;
    #[cfg(feature = "pages")]
    if options
        .pages_options
        .is_some_and(|leaf| leaf.source_limits_value() != litchi_pages::SourceLimits::default())
    {
        return true;
    }
    #[cfg(feature = "keynote")]
    if options
        .keynote_options
        .is_some_and(|leaf| leaf.source_limits_value() != litchi_keynote::SourceLimits::default())
    {
        return true;
    }
    #[cfg(feature = "numbers")]
    if options
        .numbers_options
        .is_some_and(|leaf| leaf.source_limits_value() != litchi_numbers::SourceLimits::default())
    {
        return true;
    }
    false
}

/// The unified format of an iWork detection result.
const fn iwork_format(format: litchi_iwa_detect::Format) -> FileFormat {
    match format {
        litchi_iwa_detect::Format::Pages => FileFormat::Pages,
        litchi_iwa_detect::Format::Keynote => FileFormat::Keynote,
        litchi_iwa_detect::Format::Numbers => FileFormat::Numbers,
    }
}

/// Classify an iWork package directory (the app-authored bundle form).
pub(crate) fn classify_directory(path: &Path) -> Option<FileFormat> {
    match litchi_iwa_detect::path(path, litchi_iwa_detect::Limits::default()) {
        Ok(Some(format)) => Some(iwork_format(format)),
        _ => None,
    }
}

impl fmt::Debug for Iwork {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Iwork")
            .field("format", &self.format())
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

/// A readable text property whose `null` value clears the focused leaf state.
pub(crate) const fn clearable_writable(
    name: &'static str,
    ty: crate::value::ValueType,
) -> crate::node::spec::PropSpec {
    writable(name, ty).clearable()
}

/// Runs `$body` with `$flavor` bound to the flavor type of `$format`.
macro_rules! with_flavor {
    ($format:expr, $flavor:ident => $body:expr) => {
        match $format {
            #[cfg(feature = "pages")]
            FileFormat::Pages => {
                type $flavor = pages::Pages;
                $body
            },
            #[cfg(feature = "keynote")]
            FileFormat::Keynote => {
                type $flavor = keynote::Keynote;
                $body
            },
            #[cfg(feature = "numbers")]
            FileFormat::Numbers => {
                type $flavor = numbers::Numbers;
                $body
            },
            other => Err(Error::unsupported_format(Some(other))),
        }
    };
}

/// `File::get` over an iWork tree.
pub(crate) fn get<'f>(
    file: &'f Iwork,
    path: &crate::path::PathArg<'_>,
) -> Result<crate::node::Node<'f>> {
    if let Some(selector) = chart_path(file.format(), path) {
        let (part, chart_tree) = file.checked_chart_tree(selector)?;
        let found = chart_tree.items().iter().enumerate().find(|(_, item)| {
            if item.type_name() != "chart" {
                return false;
            }
            let leaf = item.leaf();
            match selector {
                ChartPath::Pages(chart) => leaf[0] as usize == chart,
                ChartPath::Keynote { slide, chart } => {
                    leaf[0] as usize == slide && leaf[2] as usize == chart
                },
                ChartPath::Numbers { sheet, chart } => {
                    leaf[0] as usize == sheet && leaf[1] as usize == chart
                },
            }
        });
        if let Some((index, _)) = found {
            let index = u32::try_from(index).map_err(|_| {
                Error::new(
                    crate::Code::LimitExceeded,
                    "canonical chart item index exceeds the bounded namespace",
                )
                .with_format(file.format())
            })?;
            return Ok(crate::node::Node::Iwork(Node::new(
                file,
                At::chart(part, index),
            )));
        }
        return Err(
            Error::new(crate::Code::PathNotFound, "requested chart was not found")
                .with_format(file.format()),
        );
    }
    file.checked_tree()?;
    with_flavor!(file.format(), F => {
        crate::path::resolve::get(&Nav::<F>::new(file), path).map(crate::node::Node::Iwork)
    })
}

#[derive(Clone, Copy)]
enum ChartPath {
    Pages(usize),
    Keynote { slide: usize, chart: usize },
    Numbers { sheet: usize, chart: usize },
}

impl ChartPath {
    const fn container(self) -> usize {
        match self {
            Self::Pages(_) => 0,
            Self::Keynote { slide, .. } => slide,
            Self::Numbers { sheet, .. } => sheet,
        }
    }
}

fn chart_path(format: FileFormat, path: &crate::path::PathArg<'_>) -> Option<ChartPath> {
    if !path.is_absolute() {
        return None;
    }
    let steps = path.path().steps();
    let no_predicates = |step: &litchi_core::selector::Step<'_>| step.predicates().is_empty();
    match (format, steps) {
        (FileFormat::Pages, [body, chart])
            if body.name() == "body"
                && no_predicates(body)
                && chart.name() == "chart"
                && chart.predicates().len() == 1 =>
        {
            chart.index().map(|index| ChartPath::Pages(index.get()))
        },
        (FileFormat::Keynote, [slide, chart])
            if slide.name() == "slide"
                && chart.name() == "chart"
                && slide.predicates().len() == 1
                && chart.predicates().len() == 1 =>
        {
            Some(ChartPath::Keynote {
                slide: slide.index()?.get(),
                chart: chart.index()?.get(),
            })
        },
        (FileFormat::Numbers, [sheet, chart])
            if sheet.name() == "sheet"
                && chart.name() == "chart"
                && sheet.predicates().len() == 1
                && chart.predicates().len() == 1 =>
        {
            Some(ChartPath::Numbers {
                sheet: sheet.index()?.get(),
                chart: chart.index()?.get(),
            })
        },
        _ => None,
    }
}

/// `File::find` over an iWork tree.
pub(crate) fn find<'f>(
    file: &'f Iwork,
    path: &crate::path::PathArg<'_>,
) -> Result<Option<crate::node::Node<'f>>> {
    file.checked_tree()?;
    with_flavor!(file.format(), F => {
        crate::path::resolve::find(&Nav::<F>::new(file), path)
            .map(|node| node.map(crate::node::Node::Iwork))
    })
}

/// The query walk of one iWork format.
pub(crate) enum Walks<'q, 'f> {
    #[cfg(feature = "pages")]
    Pages(crate::query::walk::Walk<'q, 'f, Nav<'f, pages::Pages>>),
    #[cfg(feature = "keynote")]
    Keynote(crate::query::walk::Walk<'q, 'f, Nav<'f, keynote::Keynote>>),
    #[cfg(feature = "numbers")]
    Numbers(crate::query::walk::Walk<'q, 'f, Nav<'f, numbers::Numbers>>),
}

/// Applies `$body` to the walk of any iWork format.
macro_rules! each_walk {
    ($walks:expr, $walk:ident => $body:expr) => {
        match $walks {
            #[cfg(feature = "pages")]
            Walks::Pages($walk) => $body,
            #[cfg(feature = "keynote")]
            Walks::Keynote($walk) => $body,
            #[cfg(feature = "numbers")]
            Walks::Numbers($walk) => $body,
        }
    };
}

impl<'q, 'f> Walks<'q, 'f> {
    pub(crate) fn limit(self, n: usize) -> Self {
        match self {
            #[cfg(feature = "pages")]
            Self::Pages(walk) => Self::Pages(walk.limit(n)),
            #[cfg(feature = "keynote")]
            Self::Keynote(walk) => Self::Keynote(walk.limit(n)),
            #[cfg(feature = "numbers")]
            Self::Numbers(walk) => Self::Numbers(walk.limit(n)),
        }
    }

    pub(crate) fn plan(&self) -> &crate::query::Plan<'q> {
        each_walk!(self, walk => walk.plan())
    }

    pub(crate) fn next(&mut self) -> Option<Result<crate::node::Node<'f>>> {
        each_walk!(self, walk => walk.next().map(|node| node.map(crate::node::Node::Iwork)))
    }
}

/// `File::query` over an iWork tree.
pub(crate) fn query<'f, 'q>(
    file: &'f Iwork,
    selector: impl crate::query::IntoQuery<'q>,
) -> Result<Walks<'q, 'f>> {
    file.checked_tree()?;
    Ok(match file.format() {
        #[cfg(feature = "pages")]
        FileFormat::Pages => Walks::Pages(crate::query::run(Nav::new(file), selector)?),
        #[cfg(feature = "keynote")]
        FileFormat::Keynote => Walks::Keynote(crate::query::run(Nav::new(file), selector)?),
        #[cfg(feature = "numbers")]
        FileFormat::Numbers => Walks::Numbers(crate::query::run(Nav::new(file), selector)?),
        other => {
            let _ = selector;
            return Err(Error::unsupported_format(Some(other)));
        },
    })
}

/// Compile `query` for an iWork format.
pub(crate) fn compile<'q>(
    query: &crate::query::Query<'q>,
    format: FileFormat,
) -> Result<crate::query::Plan<'q>> {
    with_flavor!(format, F => crate::query::compile::compile::<Nav<'static, F>>(query))
}

/// The write registry of one iWork format: the elements `set`, `add`,
/// `remove` and `move` accept, and its operation gaps.
pub(crate) struct Writes {
    pub(crate) settable: &'static [&'static str],
    pub(crate) addable: &'static [&'static str],
    pub(crate) removable: &'static [&'static str],
    pub(crate) movable: &'static [&'static str],
    pub(crate) gaps: &'static [crate::node::spec::GapSpec],
}

/// The write registry of `adapter`, when it is an iWork format.
pub(crate) const fn writes(adapter: crate::read::Adapter) -> Option<&'static Writes> {
    match adapter {
        #[cfg(feature = "pages")]
        crate::read::Adapter::Pages => Some(&pages::WRITES),
        #[cfg(feature = "keynote")]
        crate::read::Adapter::Keynote => Some(&keynote::WRITES),
        #[cfg(feature = "numbers")]
        crate::read::Adapter::Numbers => Some(&numbers::WRITES),
        #[allow(
            unreachable_patterns,
            reason = "the other adapters write through their own tables"
        )]
        _ => None,
    }
}
