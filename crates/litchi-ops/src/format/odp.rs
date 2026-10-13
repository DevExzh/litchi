//! Handle for OpenDocument Presentation (`.odp`) inputs (record 0278).
//!
//! The handle retains the leaf owner it was opened as: the immutable
//! `litchi_odp::Package` with its parsed `Presentation` projection (byte
//! and reader routes, and every edited snapshot), the filesystem-backed
//! `litchi_odp::SourceBackedPresentation` (path route on `unix`/`windows`),
//! or the read-only `litchi_odp::FlatPresentation` of a flat `.fodp`
//! document. Unified reads walk the leaf's byte-range scene index of
//! `content.xml` (`litchi_odp::scene`), built once per handle on first use;
//! unified writes stage checked scene splices on one `Package::edit()`.

pub(crate) mod children;
mod nav;
mod node;
mod registry;
mod spec;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use registry::{ADDABLE, REMOVABLE, SETTABLE, WRITE_GAPS, capability_props};
pub use spec::Kind;
pub(crate) use spec::SPEC;

use std::fmt;
use std::io::Write as _;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::Metadata;
use litchi_core::detection::FileFormat;
use litchi_odf_common::PreparedPackage;
use litchi_odf_common::leniency::ToleranceReport;
use litchi_odp::scene::Scene;

use crate::{Code, Error, OpenOptions, Result};

/// An opened ODP presentation. Cheap to clone (one `Arc` bump);
/// `Send + Sync`.
#[derive(Clone)]
pub struct Odp(Arc<OdpState>);

struct OdpState {
    owner: OdpOwner,
    origin: Option<Arc<Path>>,
    odf_options: litchi_odf_common::OpenOptions,
    /// The scene index of the owner's `content.xml`, built on first use.
    scene: OnceLock<Scene>,
    /// Master pages, layouts and annotations, each parsed on first use.
    masters: OnceLock<Vec<litchi_odp::master::Summary>>,
    layouts: OnceLock<Vec<litchi_odp::layout::Layout>>,
    annotations: OnceLock<Vec<litchi_odp::annotation::Info>>,
    /// The package a source-backed owner materializes for its first write.
    materialized: OnceLock<litchi_odp::Package>,
}

/// The read owner, in the state the open route built.
#[allow(
    clippy::large_enum_variant,
    reason = "the owner lives once behind the handle's `Arc`; boxing a variant would add an allocation to that open route"
)]
pub(crate) enum OdpOwner {
    /// The immutable package and its projection (byte and reader routes,
    /// created and edited snapshots).
    Eager {
        package: litchi_odp::Package,
        presentation: litchi_odp::Presentation,
    },
    /// Filesystem-backed, with deferred package data (path route).
    #[cfg(any(unix, windows))]
    Source(litchi_odp::SourceBackedPresentation),
    /// A flat (single XML document) presentation; read-only.
    Flat(litchi_odp::FlatPresentation),
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Odp>();
    assert!(size_of::<Odp>() == size_of::<usize>());
};

/// Typed ODF limit failures stay `Source::Odf`; every other failure keeps
/// its common-core payload as `Source::Core`.
pub(crate) fn odf_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => core_error(error),
        other => Error::from(other).with_format(FileFormat::Odp),
    }
}

pub(crate) fn core_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Odp)
}

impl Odp {
    /// Adopts the package detection prepared without a second ZIP index.
    pub(crate) fn from_prepared(prepared: PreparedPackage, options: &OpenOptions) -> Result<Self> {
        let package = litchi_odp::Package::from_prepared_package(prepared).map_err(odf_error)?;
        Self::from_package(package, None, options.odf_open_options())
    }

    /// Adopts the source-backed owner the path route built over the file.
    #[cfg(any(unix, windows))]
    pub(crate) fn from_source(
        presentation: litchi_odp::SourceBackedPresentation,
        options: &OpenOptions,
    ) -> Result<Self> {
        Ok(Self::from_owner(
            OdpOwner::Source(presentation),
            None,
            options.odf_open_options(),
        ))
    }

    /// Adopts a flat (single XML) presentation for reads.
    pub(crate) fn from_flat(bytes: Vec<u8>, options: &OpenOptions) -> Result<Self> {
        let flat = litchi_odp::FlatPresentation::from_bytes(bytes).map_err(core_error)?;
        Ok(Self::from_owner(
            OdpOwner::Flat(flat),
            None,
            options.odf_open_options(),
        ))
    }

    pub(crate) fn from_package(
        package: litchi_odp::Package,
        origin: Option<Arc<Path>>,
        odf_options: litchi_odf_common::OpenOptions,
    ) -> Result<Self> {
        let presentation = package.presentation().map_err(core_error)?;
        Ok(Self::from_owner(
            OdpOwner::Eager {
                package,
                presentation,
            },
            origin,
            odf_options,
        ))
    }

    fn from_owner(
        owner: OdpOwner,
        origin: Option<Arc<Path>>,
        odf_options: litchi_odf_common::OpenOptions,
    ) -> Self {
        Self(Arc::new(OdpState {
            owner,
            origin,
            odf_options,
            scene: OnceLock::new(),
            masters: OnceLock::new(),
            layouts: OnceLock::new(),
            annotations: OnceLock::new(),
            materialized: OnceLock::new(),
        }))
    }

    /// The same snapshot, saving to `origin`.
    pub(crate) fn with_origin(self, origin: Option<Arc<Path>>) -> Self {
        let state = match Arc::try_unwrap(self.0) {
            Ok(mut state) => {
                state.origin = origin;
                state
            },
            Err(shared) => OdpState {
                owner: match &shared.owner {
                    OdpOwner::Eager {
                        package,
                        presentation,
                    } => OdpOwner::Eager {
                        package: package.clone(),
                        presentation: presentation.clone(),
                    },
                    // A source owner is adopted once by the opener; it is
                    // never shared before its origin is attached.
                    #[cfg(any(unix, windows))]
                    OdpOwner::Source(_) => {
                        return Self(shared);
                    },
                    OdpOwner::Flat(flat) => OdpOwner::Flat(flat.clone()),
                },
                origin,
                odf_options: shared.odf_options,
                scene: OnceLock::new(),
                masters: OnceLock::new(),
                layouts: OnceLock::new(),
                annotations: OnceLock::new(),
                materialized: OnceLock::new(),
            },
        };
        Self(Arc::new(state))
    }

    /// The path a path-backed open retained.
    pub(crate) fn origin(&self) -> Option<&Arc<Path>> {
        self.0.origin.as_ref()
    }

    /// Adopt the package one checked edit published, keeping the origin.
    pub(crate) fn with_package(&self, package: litchi_odp::Package) -> Result<Self> {
        Self::from_package(package, self.0.origin.clone(), self.0.odf_options)
    }

    /// Create a presentation with one blank slide.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        use crate::{Conformance, Flavor};
        let reason = if options.flavor_value() != Flavor::Document {
            Some("ODP template (.otp) creation is not supported")
        } else if options.conformance_value() == Conformance::Strict {
            Some("ODP has no OOXML Strict conformance")
        } else {
            None
        };
        if let Some(reason) = reason {
            return Err(Error::new(Code::UnsupportedCapability, reason)
                .with_format(FileFormat::Odp)
                .with_suggestion("use default CreateOptions for an ordinary presentation"));
        }
        let mut builder = litchi_odp::Builder::new();
        builder
            .add_slide_element(litchi_odp::Slide {
                title: None,
                text: String::new(),
                index: 0,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: Vec::new(),
            })
            .map_err(core_error)?;
        if let Some(metadata) = options.metadata_value() {
            let simple = Metadata {
                title: metadata.title.clone(),
                author: metadata.author.clone(),
                subject: metadata.subject.clone(),
                description: metadata.description.clone(),
                keywords: metadata.keywords.clone(),
                ..Metadata::default()
            };
            if simple != *metadata {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "initial ODP metadata supports title, author, subject, description and keywords only",
                )
                .with_format(FileFormat::Odp)
                .with_suggestion("set the other fields after creation is integrated (M6 ODP metadata)"));
            }
            builder.set_metadata(simple);
        }
        let package = builder.build().map_err(odf_error)?;
        Self::from_package(package, None, litchi_odf_common::OpenOptions::new())
    }

    /// The retained leaf owner.
    #[cfg(test)]
    pub(crate) fn owner(&self) -> &OdpOwner {
        &self.0.owner
    }

    /// Check a filesystem-backed source's revision.
    pub(crate) fn check_source(&self) -> Result<()> {
        match &self.0.owner {
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.check_source().map_err(core_error),
            _ => Ok(()),
        }
    }

    /// The owner's `content.xml` (the whole document for a flat input).
    pub(crate) fn content_xml(&self) -> Result<&str> {
        match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => Ok(presentation.content_xml()),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.content_xml().map_err(core_error),
            OdpOwner::Flat(flat) => std::str::from_utf8(flat.as_bytes()).map_err(|_| {
                Error::new(Code::CorruptFile, "the flat ODP document is not UTF-8")
                    .with_format(FileFormat::Odp)
            }),
        }
    }

    fn styles_xml(&self) -> Result<Option<&str>> {
        match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => Ok(presentation.styles_xml()),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.styles_xml().map_err(core_error),
            OdpOwner::Flat(_) => self.content_xml().map(Some),
        }
    }

    /// The scene index of the snapshot, built on first use.
    pub(crate) fn scene(&self) -> Result<&Scene> {
        if let Some(scene) = self.0.scene.get() {
            return Ok(scene);
        }
        let scene = Scene::parse(self.content_xml()?).map_err(core_error)?;
        Ok(self.0.scene.get_or_init(|| scene))
    }

    /// The master pages, listed by identity on first use.
    pub(crate) fn masters(&self) -> Result<&[litchi_odp::master::Summary]> {
        if let Some(masters) = self.0.masters.get() {
            return Ok(masters);
        }
        let masters = match self.styles_xml()? {
            Some(styles) => litchi_odp::master::list(styles).map_err(core_error)?,
            None => Vec::new(),
        };
        Ok(self.0.masters.get_or_init(|| masters))
    }

    /// The presentation page layouts, parsed on first use.
    pub(crate) fn layouts(&self) -> Result<&[litchi_odp::layout::Layout]> {
        if let Some(layouts) = self.0.layouts.get() {
            return Ok(layouts);
        }
        let layouts = match self.styles_xml()? {
            Some(styles) => {
                litchi_odp::layout::parse(styles)
                    .map_err(core_error)?
                    .layouts
            },
            None => Vec::new(),
        };
        Ok(self.0.layouts.get_or_init(|| layouts))
    }

    /// The annotation inventory, parsed on first use.
    pub(crate) fn annotations(&self) -> Result<&[litchi_odp::annotation::Info]> {
        if let Some(annotations) = self.0.annotations.get() {
            return Ok(annotations);
        }
        let annotations =
            litchi_odp::annotation::inventory(self.content_xml()?).map_err(core_error)?;
        Ok(self.0.annotations.get_or_init(|| annotations))
    }

    /// The immutable package writes start from: the eager owner, or a
    /// source-backed owner's file materialized once under its revision
    /// fence. Flat documents have no package.
    pub(crate) fn package(&self) -> Result<&litchi_odp::Package> {
        match &self.0.owner {
            OdpOwner::Eager { package, .. } => Ok(package),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => {
                if let Some(package) = self.0.materialized.get() {
                    presentation.check_source().map_err(core_error)?;
                    return Ok(package);
                }
                let origin = self.0.origin.as_ref().ok_or_else(|| {
                    Error::new(
                        Code::UnsupportedCapability,
                        "a source-backed ODP owner without a path cannot be edited",
                    )
                    .with_format(FileFormat::Odp)
                })?;
                presentation.check_source().map_err(core_error)?;
                let bytes = std::fs::read(origin).map_err(Error::from)?;
                presentation.check_source().map_err(core_error)?;
                let package =
                    litchi_odp::Package::from_bytes_with_options(&bytes, self.0.odf_options)
                        .map_err(odf_error)?;
                Ok(self.0.materialized.get_or_init(|| package))
            },
            OdpOwner::Flat(_) => Err(flat_write()),
        }
    }

    /// Whether this is a flat (single XML) document.
    pub(crate) fn is_flat(&self) -> bool {
        matches!(self.0.owner, OdpOwner::Flat(_))
    }

    /// Admit a write: the snapshot must have an editable package. A signed
    /// or encrypted package reports the refusal its leaf edit gives.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        let package = self.package()?;
        if package.security_policy().map_err(core_error)?.is_editable() {
            return Ok(());
        }
        drop(package.edit().map_err(odf_error)?);
        Ok(())
    }

    /// Whether two handles retain the same immutable snapshot.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        if Arc::ptr_eq(&self.0, &other.0) {
            return true;
        }
        match (&self.0.owner, &other.0.owner) {
            (OdpOwner::Eager { package: left, .. }, OdpOwner::Eager { package: right, .. }) => {
                std::ptr::eq(left.as_bytes(), right.as_bytes())
            },
            _ => false,
        }
    }

    /// Whether two handles publish the same bytes.
    pub(crate) fn same_content(&self, other: &Self) -> Result<bool> {
        if self.same_snapshot(other) {
            return Ok(true);
        }
        Ok(self.package()?.as_bytes() == other.package()?.as_bytes())
    }

    /// Set properties with the default unified limits.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: crate::TargetArg<'p>,
        props: impl crate::IntoProps<'v>,
    ) -> Result<write::WriteCommit> {
        self.set_with_limits(
            target,
            props,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Set properties through one checked leaf scene splice.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: crate::TargetArg<'p>,
        props: impl crate::IntoProps<'v>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::WriteCommit> {
        write::set_with_limits(self, target, props, limits)
    }

    /// Add a slide, a text box or a paragraph.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
    ) -> Result<(write::WriteCommit, crate::path::Path<'static>)> {
        self.add_with_limits(parent, node, &crate::transaction::SemanticLimits::default())
    }

    /// Add with the caller's transaction limits.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<(write::WriteCommit, crate::path::Path<'static>)> {
        write::add_with_limits(self, parent, node, limits)
    }

    /// Remove the selected slides, objects, paragraphs or runs.
    pub(crate) fn remove<'p>(&self, target: crate::TargetArg<'p>) -> Result<write::WriteCommit> {
        self.remove_with_limits(target, &crate::transaction::SemanticLimits::default())
    }

    /// Remove with the caller's transaction limits.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: crate::TargetArg<'p>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::WriteCommit> {
        write::remove_with_limits(self, target, limits)
    }

    /// Move a slide or a top-level slide object.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
    ) -> Result<write::WriteCommit> {
        self.move_to_with_limits(
            path,
            placement,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Move with the caller's transaction limits.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::WriteCommit> {
        write::move_to_with_limits(self, path, placement, limits)
    }

    /// Exchange two slides or two objects of one slide.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<write::WriteCommit> {
        self.swap_with_limits(left, right, &crate::transaction::SemanticLimits::default())
    }

    /// Swap with the caller's transaction limits.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<write::WriteCommit> {
        write::swap_with_limits(self, left, right, limits)
    }

    /// The text of every slide.
    ///
    /// # Errors
    ///
    /// Returns `corrupt_file` when a slide cannot be parsed, and `io_error`
    /// when a filesystem-backed source changed.
    pub fn text(&self) -> Result<String> {
        match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => presentation.text(),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.text(),
            OdpOwner::Flat(flat) => Ok(flat
                .slides()
                .iter()
                .map(litchi_odp::Slide::all_text)
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>()
                .join("\n\n")),
        }
        .map_err(core_error)
    }

    /// The number of slides.
    ///
    /// # Errors
    ///
    /// As for [`Self::text`].
    pub fn slide_count(&self) -> Result<usize> {
        match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => presentation.slide_count(),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.slide_count(),
            OdpOwner::Flat(flat) => Ok(flat.slide_count()),
        }
        .map_err(core_error)
    }

    /// The parsed slides, for the presentation family view.
    pub(crate) fn slides(&self) -> Result<Vec<litchi_odp::Slide>> {
        match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => presentation.slides(),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.slides(),
            OdpOwner::Flat(flat) => Ok(flat.slides().to_vec()),
        }
        .map_err(core_error)
    }

    /// One parsed slide, for the presentation family view.
    pub(crate) fn slide(&self, position: usize) -> Result<Option<litchi_odp::Slide>> {
        match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => presentation.slide(position),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.slide(position),
            OdpOwner::Flat(flat) => Ok(flat.slides().get(position).cloned()),
        }
        .map_err(core_error)
    }

    /// Document properties projected from `meta.xml`. A flat document
    /// reports the default (empty) metadata.
    ///
    /// # Errors
    ///
    /// Returns `io_error` when a filesystem-backed source changed and
    /// `corrupt_file` when `meta.xml` is malformed.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => {
                presentation.metadata().map(Some).map_err(core_error)
            },
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.metadata().map(Some).map_err(core_error),
            OdpOwner::Flat(_) => Ok(Some(Metadata::default())),
        }
    }

    /// Everything the reader repaired while opening the package and scanning
    /// the slides.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the read surface uniform across
    /// formats.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(match &self.0.owner {
            OdpOwner::Eager { presentation, .. } => presentation.tolerance_report().clone(),
            #[cfg(any(unix, windows))]
            OdpOwner::Source(presentation) => presentation.tolerance_report().clone(),
            OdpOwner::Flat(flat) => flat.tolerance_report().clone(),
        })
    }

    /// Publish the snapshot's exact package bytes.
    pub(crate) fn write_to(&self, mut sink: impl std::io::Write) -> Result<()> {
        if let OdpOwner::Flat(flat) = &self.0.owner {
            sink.write_all(flat.as_bytes()).map_err(Error::from)?;
            return Ok(());
        }
        let package = self.package()?;
        sink.write_all(package.as_bytes()).map_err(Error::from)?;
        self.check_source()
    }

    /// Atomically publish the snapshot to `path`.
    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        let bytes: &[u8] = match &self.0.owner {
            OdpOwner::Flat(flat) => flat.as_bytes(),
            _ => self.package()?.as_bytes(),
        };
        let overwrite = match options.overwrite {
            crate::Overwrite::Replace => litchi_core::atomic::Overwrite::Replace,
            crate::Overwrite::Refuse => litchi_core::atomic::Overwrite::Refuse,
        };
        litchi_core::atomic::save::<_, Error, _, _>(
            path,
            options.durability,
            overwrite,
            |staging| {
                staging.write_all(bytes).map_err(Error::from)?;
                Ok(u64::try_from(bytes.len()).unwrap_or(u64::MAX))
            },
            || self.check_source(),
        )
    }
}

fn flat_write() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "flat OpenDocument presentations (.fodp) are read-only in the unified API",
    )
    .with_format(FileFormat::Odp)
    .with_suggestion("follow-up: M6 flat ODF writes; convert to a packaged .odp to edit")
}

impl fmt::Debug for Odp {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let state = match &self.0.owner {
            OdpOwner::Eager { .. } => "eager",
            #[cfg(any(unix, windows))]
            OdpOwner::Source(_) => "source",
            OdpOwner::Flat(_) => "flat",
        };
        formatter
            .debug_struct("Odp")
            .field("owner", &state)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests;
