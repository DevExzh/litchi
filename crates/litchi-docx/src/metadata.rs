//! Document properties as a scoped transaction: the core properties
//! (`docProps/core.xml`), and the application (`docProps/app.xml`) and custom
//! (`docProps/custom.xml`) properties the same edit can stage.
//!
//! [`Package::metadata`](crate::Package::metadata) reads the typed core
//! properties, keeping a package without a core-properties part distinct from
//! one with an empty part. Every change goes through the ADR 0034 §7 family:
//! [`Package::edit_metadata`](crate::Package::edit_metadata) stages an
//! [`Edit`], [`Edit::commit`] publishes a new [`Package`] and
//! leaves the source untouched, and
//! [`Package::apply_metadata`](crate::Package::apply_metadata) replays a
//! [`Patch`] onto another snapshot with the same base value.
//!
//! The family is a thin DOCX wrapper over the shared OOXML codec
//! ([`litchi_ooxml_common::properties`]): the codec keeps the part's dialect,
//! name and relationship identity, refuses shared or ambiguous graphs, and
//! publishes nothing for a value equal to the stored one.
//!
//! [`Edit::set_app_props`] and [`Edit::set_custom_props`] stage the
//! application and custom properties (record 0342). [`Edit::remove_custom_props`]
//! explicitly removes the custom-properties graph. They are written by the
//! shared engine's exact part patch ([`litchi_ooxml_common::properties::Patch`]):
//! the application part is spliced, so every element the typed model does
//! not carry keeps its bytes, and an unchanged value publishes nothing.
//!
//! ```
//! use litchi_docx::Package;
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata();
//! let mut app = edit.app_props()?.unwrap_or_default();
//! app.company = Some("Acme".to_owned());
//! edit.set_app_props(Some(app));
//! let next = edit.commit()?.into_package();
//! assert_eq!(
//!     next.app_props()?.and_then(|props| props.company),
//!     Some("Acme".to_owned())
//! );
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ```
//! use litchi_docx::Package;
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata();
//! edit.update(|props| props.title = Some("Quarterly report".to_owned()));
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let next = commit.into_package();
//! assert_eq!(
//!     next.metadata().and_then(|props| props.title.as_deref()),
//!     Some("Quarterly report")
//! );
//! // The source snapshot is unchanged.
//! assert_eq!(package.metadata().and_then(|props| props.title.as_deref()), None);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub use litchi_core::Metadata;
pub use litchi_core::edit::Diagnostics;
use litchi_ooxml_common::custom::Props as CustomProps;
use litchi_ooxml_common::properties::{self as families, app};

use crate::Package;

/// The largest accepted text value of one core property, in UTF-8 bytes.
///
/// This is the shared OOXML codec's per-property bound; the transaction
/// checks it before it stages any package change so an oversized value is a
/// typed [`Error::Limit`] instead of a codec failure.
pub const MAX_PROPERTY_TEXT_BYTES: usize = 1_048_576;

/// Result alias of the metadata family.
pub type Result<T> = std::result::Result<T, Error>;

/// The bounded resource a metadata transaction reports.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// UTF-8 bytes of one text-valued core property.
    PropertyTextBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::PropertyTextBytes => "property text bytes",
        })
    }
}

/// A metadata transaction failure.
///
/// Messages carry structural metadata only, never property values.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The package refused the change or its core-properties graph is
    /// invalid.
    #[error(transparent)]
    Package(#[from] crate::Error),
    /// A changed write was attempted on a signed package. Remove the
    /// signature explicitly with [`Package::unsign`](crate::Package::unsign)
    /// first; an exact no-op keeps the signature.
    #[error("changing the core properties of a signed DOCX package requires Package::unsign")]
    Signed,
    /// The patch's base value does not match the package it is applied to.
    #[error("the metadata patch base does not match the package's core properties")]
    PatchConflict,
    /// The application- or custom-properties part refused the staged value,
    /// or its graph is invalid (record 0342).
    #[error(transparent)]
    Properties(#[from] families::Error),
    /// A staged value exceeds a metadata bound.
    #[error("metadata {kind} limit exceeded: {observed} > {maximum}")]
    Limit {
        /// Bounded resource.
        kind: LimitKind,
        /// Observed value.
        observed: usize,
        /// Maximum accepted value.
        maximum: usize,
    },
}

/// A replayable document-properties change: the core-properties value the
/// edit started from and the value it published (`None` is an absent
/// core-properties part), and the exact part changes of the application and
/// custom properties. The exact core delta planned over the post-parts graph
/// and the source-scoped application/custom delta are retained for
/// byte-faithful replay and inversion.
#[derive(Debug, Clone, PartialEq)]
pub struct Patch {
    base: Option<Metadata>,
    target: Option<Metadata>,
    parts: families::Patch,
    /// The exact shared-codec core patch over the graph after `parts` has
    /// been applied.
    core: families::Patch,
    /// Whether replay must apply the application/custom patch before the core
    /// patch. Forward edits are parts-first; inverses are core-first.
    parts_first: bool,
}

impl Patch {
    /// The core properties the patch expects to find.
    #[must_use]
    pub fn base(&self) -> Option<&Metadata> {
        self.base.as_ref()
    }

    /// The core properties the patch publishes.
    #[must_use]
    pub fn target(&self) -> Option<&Metadata> {
        self.target.as_ref()
    }

    /// The exact application/custom part changes planned over the source
    /// package, independent of any core transition in this patch.
    #[must_use]
    pub const fn parts(&self) -> &families::Patch {
        &self.parts
    }

    pub(crate) const fn core_patch(&self) -> &families::Patch {
        &self.core
    }

    pub(crate) const fn parts_first(&self) -> bool {
        self.parts_first
    }

    /// Whether applying the patch leaves every package byte unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.base == self.target && self.core.is_noop() && self.parts.is_noop()
    }

    /// The patch that restores the base value.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target.clone(),
            target: self.base.clone(),
            parts: self.parts.inverse(),
            core: self.core.inverse(),
            parts_first: !self.parts_first,
        }
    }
}

/// A published metadata change.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// A staged core-properties change over one package snapshot.
///
/// Dropping the edit discards it; the source package is never mutated.
#[derive(Clone)]
pub struct Edit {
    source: Package,
    base: Option<Metadata>,
    staged: Option<Metadata>,
    /// Staged application properties; `None` leaves the part untouched.
    app: Option<Option<app::Props>>,
    /// Staged custom properties; `None` leaves the part untouched.
    custom: Option<CustomProps>,
    /// Explicit custom-properties graph removal, distinct from setting an
    /// equal empty semantic collection.
    remove_custom: bool,
}

impl Edit {
    pub(crate) fn new(source: Package, base: Option<Metadata>) -> Self {
        Self {
            staged: base.clone(),
            source,
            base,
            app: None,
            custom: None,
            remove_custom: false,
        }
    }

    /// The staged core properties.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.staged.as_ref()
    }

    /// Replace the staged core properties with `props`.
    pub fn set(&mut self, props: Metadata) -> &mut Self {
        self.staged = Some(props);
        self
    }

    /// Change the staged core properties in place, starting from an empty
    /// value when the package has none.
    pub fn update(&mut self, change: impl FnOnce(&mut Metadata)) -> &mut Self {
        change(self.staged.get_or_insert_with(Metadata::default));
        self
    }

    /// Stage the removal of the core-properties part.
    pub fn clear(&mut self) -> &mut Self {
        self.staged = None;
        self
    }

    /// The staged application properties (`docProps/app.xml`): the staged
    /// value, or the package's. `None` is an absent part.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Properties`] for an invalid application-properties
    /// part or graph.
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        match &self.app {
            Some(staged) => Ok(staged.clone()),
            None => self.source.app_props(),
        }
    }

    /// Stage the complete application properties; `None` removes the part.
    pub fn set_app_props(&mut self, props: Option<app::Props>) -> &mut Self {
        self.app = Some(props);
        self
    }

    /// The staged custom properties: the staged value, or the package's.
    #[must_use]
    pub fn custom_props(&self) -> &CustomProps {
        self.custom
            .as_ref()
            .unwrap_or_else(|| self.source.custom_props())
    }

    /// Stage the complete custom properties; an empty collection removes the
    /// part when it changes the semantic collection. Setting an equal empty
    /// collection is an exact no-op, preserving a present raw part.
    pub fn set_custom_props(&mut self, props: CustomProps) -> &mut Self {
        self.custom = Some(props);
        self.remove_custom = false;
        self
    }

    /// Stage explicit removal of the custom-properties part and relationship.
    ///
    /// This is distinct from setting an equal empty collection: it removes a
    /// present raw part even when its tolerant semantic projection is empty.
    /// An absent graph remains an exact no-op.
    pub fn remove_custom_props(&mut self) -> &mut Self {
        self.custom = Some(CustomProps::new());
        self.remove_custom = true;
        self
    }

    /// Publish the staged value as a new package.
    ///
    /// An unchanged value returns a package that shares every part with the
    /// source, signature included.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for an oversized text property,
    /// [`Error::Signed`] for a change to a signed package,
    /// [`Error::Properties`] when the application or custom properties
    /// cannot be written, and [`Error::Package`] when the codec or package
    /// graph refuses the value.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            source,
            base,
            staged,
            app,
            custom,
            remove_custom,
        } = self;
        let (package, changed, parts, core) =
            source.publish_metadata(staged.as_ref(), app, custom, remove_custom)?;
        let patch = Patch {
            base,
            target: staged,
            parts,
            core,
            parts_first: true,
        };
        Ok(Commit::new(package, patch, Diagnostics::new(changed)))
    }
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("present", &self.staged.is_some())
            .field("app", &self.app.is_some())
            .field("custom", &self.custom.is_some())
            .finish_non_exhaustive()
    }
}

/// Check every text-valued property core properties store against
/// [`MAX_PROPERTY_TEXT_BYTES`].
pub(crate) fn admit(props: &Metadata) -> Result<()> {
    let texts = [
        props.title.as_deref(),
        props.subject.as_deref(),
        props.author.as_deref(),
        props.description.as_deref(),
        props.identifier.as_deref(),
        props.last_modified_by.as_deref(),
        props.category.as_deref(),
        props.content_status.as_deref(),
        props.revision.as_deref(),
        props.version.as_deref(),
        props.language.as_deref(),
    ];
    // The keywords as one `cp:keywords` text: joined with `, `.
    let keywords = props
        .keywords
        .iter()
        .map(String::len)
        .sum::<usize>()
        .saturating_add(props.keywords.len().saturating_sub(1).saturating_mul(2));
    let observed = texts
        .into_iter()
        .flatten()
        .map(str::len)
        .chain(std::iter::once(keywords))
        .max()
        .unwrap_or(0);
    if observed > MAX_PROPERTY_TEXT_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::PropertyTextBytes,
            observed,
            maximum: MAX_PROPERTY_TEXT_BYTES,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::part::BlobPart;

    fn package_with_opaque_app() -> Result<Package> {
        let package = Package::new()?;
        let mut opc = package.opc_package().clone();
        let app = PackURI::new("/docProps/app.xml").expect("static app URI");
        let part = opc.get_part_mut(&app).map_err(crate::Error::from)?;
        let mut bytes = part.blob().to_vec();
        let close = b"</Properties>";
        let offset = bytes
            .windows(close.len())
            .position(|window| window == close)
            .expect("default app root");
        bytes.splice(offset..offset, br#"<Opaque keep="yes"/>"#.iter().copied());
        part.set_blob(bytes);
        Ok(Package::from_opc_package(opc)?)
    }

    fn package_with_opaque_core() -> Result<Package> {
        let package = Package::new()?;
        let mut opc = package.opc_package().clone();
        let core = PackURI::new("/docProps/core.xml").expect("static core URI");
        let part = opc.get_part_mut(&core).map_err(crate::Error::from)?;
        let mut bytes = part.blob().to_vec();
        let close = b"/>";
        let offset = bytes
            .windows(close.len())
            .position(|window| window == close)
            .expect("default core root");
        bytes.splice(
            offset..offset + close.len(),
            br#"><!--opaque keep="yes"--></cp:coreProperties>"#.iter().copied(),
        );
        part.set_blob(bytes);
        Ok(Package::from_opc_package(opc)?)
    }

    fn package_without_parts(names: &[&str]) -> Result<Package> {
        let package = Package::new()?;
        let mut opc = package.opc_package().clone();
        for &name in names {
            let part = PackURI::new(name).expect("static property URI");
            let relationship_ids: Vec<_> = opc
                .rels()
                .iter()
                .filter(|relationship| relationship.target_partname().ok().as_ref() == Some(&part))
                .map(|relationship| relationship.r_id().to_owned())
                .collect();
            for relationship_id in relationship_ids {
                opc.rels_mut().remove(&relationship_id);
            }
            opc.remove_part(&part);
        }
        Ok(Package::from_opc_package(opc)?)
    }

    fn package_without_property_parts() -> Result<Package> {
        package_without_parts(&[
            "/docProps/core.xml",
            "/docProps/app.xml",
            "/docProps/custom.xml",
        ])
    }

    fn package_without_app_part() -> Result<Package> {
        package_without_parts(&["/docProps/app.xml"])
    }

    fn signed_package(package: &Package) -> Result<Package> {
        let mut opc = package.opc_package().clone();
        opc.try_add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").expect("static signature URI"),
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            Vec::new(),
        )))
        .map_err(crate::Error::from)?;
        opc.rels_mut().add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rSignature".to_owned(),
            false,
        );
        Ok(Package::from_opc_package(opc)?)
    }

    fn app_bytes(package: &Package) -> Vec<u8> {
        package
            .opc_package()
            .get_part(&PackURI::new("/docProps/app.xml").expect("static app URI"))
            .expect("app properties")
            .blob()
            .to_vec()
    }

    fn core_bytes(package: &Package) -> Vec<u8> {
        package
            .opc_package()
            .get_part(&PackURI::new("/docProps/core.xml").expect("static core URI"))
            .expect("core properties")
            .blob()
            .to_vec()
    }

    #[test]
    fn edit_commit_publishes_a_new_package_and_keeps_the_source() -> Result<()> {
        let package = package_with_opaque_app()?;
        let before = package.metadata().cloned();
        let app_before = package.app_props()?.expect("app properties");
        let app_bytes_before = app_bytes(&package);
        let mut edit = package.edit_metadata();
        edit.update(|props| {
            props.title = Some("T".to_owned());
            props.author = Some("A".to_owned());
        });
        let commit = edit.commit()?;
        assert!(commit.diagnostics().changed());
        assert!(!commit.patch().is_noop());
        assert_eq!(package.metadata().cloned(), before);
        let next = commit.package().clone();
        assert_eq!(
            next.metadata().and_then(|props| props.title.as_deref()),
            Some("T")
        );
        let app_after = next.app_props()?.expect("app properties");
        assert_eq!(app_after.application, app_before.application);
        assert_eq!(app_after.template, app_before.template);
        assert_eq!(app_bytes(&next), app_bytes_before);

        let bytes = next.to_bytes()?;
        let reopened = Package::from_bytes(&bytes)?;
        assert_eq!(
            reopened
                .metadata()
                .and_then(|props| props.author.as_deref()),
            Some("A")
        );
        Ok(())
    }

    #[test]
    fn unchanged_commit_is_an_exact_noop() -> Result<()> {
        let package = Package::new()?;
        let commit = package.edit_metadata().commit()?;
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            package.clone().to_bytes()?,
            commit.package().clone().to_bytes()?
        );
        let replayed = package.apply_metadata(commit.patch())?;
        assert!(!replayed.diagnostics().changed());
        assert!(replayed.patch().is_noop());
        assert_eq!(replayed.package().to_bytes()?, package.to_bytes()?);
        Ok(())
    }

    #[test]
    fn clear_removes_the_part_and_inverse_restores_it() -> Result<()> {
        let package = Package::new()?;
        let mut edit = package.edit_metadata();
        edit.clear();
        let (cleared, patch) = edit.commit()?.into_parts();
        assert!(cleared.metadata().is_none());
        let restored = cleared.apply_metadata(&patch.inverse())?.into_package();
        assert_eq!(restored.metadata(), package.metadata());
        Ok(())
    }

    #[test]
    fn clear_then_set_default_creates_core_and_inverse_removes_it() -> Result<()> {
        let package = Package::new()?;
        let mut edit = package.edit_metadata();
        edit.clear();
        let (cleared, clear_patch) = edit.commit()?.into_parts();
        assert!(clear_patch.target().is_none());
        assert!(cleared.metadata().is_none());

        let mut edit = cleared.edit_metadata();
        edit.set(Metadata::default());
        let (created, set_patch) = edit.commit()?.into_parts();
        let default = Metadata::default();
        assert!(set_patch.base().is_none());
        assert_eq!(set_patch.target(), Some(&default));
        // The DOCX wrapper exposes core transition state through `base` and
        // `target`; `parts` is the application/custom-properties patch.
        assert!(set_patch.parts().is_noop());
        assert_eq!(created.metadata(), Some(&default));

        let reopened = Package::from_bytes(&created.to_bytes()?)?;
        assert_eq!(reopened.metadata(), Some(&default));

        let reopened_cleared = Package::from_bytes(&cleared.to_bytes()?)?;
        let replayed = reopened_cleared.apply_metadata(&set_patch)?.into_package();
        assert_eq!(replayed.metadata(), Some(&default));

        let restored = created.apply_metadata(&set_patch.inverse())?.into_package();
        assert!(restored.metadata().is_none());
        assert_eq!(restored.to_bytes()?, cleared.to_bytes()?);
        Ok(())
    }

    #[test]
    fn opaque_core_replay_and_inverse_preserve_exact_bytes() -> Result<()> {
        let package = package_with_opaque_core()?;
        let original = package.clone().to_bytes()?;
        let original_core = core_bytes(&package);
        let original_app = app_bytes(&package);
        let mut edit = package.edit_metadata();
        let app = edit.app_props()?;
        edit.set_app_props(app);
        edit.update(|props| props.title = Some("Opaque core target".to_owned()));
        let (next, patch) = edit.commit()?.into_parts();
        assert!(
            patch
                .core
                .changes()
                .iter()
                .any(|change| change.family() == families::Family::Core)
        );
        patch.parts().check(package.opc_package())?;
        assert_ne!(core_bytes(&next), original_core);
        assert_eq!(app_bytes(&next), original_app);

        let reopened_source = Package::from_bytes(&package.to_bytes()?)?;
        let replayed = reopened_source.apply_metadata(&patch)?.into_package();
        assert_eq!(core_bytes(&replayed), core_bytes(&next));
        assert_eq!(replayed.to_bytes()?, next.to_bytes()?);

        let reopened_target = Package::from_bytes(&next.to_bytes()?)?;
        let restored = reopened_target
            .apply_metadata(&patch.inverse())?
            .into_package();
        assert_eq!(core_bytes(&restored), original_core);
        assert_eq!(restored.to_bytes()?, original);
        Ok(())
    }

    #[test]
    fn absent_property_parts_combined_create_replay_and_inverse() -> Result<()> {
        let package = package_without_property_parts()?;
        assert!(package.metadata().is_none());
        assert!(package.app_props()?.is_none());
        assert!(package.custom_props().is_empty());
        let original = package.clone().to_bytes()?;

        let mut edit = package.edit_metadata();
        edit.set(Metadata::default());
        let mut app = app::Props::default();
        app.company = Some("Acme".to_owned());
        edit.set_app_props(Some(app));
        let mut custom = CustomProps::new();
        custom
            .insert("Reviewed", litchi_ooxml_common::custom::Value::Bool(true))
            .map_err(families::Error::from)?;
        edit.set_custom_props(custom);
        let (next, patch) = edit.commit()?.into_parts();
        let next_bytes = next.clone().to_bytes()?;
        assert!(
            patch
                .core
                .changes()
                .iter()
                .any(|change| { change.family() == families::Family::Core })
        );
        assert!(
            patch
                .parts()
                .changes()
                .iter()
                .all(|change| change.family() != families::Family::Core)
        );
        assert_eq!(next.metadata(), Some(&Metadata::default()));
        assert_eq!(
            next.app_props()?.and_then(|props| props.company),
            Some("Acme".to_owned())
        );
        assert_eq!(
            next.custom_props().get("Reviewed"),
            Some(&litchi_ooxml_common::custom::Value::Bool(true))
        );

        let projected = patch.parts().apply(package.opc_package())?.into_package();
        let projected = Package::from_opc_package(projected)?;
        assert!(projected.metadata().is_none());
        assert_eq!(
            projected.app_props()?.and_then(|props| props.company),
            Some("Acme".to_owned())
        );
        assert_eq!(
            projected.custom_props().get("Reviewed"),
            Some(&litchi_ooxml_common::custom::Value::Bool(true))
        );

        let reopened_source = Package::from_bytes(&original)?;
        let replayed = reopened_source.apply_metadata(&patch)?.into_package();
        assert_eq!(replayed.to_bytes()?, next_bytes);

        let reopened_target = Package::from_bytes(&next_bytes)?;
        let restored = reopened_target
            .apply_metadata(&patch.inverse())?
            .into_package();
        assert_eq!(restored.to_bytes()?, original);
        Ok(())
    }

    #[test]
    fn apply_refuses_a_conflicting_base() -> Result<()> {
        let package = Package::new()?;
        let mut first = package.edit_metadata();
        first.update(|props| props.title = Some("one".to_owned()));
        let (one, _) = first.commit()?.into_parts();
        let mut second = package.edit_metadata();
        second.update(|props| props.title = Some("two".to_owned()));
        let patch = second.commit()?.patch().clone();
        assert!(matches!(
            one.apply_metadata(&patch),
            Err(Error::PatchConflict)
        ));
        let replayed = package.apply_metadata(&patch)?;
        assert!(replayed.diagnostics().changed());
        let replayed = replayed.into_package();
        assert_eq!(
            replayed.metadata().and_then(|props| props.title.as_deref()),
            Some("two")
        );
        Ok(())
    }

    #[test]
    fn replay_reports_property_graph_refusal_as_properties() -> Result<()> {
        let source = Package::new()?;
        let mut edit = source.edit_metadata();
        let mut app = edit.app_props()?.expect("app properties");
        app.company = Some("Acme".to_owned());
        edit.set_app_props(Some(app));
        let patch = edit.commit()?.patch().clone();

        let target = package_without_app_part()?;
        assert_eq!(target.metadata(), source.metadata());
        assert!(matches!(
            target.apply_metadata(&patch),
            Err(Error::Properties(_))
        ));
        Ok(())
    }

    #[test]
    fn replay_refuses_a_changed_signed_package_with_metadata_error() -> Result<()> {
        let source = Package::new()?;
        let mut edit = source.edit_metadata();
        edit.update(|props| props.title = Some("Signed replay".to_owned()));
        let patch = edit.commit()?.patch().clone();
        let signed = signed_package(&source)?;
        assert!(matches!(signed.apply_metadata(&patch), Err(Error::Signed)));
        Ok(())
    }

    #[test]
    fn app_and_custom_props_publish_replay_and_invert() -> Result<()> {
        let package = package_with_opaque_app()?;
        let original = package.clone().to_bytes()?;
        let original_metadata = package.metadata().cloned();
        let app_before = package.app_props()?.expect("app properties");
        let mut edit = package.edit_metadata();
        edit.update(|props| props.title = Some("Quarterly".to_owned()));
        let mut app = edit.app_props()?.unwrap_or_default();
        app.company = Some("Acme".to_owned());
        app.manager = Some("Dana".to_owned());
        edit.set_app_props(Some(app));
        let mut custom = edit.custom_props().clone();
        custom
            .insert("Reviewed", litchi_ooxml_common::custom::Value::Bool(true))
            .map_err(families::Error::from)?;
        edit.set_custom_props(custom);
        let commit = edit.commit()?;
        assert!(commit.diagnostics().changed());
        assert!(!commit.patch().is_noop());
        let (next, patch) = commit.into_parts();
        assert!(
            patch
                .core
                .changes()
                .iter()
                .any(|change| change.family() == families::Family::Core)
        );
        let app = next.app_props()?.unwrap_or_default();
        assert_eq!(app.company.as_deref(), Some("Acme"));
        assert_eq!(app.manager.as_deref(), Some("Dana"));
        assert_eq!(app.application, app_before.application);
        assert_eq!(app.template, app_before.template);
        assert!(
            app_bytes(&next)
                .windows(b"<Opaque keep=\"yes\"/>".len())
                .any(|window| { window == b"<Opaque keep=\"yes\"/>" })
        );
        assert_eq!(
            next.custom_props().get("Reviewed"),
            Some(&litchi_ooxml_common::custom::Value::Bool(true))
        );
        // All three property families were published, and the source is unchanged.
        assert_eq!(
            next.metadata().and_then(|props| props.title.as_deref()),
            Some("Quarterly")
        );
        // `Package::new` owns an explicit, empty core-properties part.  The
        // edit must preserve that present-vs-absent distinction on the source
        // snapshot while changing the staged value in the new package.
        assert_eq!(package.metadata(), original_metadata.as_ref());
        assert_eq!(package.clone().to_bytes()?, original);
        let reopened = Package::from_bytes(&next.clone().to_bytes()?)?;
        assert_eq!(
            reopened.app_props()?.and_then(|props| props.company),
            Some("Acme".to_owned())
        );
        let replayed = package.apply_metadata(&patch)?.into_package();
        assert_eq!(replayed.app_props()?, next.app_props()?);
        assert_eq!(replayed.metadata(), next.metadata());
        let restored = next.apply_metadata(&patch.inverse())?.into_package();
        assert_eq!(restored.to_bytes()?, original);
        Ok(())
    }

    #[test]
    fn unchanged_app_props_are_an_exact_noop() -> Result<()> {
        let package = Package::new()?;
        let mut edit = package.edit_metadata();
        let app = edit.app_props()?;
        edit.set_app_props(app);
        edit.set_custom_props(package.custom_props().clone());
        let commit = edit.commit()?;
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        Ok(())
    }

    #[test]
    fn oversized_text_is_a_typed_limit() -> Result<()> {
        let package = Package::new()?;
        let mut edit = package.edit_metadata();
        edit.update(|props| props.title = Some("x".repeat(MAX_PROPERTY_TEXT_BYTES + 1)));
        assert!(matches!(
            edit.commit(),
            Err(Error::Limit {
                kind: LimitKind::PropertyTextBytes,
                ..
            })
        ));
        Ok(())
    }
}
