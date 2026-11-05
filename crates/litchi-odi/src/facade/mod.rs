//! Concise family entry points: the immutable [`Package`] owner, its
//! read-only [`Image`] projection, and the primary transaction family.

use litchi_core::{Error as CoreError, Metadata, Result};
use litchi_odf_common::{
    SemanticLimitKind, compact_xml,
    core::{MetaXmlPatch, metadata::Metadata as OdfMetadata, patch_meta_xml},
};
use std::path::Path;

pub use crate::authoring::Builder;
pub use litchi_core::edit::Diagnostics;
use litchi_odf_common::{Error as PackageError, OpenOptions};

type PackageResult<T> = std::result::Result<T, PackageError>;

/// A committed image package: the published package, patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// The immutable, cheap-to-clone ODI owner snapshot.
///
/// It owns every state verb: opening, the primary [`Edit`] transaction,
/// patch application and publication. Reads go through the [`Image`]
/// projection returned by [`Package::image`].
#[derive(Clone)]
pub struct Package {
    image: Image,
}

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("bytes", &self.as_bytes().len())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// The retained physical package (metadata transactions, record 0288).
    pub(crate) fn owned_package(&self) -> &litchi_odf_common::OwnedPackage {
        self.image.package.owned_package()
    }

    /// The minimal valid image package: one frame holding an inline 1x1 PNG.
    /// Deterministic: no clock or randomness is consulted.
    ///
    /// # Errors
    ///
    /// Returns an error if the builder output fails validation.
    pub fn new() -> std::result::Result<Self, PackageError> {
        Builder::new().build()
    }

    /// Opens an image package from a file path under the default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, PackageError> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Opens an image package from a file path under caller-selected limits.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Limit`] when a selected limit is exceeded,
    /// otherwise as for [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, PackageError> {
        Self::adopt(
            crate::package::Snapshot::open_with_options(path, options)?,
            options,
        )
    }

    /// Opens a password-encrypted image package from a file for inert
    /// inspection. Protected snapshots remain rewrite-refused.
    ///
    /// # Errors
    ///
    /// As for [`Self::open`].
    pub fn open_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> std::result::Result<Self, PackageError> {
        Self::open_with_password_and_options(path, password, OpenOptions::new())
    }

    /// Opens a password-encrypted image package from a file under
    /// caller-selected limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open_with_options`].
    pub fn open_with_password_and_options(
        path: impl AsRef<Path>,
        password: impl Into<String>,
        options: OpenOptions,
    ) -> std::result::Result<Self, PackageError> {
        Self::adopt(
            crate::package::Snapshot::open_with_password_and_options(path, password, options)?,
            options,
        )
    }

    /// Opens an image package from borrowed bytes (copied once).
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, PackageError> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Opens an image package from borrowed bytes under caller-selected
    /// limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open_with_options`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, PackageError> {
        Self::adopt(
            crate::package::Snapshot::from_bytes_with_options(bytes.to_vec(), options)?,
            options,
        )
    }

    /// Opens password-encrypted package bytes for inert inspection.
    ///
    /// # Errors
    ///
    /// As for [`Self::from_bytes`].
    pub fn from_bytes_with_password(
        bytes: &[u8],
        password: impl Into<String>,
    ) -> std::result::Result<Self, PackageError> {
        Self::from_bytes_with_password_and_options(bytes, password, OpenOptions::new())
    }

    /// Opens password-encrypted package bytes under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open_with_options`].
    pub fn from_bytes_with_password_and_options(
        bytes: &[u8],
        password: impl Into<String>,
        options: OpenOptions,
    ) -> std::result::Result<Self, PackageError> {
        Self::adopt(
            crate::package::Snapshot::from_bytes_with_password_and_options(
                bytes.to_vec(),
                password,
                options,
            )?,
            options,
        )
    }

    /// Reads an image package from a stream under the default limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open`].
    pub fn from_reader(reader: impl std::io::Read) -> std::result::Result<Self, PackageError> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Reads an image package from a stream under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open_with_options`].
    pub fn from_reader_with_options(
        reader: impl std::io::Read,
        options: OpenOptions,
    ) -> std::result::Result<Self, PackageError> {
        Self::adopt(
            crate::package::Snapshot::from_reader_with_options(reader, options)?,
            options,
        )
    }

    /// Adopts the packaged ODF result smart detection prepared, reusing its
    /// archive index. Source limits were applied when it was prepared.
    ///
    /// # Errors
    ///
    /// Returns an error when the prepared package is not an ODI package or
    /// fails validation.
    pub fn from_prepared_package(
        prepared: litchi_odf_common::PreparedPackage,
    ) -> std::result::Result<Self, PackageError> {
        Self::from_prepared_package_with_options(prepared, OpenOptions::new())
    }

    /// Adopt a prepared package under the semantic limits in `options`.
    /// The source limits were applied when the archive was prepared; its
    /// existing index and source allocation are reused.
    ///
    /// # Errors
    ///
    /// As [`Self::from_prepared_package`], plus a typed semantic limit failure.
    pub fn from_prepared_package_with_options(
        prepared: litchi_odf_common::PreparedPackage,
        options: OpenOptions,
    ) -> std::result::Result<Self, PackageError> {
        if prepared.format() != litchi_core::detection::FileFormat::Odi {
            return Err(PackageError::from(CoreError::UnexpectedFormat {
                detected: prepared.format(),
            }));
        }
        Self::adopt(
            crate::package::Snapshot::from_owned_package_with_options(
                prepared.into_package(),
                options,
            )?,
            options,
        )
    }

    fn adopt(
        package: crate::package::Snapshot,
        options: OpenOptions,
    ) -> std::result::Result<Self, PackageError> {
        options.semantic_limits_value().check(
            SemanticLimitKind::Units,
            u64::try_from(package.frames().len()).unwrap_or(u64::MAX),
        )?;
        Ok(Self::from_snapshot(package))
    }

    pub(crate) const fn from_snapshot(package: crate::package::Snapshot) -> Self {
        Self {
            image: Image { package },
        }
    }

    /// Reparses package bytes produced inside this crate.
    pub(crate) fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(Self::from_snapshot)
    }

    /// Returns the read-only image projection (a cheap shared clone).
    #[must_use]
    pub fn image(&self) -> Image {
        self.image.clone()
    }

    /// Borrows the read-only projection.
    pub(crate) const fn image_ref(&self) -> &Image {
        &self.image
    }

    /// Everything a tolerant read repaired while this package was opened.
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        self.image.package.tolerance_report()
    }

    /// Returns the exact package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.image.package.as_bytes()
    }

    /// Returns a copy of the exact package bytes.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the publication surface uniform.
    pub fn to_bytes(&self) -> std::result::Result<Vec<u8>, PackageError> {
        Ok(self.as_bytes().to_vec())
    }

    /// Writes the exact package bytes into `sink`.
    ///
    /// # Errors
    ///
    /// Returns an error when the sink fails.
    pub fn write_to(&self, mut sink: impl std::io::Write) -> std::result::Result<(), PackageError> {
        sink.write_all(self.as_bytes())
            .map_err(PackageError::from)?;
        Ok(())
    }

    /// Atomically saves the package to `path` (sibling temporary file and
    /// rename, with file and directory `fsync`).
    ///
    /// # Errors
    ///
    /// Returns an error when staging or replacement fails; the destination
    /// is then left untouched.
    pub fn save(&self, path: impl AsRef<Path>) -> std::result::Result<(), PackageError> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically saves the package to `path` with an explicit durability.
    ///
    /// # Errors
    ///
    /// As for [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> std::result::Result<(), PackageError> {
        litchi_odf_common::publish::save_bytes(path.as_ref(), self.as_bytes(), durability)
    }

    /// Starts a source-bound package image transaction.
    ///
    /// This supports the same lossless existing-name and existing-source edits
    /// as [`crate::flat::Image`], while rebuilding a validated ODI package and
    /// preserving all untouched member payloads.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            transaction: self.image.package.content_snapshot().edit(),
            resource_changes: Vec::new(),
            metadata: None,
            styles: None,
        }
    }

    /// Applies a patch to its exact source package and returns the target.
    ///
    /// # Errors
    ///
    /// Returns an error when this package differs byte-for-byte from the
    /// patch's source.
    pub fn apply(&self, patch: &Patch) -> std::result::Result<Commit, PackageError> {
        if !patch.is_applicable_to(self) {
            return Err(PackageError::patch_conflict(
                "odi.apply",
                "the package is not the patch source",
            ));
        }
        Ok(Commit::new(
            patch.apply(self)?,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}

/// The read-only semantic image projection, obtained from
/// [`Package::image`]. Cheap to clone; it has no state verbs.
#[derive(Clone)]
pub struct Image {
    package: crate::package::Snapshot,
}

impl Image {
    /// Returns the `content.xml` document.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Returns the `styles.xml` document, if present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Returns the document metadata, if present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    /// Returns the exact UTF-8 `meta.xml` part, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the retained package member cannot be read.
    pub fn meta_xml(&self) -> PackageResult<Option<String>> {
        self.package.meta_xml().map_err(PackageError::from)
    }

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> PackageResult<Vec<String>> {
        self.package.files().map_err(PackageError::from)
    }

    /// Returns the inert semantic frame inventory from `content.xml`.
    ///
    /// Links and embedded bytes are reported only. They are never fetched,
    /// executed, or otherwise activated.
    #[must_use]
    pub fn frames(&self) -> &[crate::frame::Frame] {
        self.package.frames()
    }

    /// Returns the package's single normative image frame.
    #[must_use]
    pub fn frame(&self) -> Option<&crate::frame::Frame> {
        self.package.frames().first()
    }

    /// Returns package-local image resources referenced by `content.xml`.
    ///
    /// External links and inline `office:binary-data` are not package
    /// resources. Missing safe package references remain visible with
    /// [`crate::resource::Resource::is_present`] set to `false`.
    #[must_use]
    pub fn resources(&self) -> &[crate::resource::Resource] {
        self.package.resources()
    }

    /// Returns the package resource graph, including unreferenced inert files
    /// and safely resolved missing image targets.
    #[must_use]
    pub fn resource_graph(&self) -> &crate::resource::Graph {
        self.package.resource_graph()
    }

    /// Returns inert XML active-content constructs and script-container members.
    ///
    /// Values are inventoried only; no script, event, macro, DDE target, or
    /// package member is loaded or activated.
    #[must_use]
    pub fn active_content(&self) -> &[crate::active::ActiveContent] {
        self.package.active_content()
    }

    /// Returns inert ODF form and `XForms` occurrences from content and styles.
    #[must_use]
    pub fn forms(&self) -> &[crate::surface::Surface] {
        self.package.forms()
    }

    /// Returns inert producer-extension elements and attributes from content and styles.
    #[must_use]
    pub fn extensions(&self) -> &[crate::surface::Surface] {
        self.package.extensions()
    }

    /// Returns exact signature-member and encrypted-manifest-entry inventory.
    #[must_use]
    pub fn protection(&self) -> &crate::ProtectionInventory {
        self.package.protection()
    }

    /// Reads inert document and macro signature metadata without executing content.
    pub fn digital_signatures(
        &self,
    ) -> PackageResult<litchi_odf_common::signature::DigitalSignatures> {
        self.package
            .digital_signatures()
            .map_err(PackageError::from)
    }

    /// Verifies document-signature math without making a certificate trust decision.
    pub fn verify_document_signatures(
        &self,
    ) -> PackageResult<Vec<litchi_odf_common::signature::SignatureVerification>> {
        self.package
            .verify_document_signatures()
            .map_err(PackageError::from)
    }

    /// Classifies one style name/family dependency across content and styles.
    pub fn style_dependency_state(
        &self,
        name: &str,
        family: &str,
    ) -> PackageResult<crate::StyleDependencyState> {
        crate::semantic::inspect_style_dependency(
            self.content_xml(),
            self.styles_xml(),
            name,
            family,
        )
        .map_err(PackageError::from)
    }

    /// Reads one inventoried package-local image resource.
    ///
    /// # Errors
    ///
    /// Returns an error when `index` is out of bounds or the archive member
    /// cannot be read. A referenced but absent member returns `Ok(None)`.
    pub fn resource_bytes(&self, index: usize) -> PackageResult<Option<Vec<u8>>> {
        self.package
            .resource_bytes(index)
            .map_err(PackageError::from)
    }

    /// Reads one safely named auxiliary package member without activating it.
    pub fn member_bytes(&self, path: &str) -> PackageResult<Option<Vec<u8>>> {
        validate_resource_path(path)?;
        self.package.resource_file(path).map_err(PackageError::from)
    }

    /// Audits whether this exact package can enter changed-byte publication.
    ///
    /// The audit is inert: it does not verify signatures, decrypt content, or
    /// mutate bytes. Every independently detected blocker is returned.
    pub fn rewrite_capability(&self) -> PackageResult<crate::RewriteCapability> {
        self.package
            .rewrite_capability()
            .map_err(PackageError::from)
    }
}

/// A source-bound package image and resource transaction.
pub struct Edit<'a> {
    source: &'a Package,
    transaction: crate::flat::Edit,
    resource_changes: Vec<ResourceEdit>,
    metadata: Option<MetadataEdit>,
    styles: Option<StyleEdit>,
}

impl Edit<'_> {
    /// Stages a replacement for the document frame's optional name.
    ///
    /// # Errors
    ///
    /// Returns an error if the lossless frame site is unavailable.
    pub fn set_name(&mut self, name: Option<String>) -> PackageResult<()> {
        self.transaction.set_name(name).map_err(PackageError::from)
    }

    /// Stages a replacement for the document's linked or inline source.
    ///
    /// # Errors
    ///
    /// Returns an error when changing source representation would be lossy.
    pub fn set_image_source(&mut self, source: crate::source::Source) -> PackageResult<()> {
        self.transaction
            .set_image_source(source)
            .map_err(PackageError::from)
    }

    /// Stages a replacement for one frame's optional `draw:name`.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or an image without a
    /// losslessly editable `draw:frame` owner.
    pub fn set_frame_name(&mut self, frame: usize, name: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_frame_name(frame, name)
            .map_err(PackageError::from)
    }

    /// Stages replacement of an existing linked URI or inline image payload.
    ///
    /// Cross-kind changes are refused rather than reconstructing unknown XML.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or lossy source representation.
    pub fn set_source(&mut self, frame: usize, source: crate::source::Source) -> PackageResult<()> {
        self.transaction
            .set_source(frame, source)
            .map_err(PackageError::from)
    }

    /// Stages the optional stable XML identifier on `draw:frame`.
    pub fn set_frame_xml_id(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_frame_xml_id(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional accessible frame title.
    pub fn set_frame_title(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_frame_title(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional accessible frame description.
    pub fn set_frame_description(
        &mut self,
        frame: usize,
        value: Option<String>,
    ) -> PackageResult<()> {
        self.transaction
            .set_frame_description(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional declared image media type.
    pub fn set_image_media_type(
        &mut self,
        frame: usize,
        value: Option<String>,
    ) -> PackageResult<()> {
        self.transaction
            .set_image_media_type(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional stable XML identifier on `draw:image`.
    pub fn set_image_xml_id(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_image_xml_id(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional inert producer filter hint.
    pub fn set_filter_name(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_filter_name(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional image `XLink` type.
    pub fn set_link_type(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_link_type(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional image `XLink` presentation behavior.
    pub fn set_show(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_show(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional image `XLink` activation behavior.
    pub fn set_actuate(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_actuate(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional graphic style reference.
    pub fn set_style_name(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_style_name(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional paragraph style reference used by frame text.
    pub fn set_text_style_name(
        &mut self,
        frame: usize,
        value: Option<String>,
    ) -> PackageResult<()> {
        self.transaction
            .set_text_style_name(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional drawing layer name.
    pub fn set_layer(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_layer(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional non-negative stacking order.
    pub fn set_z_index(&mut self, frame: usize, value: Option<u32>) -> PackageResult<()> {
        self.transaction
            .set_z_index(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the lexical drawing transform.
    pub fn set_transform(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_transform(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the optional text anchoring mode.
    pub fn set_anchor_type(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_anchor_type(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages lexical frame position and size values.
    pub fn set_geometry(
        &mut self,
        frame: usize,
        x: Option<String>,
        y: Option<String>,
        width: Option<String>,
        height: Option<String>,
    ) -> PackageResult<()> {
        self.transaction
            .set_geometry(frame, x, y, width, height)
            .map_err(PackageError::from)
    }

    /// Stages lexical relative frame width and height values.
    pub fn set_relative_size(
        &mut self,
        frame: usize,
        width: Option<String>,
        height: Option<String>,
    ) -> PackageResult<()> {
        self.transaction
            .set_relative_size(frame, width, height)
            .map_err(PackageError::from)
    }

    /// Replaces, inserts, or removes the document frame's typed image map.
    pub fn set_image_map(
        &mut self,
        frame: usize,
        value: Option<crate::map::ImageMap>,
    ) -> PackageResult<()> {
        self.transaction
            .set_image_map(frame, value)
            .map_err(PackageError::from)
    }

    /// Replaces the lexical horizontal position.
    pub fn set_x(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_x(frame, value)
            .map_err(PackageError::from)
    }

    /// Replaces the lexical vertical position.
    pub fn set_y(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_y(frame, value)
            .map_err(PackageError::from)
    }

    /// Replaces the lexical width.
    pub fn set_width(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_width(frame, value)
            .map_err(PackageError::from)
    }

    /// Replaces the lexical height.
    pub fn set_height(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_height(frame, value)
            .map_err(PackageError::from)
    }

    /// Replaces the lexical relative width.
    pub fn set_relative_width(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_relative_width(frame, value)
            .map_err(PackageError::from)
    }

    /// Replaces the lexical relative height.
    pub fn set_relative_height(
        &mut self,
        frame: usize,
        value: Option<String>,
    ) -> PackageResult<()> {
        self.transaction
            .set_relative_height(frame, value)
            .map_err(PackageError::from)
    }

    /// Replaces the optional `draw:copy-of` frame reference.
    pub fn set_copy_of(&mut self, frame: usize, value: Option<String>) -> PackageResult<()> {
        self.transaction
            .set_copy_of(frame, value)
            .map_err(PackageError::from)
    }

    /// Stages the document title while preserving every unedited metadata node.
    pub fn set_title(&mut self, value: Option<String>) -> PackageResult<()> {
        self.metadata_mut()?.after.title = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the document author while preserving every unedited metadata node.
    pub fn set_author(&mut self, value: Option<String>) -> PackageResult<()> {
        self.metadata_mut()?.after.author = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the document subject while preserving every unedited metadata node.
    pub fn set_subject(&mut self, value: Option<String>) -> PackageResult<()> {
        self.metadata_mut()?.after.subject = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the document description while preserving every unedited metadata node.
    pub fn set_description(&mut self, value: Option<String>) -> PackageResult<()> {
        self.metadata_mut()?.after.description = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the keyword list, one `meta:keyword` each; an empty list
    /// removes them.
    pub fn set_keywords(&mut self, keywords: Vec<String>) -> PackageResult<()> {
        self.metadata_mut()?.after.keywords = keywords;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Replaces or removes the exact compact `styles.xml` part.
    ///
    /// The supplied XML is treated as newly authored and must pass the shared
    /// compact publication boundary before it is staged.
    pub fn set_styles_xml(&mut self, value: Option<String>) -> PackageResult<()> {
        if let Some(xml) = &value {
            compact_xml::validate(xml.as_bytes(), compact_xml::Limits::default())
                .map_err(CoreError::from)?;
        }
        let before = self.source.image.styles_xml().map(str::to_owned);
        if before == value {
            self.styles = None;
        } else {
            self.styles = Some(StyleEdit {
                before,
                after: value,
            });
        }
        Ok(())
    }

    /// Adds or replaces one referenced package-local image resource.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid resource selector, media type, or
    /// resource that is not package-local.
    pub fn set_resource(
        &mut self,
        resource: usize,
        media_type: String,
        bytes: Vec<u8>,
    ) -> PackageResult<()> {
        validate_media_type(&media_type).map_err(PackageError::from)?;
        self.stage_resource(resource, Some(media_type), Some(bytes))
            .map_err(PackageError::from)
    }

    /// Removes one existing package-local image member while preserving its
    /// inert reference as a typed missing resource.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector.
    pub fn remove_resource(&mut self, resource: usize) -> PackageResult<()> {
        self.stage_resource(resource, None, None)
            .map_err(PackageError::from)
    }

    /// Adds or replaces an auxiliary package member by safe path.
    pub fn put_member(
        &mut self,
        path: String,
        media_type: String,
        bytes: Vec<u8>,
    ) -> PackageResult<()> {
        validate_resource_path(&path).map_err(PackageError::from)?;
        validate_media_type(&media_type).map_err(PackageError::from)?;
        self.stage_member(path, Some(media_type), Some(bytes))
            .map_err(PackageError::from)
    }

    /// Removes an auxiliary package member by safe path.
    pub fn remove_member(&mut self, path: String) -> PackageResult<()> {
        validate_resource_path(&path).map_err(PackageError::from)?;
        self.stage_member(path, None, None)
            .map_err(PackageError::from)
    }

    /// Atomically validates, rebuilds, and publishes the package edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be safely rebuilt, including
    /// signed or non-compact source XML, or if semantic readback fails.
    pub fn commit(self) -> std::result::Result<Commit, PackageError> {
        Ok(self.commit_inner()?)
    }

    fn commit_inner(self) -> Result<Commit> {
        let content = self.transaction.commit_core()?;
        let changes = content.patch().changes().to_vec();
        let inverse_changes = content.patch().inverse().changes().to_vec();
        let mut resource_edits = self.resource_changes;
        resource_edits.sort_unstable_by(|left, right| left.path.cmp(&right.path));
        let resource_changes = resource_edits
            .iter()
            .map(ResourceEdit::change)
            .collect::<Vec<_>>();
        let inverse_resource_changes = resource_changes
            .iter()
            .map(ResourceChange::inverse)
            .collect::<Vec<_>>();
        let style_change = self.styles.map(|edit| StyleChange {
            before: edit.before,
            after: edit.after,
        });
        let metadata_change = self.metadata.map(|edit| MetadataChange {
            before: edit.before,
            after: edit.after,
        });
        let replacement_meta_xml = metadata_change
            .as_ref()
            .map(|change| patch_metadata(&self.source.image, &change.after))
            .transpose()?;
        let snapshot = if changes.is_empty()
            && resource_edits.is_empty()
            && metadata_change.is_none()
            && style_change.is_none()
        {
            self.source.clone()
        } else {
            let xml = std::str::from_utf8(content.package().as_bytes()).map_err(|error| {
                CoreError::InvalidFormat(format!("ODI edited content.xml is not UTF-8: {error}"))
            })?;
            let replacements = resource_edits
                .iter()
                .map(|change| crate::package::ResourceReplacement {
                    path: &change.path,
                    media_type: change.after_media_type.as_deref().unwrap_or_default(),
                    bytes: change.after_bytes.as_deref(),
                })
                .collect::<Vec<_>>();
            let styles_replacement = match &style_change {
                None => crate::package::StylesReplacement::Preserve,
                Some(change) => match change.after.as_deref() {
                    Some(styles_xml) => crate::package::StylesReplacement::Set(styles_xml),
                    None => crate::package::StylesReplacement::Remove,
                },
            };
            Package::from_snapshot(self.source.image.package.rebuild(
                xml,
                styles_replacement,
                replacement_meta_xml.as_deref(),
                &replacements,
            )?)
        };
        for change in &changes {
            let actual = snapshot.image.frames().get(change.frame()).ok_or_else(|| {
                CoreError::InvalidFormat("ODI edited frame disappeared during readback".to_string())
            })?;
            if actual != change.after() {
                return Err(CoreError::InvalidFormat(
                    "ODI package edit failed semantic readback".to_string(),
                ));
            }
        }
        for change in &resource_edits {
            if snapshot
                .image
                .package
                .resource_file(&change.path)?
                .as_deref()
                != change.after_bytes.as_deref()
            {
                return Err(CoreError::InvalidFormat(
                    "ODI package resource edit failed byte readback".to_string(),
                ));
            }
            if let Some(after_media_type) = change.after_media_type.as_deref()
                && snapshot
                    .image
                    .resource_graph()
                    .nodes()
                    .iter()
                    .find(|node| node.path() == change.path)
                    .and_then(crate::resource::Node::media_type)
                    != Some(after_media_type)
            {
                return Err(CoreError::InvalidFormat(
                    "ODI package resource edit failed media-type readback".to_string(),
                ));
            }
        }
        if let Some(change) = &metadata_change
            && MetadataFields::from(snapshot.image.metadata()) != change.after
        {
            return Err(CoreError::InvalidFormat(
                "ODI package metadata edit failed semantic readback".to_string(),
            ));
        }
        if let Some(change) = &style_change
            && snapshot.image.styles_xml() != change.after.as_deref()
        {
            return Err(CoreError::InvalidFormat(
                "ODI styles.xml edit failed semantic readback".to_string(),
            ));
        }
        let changed = !changes.is_empty()
            || !resource_changes.is_empty()
            || style_change.is_some()
            || metadata_change.is_some();
        Ok(Commit::new(
            snapshot.clone(),
            Patch {
                source: self.source.clone(),
                target: snapshot,
                changes,
                inverse_changes,
                resource_changes,
                inverse_resource_changes,
                style_change: style_change.clone(),
                inverse_style_change: style_change.as_ref().map(StyleChange::inverse),
                metadata_change: metadata_change.clone(),
                inverse_metadata_change: metadata_change.as_ref().map(MetadataChange::inverse),
            },
            Diagnostics::new(changed),
        ))
    }

    fn stage_resource(
        &mut self,
        resource: usize,
        after_media_type: Option<String>,
        after_bytes: Option<Vec<u8>>,
    ) -> Result<()> {
        let selected = self.source.image.resources().get(resource).ok_or_else(|| {
            CoreError::InvalidFormat("ODI resource selector is out of bounds".to_string())
        })?;
        let before_bytes = self.source.image.resource_bytes(resource)?;
        let before_media_type = selected.media_type().map(str::to_owned);
        if let Some(change) = self
            .resource_changes
            .iter_mut()
            .find(|change| change.resource == Some(resource))
        {
            change.after_media_type = after_media_type;
            change.after_bytes = after_bytes;
        } else {
            self.resource_changes.push(ResourceEdit {
                resource: Some(resource),
                path: selected.path().to_string(),
                before_media_type: before_media_type.clone(),
                after_media_type,
                before_size: before_bytes.as_ref().map(Vec::len),
                before_bytes,
                after_bytes,
            });
        }
        self.resource_changes.retain(|change| {
            change.resource != Some(resource)
                || change.before_media_type != change.after_media_type
                || change.before_bytes.as_deref() != change.after_bytes.as_deref()
        });
        Ok(())
    }

    fn stage_member(
        &mut self,
        path: String,
        after_media_type: Option<String>,
        after_bytes: Option<Vec<u8>>,
    ) -> Result<()> {
        let before_bytes = self.source.image.package.resource_file(&path)?;
        let before_media_type = self
            .source
            .image
            .resource_graph()
            .nodes()
            .iter()
            .find(|node| node.path() == path)
            .and_then(crate::resource::Node::media_type)
            .map(str::to_owned);
        let resource = self
            .source
            .image
            .resources()
            .iter()
            .position(|candidate| candidate.path() == path);
        if let Some(change) = self
            .resource_changes
            .iter_mut()
            .find(|change| change.path == path)
        {
            change.after_media_type = after_media_type;
            change.after_bytes = after_bytes;
        } else {
            self.resource_changes.push(ResourceEdit {
                resource,
                path,
                before_media_type,
                after_media_type,
                before_size: before_bytes.as_ref().map(Vec::len),
                before_bytes,
                after_bytes,
            });
        }
        self.resource_changes.retain(|change| {
            change.before_media_type != change.after_media_type
                || change.before_bytes.as_deref() != change.after_bytes.as_deref()
        });
        Ok(())
    }

    fn metadata_mut(&mut self) -> Result<&mut MetadataEdit> {
        if self.source.image.metadata().is_none() {
            return Err(CoreError::InvalidFormat(
                "ODI metadata editing requires an existing meta.xml part".to_string(),
            ));
        }
        if self.metadata.is_none() {
            let before = MetadataFields::from(self.source.image.metadata());
            self.metadata = Some(MetadataEdit {
                before: before.clone(),
                after: before,
            });
        }
        self.metadata.as_mut().ok_or_else(|| {
            CoreError::InvalidFormat("ODI metadata edit state disappeared".to_string())
        })
    }

    fn remove_metadata_noop(&mut self) {
        if self
            .metadata
            .as_ref()
            .is_some_and(|edit| edit.before == edit.after)
        {
            self.metadata = None;
        }
    }
}

struct MetadataEdit {
    before: MetadataFields,
    after: MetadataFields,
}

struct StyleEdit {
    before: Option<String>,
    after: Option<String>,
}

struct ResourceEdit {
    resource: Option<usize>,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
    before_bytes: Option<Vec<u8>>,
    after_bytes: Option<Vec<u8>>,
}

impl ResourceEdit {
    fn change(&self) -> ResourceChange {
        ResourceChange {
            resource: self.resource,
            path: self.path.clone(),
            before_media_type: self.before_media_type.clone(),
            after_media_type: self.after_media_type.clone(),
            before_size: self.before_size,
            after_size: self.after_bytes.as_ref().map(Vec::len),
        }
    }
}

/// A source-checked reversible ODI package image/resource patch.
#[derive(Clone)]
pub struct Patch {
    source: Package,
    target: Package,
    changes: Vec<crate::FrameChange>,
    inverse_changes: Vec<crate::FrameChange>,
    resource_changes: Vec<ResourceChange>,
    inverse_resource_changes: Vec<ResourceChange>,
    style_change: Option<StyleChange>,
    inverse_style_change: Option<StyleChange>,
    metadata_change: Option<MetadataChange>,
    inverse_metadata_change: Option<MetadataChange>,
}

impl Patch {
    /// Returns the exact source package.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.source
    }

    /// Returns the exact target package.
    #[must_use]
    pub const fn target(&self) -> &Package {
        &self.target
    }

    /// Returns whether the patch applies to this exact source byte sequence.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Package) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Whether the patch changes nothing (its target is its source).
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.changes.is_empty()
            && self.resource_changes.is_empty()
            && self.style_change.is_none()
            && self.metadata_change.is_none()
    }

    /// Applies this patch only to its exact immutable source
    /// (public route: [`Package::apply`]).
    pub(crate) fn apply(&self, source: &Package) -> Result<Package> {
        if !self.is_applicable_to(source) {
            return Err(CoreError::InvalidFormat(
                "ODI package patch source does not match its expected snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Package::from_vec(self.target.as_bytes().to_vec())
    }

    /// Returns the semantic changes in source order.
    #[must_use]
    pub fn changes(&self) -> &[crate::FrameChange] {
        &self.changes
    }

    /// Returns package-resource changes in source order.
    #[must_use]
    pub fn resource_changes(&self) -> &[ResourceChange] {
        &self.resource_changes
    }

    /// Returns the whole compact style-part change, if any.
    #[must_use]
    pub const fn style_change(&self) -> Option<&StyleChange> {
        self.style_change.as_ref()
    }

    /// Returns the simple metadata change, if any.
    #[must_use]
    pub const fn metadata_change(&self) -> Option<&MetadataChange> {
        self.metadata_change.as_ref()
    }

    /// Returns the patch that restores the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.inverse_changes.clone(),
            inverse_changes: self.changes.clone(),
            resource_changes: self.inverse_resource_changes.clone(),
            inverse_resource_changes: self.resource_changes.clone(),
            style_change: self.inverse_style_change.clone(),
            inverse_style_change: self.style_change.clone(),
            metadata_change: self.inverse_metadata_change.clone(),
            inverse_metadata_change: self.metadata_change.clone(),
        }
    }
}

/// The simple ODF metadata values editable without normalizing opaque nodes.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct MetadataFields {
    title: Option<String>,
    author: Option<String>,
    subject: Option<String>,
    description: Option<String>,
    keywords: Vec<String>,
}

impl MetadataFields {
    /// Returns the title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }
    /// Returns the author.
    #[must_use]
    pub fn author(&self) -> Option<&str> {
        self.author.as_deref()
    }
    /// Returns the subject.
    #[must_use]
    pub fn subject(&self) -> Option<&str> {
        self.subject.as_deref()
    }
    /// Returns the description.
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }
    /// Returns the keywords, one `meta:keyword` each.
    #[must_use]
    pub fn keywords(&self) -> &[String] {
        &self.keywords
    }

    pub(crate) fn set_title(&mut self, value: Option<String>) {
        self.title = value;
    }

    pub(crate) fn set_author(&mut self, value: Option<String>) {
        self.author = value;
    }

    pub(crate) fn set_subject(&mut self, value: Option<String>) {
        self.subject = value;
    }

    pub(crate) fn set_description(&mut self, value: Option<String>) {
        self.description = value;
    }

    pub(crate) fn set_keywords(&mut self, keywords: Vec<String>) {
        self.keywords = keywords;
    }
}

impl From<Option<&Metadata>> for MetadataFields {
    fn from(metadata: Option<&Metadata>) -> Self {
        metadata.map_or_else(Self::default, |value| Self {
            title: value.title.clone(),
            author: value.author.clone(),
            subject: value.subject.clone(),
            description: value.description.clone(),
            keywords: value.keywords.clone(),
        })
    }
}

/// One reversible whole `styles.xml` part change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StyleChange {
    before: Option<String>,
    after: Option<String>,
}

impl StyleChange {
    /// Returns the exact source style XML, if present.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Returns the exact target style XML, if present.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// One reversible metadata change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MetadataChange {
    before: MetadataFields,
    after: MetadataFields,
}

impl MetadataChange {
    pub(crate) fn new(before: MetadataFields, after: MetadataFields) -> Self {
        Self { before, after }
    }
    /// Returns source metadata.
    #[must_use]
    pub const fn before(&self) -> &MetadataFields {
        &self.before
    }
    /// Returns target metadata.
    #[must_use]
    pub const fn after(&self) -> &MetadataFields {
        &self.after
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

fn patch_metadata(source: &Image, values: &MetadataFields) -> Result<String> {
    let source_xml = source.meta_xml()?.ok_or_else(|| {
        CoreError::InvalidFormat(
            "ODI metadata editing requires an existing meta.xml part".to_string(),
        )
    })?;
    let parsed = OdfMetadata::from_xml(&source_xml, source.package.xml_limits_value())?;
    let mut target = Metadata::from(parsed.clone());
    target.title.clone_from(&values.title);
    target.author.clone_from(&values.author);
    target.subject.clone_from(&values.subject);
    target.description.clone_from(&values.description);
    target.keywords.clone_from(&values.keywords);
    let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&parsed, &target);
    patch_meta_xml(&source_xml, &patch, source.package.xml_limits_value())?.ok_or_else(|| {
        CoreError::InvalidFormat(
            "ODI metadata editing requires an office:meta container".to_string(),
        )
    })
}

/// One reversible package-local image-resource change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceChange {
    resource: Option<usize>,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
    after_size: Option<usize>,
}

impl ResourceChange {
    /// Returns the source resource selector.
    #[must_use]
    pub const fn resource(&self) -> Option<usize> {
        self.resource
    }

    /// Returns the safe package path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Returns the source manifest media type.
    #[must_use]
    pub fn before_media_type(&self) -> Option<&str> {
        self.before_media_type.as_deref()
    }

    /// Returns the target manifest media type.
    #[must_use]
    pub fn after_media_type(&self) -> Option<&str> {
        self.after_media_type.as_deref()
    }

    /// Returns the source member size, or `None` when it was missing.
    #[must_use]
    pub fn before_size(&self) -> Option<usize> {
        self.before_size
    }

    /// Returns the target member size, or `None` when it is removed.
    #[must_use]
    pub fn after_size(&self) -> Option<usize> {
        self.after_size
    }

    fn inverse(&self) -> Self {
        Self {
            resource: self.resource,
            path: self.path.clone(),
            before_media_type: self.after_media_type.clone(),
            after_media_type: self.before_media_type.clone(),
            before_size: self.after_size,
            after_size: self.before_size,
        }
    }
}

fn validate_media_type(media_type: &str) -> Result<()> {
    if media_type.is_empty()
        || media_type.len() > 1_024
        || !media_type.is_ascii()
        || media_type
            .bytes()
            .any(|byte| byte.is_ascii_whitespace() || byte.is_ascii_control())
        || !media_type.contains('/')
    {
        return Err(CoreError::InvalidFormat(
            "ODI resource media type is invalid".to_string(),
        ));
    }
    Ok(())
}

fn validate_resource_path(path: &str) -> Result<()> {
    let reserved = matches!(
        path,
        "mimetype"
            | "content.xml"
            | "styles.xml"
            | "meta.xml"
            | "settings.xml"
            | "META-INF/manifest.xml"
    );
    if path.is_empty()
        || path.len() > 4_096
        || path.starts_with('/')
        || path.contains('\\')
        || path
            .split('/')
            .any(|part| part.is_empty() || matches!(part, "." | ".."))
        || path
            .strip_prefix("META-INF/")
            .is_some_and(|name| name.contains("signatures"))
        || reserved
    {
        return Err(CoreError::InvalidFormat(
            "ODI auxiliary resource path is unsafe or reserved".to_string(),
        ));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]
mod tests {
    use super::{Builder, Package};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let built = Builder::new().build().unwrap();
        let package = Package::from_bytes(built.as_bytes()).unwrap();
        assert!(package.image().content_xml().contains("<office:image"));
        assert!(!package.as_bytes().is_empty());
    }
}
