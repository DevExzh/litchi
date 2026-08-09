//! Concise family entry points.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::{MetaXmlPatch, metadata::Metadata as OdfMetadata, patch_meta_xml};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Image {
    package: crate::package::Snapshot,
}

impl Image {
    /// Opens an image package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens an image package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(|package| Self { package })
    }

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
    pub fn meta_xml(&self) -> Result<Option<String>> {
        self.package.meta_xml()
    }

    /// Returns the raw package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
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

    /// Reads one inventoried package-local image resource.
    ///
    /// # Errors
    ///
    /// Returns an error when `index` is out of bounds or the archive member
    /// cannot be read. A referenced but absent member returns `Ok(None)`.
    pub fn resource_bytes(&self, index: usize) -> Result<Option<Vec<u8>>> {
        self.package.resource_bytes(index)
    }

    /// Starts a source-bound package image transaction.
    ///
    /// This supports the same lossless existing-name and existing-source edits
    /// as [`crate::FlatImage`], while rebuilding a validated ODI package and
    /// preserving all untouched member payloads.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            transaction: self.package.content_snapshot().transaction(),
            resource_changes: Vec::new(),
            metadata: None,
        }
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

/// A source-bound package image and resource transaction.
pub struct Edit<'a> {
    source: &'a Image,
    transaction: crate::FlatImageTransaction,
    resource_changes: Vec<ResourceEdit>,
    metadata: Option<MetadataEdit>,
}

impl Edit<'_> {
    /// Stages a replacement for the document frame's optional name.
    ///
    /// # Errors
    ///
    /// Returns an error if the lossless frame site is unavailable.
    pub fn set_name(&mut self, name: Option<String>) -> Result<()> {
        self.transaction.set_name(name)
    }

    /// Stages a replacement for the document's linked or inline source.
    ///
    /// # Errors
    ///
    /// Returns an error when changing source representation would be lossy.
    pub fn set_image_source(&mut self, source: crate::source::Source) -> Result<()> {
        self.transaction.set_image_source(source)
    }

    /// Stages a replacement for one frame's optional `draw:name`.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or an image without a
    /// losslessly editable `draw:frame` owner.
    pub fn set_frame_name(&mut self, frame: usize, name: Option<String>) -> Result<()> {
        self.transaction.set_frame_name(frame, name)
    }

    /// Stages replacement of an existing linked URI or inline image payload.
    ///
    /// Cross-kind changes are refused rather than reconstructing unknown XML.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or lossy source representation.
    pub fn set_source(&mut self, frame: usize, source: crate::source::Source) -> Result<()> {
        self.transaction.set_source(frame, source)
    }

    /// Stages the optional graphic style reference.
    pub fn set_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.transaction.set_style_name(frame, value)
    }

    /// Stages the optional paragraph style reference used by frame text.
    pub fn set_text_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.transaction.set_text_style_name(frame, value)
    }

    /// Stages the optional drawing layer name.
    pub fn set_layer(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.transaction.set_layer(frame, value)
    }

    /// Stages the optional non-negative stacking order.
    pub fn set_z_index(&mut self, frame: usize, value: Option<u32>) -> Result<()> {
        self.transaction.set_z_index(frame, value)
    }

    /// Stages the lexical drawing transform.
    pub fn set_transform(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.transaction.set_transform(frame, value)
    }

    /// Stages the optional text anchoring mode.
    pub fn set_anchor_type(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.transaction.set_anchor_type(frame, value)
    }

    /// Stages lexical frame position and size values.
    pub fn set_geometry(
        &mut self,
        frame: usize,
        x: Option<String>,
        y: Option<String>,
        width: Option<String>,
        height: Option<String>,
    ) -> Result<()> {
        self.transaction.set_geometry(frame, x, y, width, height)
    }

    /// Stages lexical relative frame width and height values.
    pub fn set_relative_size(
        &mut self,
        frame: usize,
        width: Option<String>,
        height: Option<String>,
    ) -> Result<()> {
        self.transaction.set_relative_size(frame, width, height)
    }

    /// Stages the document title while preserving every unedited metadata node.
    pub fn set_title(&mut self, value: Option<String>) -> Result<()> {
        self.metadata_mut()?.after.title = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the document author while preserving every unedited metadata node.
    pub fn set_author(&mut self, value: Option<String>) -> Result<()> {
        self.metadata_mut()?.after.author = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the document subject while preserving every unedited metadata node.
    pub fn set_subject(&mut self, value: Option<String>) -> Result<()> {
        self.metadata_mut()?.after.subject = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the document description while preserving every unedited metadata node.
    pub fn set_description(&mut self, value: Option<String>) -> Result<()> {
        self.metadata_mut()?.after.description = value;
        self.remove_metadata_noop();
        Ok(())
    }

    /// Stages the comma-separated keyword value.
    pub fn set_keywords(&mut self, value: Option<String>) -> Result<()> {
        self.metadata_mut()?.after.keywords = value;
        self.remove_metadata_noop();
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
    ) -> Result<()> {
        validate_media_type(&media_type)?;
        self.stage_resource(resource, Some(media_type), Some(bytes))
    }

    /// Removes one existing package-local image member while preserving its
    /// inert reference as a typed missing resource.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector.
    pub fn remove_resource(&mut self, resource: usize) -> Result<()> {
        self.stage_resource(resource, None, None)
    }

    /// Atomically validates, rebuilds, and publishes the package edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be safely rebuilt, including
    /// signed or non-compact source XML, or if semantic readback fails.
    pub fn commit(self) -> Result<Commit> {
        let content = self.transaction.commit()?;
        let changes = content.patch().changes().to_vec();
        let inverse_changes = content.patch().inverse().changes().to_vec();
        let mut resource_edits = self.resource_changes;
        resource_edits.sort_unstable_by_key(|change| change.resource);
        let resource_changes = resource_edits
            .iter()
            .map(ResourceEdit::change)
            .collect::<Vec<_>>();
        let inverse_resource_changes = resource_changes
            .iter()
            .map(ResourceChange::inverse)
            .collect::<Vec<_>>();
        let metadata_change = self.metadata.map(|edit| MetadataChange {
            before: edit.before,
            after: edit.after,
        });
        let replacement_meta_xml = metadata_change
            .as_ref()
            .map(|change| patch_metadata(self.source, &change.after))
            .transpose()?;
        let snapshot =
            if changes.is_empty() && resource_edits.is_empty() && metadata_change.is_none() {
                self.source.clone()
            } else {
                let xml = std::str::from_utf8(content.snapshot().as_bytes()).map_err(|error| {
                    Error::InvalidFormat(format!("ODI edited content.xml is not UTF-8: {error}"))
                })?;
                let replacements = resource_edits
                    .iter()
                    .map(|change| crate::package::ResourceReplacement {
                        path: &change.path,
                        media_type: change.after_media_type.as_deref().unwrap_or_default(),
                        bytes: change.after_bytes.as_deref(),
                    })
                    .collect::<Vec<_>>();
                Image {
                    package: self.source.package.rebuild(
                        xml,
                        replacement_meta_xml.as_deref(),
                        &replacements,
                    )?,
                }
            };
        for change in &changes {
            let actual = snapshot.frames().get(change.frame()).ok_or_else(|| {
                Error::InvalidFormat("ODI edited frame disappeared during readback".to_string())
            })?;
            if actual.name() != change.after_name() || actual.source() != change.after_source() {
                return Err(Error::InvalidFormat(
                    "ODI package edit failed semantic readback".to_string(),
                ));
            }
        }
        for change in &resource_edits {
            if snapshot.package.resource_file(&change.path)?.as_deref()
                != change.after_bytes.as_deref()
            {
                return Err(Error::InvalidFormat(
                    "ODI package resource edit failed byte readback".to_string(),
                ));
            }
            if let Some(after_media_type) = change.after_media_type.as_deref()
                && snapshot
                    .resources()
                    .iter()
                    .find(|resource| resource.path() == change.path)
                    .and_then(crate::resource::Resource::media_type)
                    != Some(after_media_type)
            {
                return Err(Error::InvalidFormat(
                    "ODI package resource edit failed media-type readback".to_string(),
                ));
            }
        }
        if let Some(change) = &metadata_change
            && MetadataFields::from(snapshot.metadata()) != change.after
        {
            return Err(Error::InvalidFormat(
                "ODI package metadata edit failed semantic readback".to_string(),
            ));
        }
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                changes,
                inverse_changes,
                resource_changes,
                inverse_resource_changes,
                metadata_change: metadata_change.clone(),
                inverse_metadata_change: metadata_change.as_ref().map(MetadataChange::inverse),
            },
        })
    }

    fn stage_resource(
        &mut self,
        resource: usize,
        after_media_type: Option<String>,
        after_bytes: Option<Vec<u8>>,
    ) -> Result<()> {
        let selected = self.source.resources().get(resource).ok_or_else(|| {
            Error::InvalidFormat("ODI resource selector is out of bounds".to_string())
        })?;
        let before_bytes = self.source.resource_bytes(resource)?;
        let before_media_type = selected.media_type().map(str::to_owned);
        if let Some(change) = self
            .resource_changes
            .iter_mut()
            .find(|change| change.resource == resource)
        {
            change.after_media_type = after_media_type;
            change.after_bytes = after_bytes;
        } else {
            self.resource_changes.push(ResourceEdit {
                resource,
                path: selected.path().to_string(),
                before_media_type: before_media_type.clone(),
                after_media_type,
                before_size: before_bytes.as_ref().map(Vec::len),
                after_bytes,
            });
        }
        self.resource_changes.retain(|change| {
            change.resource != resource
                || change.before_media_type != change.after_media_type
                || before_bytes.as_deref() != change.after_bytes.as_deref()
        });
        Ok(())
    }

    fn metadata_mut(&mut self) -> Result<&mut MetadataEdit> {
        if self.source.metadata().is_none() {
            return Err(Error::InvalidFormat(
                "ODI metadata editing requires an existing meta.xml part".to_string(),
            ));
        }
        if self.metadata.is_none() {
            let before = MetadataFields::from(self.source.metadata());
            self.metadata = Some(MetadataEdit {
                before: before.clone(),
                after: before,
            });
        }
        self.metadata
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("ODI metadata edit state disappeared".to_string()))
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

struct ResourceEdit {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
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

/// A committed immutable image package and its exact-source patch.
pub struct Commit {
    snapshot: Image,
    patch: Patch,
}

impl Commit {
    /// Returns whether the package bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.changes.is_empty()
            || !self.patch.resource_changes.is_empty()
            || self.patch.metadata_change.is_some()
    }

    /// Returns the committed image snapshot.
    #[must_use]
    pub fn image(&self) -> &Image {
        &self.snapshot
    }

    /// Returns the reversible exact-source patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes this commit into its published snapshot.
    #[must_use]
    pub fn into_image(self) -> Image {
        self.snapshot
    }
}

/// A source-checked reversible ODI package image/resource patch.
#[derive(Clone)]
pub struct Patch {
    source: Image,
    target: Image,
    changes: Vec<crate::FrameChange>,
    inverse_changes: Vec<crate::FrameChange>,
    resource_changes: Vec<ResourceChange>,
    inverse_resource_changes: Vec<ResourceChange>,
    metadata_change: Option<MetadataChange>,
    inverse_metadata_change: Option<MetadataChange>,
}

impl Patch {
    /// Returns whether the patch applies to this exact source byte sequence.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Image) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source differs byte-for-byte.
    pub fn apply(&self, source: &Image) -> Result<Image> {
        if !self.is_applicable_to(source) {
            return Err(Error::InvalidFormat(
                "ODI package patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.target.clone())
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
    keywords: Option<String>,
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
    /// Returns comma-separated keywords.
    #[must_use]
    pub fn keywords(&self) -> Option<&str> {
        self.keywords.as_deref()
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

/// One reversible metadata change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MetadataChange {
    before: MetadataFields,
    after: MetadataFields,
}

impl MetadataChange {
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

    fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

fn patch_metadata(source: &Image, values: &MetadataFields) -> Result<String> {
    let source_xml = source.meta_xml()?.ok_or_else(|| {
        Error::InvalidFormat("ODI metadata editing requires an existing meta.xml part".to_string())
    })?;
    let parsed = OdfMetadata::from_xml(&source_xml)?;
    let mut target = Metadata::from(parsed.clone());
    target.title.clone_from(&values.title);
    target.author.clone_from(&values.author);
    target.subject.clone_from(&values.subject);
    target.description.clone_from(&values.description);
    target.keywords.clone_from(&values.keywords);
    let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&parsed, &target);
    patch_meta_xml(&source_xml, &patch)?.ok_or_else(|| {
        Error::InvalidFormat("ODI metadata editing requires an office:meta container".to_string())
    })
}

/// One reversible package-local image-resource change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceChange {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
    after_size: Option<usize>,
}

impl ResourceChange {
    /// Returns the source resource selector.
    #[must_use]
    pub fn resource(&self) -> usize {
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
        return Err(Error::InvalidFormat(
            "ODI resource media type is invalid".to_string(),
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
    use super::{Builder, Image};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Image::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:image"));
        assert!(!document.as_bytes().is_empty());
    }
}
