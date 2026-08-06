//! Detached, source-checked ActiveX/OCX metadata transactions.

use std::sync::Arc;

use litchi_opc::{PackURI, TargetMode};

use super::super::model::{Control, Descriptor, Persistence};
use super::codec::{ControlChanges, DescriptorChanges, rewrite_control, rewrite_descriptor};
use super::validation::{validate_binary, validate_source_size, validate_text};
use crate::{Error, Result};

/// Stable identity of the complete detached slide/control source graph.
pub type Revision = u64;

/// Relationship metadata retained for exact source checks and safe lifecycle
/// publication.  The target is never opened or executed by this module.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RelationshipState {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
}

/// Opaque binary state referenced by an OCX descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct BinarySource {
    pub(crate) relationship_id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
    pub(crate) part_name: PackURI,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
    pub(crate) relationships: Arc<Vec<RelationshipState>>,
}

/// An immutable, detached view of one control and its owning slide graph.
#[derive(Debug, Clone)]
pub struct Snapshot {
    pub(crate) slide_index: usize,
    pub(crate) control_index: usize,
    pub(crate) slide_part_name: PackURI,
    pub(crate) control: Control,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) slide_relationships: Arc<Vec<RelationshipState>>,
    pub(crate) descriptor_xml: Option<Arc<Vec<u8>>>,
    pub(crate) descriptor_relationships: Option<Arc<Vec<RelationshipState>>>,
    pub(crate) binary: Option<BinarySource>,
    pub(crate) revision: Revision,
}

impl Snapshot {
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn from_parts(
        slide_index: usize,
        control_index: usize,
        slide_part_name: PackURI,
        control: Control,
        source_xml: Arc<Vec<u8>>,
        slide_relationships: Vec<RelationshipState>,
        descriptor_xml: Option<Arc<Vec<u8>>>,
        descriptor_relationships: Option<Vec<RelationshipState>>,
        binary: Option<BinarySource>,
    ) -> Result<Self> {
        validate_source_size(&source_xml, "control slide XML bytes")?;
        if let Some(xml) = descriptor_xml.as_ref() {
            validate_source_size(xml, "control descriptor XML bytes")?;
        }
        if let Some(binary) = binary.as_ref() {
            validate_binary(binary.bytes.as_ref())?;
            if binary.bytes.len()
                != control
                    .descriptor
                    .as_ref()
                    .and_then(|d| d.binary())
                    .map_or(binary.bytes.len(), |b| b.byte_length())
            {
                return Err(invalid(
                    "ActiveX binary metadata length does not match its part",
                ));
            }
        }
        let slide_relationships = Arc::new(sorted_relationships(slide_relationships)?);
        let descriptor_relationships = descriptor_relationships
            .map(|value| sorted_relationships(value).map(Arc::new))
            .transpose()?;
        let revision = fingerprint(
            &source_xml,
            &slide_relationships,
            descriptor_xml.as_deref(),
            descriptor_relationships.as_deref(),
            binary.as_ref(),
        );
        Ok(Self {
            slide_index,
            control_index,
            slide_part_name,
            control,
            source_xml,
            slide_relationships,
            descriptor_xml,
            descriptor_relationships,
            binary,
            revision,
        })
    }

    /// Zero-based presentation position of the owning slide.
    #[inline]
    pub const fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Zero-based control position in the selected slide's semantic control list.
    #[inline]
    pub const fn control_index(&self) -> usize {
        self.control_index
    }

    /// Absolute OPC part name of the owning slide.
    #[inline]
    pub fn slide_part_name(&self) -> &PackURI {
        &self.slide_part_name
    }

    /// Typed inert control metadata captured by this snapshot.
    #[inline]
    pub fn control(&self) -> &Control {
        &self.control
    }

    /// Exact owning slide XML bytes captured by this snapshot.
    #[inline]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Exact OCX descriptor XML, when the control has a descriptor relationship.
    #[inline]
    pub fn descriptor_xml(&self) -> Option<&[u8]> {
        self.descriptor_xml.as_deref().map(Vec::as_slice)
    }

    /// Exact opaque ActiveX binary payload, when the descriptor has one.
    #[inline]
    pub fn binary_bytes(&self) -> Option<&[u8]> {
        self.binary.as_ref().map(|value| value.bytes.as_slice())
    }

    /// Compact source revision used for optimistic stale-source checks.
    #[inline]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start an isolated transaction over this snapshot.
    #[inline]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            control: self.control.clone(),
            binary: self.binary.as_ref().map(|value| Arc::clone(&value.bytes)),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.slide_index == other.slide_index
            && self.control_index == other.control_index
            && self.slide_part_name == other.slide_part_name
            && self.source_xml == other.source_xml
            && self.slide_relationships == other.slide_relationships
            && self.descriptor_xml == other.descriptor_xml
            && self.descriptor_relationships == other.descriptor_relationships
            && self.binary == other.binary
    }
}

/// Failure-atomic editor over typed control and descriptor metadata.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    control: Control,
    binary: Option<Arc<Vec<u8>>>,
}

impl Transaction {
    /// Borrow the immutable source snapshot used by this edit.
    #[inline]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrow the currently staged typed control metadata.
    #[inline]
    pub const fn control(&self) -> &Control {
        &self.control
    }

    /// Whether any staged metadata or payload operation changes the source.
    pub fn is_changed(&self) -> bool {
        self.control != self.source.control
            || self.binary.as_ref().map(|value| value.as_slice())
                != self
                    .source
                    .binary
                    .as_ref()
                    .map(|value| value.bytes.as_slice())
    }

    /// Replace or remove the owning slide control's display name.
    pub fn set_name(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = value.as_deref() {
            validate_text(value, "control name", true)?;
        }
        self.control.name = value;
        Ok(())
    }

    /// Replace or remove the optional `showAsIcon` flag.
    pub const fn set_show_as_icon(&mut self, value: Option<bool>) {
        self.control.show_as_icon = value;
    }

    /// Replace or remove the optional image width in EMUs.
    pub const fn set_image_width(&mut self, value: Option<u32>) {
        self.control.image_width = value;
    }

    /// Replace or remove the optional image height in EMUs.
    pub const fn set_image_height(&mut self, value: Option<u32>) {
        self.control.image_height = value;
    }

    /// Replace the descriptor's required `ax:classid` value.
    pub fn set_class_id(&mut self, value: impl Into<String>) -> Result<()> {
        let value = value.into();
        validate_text(&value, "ActiveX class ID", false)?;
        self.descriptor_mut()?.class_id = value;
        Ok(())
    }

    /// Replace or remove the descriptor's optional `ax:license` value.
    pub fn set_license(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = value.as_deref() {
            validate_text(value, "ActiveX license", true)?;
        }
        self.descriptor_mut()?.license = value;
        Ok(())
    }

    /// Set the descriptor persistence mode. `Unknown` removes a known
    /// persistence attribute and leaves the descriptor in its inert default
    /// state.
    pub fn set_persistence(&mut self, value: Persistence) -> Result<()> {
        self.descriptor_mut()?.persistence = value;
        Ok(())
    }

    /// Replace the opaque binary state without interpreting or activating it.
    pub fn replace_binary(&mut self, bytes: impl Into<Vec<u8>>) -> Result<()> {
        let bytes = bytes.into();
        validate_binary(&bytes)?;
        if self.binary.is_none() {
            return Err(invalid(
                "cannot create an ActiveX binary relationship in this transaction",
            ));
        }
        let byte_length = bytes.len();
        self.binary = Some(Arc::new(bytes));
        if let Some(binary) = self.descriptor_mut()?.binary.as_mut() {
            binary.byte_length = byte_length;
        }
        Ok(())
    }

    /// Remove the descriptor's binary relationship. An unshared orphan part
    /// is collected during package publication; a shared part is retained.
    pub fn remove_binary(&mut self) -> Result<()> {
        if self.binary.is_none() {
            return Ok(());
        }
        self.binary = None;
        self.descriptor_mut()?.binary = None;
        Ok(())
    }

    /// Validate and consume this edit into a reversible package patch.
    pub fn commit(self) -> Result<Commit> {
        validate_control(&self.control)?;
        let relationship_id = self
            .source
            .control
            .relationship_id
            .as_deref()
            .ok_or_else(|| invalid("ActiveX transaction requires a control relationship ID"))?;
        if self.control.relationship_id.as_deref() != Some(relationship_id) {
            return Err(invalid("control relationship identity cannot change"));
        }

        let control_changes = ControlChanges {
            name: (self.control.name != self.source.control.name)
                .then(|| self.control.name.clone()),
            show_as_icon: (self.control.show_as_icon != self.source.control.show_as_icon)
                .then(|| self.control.show_as_icon),
            image_width: (self.control.image_width != self.source.control.image_width)
                .then(|| self.control.image_width),
            image_height: (self.control.image_height != self.source.control.image_height)
                .then(|| self.control.image_height),
        };
        let source_xml = if control_changes != ControlChanges::default() {
            Arc::new(rewrite_control(
                self.source.source_xml.as_slice(),
                relationship_id,
                &control_changes,
            )?)
        } else {
            Arc::clone(&self.source.source_xml)
        };

        let source_descriptor = self.source.control.descriptor.as_ref();
        let descriptor = self.control.descriptor.as_ref();
        let descriptor_changes = DescriptorChanges {
            class_id: descriptor
                .zip(source_descriptor)
                .and_then(|(after, before)| {
                    (after.class_id != before.class_id).then(|| after.class_id.clone())
                }),
            license: descriptor
                .zip(source_descriptor)
                .and_then(|(after, before)| {
                    (after.license != before.license).then(|| after.license.clone())
                }),
            persistence: descriptor
                .zip(source_descriptor)
                .and_then(|(after, before)| {
                    (after.persistence != before.persistence).then_some(after.persistence)
                }),
            remove_binary_relationship: self.source.binary.is_some() && self.binary.is_none(),
        };
        let descriptor_xml = match self.source.descriptor_xml.as_ref() {
            Some(source) if descriptor_changes != DescriptorChanges::default() => Some(Arc::new(
                rewrite_descriptor(source.as_slice(), &descriptor_changes)?,
            )),
            Some(source) => Some(Arc::clone(source)),
            None => None,
        };

        let mut descriptor_relationships = self
            .source
            .descriptor_relationships
            .as_ref()
            .map(|value| value.as_ref().clone());
        if descriptor_changes.remove_binary_relationship {
            if let Some(binary) = self.source.binary.as_ref() {
                if let Some(relationships) = descriptor_relationships.as_mut() {
                    relationships.retain(|value| value.id != binary.relationship_id);
                }
            }
        }
        let binary = match (&self.source.binary, &self.binary) {
            (Some(source), Some(bytes)) => {
                let mut value = source.clone();
                value.bytes = Arc::clone(bytes);
                Some(value)
            },
            (Some(_), None) | (None, None) => None,
            (None, Some(_)) => {
                return Err(invalid(
                    "cannot create an ActiveX binary relationship in this transaction",
                ));
            },
        };
        let snapshot = Snapshot::from_parts(
            self.source.slide_index,
            self.source.control_index,
            self.source.slide_part_name.clone(),
            self.control,
            source_xml,
            self.source.slide_relationships.as_ref().clone(),
            descriptor_xml,
            descriptor_relationships,
            binary,
        )?;
        let patch = Patch {
            before: self.source,
            after: snapshot.clone(),
        };
        Ok(Commit { snapshot, patch })
    }

    /// Discard staged edits and return the exact source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn descriptor_mut(&mut self) -> Result<&mut Descriptor> {
        self.control
            .descriptor
            .as_mut()
            .ok_or_else(|| invalid("selected control has no OCX descriptor"))
    }
}

/// Successful publication candidate and its reversible package patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the validated detached target snapshot.
    #[inline]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the source-checked package patch.
    #[inline]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether the transaction changes any serialized graph bytes or edges.
    #[inline]
    pub fn is_changed(&self) -> bool {
        !self.patch.is_empty()
    }

    /// Consume this commit into its target snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    /// Consume this commit into its source-checked package patch.
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible, source-checked replacement of one slide-owned ActiveX graph.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Source snapshot required for forward application.
    #[inline]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Target snapshot produced by forward application.
    #[inline]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact graph no-op.
    #[inline]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Public spelling for callers that inspect transaction changes.
    #[inline]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Build the exact inverse patch.
    #[inline]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Publish this patch atomically to its owning OPC package.
    pub fn apply(&self, package: &mut litchi_opc::OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(package, self)
    }

    /// Alias for undoing a committed patch against its exact target graph.
    pub fn undo(&self, package: &mut litchi_opc::OpcPackage) -> Result<Snapshot> {
        self.inverse().apply(package)
    }
}

fn validate_control(control: &Control) -> Result<()> {
    if let Some(value) = control.name.as_deref() {
        validate_text(value, "control name", true)?;
    }
    if let Some(descriptor) = control.descriptor.as_ref() {
        validate_text(&descriptor.class_id, "ActiveX class ID", false)?;
        if let Some(value) = descriptor.license.as_deref() {
            validate_text(value, "ActiveX license", true)?;
        }
    }
    Ok(())
}

fn sorted_relationships(
    mut relationships: Vec<RelationshipState>,
) -> Result<Vec<RelationshipState>> {
    for relationship in &relationships {
        validate_text(&relationship.id, "ActiveX relationship ID", false)?;
        validate_text(
            &relationship.relationship_type,
            "ActiveX relationship type",
            false,
        )?;
        validate_text(
            &relationship.target_ref,
            "ActiveX relationship target",
            false,
        )?;
    }
    relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    for pair in relationships.windows(2) {
        if pair[0].id == pair[1].id {
            return Err(invalid("duplicate ActiveX relationship ID"));
        }
    }
    Ok(relationships)
}

fn fingerprint(
    source_xml: &[u8],
    slide_relationships: &[RelationshipState],
    descriptor_xml: Option<&Vec<u8>>,
    descriptor_relationships: Option<&Vec<RelationshipState>>,
    binary: Option<&BinarySource>,
) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    feed_bytes(&mut hash, source_xml);
    feed_relationships(&mut hash, slide_relationships);
    match descriptor_xml {
        Some(value) => {
            hash ^= 1;
            feed_bytes(&mut hash, value);
        },
        None => hash ^= 0,
    }
    if let Some(value) = descriptor_relationships {
        feed_relationships(&mut hash, value);
    }
    if let Some(binary) = binary {
        feed_text(&mut hash, &binary.relationship_id);
        feed_text(&mut hash, &binary.relationship_type);
        feed_text(&mut hash, &binary.target_ref);
        feed_text(&mut hash, binary.part_name.as_str());
        feed_text(&mut hash, &binary.content_type);
        feed_bytes(&mut hash, &binary.bytes);
        feed_relationships(&mut hash, &binary.relationships);
    }
    hash
}

fn feed_relationships(hash: &mut u64, relationships: &[RelationshipState]) {
    for relationship in relationships {
        feed_text(hash, &relationship.id);
        feed_text(hash, &relationship.relationship_type);
        feed_text(hash, &relationship.target_ref);
        *hash ^= match relationship.target_mode {
            TargetMode::Internal => 0,
            TargetMode::External => 1,
        };
        *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
}

fn feed_text(hash: &mut u64, value: &str) {
    feed_bytes(hash, value.as_bytes());
}

fn feed_bytes(hash: &mut u64, value: &[u8]) {
    for byte in value {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    *hash ^= 0xff;
    *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
