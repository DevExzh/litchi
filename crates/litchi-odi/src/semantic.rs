//! Deterministic semantic patches, transfer, and non-mutating merge planning.

use crate::{
    Commit, Edit, FlatImage, FlatImageCommit, FrameChange, Image, MetadataFields, ResourceChange,
    frame::Frame, map::ImageMap, source::Source,
};
use litchi_core::{Error, Result};
use std::{collections::BTreeMap, sync::Arc};

const MIB: usize = 1024 * 1024;

/// Security and allocation limits for semantic planning and publication.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SecurityPolicy {
    max_operations: usize,
    max_patch_bytes: usize,
    max_resource_bytes: usize,
    max_xml_bytes: usize,
    max_text_bytes: usize,
    max_map_areas: usize,
    allow_external_links: bool,
    allow_package_members: bool,
}

impl SecurityPolicy {
    /// Returns a policy suited to hostile, untrusted edits.
    #[must_use]
    pub const fn strict() -> Self {
        Self {
            max_operations: 1_024,
            max_patch_bytes: 64 * MIB,
            max_resource_bytes: 16 * MIB,
            max_xml_bytes: 4 * MIB,
            max_text_bytes: 64 * 1024,
            max_map_areas: 10_000,
            allow_external_links: false,
            allow_package_members: true,
        }
    }

    /// Sets whether newly transferred external URI values are accepted.
    #[must_use]
    pub const fn with_external_links(mut self, allow: bool) -> Self {
        self.allow_external_links = allow;
        self
    }

    /// Sets whether package-member operations are accepted.
    #[must_use]
    pub const fn with_package_members(mut self, allow: bool) -> Self {
        self.allow_package_members = allow;
        self
    }

    /// Sets the maximum retained semantic-patch byte total.
    #[must_use]
    pub const fn with_patch_bytes(mut self, bytes: usize) -> Self {
        self.max_patch_bytes = bytes;
        self
    }

    /// Sets the maximum bytes in one package-resource value.
    #[must_use]
    pub const fn with_resource_bytes(mut self, bytes: usize) -> Self {
        self.max_resource_bytes = bytes;
        self
    }

    /// Sets the maximum bytes in one authored XML part value.
    #[must_use]
    pub const fn with_xml_bytes(mut self, bytes: usize) -> Self {
        self.max_xml_bytes = bytes;
        self
    }

    /// Sets the maximum bytes in one semantic text value.
    #[must_use]
    pub const fn with_text_bytes(mut self, bytes: usize) -> Self {
        self.max_text_bytes = bytes;
        self
    }

    /// Sets the maximum number of semantic operations.
    #[must_use]
    pub const fn with_operations(mut self, count: usize) -> Self {
        self.max_operations = count;
        self
    }

    /// Sets the maximum number of areas in one image-map value.
    #[must_use]
    pub const fn with_map_areas(mut self, count: usize) -> Self {
        self.max_map_areas = count;
        self
    }
}

impl Default for SecurityPolicy {
    fn default() -> Self {
        Self {
            max_operations: 100_000,
            max_patch_bytes: 512 * MIB,
            max_resource_bytes: 256 * MIB,
            max_xml_bytes: 64 * MIB,
            max_text_bytes: 1024 * 1024,
            max_map_areas: 100_000,
            allow_external_links: true,
            allow_package_members: true,
        }
    }
}

/// Serialized artifact kind bound to an exact semantic patch.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ArtifactKind {
    /// Flat single-XML ODI.
    Flat,
    /// Packaged ODI archive.
    Package,
}

/// One editable frame property.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum FrameProperty {
    Name,
    XmlId,
    Title,
    Description,
    Source,
    MediaType,
    ImageXmlId,
    FilterName,
    LinkType,
    Show,
    Actuate,
    StyleName,
    TextStyleName,
    Layer,
    ZIndex,
    Transform,
    AnchorType,
    X,
    Y,
    Width,
    Height,
    RelativeWidth,
    RelativeHeight,
    CopyOf,
    ImageMap,
}

/// One editable common metadata property.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum MetadataProperty {
    Title,
    Author,
    Subject,
    Description,
    Keywords,
}

/// Stable operation identity used for deterministic ordering and conflicts.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum OperationKey {
    Frame {
        frame: usize,
        property: FrameProperty,
    },
    Metadata(MetadataProperty),
    Styles,
    Resource(String),
}

/// One package resource state retained by a semantic operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceValue {
    media_type: String,
    bytes: Vec<u8>,
}

impl ResourceValue {
    /// Creates a typed resource payload.
    #[must_use]
    pub fn new(media_type: String, bytes: Vec<u8>) -> Self {
        Self { media_type, bytes }
    }

    /// Returns the manifest media type.
    #[must_use]
    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    /// Returns the inert member bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
}

/// A typed before/after semantic value.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticValue {
    Text(Option<String>),
    Unsigned(Option<u32>),
    Source(Source),
    ImageMap(Option<ImageMap>),
    Xml(Option<String>),
    Resource(Option<ResourceValue>),
}

/// One deterministic compare-and-set semantic operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SemanticOperation {
    key: OperationKey,
    before: SemanticValue,
    after: SemanticValue,
}

impl SemanticOperation {
    /// Returns the stable operation identity.
    #[must_use]
    pub const fn key(&self) -> &OperationKey {
        &self.key
    }

    /// Returns the required source value.
    #[must_use]
    pub const fn before(&self) -> &SemanticValue {
        &self.before
    }

    /// Returns the desired target value.
    #[must_use]
    pub const fn after(&self) -> &SemanticValue {
        &self.after
    }

    fn inverse(&self) -> Self {
        Self {
            key: self.key.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Why one operation could not be joined or transferred.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ConflictKind {
    Diverged,
    IncompatibleBase,
    Unsupported,
    Policy,
}

/// One deterministic planning conflict. Planning never mutates an artifact.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Conflict {
    key: Option<OperationKey>,
    kind: ConflictKind,
    expected: Option<SemanticValue>,
    actual: Option<SemanticValue>,
    desired: Option<SemanticValue>,
}

impl Conflict {
    /// Returns the affected key, or `None` for an artifact-wide conflict.
    #[must_use]
    pub const fn key(&self) -> Option<&OperationKey> {
        self.key.as_ref()
    }

    /// Returns the conflict category.
    #[must_use]
    pub const fn kind(&self) -> ConflictKind {
        self.kind
    }

    /// Returns the expected common-base value.
    #[must_use]
    pub const fn expected(&self) -> Option<&SemanticValue> {
        self.expected.as_ref()
    }

    /// Returns the actual current or competing value.
    #[must_use]
    pub const fn actual(&self) -> Option<&SemanticValue> {
        self.actual.as_ref()
    }

    /// Returns the desired value.
    #[must_use]
    pub const fn desired(&self) -> Option<&SemanticValue> {
        self.desired.as_ref()
    }
}

/// A deterministic non-mutating merge or transfer plan.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SemanticPlan {
    operations: Vec<SemanticOperation>,
    conflicts: Vec<Conflict>,
}

impl SemanticPlan {
    /// Returns applicable operations in stable key order.
    #[must_use]
    pub fn operations(&self) -> &[SemanticOperation] {
        &self.operations
    }

    /// Returns conflicts in stable key order.
    #[must_use]
    pub fn conflicts(&self) -> &[Conflict] {
        &self.conflicts
    }

    /// Returns whether this plan can be published atomically.
    #[must_use]
    pub fn is_conflict_free(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// Applies this plan to flat XML and fully reopens the result.
    pub fn commit_flat(
        &self,
        source: &FlatImage,
        policy: &SecurityPolicy,
    ) -> Result<FlatImageCommit> {
        self.ensure_publishable(policy)?;
        ensure_current_values(&self.operations, |key| flat_value(source, key))?;
        let mut edit = source.transaction();
        for operation in &self.operations {
            apply_flat_operation(&mut edit, operation)?;
        }
        let commit = edit.commit()?;
        let reopened = FlatImage::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        if reopened.as_bytes() != commit.snapshot().as_bytes() {
            return Err(invalid("ODI flat plan failed full byte reopen"));
        }
        Ok(commit)
    }

    /// Applies this plan to a package and fully reopens the result.
    pub fn commit_package(&self, source: &Image, policy: &SecurityPolicy) -> Result<Commit> {
        self.ensure_publishable(policy)?;
        ensure_current_values(&self.operations, |key| package_value(source, key))?;
        let mut edit = source.edit();
        for operation in &self.operations {
            apply_package_operation(&mut edit, operation)?;
        }
        let commit = edit.commit()?;
        let reopened = Image::from_bytes(commit.image().as_bytes().to_vec())?;
        if reopened.as_bytes() != commit.image().as_bytes() {
            return Err(invalid("ODI package plan failed full byte reopen"));
        }
        Ok(commit)
    }

    fn ensure_publishable(&self, policy: &SecurityPolicy) -> Result<()> {
        if !self.conflicts.is_empty() {
            return Err(invalid("ODI semantic plan contains unresolved conflicts"));
        }
        validate_operations(&self.operations, policy)
    }
}

fn ensure_current_values(
    operations: &[SemanticOperation],
    mut current: impl FnMut(&OperationKey) -> Result<Option<SemanticValue>>,
) -> Result<()> {
    for operation in operations {
        if current(operation.key())?.as_ref() != Some(operation.before()) {
            return Err(invalid("stale ODI semantic plan source"));
        }
    }
    Ok(())
}

/// An exact-source, deterministic semantic patch with exact target bytes.
#[derive(Clone, Debug)]
pub struct SemanticPatch {
    kind: ArtifactKind,
    source: Arc<Vec<u8>>,
    target: Arc<Vec<u8>>,
    operations: Vec<SemanticOperation>,
}

impl SemanticPatch {
    pub(crate) fn from_flat_commit(
        commit: &FlatImageCommit,
        policy: &SecurityPolicy,
    ) -> Result<Self> {
        let mut operations = frame_operations(commit.patch().changes());
        if let Some(change) = commit.patch().metadata_change() {
            operations.extend(metadata_operations(change.before(), change.after()));
        }
        Self::new(
            ArtifactKind::Flat,
            commit.source().as_bytes(),
            commit.snapshot().as_bytes(),
            operations,
            policy,
        )
    }

    pub(crate) fn from_package_commit(commit: &Commit, policy: &SecurityPolicy) -> Result<Self> {
        let mut operations = frame_operations(commit.patch().changes());
        if let Some(change) = commit.patch().metadata_change() {
            operations.extend(metadata_operations(change.before(), change.after()));
        }
        if let Some(change) = commit.patch().style_change() {
            operations.push(SemanticOperation {
                key: OperationKey::Styles,
                before: SemanticValue::Xml(change.before().map(str::to_owned)),
                after: SemanticValue::Xml(change.after().map(str::to_owned)),
            });
        }
        for change in commit.patch().resource_changes() {
            operations.push(resource_operation(commit.source(), commit.image(), change)?);
        }
        Self::new(
            ArtifactKind::Package,
            commit.source().as_bytes(),
            commit.image().as_bytes(),
            operations,
            policy,
        )
    }

    fn new(
        kind: ArtifactKind,
        source: &[u8],
        target: &[u8],
        mut operations: Vec<SemanticOperation>,
        policy: &SecurityPolicy,
    ) -> Result<Self> {
        operations.sort_unstable_by(|left, right| left.key.cmp(&right.key));
        operations.dedup();
        validate_operations(&operations, policy)?;
        let retained = source
            .len()
            .checked_add(target.len())
            .and_then(|size| {
                operations.iter().try_fold(size, |total, operation| {
                    total.checked_add(operation_size(operation))
                })
            })
            .ok_or_else(|| invalid("ODI semantic patch byte count overflow"))?;
        if retained > policy.max_patch_bytes {
            return Err(invalid("ODI semantic patch exceeds the byte policy"));
        }
        Ok(Self {
            kind,
            source: Arc::new(source.to_vec()),
            target: Arc::new(target.to_vec()),
            operations,
        })
    }

    /// Returns the bound artifact kind.
    #[must_use]
    pub const fn kind(&self) -> ArtifactKind {
        self.kind
    }

    /// Returns stable compare-and-set operations.
    #[must_use]
    pub fn operations(&self) -> &[SemanticOperation] {
        &self.operations
    }

    /// Returns a patch that restores the exact source bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            kind: self.kind,
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            operations: self
                .operations
                .iter()
                .map(SemanticOperation::inverse)
                .collect(),
        }
    }

    /// Applies only to the exact flat source and reparses the target bytes.
    pub fn apply_flat(&self, source: &FlatImage) -> Result<FlatImage> {
        if self.kind != ArtifactKind::Flat || self.source.as_slice() != source.as_bytes() {
            return Err(invalid("stale or wrong-kind ODI semantic patch source"));
        }
        FlatImage::from_bytes(self.target.as_ref().clone())
    }

    /// Applies only to the exact package source and reparses the target bytes.
    pub fn apply_package(&self, source: &Image) -> Result<Image> {
        if self.kind != ArtifactKind::Package || self.source.as_slice() != source.as_bytes() {
            return Err(invalid("stale or wrong-kind ODI semantic patch source"));
        }
        Image::from_bytes(self.target.as_ref().clone())
    }

    /// Joins two patches derived from the same exact source without mutation.
    #[must_use]
    pub fn join(&self, other: &Self) -> SemanticPlan {
        if self.kind != other.kind || self.source != other.source {
            return SemanticPlan {
                operations: Vec::new(),
                conflicts: vec![artifact_conflict(ConflictKind::IncompatibleBase)],
            };
        }
        join_operations(&self.operations, &other.operations)
    }

    /// Plans a three-way application against current flat semantics.
    #[must_use]
    pub fn plan_flat(&self, current: &FlatImage) -> SemanticPlan {
        plan_operations(&self.operations, |key| flat_value(current, key))
    }

    /// Plans a three-way application or cross-artifact transfer to a package.
    #[must_use]
    pub fn plan_package(&self, current: &Image) -> SemanticPlan {
        plan_operations(&self.operations, |key| package_value(current, key))
    }
}

fn frame_operations(changes: &[FrameChange]) -> Vec<SemanticOperation> {
    let mut operations = Vec::new();
    for change in changes {
        for property in [
            FrameProperty::Name,
            FrameProperty::XmlId,
            FrameProperty::Title,
            FrameProperty::Description,
            FrameProperty::Source,
            FrameProperty::MediaType,
            FrameProperty::ImageXmlId,
            FrameProperty::FilterName,
            FrameProperty::LinkType,
            FrameProperty::Show,
            FrameProperty::Actuate,
            FrameProperty::StyleName,
            FrameProperty::TextStyleName,
            FrameProperty::Layer,
            FrameProperty::ZIndex,
            FrameProperty::Transform,
            FrameProperty::AnchorType,
            FrameProperty::X,
            FrameProperty::Y,
            FrameProperty::Width,
            FrameProperty::Height,
            FrameProperty::RelativeWidth,
            FrameProperty::RelativeHeight,
            FrameProperty::CopyOf,
            FrameProperty::ImageMap,
        ] {
            let before = frame_value(change.before(), property);
            let after = frame_value(change.after(), property);
            if before != after {
                operations.push(SemanticOperation {
                    key: OperationKey::Frame {
                        frame: change.frame(),
                        property,
                    },
                    before,
                    after,
                });
            }
        }
    }
    operations
}

fn metadata_operations(before: &MetadataFields, after: &MetadataFields) -> Vec<SemanticOperation> {
    let mut operations = Vec::new();
    for property in [
        MetadataProperty::Title,
        MetadataProperty::Author,
        MetadataProperty::Subject,
        MetadataProperty::Description,
        MetadataProperty::Keywords,
    ] {
        let old = metadata_value(before, property);
        let new = metadata_value(after, property);
        if old != new {
            operations.push(SemanticOperation {
                key: OperationKey::Metadata(property),
                before: old,
                after: new,
            });
        }
    }
    operations
}

fn resource_operation(
    source: &Image,
    target: &Image,
    change: &ResourceChange,
) -> Result<SemanticOperation> {
    Ok(SemanticOperation {
        key: OperationKey::Resource(change.path().to_owned()),
        before: resource_value(source, change.path())?,
        after: resource_value(target, change.path())?,
    })
}

fn resource_value(image: &Image, path: &str) -> Result<SemanticValue> {
    let node = image
        .resource_graph()
        .nodes()
        .iter()
        .find(|node| node.path() == path);
    let Some(bytes) = image.member_bytes(path)? else {
        return Ok(SemanticValue::Resource(None));
    };
    let media_type = node
        .and_then(crate::resource::Node::media_type)
        .ok_or_else(|| invalid("ODI semantic resource has no manifest media type"))?;
    Ok(SemanticValue::Resource(Some(ResourceValue::new(
        media_type.to_owned(),
        bytes,
    ))))
}

fn join_operations(left: &[SemanticOperation], right: &[SemanticOperation]) -> SemanticPlan {
    let mut joined = BTreeMap::<OperationKey, SemanticOperation>::new();
    let mut conflicts = Vec::new();
    for operation in left.iter().chain(right) {
        if let Some(existing) = joined.get(operation.key()) {
            if existing != operation {
                conflicts.push(Conflict {
                    key: Some(operation.key.clone()),
                    kind: ConflictKind::Diverged,
                    expected: Some(existing.before.clone()),
                    actual: Some(operation.before.clone()),
                    desired: Some(operation.after.clone()),
                });
            }
        } else {
            joined.insert(operation.key.clone(), operation.clone());
        }
    }
    SemanticPlan {
        operations: joined.into_values().collect(),
        conflicts,
    }
}

fn plan_operations(
    operations: &[SemanticOperation],
    mut current: impl FnMut(&OperationKey) -> Result<Option<SemanticValue>>,
) -> SemanticPlan {
    let mut applicable = Vec::new();
    let mut conflicts = Vec::new();
    for operation in operations {
        match current(operation.key()) {
            Ok(Some(value)) if value == operation.before => applicable.push(operation.clone()),
            Ok(Some(value)) if value == operation.after => {},
            Ok(Some(value)) => conflicts.push(Conflict {
                key: Some(operation.key.clone()),
                kind: ConflictKind::Diverged,
                expected: Some(operation.before.clone()),
                actual: Some(value),
                desired: Some(operation.after.clone()),
            }),
            Ok(None) | Err(_) => conflicts.push(Conflict {
                key: Some(operation.key.clone()),
                kind: ConflictKind::Unsupported,
                expected: Some(operation.before.clone()),
                actual: None,
                desired: Some(operation.after.clone()),
            }),
        }
    }
    SemanticPlan {
        operations: applicable,
        conflicts,
    }
}

fn flat_value(image: &FlatImage, key: &OperationKey) -> Result<Option<SemanticValue>> {
    match key {
        OperationKey::Frame { frame, property } => image
            .frames()
            .get(*frame)
            .map(|value| frame_value(value, *property))
            .map(Some)
            .ok_or_else(|| invalid("ODI semantic frame selector is out of bounds")),
        OperationKey::Metadata(property) => image.metadata().map_or(Ok(None), |metadata| {
            Ok(Some(metadata_value(
                &MetadataFields::from(Some(metadata)),
                *property,
            )))
        }),
        OperationKey::Styles | OperationKey::Resource(_) => Ok(None),
    }
}

fn package_value(image: &Image, key: &OperationKey) -> Result<Option<SemanticValue>> {
    match key {
        OperationKey::Frame { frame, property } => image
            .frames()
            .get(*frame)
            .map(|value| frame_value(value, *property))
            .map(Some)
            .ok_or_else(|| invalid("ODI semantic frame selector is out of bounds")),
        OperationKey::Metadata(property) => image.metadata().map_or(Ok(None), |metadata| {
            Ok(Some(metadata_value(
                &MetadataFields::from(Some(metadata)),
                *property,
            )))
        }),
        OperationKey::Styles => Ok(Some(SemanticValue::Xml(
            image.styles_xml().map(str::to_owned),
        ))),
        OperationKey::Resource(path) => resource_value(image, path).map(Some),
    }
}

fn frame_value(frame: &Frame, property: FrameProperty) -> SemanticValue {
    match property {
        FrameProperty::Name => text(frame.name()),
        FrameProperty::XmlId => text(frame.xml_id()),
        FrameProperty::Title => text(frame.title()),
        FrameProperty::Description => text(frame.description()),
        FrameProperty::Source => SemanticValue::Source(frame.source().clone()),
        FrameProperty::MediaType => text(frame.media_type()),
        FrameProperty::ImageXmlId => text(frame.image_xml_id()),
        FrameProperty::FilterName => text(frame.filter_name()),
        FrameProperty::LinkType => text(frame.link_type()),
        FrameProperty::Show => text(frame.show()),
        FrameProperty::Actuate => text(frame.actuate()),
        FrameProperty::StyleName => text(frame.style_name()),
        FrameProperty::TextStyleName => text(frame.text_style_name()),
        FrameProperty::Layer => text(frame.layer()),
        FrameProperty::ZIndex => SemanticValue::Unsigned(frame.z_index()),
        FrameProperty::Transform => text(frame.transform()),
        FrameProperty::AnchorType => text(frame.anchor_type()),
        FrameProperty::X => text(frame.x()),
        FrameProperty::Y => text(frame.y()),
        FrameProperty::Width => text(frame.width()),
        FrameProperty::Height => text(frame.height()),
        FrameProperty::RelativeWidth => text(frame.relative_width()),
        FrameProperty::RelativeHeight => text(frame.relative_height()),
        FrameProperty::CopyOf => text(frame.copy_of()),
        FrameProperty::ImageMap => SemanticValue::ImageMap(frame.image_map().cloned()),
    }
}

fn metadata_value(fields: &MetadataFields, property: MetadataProperty) -> SemanticValue {
    let value = match property {
        MetadataProperty::Title => fields.title(),
        MetadataProperty::Author => fields.author(),
        MetadataProperty::Subject => fields.subject(),
        MetadataProperty::Description => fields.description(),
        MetadataProperty::Keywords => fields.keywords(),
    };
    text(value)
}

fn text(value: Option<&str>) -> SemanticValue {
    SemanticValue::Text(value.map(str::to_owned))
}

fn apply_flat_operation(
    edit: &mut crate::FlatImageTransaction,
    operation: &SemanticOperation,
) -> Result<()> {
    match operation.key() {
        OperationKey::Frame { frame, property } => {
            apply_frame_operation(edit, *frame, *property, operation.after())
        },
        OperationKey::Metadata(property) => {
            let SemanticValue::Text(value) = operation.after() else {
                return Err(invalid("ODI metadata operation has the wrong value type"));
            };
            match property {
                MetadataProperty::Title => edit.set_title(value.clone()),
                MetadataProperty::Author => edit.set_author(value.clone()),
                MetadataProperty::Subject => edit.set_subject(value.clone()),
                MetadataProperty::Description => edit.set_description(value.clone()),
                MetadataProperty::Keywords => edit.set_keywords(value.clone()),
            }
        },
        OperationKey::Styles => Err(invalid("ODI flat plan contains a package style operation")),
        OperationKey::Resource(_) => Err(invalid(
            "ODI flat plan contains a package-resource operation",
        )),
    }
}

fn apply_package_operation(edit: &mut Edit<'_>, operation: &SemanticOperation) -> Result<()> {
    match operation.key() {
        OperationKey::Frame { frame, property } => {
            apply_frame_operation(edit, *frame, *property, operation.after())
        },
        OperationKey::Metadata(property) => {
            let SemanticValue::Text(value) = operation.after() else {
                return Err(invalid("ODI metadata operation has the wrong value type"));
            };
            match property {
                MetadataProperty::Title => edit.set_title(value.clone()),
                MetadataProperty::Author => edit.set_author(value.clone()),
                MetadataProperty::Subject => edit.set_subject(value.clone()),
                MetadataProperty::Description => edit.set_description(value.clone()),
                MetadataProperty::Keywords => edit.set_keywords(value.clone()),
            }
        },
        OperationKey::Styles => {
            let SemanticValue::Xml(value) = operation.after() else {
                return Err(invalid("ODI style operation has the wrong value type"));
            };
            edit.set_styles_xml(value.clone())
        },
        OperationKey::Resource(path) => {
            let SemanticValue::Resource(value) = operation.after() else {
                return Err(invalid("ODI resource operation has the wrong value type"));
            };
            if let Some(resource) = value {
                edit.put_member(
                    path.clone(),
                    resource.media_type.clone(),
                    resource.bytes.clone(),
                )
            } else {
                edit.remove_member(path.clone())
            }
        },
    }
}

fn apply_frame_operation(
    edit: &mut impl crate::FrameEditor,
    frame: usize,
    property: FrameProperty,
    value: &SemanticValue,
) -> Result<()> {
    match (property, value) {
        (FrameProperty::Name, SemanticValue::Text(value)) => {
            edit.set_frame_name(frame, value.clone())
        },
        (FrameProperty::XmlId, SemanticValue::Text(value)) => {
            edit.set_frame_xml_id(frame, value.clone())
        },
        (FrameProperty::Title, SemanticValue::Text(value)) => {
            edit.set_frame_title(frame, value.clone())
        },
        (FrameProperty::Description, SemanticValue::Text(value)) => {
            edit.set_frame_description(frame, value.clone())
        },
        (FrameProperty::Source, SemanticValue::Source(value)) => {
            edit.set_source(frame, value.clone())
        },
        (FrameProperty::MediaType, SemanticValue::Text(value)) => {
            edit.set_image_media_type(frame, value.clone())
        },
        (FrameProperty::ImageXmlId, SemanticValue::Text(value)) => {
            edit.set_image_xml_id(frame, value.clone())
        },
        (FrameProperty::FilterName, SemanticValue::Text(value)) => {
            edit.set_filter_name(frame, value.clone())
        },
        (FrameProperty::LinkType, SemanticValue::Text(value)) => {
            edit.set_link_type(frame, value.clone())
        },
        (FrameProperty::Show, SemanticValue::Text(value)) => edit.set_show(frame, value.clone()),
        (FrameProperty::Actuate, SemanticValue::Text(value)) => {
            edit.set_actuate(frame, value.clone())
        },
        (FrameProperty::StyleName, SemanticValue::Text(value)) => {
            edit.set_style_name(frame, value.clone())
        },
        (FrameProperty::TextStyleName, SemanticValue::Text(value)) => {
            edit.set_text_style_name(frame, value.clone())
        },
        (FrameProperty::Layer, SemanticValue::Text(value)) => edit.set_layer(frame, value.clone()),
        (FrameProperty::ZIndex, SemanticValue::Unsigned(value)) => edit.set_z_index(frame, *value),
        (FrameProperty::Transform, SemanticValue::Text(value)) => {
            edit.set_transform(frame, value.clone())
        },
        (FrameProperty::AnchorType, SemanticValue::Text(value)) => {
            edit.set_anchor_type(frame, value.clone())
        },
        (FrameProperty::ImageMap, SemanticValue::ImageMap(value)) => {
            edit.set_image_map(frame, value.clone())
        },
        (FrameProperty::X, SemanticValue::Text(value)) => edit.set_x(frame, value.clone()),
        (FrameProperty::Y, SemanticValue::Text(value)) => edit.set_y(frame, value.clone()),
        (FrameProperty::Width, SemanticValue::Text(value)) => edit.set_width(frame, value.clone()),
        (FrameProperty::Height, SemanticValue::Text(value)) => {
            edit.set_height(frame, value.clone())
        },
        (FrameProperty::RelativeWidth, SemanticValue::Text(value)) => {
            edit.set_relative_width(frame, value.clone())
        },
        (FrameProperty::RelativeHeight, SemanticValue::Text(value)) => {
            edit.set_relative_height(frame, value.clone())
        },
        (FrameProperty::CopyOf, SemanticValue::Text(value)) => {
            edit.set_copy_of(frame, value.clone())
        },
        _ => Err(invalid("ODI frame operation has the wrong value type")),
    }
}

fn validate_operations(operations: &[SemanticOperation], policy: &SecurityPolicy) -> Result<()> {
    if operations.len() > policy.max_operations {
        return Err(invalid("ODI semantic operation count exceeds policy"));
    }
    for operation in operations {
        for value in [operation.before(), operation.after()] {
            match value {
                SemanticValue::Text(Some(text)) if text.len() > policy.max_text_bytes => {
                    return Err(invalid("ODI semantic text exceeds policy"));
                },
                SemanticValue::Source(Source::Embedded(bytes))
                    if bytes.len() > policy.max_resource_bytes =>
                {
                    return Err(invalid("ODI inline image exceeds resource policy"));
                },
                SemanticValue::Source(Source::Linked(href))
                    if !policy.allow_external_links && is_external(href) =>
                {
                    return Err(invalid("ODI external link is refused by policy"));
                },
                SemanticValue::ImageMap(Some(map)) => validate_map(map, policy)?,
                SemanticValue::Xml(Some(xml)) if xml.len() > policy.max_xml_bytes => {
                    return Err(invalid("ODI authored XML part exceeds policy"));
                },
                SemanticValue::Xml(_) if !policy.allow_package_members => {
                    return Err(invalid("ODI package XML part is refused by policy"));
                },
                SemanticValue::Resource(Some(resource))
                    if !policy.allow_package_members
                        || resource.bytes.len() > policy.max_resource_bytes =>
                {
                    return Err(invalid("ODI package resource is refused by policy"));
                },
                SemanticValue::Text(_)
                | SemanticValue::Unsigned(_)
                | SemanticValue::Source(_)
                | SemanticValue::ImageMap(_)
                | SemanticValue::Xml(_)
                | SemanticValue::Resource(_) => {},
            }
        }
    }
    Ok(())
}

fn operation_size(operation: &SemanticOperation) -> usize {
    value_size(operation.before()).saturating_add(value_size(operation.after()))
}

fn value_size(value: &SemanticValue) -> usize {
    match value {
        SemanticValue::Text(value) | SemanticValue::Xml(value) => {
            value.as_ref().map_or(0, String::len)
        },
        SemanticValue::Unsigned(_) => size_of::<u32>(),
        SemanticValue::Source(Source::Linked(value)) => String::len(value),
        SemanticValue::Source(Source::Embedded(value)) => Vec::len(value),
        SemanticValue::ImageMap(value) => value.as_ref().map_or(0, |map| map.areas().len() * 128),
        SemanticValue::Resource(value) => value.as_ref().map_or(0, |resource| {
            resource.media_type.len() + resource.bytes.len()
        }),
    }
}

fn validate_map(map: &ImageMap, policy: &SecurityPolicy) -> Result<()> {
    if map.areas().len() > policy.max_map_areas {
        return Err(invalid("ODI image-map area count exceeds policy"));
    }
    for area in map.areas() {
        for value in [
            area.href(),
            area.target_frame_name(),
            area.name(),
            area.link_type(),
            area.show(),
            area.title(),
            area.description(),
        ]
        .into_iter()
        .flatten()
        {
            if value.len() > policy.max_text_bytes {
                return Err(invalid("ODI image-map text exceeds policy"));
            }
        }
        if !policy.allow_external_links && area.href().is_some_and(is_external) {
            return Err(invalid("ODI image-map external link is refused by policy"));
        }
        let geometry = match area.kind() {
            crate::map::AreaKind::Rectangle {
                x,
                y,
                width,
                height,
            }
            | crate::map::AreaKind::Polygon {
                x,
                y,
                width,
                height,
                ..
            } => [x.as_str(), y.as_str(), width.as_str(), height.as_str()],
            crate::map::AreaKind::Circle {
                center_x,
                center_y,
                radius,
            } => [center_x.as_str(), center_y.as_str(), radius.as_str(), ""],
        };
        if geometry
            .into_iter()
            .any(|value| value.len() > policy.max_text_bytes)
        {
            return Err(invalid("ODI image-map geometry exceeds text policy"));
        }
        if let crate::map::AreaKind::Polygon {
            view_box, points, ..
        } = area.kind()
            && (view_box.len() > policy.max_text_bytes || points.len() > policy.max_text_bytes)
        {
            return Err(invalid("ODI image-map geometry exceeds text policy"));
        }
    }
    Ok(())
}

fn is_external(href: &str) -> bool {
    href.starts_with('/')
        || href.starts_with("//")
        || href.find(':').is_some_and(|colon| {
            let scheme = &href[..colon];
            scheme
                .as_bytes()
                .first()
                .is_some_and(u8::is_ascii_alphabetic)
                && scheme
                    .bytes()
                    .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
        })
}

fn artifact_conflict(kind: ConflictKind) -> Conflict {
    Conflict {
        key: None,
        kind,
        expected: None,
        actual: None,
        desired: None,
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
