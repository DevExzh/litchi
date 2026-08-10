//! Deterministic semantic patches, transfer, and non-mutating merge planning.

use crate::{
    Commit, Edit, FlatImage, FlatImageCommit, FrameChange, Image, MetadataFields, ResourceChange,
    frame::Frame, map::ImageMap, source::Source,
};
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    collections::{BTreeMap, BTreeSet},
    sync::Arc,
};

const MIB: usize = 1024 * 1024;
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_STYLE_CLOSURE: usize = 1_024;

/// Whether a security-sensitive semantic capability is available.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CapabilityState {
    Allowed,
    Refused,
}

/// Security-sensitive capabilities selected by a [`SecurityPolicy`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityCapabilities {
    external_links: CapabilityState,
    package_members: CapabilityState,
    active_content: CapabilityState,
}

impl SecurityCapabilities {
    /// Returns whether external URI values may enter a semantic plan.
    #[must_use]
    pub const fn external_links(self) -> CapabilityState {
        self.external_links
    }

    /// Returns whether package XML and resource operations may enter a plan.
    #[must_use]
    pub const fn package_members(self) -> CapabilityState {
        self.package_members
    }

    /// Returns whether operations containing active-content constructs are accepted.
    #[must_use]
    pub const fn active_content(self) -> CapabilityState {
        self.active_content
    }
}

/// Security-relevant package conditions that prevent source rewriting.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum RewriteBlocker {
    Signature,
    Encryption,
    NonCompactXml,
    UnreadableXml,
}

/// Explicit rewrite capability for one immutable package snapshot.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct RewriteCapability {
    blockers: Vec<RewriteBlocker>,
}

impl RewriteCapability {
    pub(crate) fn new(mut blockers: Vec<RewriteBlocker>) -> Self {
        blockers.sort_unstable();
        blockers.dedup();
        Self { blockers }
    }

    /// Returns whether changed-byte publication is currently supported.
    #[must_use]
    pub fn is_ready(&self) -> bool {
        self.blockers.is_empty()
    }

    /// Returns every independently detected rewrite blocker.
    #[must_use]
    pub fn blockers(&self) -> &[RewriteBlocker] {
        &self.blockers
    }
}

/// Current lifecycle state of a non-mutating semantic plan.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PublicationState {
    Ready,
    ConflictRefused,
    ActiveContentRefused,
    PolicyRefused,
}

/// Lifecycle state of one exact style name/family dependency.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StyleDependencyState {
    Absent,
    Named,
    Automatic,
    EquivalentDuplicate,
    Collision,
}

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
    allow_active_content: bool,
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
            allow_active_content: false,
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

    /// Sets whether semantic operations may carry active-content constructs.
    #[must_use]
    pub const fn with_active_content(mut self, allow: bool) -> Self {
        self.allow_active_content = allow;
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

    /// Returns explicit states for security-sensitive capabilities.
    #[must_use]
    pub const fn capabilities(&self) -> SecurityCapabilities {
        SecurityCapabilities {
            external_links: if self.allow_external_links {
                CapabilityState::Allowed
            } else {
                CapabilityState::Refused
            },
            package_members: if self.allow_package_members {
                CapabilityState::Allowed
            } else {
                CapabilityState::Refused
            },
            active_content: if self.allow_active_content {
                CapabilityState::Allowed
            } else {
                CapabilityState::Refused
            },
        }
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
            allow_active_content: true,
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
    /// A referenced style or package member cannot be supplied losslessly.
    MissingDependency,
    /// A dependency key exists with different bytes or namespace bindings.
    DependencyCollision,
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

    /// Classifies the plan before any publication is attempted.
    #[must_use]
    pub fn publication_state(&self, policy: &SecurityPolicy) -> PublicationState {
        if !self.conflicts.is_empty() {
            PublicationState::ConflictRefused
        } else if !policy.allow_active_content
            && self
                .operations
                .iter()
                .any(|operation| operation_contains_active_content(operation).unwrap_or(true))
        {
            PublicationState::ActiveContentRefused
        } else if validate_operations(&self.operations, policy).is_err() {
            PublicationState::PolicyRefused
        } else {
            PublicationState::Ready
        }
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
        let mut plan = plan_operations(&self.operations, |key| package_value(current, key));
        complete_package_dependencies(self, current, &mut plan);
        plan.operations
            .sort_unstable_by(|left, right| left.key.cmp(&right.key));
        plan.conflicts
            .sort_by(|left, right| left.key.cmp(&right.key));
        plan
    }
}

fn complete_package_dependencies(patch: &SemanticPatch, current: &Image, plan: &mut SemanticPlan) {
    let target_package = if patch.kind == ArtifactKind::Package {
        Image::from_bytes(patch.target.as_ref().clone()).ok()
    } else {
        None
    };
    for operation in &patch.operations {
        let OperationKey::Frame { frame, property } = operation.key() else {
            continue;
        };
        match (property, operation.after()) {
            (FrameProperty::Source, SemanticValue::Source(Source::Linked(href))) => {
                complete_resource_dependency(
                    patch,
                    current,
                    target_package.as_ref(),
                    *frame,
                    href,
                    operation.key(),
                    plan,
                );
            },
            (FrameProperty::StyleName, SemanticValue::Text(Some(name))) => {
                complete_style_dependency(
                    current,
                    target_package.as_ref(),
                    name,
                    "graphic",
                    operation.key(),
                    plan,
                );
            },
            (FrameProperty::TextStyleName, SemanticValue::Text(Some(name))) => {
                complete_style_dependency(
                    current,
                    target_package.as_ref(),
                    name,
                    "paragraph",
                    operation.key(),
                    plan,
                );
            },
            _ => {},
        }
    }
}

fn complete_resource_dependency(
    semantic_delta: &SemanticPatch,
    current: &Image,
    target: Option<&Image>,
    frame: usize,
    href: &str,
    frame_key: &OperationKey,
    plan: &mut SemanticPlan,
) {
    let target_resource = target.and_then(|image| {
        image
            .resources()
            .iter()
            .find(|resource| resource.frame() == frame && resource.href() == href)
    });
    let path = target_resource
        .map(crate::resource::Resource::path)
        .map(str::to_owned)
        .or_else(|| local_package_path(href));
    let Some(path) = path else {
        return;
    };
    let key = OperationKey::Resource(path.clone());
    if let Some(explicit) = semantic_delta
        .operations
        .iter()
        .find(|item| item.key() == &key)
    {
        if matches!(explicit.after(), SemanticValue::Resource(Some(_))) {
            return;
        }
        push_missing_dependency(plan, frame_key, explicit.after().clone());
        return;
    }
    let current_value = resource_value(current, &path).ok();
    let desired = target.and_then(|image| resource_value(image, &path).ok());
    match (current_value, desired) {
        (Some(actual), Some(expected)) if actual == expected => {},
        (
            Some(SemanticValue::Resource(None)),
            Some(expected @ SemanticValue::Resource(Some(_))),
        ) => {
            plan.operations.push(SemanticOperation {
                key,
                before: SemanticValue::Resource(None),
                after: expected,
            });
        },
        (Some(actual), Some(expected)) => plan.conflicts.push(Conflict {
            key: Some(key),
            kind: ConflictKind::MissingDependency,
            expected: None,
            actual: Some(actual),
            desired: Some(expected),
        }),
        (Some(SemanticValue::Resource(Some(_))), None) => {},
        (Some(actual), None) => plan.conflicts.push(Conflict {
            key: Some(key),
            kind: ConflictKind::MissingDependency,
            expected: None,
            actual: Some(actual),
            desired: Some(SemanticValue::Text(Some(href.to_owned()))),
        }),
        (None, _) => {
            push_missing_dependency(plan, frame_key, SemanticValue::Text(Some(href.into())));
        },
    }
}

fn complete_style_dependency(
    current: &Image,
    target: Option<&Image>,
    name: &str,
    family: &str,
    frame_key: &OperationKey,
    plan: &mut SemanticPlan,
) {
    let Some(target) = target else {
        match inspect_style_dependency(current.content_xml(), current.styles_xml(), name, family) {
            Ok(
                StyleDependencyState::Named
                | StyleDependencyState::Automatic
                | StyleDependencyState::EquivalentDuplicate,
            ) => {},
            Ok(StyleDependencyState::Collision) => plan.conflicts.push(Conflict {
                key: Some(frame_key.clone()),
                kind: ConflictKind::DependencyCollision,
                expected: None,
                actual: Some(SemanticValue::Text(Some(name.to_owned()))),
                desired: Some(SemanticValue::Text(Some(name.to_owned()))),
            }),
            Ok(StyleDependencyState::Absent) | Err(_) => {
                push_missing_dependency(
                    plan,
                    frame_key,
                    SemanticValue::Text(Some(name.to_owned())),
                );
            },
        }
        return;
    };
    let Ok(closure) = collect_style_closure(target, name, family) else {
        push_missing_dependency(plan, frame_key, SemanticValue::Text(Some(name.to_owned())));
        return;
    };
    for desired in &closure {
        match merge_style_dependency(current, plan, desired) {
            Ok(()) => {},
            Err(actual) => {
                plan.conflicts.push(Conflict {
                    key: Some(frame_key.clone()),
                    kind: ConflictKind::DependencyCollision,
                    expected: None,
                    actual: Some(actual),
                    desired: Some(SemanticValue::Xml(Some(desired.xml.clone()))),
                });
                return;
            },
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct StyleKey {
    name: String,
    family: String,
}

struct TransferStyleDefinition {
    key: StyleKey,
    container: StyleContainer,
    xml: String,
    namespaces: BTreeMap<String, String>,
    declaration: Option<String>,
}

fn collect_style_closure(
    source: &Image,
    name: &str,
    family: &str,
) -> Result<Vec<TransferStyleDefinition>> {
    let mut pending = vec![StyleKey {
        name: name.to_owned(),
        family: family.to_owned(),
    }];
    let mut seen = BTreeSet::new();
    let mut closure = Vec::new();
    while let Some(key) = pending.pop() {
        if !seen.insert(key.clone()) {
            continue;
        }
        if seen.len() > MAX_STYLE_CLOSURE {
            return Err(invalid(
                "ODI transitive style closure exceeds the item limit",
            ));
        }
        let definition = source_style_definition(source, &key)?;
        let mut dependencies = style_dependencies(&definition)?;
        dependencies.reverse();
        pending.extend(dependencies);
        closure.push(definition);
    }
    Ok(closure)
}

fn source_style_definition(source: &Image, key: &StyleKey) -> Result<TransferStyleDefinition> {
    let content = scan_style_document(source.content_xml(), &key.name, &key.family)?;
    let styles = source
        .styles_xml()
        .map(|xml| scan_style_document(xml, &key.name, &key.family))
        .transpose()?;
    let content_definition = content.definition.as_ref();
    let styles_definition = styles
        .as_ref()
        .and_then(|document| document.definition.as_ref());
    let (definition, namespaces, document_xml) = match (content_definition, styles_definition) {
        (None, None) => return Err(invalid("ODI transitive style dependency is missing")),
        (Some(definition), None) => (definition, &content.namespaces, source.content_xml()),
        (None, Some(definition)) => (
            definition,
            &styles
                .as_ref()
                .ok_or_else(|| invalid("ODI style inventory disappeared"))?
                .namespaces,
            source
                .styles_xml()
                .ok_or_else(|| invalid("ODI styles part disappeared"))?,
        ),
        (Some(left), Some(right)) if definitions_match(left, right) => (
            right,
            &styles
                .as_ref()
                .ok_or_else(|| invalid("ODI style inventory disappeared"))?
                .namespaces,
            source
                .styles_xml()
                .ok_or_else(|| invalid("ODI styles part disappeared"))?,
        ),
        (Some(_), Some(_)) => {
            return Err(invalid(
                "ODI transitive style dependency collides across parts",
            ));
        },
    };
    Ok(TransferStyleDefinition {
        key: key.clone(),
        container: definition.container,
        xml: definition.xml.to_owned(),
        namespaces: namespaces.clone(),
        declaration: leading_xml_declaration(document_xml).map(str::to_owned),
    })
}

fn leading_xml_declaration(xml: &str) -> Option<&str> {
    xml.strip_prefix("<?xml")
        .and_then(|suffix| suffix.find("?>").map(|end| &xml[..end.saturating_add(7)]))
}

fn style_dependencies(definition: &TransferStyleDefinition) -> Result<Vec<StyleKey>> {
    let mut reader = NsReader::from_reader(definition.xml.as_bytes());
    loop {
        let (_, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODI style dependency fragment: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let values = inherited_style_reference_attributes(
                    &reader,
                    &element,
                    &definition.namespaces,
                )?;
                let mut dependencies = Vec::new();
                for (kind, name) in values {
                    let family = if kind == b"linked-style-name" {
                        match definition.key.family.as_str() {
                            "paragraph" => "text",
                            "text" => "paragraph",
                            value => value,
                        }
                    } else {
                        &definition.key.family
                    };
                    dependencies.push(StyleKey {
                        name,
                        family: family.to_owned(),
                    });
                }
                return Ok(dependencies);
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI style XML")),
            Event::Eof => return Err(invalid("ODI style dependency fragment is empty")),
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn inherited_style_reference_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespaces: &BTreeMap<String, String>,
) -> Result<Vec<(&'static [u8], String)>> {
    let mut values = Vec::new();
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI style attribute: {error}")))?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("ODI style attribute name is not UTF-8: {error}")))?;
        let Some((prefix, local)) = key.split_once(':') else {
            continue;
        };
        if namespaces.get(prefix).map(String::as_bytes) != Some(STYLE_NAMESPACE) {
            continue;
        }
        let kind = match local.as_bytes() {
            b"parent-style-name" => b"parent-style-name".as_slice(),
            b"next-style-name" => b"next-style-name".as_slice(),
            b"linked-style-name" => b"linked-style-name".as_slice(),
            _ => continue,
        };
        if values.iter().any(|(present, _)| *present == kind) {
            return Err(invalid("duplicate ODI transitive style reference"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid ODI style reference value: {error}")))?
            .into_owned();
        values.push((kind, value));
    }
    Ok(values)
}

fn merge_style_dependency(
    current: &Image,
    plan: &mut SemanticPlan,
    desired: &TransferStyleDefinition,
) -> std::result::Result<(), SemanticValue> {
    let content = scan_style_document(
        current.content_xml(),
        &desired.key.name,
        &desired.key.family,
    )
    .map_err(|error| SemanticValue::Text(Some(error.to_string())))?;
    if let Some(actual) = content.definition.as_ref() {
        return if definition_matches_transfer(actual, desired) {
            Ok(())
        } else {
            Err(SemanticValue::Xml(Some(actual.xml.to_owned())))
        };
    }
    let office_version = content.office_version.as_deref().unwrap_or("1.4");
    let base = planned_styles_xml(current, plan).map(str::to_owned);
    if let Some(base) = base.as_deref() {
        let inventory = scan_style_document(base, &desired.key.name, &desired.key.family)
            .map_err(|error| SemanticValue::Text(Some(error.to_string())))?;
        if let Some(actual) = inventory.definition.as_ref() {
            return if definition_matches_transfer(actual, desired) {
                Ok(())
            } else {
                Err(SemanticValue::Xml(Some(actual.xml.to_owned())))
            };
        }
        let merged = merge_style(base, &inventory, desired)
            .map_err(|error| SemanticValue::Text(Some(error.to_string())))?;
        set_planned_styles(current, plan, merged);
    } else {
        let created = new_style_document(desired, office_version);
        set_planned_styles(current, plan, created);
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum StyleContainer {
    Named,
    Automatic,
}

pub(crate) fn inspect_style_dependency(
    content_xml: &str,
    styles_xml: Option<&str>,
    name: &str,
    family: &str,
) -> Result<StyleDependencyState> {
    let content = scan_style_document(content_xml, name, family)?;
    let styles = styles_xml
        .map(|xml| scan_style_document(xml, name, family))
        .transpose()?;
    match (
        content.definition.as_ref(),
        styles
            .as_ref()
            .and_then(|document| document.definition.as_ref()),
    ) {
        (None, None) => Ok(StyleDependencyState::Absent),
        (Some(definition), None) | (None, Some(definition)) => Ok(match definition.container {
            StyleContainer::Named => StyleDependencyState::Named,
            StyleContainer::Automatic => StyleDependencyState::Automatic,
        }),
        (Some(left), Some(right)) if definitions_match(left, right) => {
            Ok(StyleDependencyState::EquivalentDuplicate)
        },
        (Some(_), Some(_)) => Ok(StyleDependencyState::Collision),
    }
}

#[derive(Clone, Copy)]
struct StyleDefinition<'a> {
    container: StyleContainer,
    xml: &'a str,
}

struct StyleDocument<'a> {
    definition: Option<StyleDefinition<'a>>,
    named_close: Option<usize>,
    automatic_close: Option<usize>,
    root_close: Option<usize>,
    office_version: Option<String>,
    namespaces: BTreeMap<String, String>,
}

fn scan_style_document<'a>(xml: &'a str, name: &str, family: &str) -> Result<StyleDocument<'a>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut container = None::<(usize, StyleContainer)>;
    let mut active = None::<(usize, usize, StyleContainer)>;
    let mut definition = None;
    let mut named_close = None;
    let mut automatic_close = None;
    let mut root_close = None;
    let mut office_version = None;
    let mut namespaces = BTreeMap::new();
    let mut buffer = Vec::new();
    loop {
        buffer.clear();
        let start = usize::try_from(reader.buffer_position())
            .map_err(|error| invalid(format!("ODI style position exceeds usize: {error}")))?;
        let (is_office, is_style, event) = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid(format!("invalid ODI style dependency XML: {error}")))?;
            (
                bound_to(&namespace, OFFICE_NAMESPACE),
                bound_to(&namespace, STYLE_NAMESPACE),
                event,
            )
        };
        let end = usize::try_from(reader.buffer_position())
            .map_err(|error| invalid(format!("ODI style position exceeds usize: {error}")))?;
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("ODI style dependency depth overflow"))?;
                if depth == 1 {
                    namespaces = namespace_declarations(&reader, &element)?;
                    office_version =
                        namespaced_attribute(&reader, &element, OFFICE_NAMESPACE, b"version")?;
                }
                if is_office {
                    let kind = match element.local_name().as_ref() {
                        b"styles" => Some(StyleContainer::Named),
                        b"automatic-styles" => Some(StyleContainer::Automatic),
                        _ => None,
                    };
                    if let Some(kind) = kind
                        && depth == 2
                    {
                        if container.is_some() {
                            return Err(invalid("nested ODI style containers are not allowed"));
                        }
                        container = Some((depth, kind));
                    }
                } else if is_style
                    && element.local_name().as_ref() == b"style"
                    && container.is_some_and(|(owner, _)| owner + 1 == depth)
                    && style_key(&reader, &element)?
                        == (Some(name.to_owned()), Some(family.to_owned()))
                {
                    let kind = container
                        .map(|(_, kind)| kind)
                        .ok_or_else(|| invalid("ODI style owner disappeared"))?;
                    active = Some((start, depth, kind));
                }
            },
            Event::Empty(element)
                if is_style
                    && element.local_name().as_ref() == b"style"
                    && container.is_some_and(|(owner, _)| owner == depth)
                    && style_key(&reader, &element)?
                        == (Some(name.to_owned()), Some(family.to_owned())) =>
            {
                let kind = container
                    .map(|(_, kind)| kind)
                    .ok_or_else(|| invalid("ODI style owner disappeared"))?;
                record_style_definition(
                    &mut definition,
                    StyleDefinition {
                        container: kind,
                        xml: xml
                            .get(start..end)
                            .ok_or_else(|| invalid("ODI style range is invalid"))?,
                    },
                )?;
            },
            Event::End(_) => {
                if active.is_some_and(|(_, owner, _)| owner == depth) {
                    let (definition_start, _, kind) = active
                        .take()
                        .ok_or_else(|| invalid("ODI style definition state disappeared"))?;
                    record_style_definition(
                        &mut definition,
                        StyleDefinition {
                            container: kind,
                            xml: xml
                                .get(definition_start..end)
                                .ok_or_else(|| invalid("ODI style range is invalid"))?,
                        },
                    )?;
                }
                if let Some((owner, kind)) = container
                    && owner == depth
                {
                    match kind {
                        StyleContainer::Named => named_close = Some(start),
                        StyleContainer::Automatic => automatic_close = Some(start),
                    }
                    container = None;
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODI style dependency depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI style XML")),
            Event::Eof => {
                if depth != 0 || active.is_some() || container.is_some() {
                    return Err(invalid("unterminated ODI style dependency XML"));
                }
                return Ok(StyleDocument {
                    definition,
                    named_close,
                    automatic_close,
                    root_close,
                    office_version,
                    namespaces,
                });
            },
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn style_key(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<(Option<String>, Option<String>)> {
    let mut name = None;
    let mut family = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI style attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if !bound_to(&resolved, STYLE_NAMESPACE) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid ODI style value: {error}")))?
            .into_owned();
        match local.as_ref() {
            b"name" => name = Some(value),
            b"family" => family = Some(value),
            _ => {},
        }
    }
    Ok((name, family))
}

fn namespaced_attribute(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut result = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI style attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if bound_to(&resolved, namespace) && name.as_ref() == local {
            if result.is_some() {
                return Err(invalid("duplicate expanded ODI style attribute"));
            }
            result = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| invalid(format!("invalid ODI style value: {error}")))?
                    .into_owned(),
            );
        }
    }
    Ok(result)
}

fn namespace_declarations(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<BTreeMap<String, String>> {
    let mut declarations = BTreeMap::new();
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI namespace declaration: {error}")))?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("ODI namespace name is not UTF-8: {error}")))?;
        let prefix = if key == "xmlns" {
            Some("")
        } else {
            key.strip_prefix("xmlns:")
        };
        let Some(prefix) = prefix else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid ODI namespace value: {error}")))?
            .into_owned();
        if declarations.insert(prefix.to_owned(), value).is_some() {
            return Err(invalid("duplicate ODI root namespace declaration"));
        }
    }
    Ok(declarations)
}

fn record_style_definition<'a>(
    slot: &mut Option<StyleDefinition<'a>>,
    definition: StyleDefinition<'a>,
) -> Result<()> {
    if slot.replace(definition).is_some() {
        return Err(invalid("duplicate ODI style dependency key"));
    }
    Ok(())
}

fn definitions_match(left: &StyleDefinition<'_>, right: &StyleDefinition<'_>) -> bool {
    left.container == right.container && left.xml == right.xml
}

fn definition_matches_transfer(
    actual: &StyleDefinition<'_>,
    desired: &TransferStyleDefinition,
) -> bool {
    actual.container == desired.container && actual.xml == desired.xml
}

fn planned_styles_xml<'a>(current: &'a Image, plan: &'a SemanticPlan) -> Option<&'a str> {
    plan.operations
        .iter()
        .find(|operation| operation.key() == &OperationKey::Styles)
        .and_then(|operation| match operation.after() {
            SemanticValue::Xml(value) => Some(value.as_deref()),
            SemanticValue::Text(_)
            | SemanticValue::Unsigned(_)
            | SemanticValue::Source(_)
            | SemanticValue::ImageMap(_)
            | SemanticValue::Resource(_) => None,
        })
        .unwrap_or_else(|| current.styles_xml())
}

fn merge_style(
    base: &str,
    destination_inventory: &StyleDocument<'_>,
    definition: &TransferStyleDefinition,
) -> Result<String> {
    for prefix in used_prefixes(&definition.xml)? {
        if prefix == "xml" {
            continue;
        }
        let source_binding = definition.namespaces.get(&prefix);
        let destination_binding = destination_inventory.namespaces.get(&prefix);
        let compatible = if prefix.is_empty() {
            source_binding == destination_binding
        } else {
            source_binding.is_some() && source_binding == destination_binding
        };
        if !compatible {
            return Err(invalid(
                "ODI granular style merge has a namespace collision",
            ));
        }
    }
    let close = match definition.container {
        StyleContainer::Named => destination_inventory.named_close,
        StyleContainer::Automatic => destination_inventory.automatic_close,
    };
    let insertion = if close.is_some() {
        definition.xml.clone()
    } else {
        let root_close = destination_inventory
            .root_close
            .ok_or_else(|| invalid("ODI destination has no style-document close site"))?;
        let office_prefix = office_prefix(&destination_inventory.namespaces)?;
        let local = match definition.container {
            StyleContainer::Named => "styles",
            StyleContainer::Automatic => "automatic-styles",
        };
        return insert_style_container(base, root_close, office_prefix, local, &definition.xml);
    };
    let close = close.ok_or_else(|| invalid("ODI style merge site disappeared"))?;
    let mut merged = String::with_capacity(base.len().saturating_add(insertion.len()));
    merged.push_str(
        base.get(..close)
            .ok_or_else(|| invalid("ODI style insertion range is invalid"))?,
    );
    merged.push_str(&insertion);
    merged.push_str(
        base.get(close..)
            .ok_or_else(|| invalid("ODI style insertion range is invalid"))?,
    );
    Ok(merged)
}

fn insert_style_container(
    base: &str,
    root_close: usize,
    office_prefix: &str,
    local: &str,
    definition: &str,
) -> Result<String> {
    let mut merged = String::with_capacity(
        base.len()
            .saturating_add(definition.len())
            .saturating_add(64),
    );
    merged.push_str(
        base.get(..root_close)
            .ok_or_else(|| invalid("ODI style insertion range is invalid"))?,
    );
    merged.push('<');
    merged.push_str(office_prefix);
    merged.push(':');
    merged.push_str(local);
    merged.push('>');
    merged.push_str(definition);
    merged.push_str("</");
    merged.push_str(office_prefix);
    merged.push(':');
    merged.push_str(local);
    merged.push('>');
    merged.push_str(
        base.get(root_close..)
            .ok_or_else(|| invalid("ODI style insertion range is invalid"))?,
    );
    Ok(merged)
}

fn new_style_document(definition: &TransferStyleDefinition, office_version: &str) -> String {
    let mut namespaces = definition.namespaces.clone();
    let office_prefix = ensure_office_prefix(&mut namespaces);
    let mut xml = definition.declaration.clone().unwrap_or_default();
    xml.push('<');
    xml.push_str(&office_prefix);
    xml.push_str(":document-styles");
    for (prefix, namespace) in &namespaces {
        xml.push_str(" xmlns");
        if !prefix.is_empty() {
            xml.push(':');
            xml.push_str(prefix);
        }
        xml.push_str("=\"");
        push_escaped_attribute(&mut xml, namespace);
        xml.push('"');
    }
    xml.push(' ');
    xml.push_str(&office_prefix);
    xml.push_str(":version=\"");
    push_escaped_attribute(&mut xml, office_version);
    xml.push_str("\">");
    let local = match definition.container {
        StyleContainer::Named => "styles",
        StyleContainer::Automatic => "automatic-styles",
    };
    xml.push('<');
    xml.push_str(&office_prefix);
    xml.push(':');
    xml.push_str(local);
    xml.push('>');
    xml.push_str(&definition.xml);
    xml.push_str("</");
    xml.push_str(&office_prefix);
    xml.push(':');
    xml.push_str(local);
    xml.push_str("></");
    xml.push_str(&office_prefix);
    xml.push_str(":document-styles>");
    xml
}

fn ensure_office_prefix(namespaces: &mut BTreeMap<String, String>) -> String {
    let office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    if let Some((prefix, _)) = namespaces
        .iter()
        .find(|(prefix, value)| !prefix.is_empty() && value.as_str() == office)
    {
        return prefix.clone();
    }
    let mut prefix = "office".to_owned();
    let mut suffix = 0usize;
    while namespaces.contains_key(&prefix) {
        suffix = suffix.saturating_add(1);
        prefix = format!("odi-office-{suffix}");
    }
    namespaces.insert(prefix.clone(), office.to_owned());
    prefix
}

fn office_prefix(namespaces: &BTreeMap<String, String>) -> Result<&str> {
    let office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    namespaces
        .iter()
        .find(|(prefix, value)| !prefix.is_empty() && value.as_str() == office)
        .map(|(prefix, _)| prefix.as_str())
        .ok_or_else(|| invalid("ODI destination has no prefixed office namespace binding"))
}

fn push_escaped_attribute(target: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => target.push_str("&amp;"),
            '<' => target.push_str("&lt;"),
            '"' => target.push_str("&quot;"),
            _ => target.push(character),
        }
    }
}

fn used_prefixes(xml: &str) -> Result<Vec<String>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut prefixes = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODI style fragment: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                push_qname_prefix(&mut prefixes, element.name().as_ref(), true)?;
                for raw in element.attributes() {
                    let attribute = raw.map_err(|error| {
                        invalid(format!("invalid ODI style fragment attribute: {error}"))
                    })?;
                    let key = attribute.key.as_ref();
                    if key != b"xmlns" && !key.starts_with(b"xmlns:") {
                        push_qname_prefix(&mut prefixes, attribute.key.as_ref(), false)?;
                    }
                }
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI style XML")),
            Event::Eof => {
                prefixes.sort_unstable();
                prefixes.dedup();
                return Ok(prefixes);
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn push_qname_prefix(target: &mut Vec<String>, qname: &[u8], element: bool) -> Result<()> {
    let name = std::str::from_utf8(qname)
        .map_err(|error| invalid(format!("ODI style QName is not UTF-8: {error}")))?;
    if let Some((prefix, _)) = name.split_once(':') {
        target.push(prefix.to_owned());
    } else if element {
        target.push(String::new());
    }
    Ok(())
}

fn set_planned_styles(current: &Image, plan: &mut SemanticPlan, merged: String) {
    if let Some(operation) = plan
        .operations
        .iter_mut()
        .find(|operation| operation.key() == &OperationKey::Styles)
    {
        operation.after = SemanticValue::Xml(Some(merged));
    } else {
        plan.operations.push(SemanticOperation {
            key: OperationKey::Styles,
            before: SemanticValue::Xml(current.styles_xml().map(str::to_owned)),
            after: SemanticValue::Xml(Some(merged)),
        });
    }
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn local_package_path(href: &str) -> Option<String> {
    let path = href.strip_prefix("./").unwrap_or(href);
    (!path.is_empty()
        && !is_external(path)
        && !path.starts_with('#')
        && !path
            .chars()
            .any(|character| matches!(character, '?' | '#' | '\\'))
        && path
            .split('/')
            .all(|part| !part.is_empty() && !matches!(part, "." | "..")))
    .then(|| path.to_owned())
}

fn push_missing_dependency(plan: &mut SemanticPlan, key: &OperationKey, desired: SemanticValue) {
    if plan.conflicts.iter().any(|conflict| {
        conflict.kind == ConflictKind::MissingDependency && conflict.key.as_ref() == Some(key)
    }) {
        return;
    }
    plan.conflicts.push(Conflict {
        key: Some(key.clone()),
        kind: ConflictKind::MissingDependency,
        expected: None,
        actual: None,
        desired: Some(desired),
    });
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
    let operation_bytes = operations.iter().try_fold(0usize, |total, operation| {
        total.checked_add(operation_size(operation))
    });
    if operation_bytes.is_none_or(|bytes| bytes > policy.max_patch_bytes) {
        return Err(invalid("ODI semantic operation bytes exceed patch policy"));
    }
    for operation in operations {
        if !policy.allow_active_content && operation_contains_active_content(operation)? {
            return Err(invalid("ODI active content is refused by policy"));
        }
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

fn operation_contains_active_content(operation: &SemanticOperation) -> Result<bool> {
    if let OperationKey::Resource(path) = operation.key()
        && crate::active::is_package_script_member(path)
        && matches!(operation.after(), SemanticValue::Resource(Some(_)))
    {
        return Ok(true);
    }
    if let SemanticValue::Xml(Some(xml)) = operation.after()
        && !crate::active::scan_xml(xml, crate::active::ActiveContentLocation::StylesXml)?
            .is_empty()
    {
        return Ok(true);
    }
    Ok(false)
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
            area.actuate(),
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

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

    use super::*;

    const EMPTY: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"/>"#;
    const NAMED: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles><style:style style:name="gr1" style:family="graphic"/></office:styles></office:document-styles>"#;
    const AUTOMATIC: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:automatic-styles><style:style style:name="gr1" style:family="graphic"/></office:automatic-styles></office:document-styles>"#;
    const COLLIDING: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles><style:style style:name="gr1" style:family="graphic" style:display-name="Other"/></office:styles></office:document-styles>"#;
    const ACTIVE: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"><office:script script:language="python"/></office:document-styles>"#;

    #[test]
    fn style_dependency_states_cover_the_exact_cross_part_lifecycle() {
        assert_eq!(
            inspect_style_dependency(EMPTY, None, "gr1", "graphic").unwrap(),
            StyleDependencyState::Absent
        );
        assert_eq!(
            inspect_style_dependency(NAMED, None, "gr1", "graphic").unwrap(),
            StyleDependencyState::Named
        );
        assert_eq!(
            inspect_style_dependency(AUTOMATIC, None, "gr1", "graphic").unwrap(),
            StyleDependencyState::Automatic
        );
        assert_eq!(
            inspect_style_dependency(AUTOMATIC, Some(AUTOMATIC), "gr1", "graphic").unwrap(),
            StyleDependencyState::EquivalentDuplicate
        );
        assert_eq!(
            inspect_style_dependency(NAMED, Some(COLLIDING), "gr1", "graphic").unwrap(),
            StyleDependencyState::Collision
        );
    }

    #[test]
    fn strict_active_content_lifecycle_refuses_introduction_but_allows_removal() {
        let introduction = SemanticOperation {
            key: OperationKey::Styles,
            before: SemanticValue::Xml(None),
            after: SemanticValue::Xml(Some(ACTIVE.to_owned())),
        };
        let removal = introduction.inverse();
        let introducing_plan = SemanticPlan {
            operations: vec![introduction],
            conflicts: Vec::new(),
        };
        let removal_plan = SemanticPlan {
            operations: vec![removal],
            conflicts: Vec::new(),
        };
        assert_eq!(
            SecurityPolicy::strict().capabilities().active_content(),
            CapabilityState::Refused
        );
        assert_eq!(
            introducing_plan.publication_state(&SecurityPolicy::strict()),
            PublicationState::ActiveContentRefused
        );
        assert_eq!(
            introducing_plan.publication_state(&SecurityPolicy::default()),
            PublicationState::Ready
        );
        assert_eq!(
            removal_plan.publication_state(&SecurityPolicy::strict()),
            PublicationState::Ready
        );
    }
}
