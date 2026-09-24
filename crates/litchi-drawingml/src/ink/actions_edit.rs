//! Detached construction and source-backed editing for the strict ink-action
//! profile.
//!
//! This module deliberately sits below the package owners.  It emits and
//! edits one `iact:actions` value; it does not know about a PresentationML
//! relationship, an OPC part name, or action execution.  Fresh values use a
//! small canonical namespace policy.  Edits operate on the retained source
//! spans from [`super::Profile`], so untouched declarations, comments, and
//! opaque InkML payloads remain byte-for-byte unchanged.

use std::{
    collections::{BinaryHeap, HashMap, HashSet},
    sync::Arc,
};

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::{SeenNames, first_wins};
use litchi_ooxml_common::xml_name::{is_ncname, is_qualified_name};
use quick_xml::{
    Reader, XmlVersion,
    events::{BytesRef, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use crate::{Error, Result};

use super::{
    ACTION_NAMESPACE, Action, ActionChild, ActionData, ActionDataGroup, ActionGroup,
    ActionProperty, ActionType, DataChild, LengthUnit, Profile, RootChild, SourceSpan, TimeUnit,
    read_profile, read_profile_owned,
};

impl ActionType {
    /// Construct a checked custom action type.
    ///
    /// Empty `ST_ActionTypeUser` values are allowed by the unrestricted
    /// `xsd:string` branch and remain source-preserving through the strict
    /// profile.
    pub fn custom(value: impl AsRef<str>) -> Result<Self> {
        Ok(Self::Custom(scalar(value.as_ref(), "action type")?))
    }
}

const MAX_EDIT_OPERATIONS: usize = 65_536;
const MAX_SCALAR_BYTES: usize = 256;
const MAX_PAYLOAD_BYTES: usize = super::MAX_SOURCE_BYTES;

/// Limits for one detached construction or source-backed edit.
///
/// The limits are caller-facing budgets below the strict reader's hard
/// ceilings.  An edit checks the complete resulting source before allocating
/// its replacement buffer.  Opaque payload bytes are bounded here and are
/// still checked again by `read_profile` when the value is finished.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Limits {
    /// Maximum complete output/source bytes.
    pub max_output_bytes: usize,
    /// Maximum one opaque payload byte length.
    pub max_payload_bytes: usize,
    /// Maximum XML nodes in the authored structural projection.
    pub max_nodes: usize,
    /// Maximum authored depth.
    pub max_depth: usize,
    /// Maximum action records.
    pub max_actions: usize,
    /// Maximum action-group records.
    pub max_action_groups: usize,
    /// Maximum bytes in a typed scalar or identifier.
    pub max_scalar_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_output_bytes: super::MAX_SOURCE_BYTES,
            max_payload_bytes: MAX_PAYLOAD_BYTES,
            max_nodes: super::MAX_NODES,
            max_depth: super::MAX_DEPTH,
            max_actions: super::MAX_ACTIONS,
            max_action_groups: super::MAX_ACTION_GROUPS,
            max_scalar_bytes: MAX_SCALAR_BYTES,
        }
    }
}

impl Limits {
    /// Validate a caller budget before any builder state is retained.
    pub fn validate(self) -> Result<()> {
        if self.max_output_bytes == 0 || self.max_output_bytes > super::MAX_SOURCE_BYTES {
            return Err(limit("ink action output bytes", super::MAX_SOURCE_BYTES));
        }
        if self.max_payload_bytes == 0 || self.max_payload_bytes > MAX_PAYLOAD_BYTES {
            return Err(limit("ink action payload bytes", MAX_PAYLOAD_BYTES));
        }
        if self.max_nodes == 0 || self.max_nodes > super::MAX_NODES {
            return Err(limit("ink action XML nodes", super::MAX_NODES));
        }
        if self.max_depth == 0 || self.max_depth > super::MAX_DEPTH {
            return Err(limit("ink action XML depth", super::MAX_DEPTH));
        }
        if self.max_actions == 0 || self.max_actions > super::MAX_ACTIONS {
            return Err(limit("ink action records", super::MAX_ACTIONS));
        }
        if self.max_action_groups == 0 || self.max_action_groups > super::MAX_ACTION_GROUPS {
            return Err(limit("ink action groups", super::MAX_ACTION_GROUPS));
        }
        if self.max_scalar_bytes == 0 || self.max_scalar_bytes > MAX_SCALAR_BYTES {
            return Err(limit("ink action scalar bytes", MAX_SCALAR_BYTES));
        }
        Ok(())
    }
}

/// A bounded complete XML element retained as an opaque action payload.
///
/// `definitions`, `transform`, `trace`, and `traceView` values are deliberately
/// opaque at this layer.  The payload must contain the complete element,
/// including any namespace declarations it needs.  Detached authoring also
/// admits the canonical `iact:` and `inkml:` roots without repeating those
/// declarations because the generated actions root supplies them; other
/// prefixes still need a declaration in the payload.  The enclosing profile
/// readback validates its admitted wrapper and preserves its descendants.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OpaquePayload(Box<[u8]>);

impl OpaquePayload {
    /// Copy one bounded complete XML payload.
    pub fn new(bytes: impl AsRef<[u8]>) -> Result<Self> {
        let bytes = bytes.as_ref();
        if bytes.len() > MAX_PAYLOAD_BYTES {
            return Err(limit("ink action payload bytes", MAX_PAYLOAD_BYTES));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|_| invalid("ink action payload allocation failed"))?;
        owned.extend_from_slice(bytes);
        Ok(Self(owned.into_boxed_slice()))
    }

    /// Alias for [`OpaquePayload::new`] when the caller already has bytes.
    pub fn from_bytes(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::new(bytes)
    }

    /// Borrow the exact payload bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }
}

/// A property supplied to a detached action builder.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PropertyDraft {
    name: Box<str>,
    value: Box<str>,
}

impl PropertyDraft {
    /// Construct one property.  The schema requires `name`; empty custom
    /// strings remain representable, subject to the scalar/XML bounds.
    pub fn new(name: impl AsRef<str>, value: impl AsRef<str>) -> Result<Self> {
        Ok(Self {
            name: scalar(name.as_ref(), "property name")?,
            value: scalar(value.as_ref(), "property value")?,
        })
    }

    /// Name lexical value.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Value lexical value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Alias for [`PropertyDraft::new`].
    pub fn with_value(name: impl AsRef<str>, value: impl AsRef<str>) -> Result<Self> {
        Self::new(name, value)
    }
}

/// One opaque child of an action-data builder.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum DataChildDraft {
    /// A complete `iact:transform` element.
    Transform(OpaquePayload),
    /// A complete `inkml:trace` element.
    Trace(OpaquePayload),
    /// A complete `inkml:traceView` element.
    TraceView(OpaquePayload),
}

/// A detached `CT_ActionData` builder.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct ActionDataDraft {
    xml_id: Option<Box<str>>,
    name: Box<str>,
    reference: Option<Box<str>>,
    children: Vec<DataChildDraft>,
    transform_seen: bool,
    trace_seen: bool,
}

impl ActionDataDraft {
    /// Construct data with the schema default name (`stroke`).
    pub fn new() -> Self {
        Self {
            name: "stroke".into(),
            ..Self::default()
        }
    }

    /// Set an optional XML identifier.
    pub fn with_xml_id(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.xml_id = Some(xml_id(value.as_ref())?);
        Ok(self)
    }

    /// Set the data name.
    pub fn with_name(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.name = scalar(value.as_ref(), "action data name")?;
        Ok(self)
    }

    /// Set the optional `ref` lexical value.
    pub fn with_reference(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.reference = Some(scalar(value.as_ref(), "action data reference")?);
        Ok(self)
    }

    /// Add the unique first transform child.
    pub fn transform(mut self, payload: OpaquePayload) -> Result<Self> {
        if self.transform_seen || self.trace_seen {
            return Err(invalid(
                "ink action transform must be the first and unique data child",
            ));
        }
        self.transform_seen = true;
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action data allocation failed"))?;
        self.children.push(DataChildDraft::Transform(payload));
        Ok(self)
    }

    /// Add a trace child after any transform.
    pub fn trace(mut self, payload: OpaquePayload) -> Result<Self> {
        self.trace_seen = true;
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action data allocation failed"))?;
        self.children.push(DataChildDraft::Trace(payload));
        Ok(self)
    }

    /// Add a trace-view child after any transform.
    pub fn trace_view(mut self, payload: OpaquePayload) -> Result<Self> {
        self.trace_seen = true;
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action data allocation failed"))?;
        self.children.push(DataChildDraft::TraceView(payload));
        Ok(self)
    }
}

/// A detached nonempty `CT_ActionDataGroup` builder.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DataGroupDraft {
    xml_id: Option<Box<str>>,
    name: Box<str>,
    data: Vec<ActionDataDraft>,
}

impl DataGroupDraft {
    /// Construct a group with its required first data child.
    pub fn new(data: ActionDataDraft) -> Result<Self> {
        let mut values = Vec::new();
        values
            .try_reserve(1)
            .map_err(|_| invalid("ink action data-group allocation failed"))?;
        values.push(data);
        Ok(Self {
            xml_id: None,
            name: "stroke".into(),
            data: values,
        })
    }

    /// Set an optional XML identifier.
    pub fn with_xml_id(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.xml_id = Some(xml_id(value.as_ref())?);
        Ok(self)
    }

    /// Set the group name.
    pub fn with_name(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.name = scalar(value.as_ref(), "action data-group name")?;
        Ok(self)
    }

    /// Append a required data child.
    pub fn data(mut self, data: ActionDataDraft) -> Result<Self> {
        self.data
            .try_reserve(1)
            .map_err(|_| invalid("ink action data-group allocation failed"))?;
        self.data.push(data);
        Ok(self)
    }

    /// Alias for [`DataGroupDraft::data`].
    pub fn add_data(self, data: ActionDataDraft) -> Result<Self> {
        self.data(data)
    }
}

/// A detached `CT_Action` builder.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ActionDraft {
    xml_id: Option<Box<str>>,
    action_type: ActionType,
    start_time: Box<str>,
    properties: Vec<PropertyDraft>,
    children: Vec<ActionChildDraft>,
    data_seen: bool,
}

#[derive(Clone, Debug, Eq, PartialEq)]
enum ActionChildDraft {
    Property(PropertyDraft),
    Data(ActionDataDraft),
    DataGroup(DataGroupDraft),
}

impl ActionDraft {
    /// Construct an action with required type and decimal start time.
    pub fn new(action_type: ActionType, start_time: impl AsRef<str>) -> Result<Self> {
        let action_type = checked_action_type(action_type)?;
        let start_time = decimal(start_time.as_ref())?;
        Ok(Self {
            xml_id: None,
            action_type,
            start_time,
            properties: Vec::new(),
            children: Vec::new(),
            data_seen: false,
        })
    }

    /// Set an optional XML identifier.
    pub fn with_xml_id(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.xml_id = Some(xml_id(value.as_ref())?);
        Ok(self)
    }

    /// Append an action property before any data child.
    pub fn property(mut self, property: PropertyDraft) -> Result<Self> {
        if self.data_seen {
            return Err(invalid("ink action properties must precede action data"));
        }
        self.properties
            .try_reserve(1)
            .map_err(|_| invalid("ink action property allocation failed"))?;
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action allocation failed"))?;
        self.properties.push(property.clone());
        self.children.push(ActionChildDraft::Property(property));
        Ok(self)
    }

    /// Alias for [`ActionDraft::property`].
    pub fn add_property(self, property: PropertyDraft) -> Result<Self> {
        self.property(property)
    }

    /// Append one action-data child.
    pub fn data(mut self, data: ActionDataDraft) -> Result<Self> {
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action allocation failed"))?;
        self.data_seen = true;
        self.children.push(ActionChildDraft::Data(data));
        Ok(self)
    }

    /// Alias for [`ActionDraft::data`].
    pub fn add_data(self, data: ActionDataDraft) -> Result<Self> {
        self.data(data)
    }

    /// Append one nonempty action-data group.
    pub fn data_group(mut self, group: DataGroupDraft) -> Result<Self> {
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action allocation failed"))?;
        self.data_seen = true;
        self.children.push(ActionChildDraft::DataGroup(group));
        Ok(self)
    }

    /// Alias for [`ActionDraft::data_group`].
    pub fn add_data_group(self, group: DataGroupDraft) -> Result<Self> {
        self.data_group(group)
    }
}

/// A detached nonempty `CT_ActionGroup` builder.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ActionGroupDraft {
    xml_id: Option<Box<str>>,
    action_type: ActionType,
    start_time: Box<str>,
    actions: Vec<ActionDraft>,
}

impl ActionGroupDraft {
    /// Construct a group with its required first action.
    pub fn new(action: ActionDraft) -> Result<Self> {
        let action_type = action.action_type.clone();
        let start_time = action.start_time.clone();
        let mut actions = Vec::new();
        actions
            .try_reserve(1)
            .map_err(|_| invalid("ink action-group allocation failed"))?;
        actions.push(action);
        Ok(Self {
            xml_id: None,
            action_type,
            start_time,
            actions,
        })
    }

    /// Construct a group with explicit group metadata and one required child.
    pub fn with_metadata(
        action_type: ActionType,
        start_time: impl AsRef<str>,
        action: ActionDraft,
    ) -> Result<Self> {
        let mut actions = Vec::new();
        actions
            .try_reserve(1)
            .map_err(|_| invalid("ink action-group allocation failed"))?;
        actions.push(action);
        Ok(Self {
            xml_id: None,
            action_type: checked_action_type(action_type)?,
            start_time: decimal(start_time.as_ref())?,
            actions,
        })
    }

    /// Set an optional XML identifier.
    pub fn with_xml_id(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.xml_id = Some(xml_id(value.as_ref())?);
        Ok(self)
    }

    /// Append an action.
    pub fn action(mut self, action: ActionDraft) -> Result<Self> {
        self.actions
            .try_reserve(1)
            .map_err(|_| invalid("ink action-group allocation failed"))?;
        self.actions.push(action);
        Ok(self)
    }

    /// Alias for [`ActionGroupDraft::action`].
    pub fn add_action(self, action: ActionDraft) -> Result<Self> {
        self.action(action)
    }
}

/// Root insertion parent for a source-backed edit.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ActionParent {
    /// Append to the root action sequence.
    Root,
    /// Append to the direct action sequence of the semantic group ordinal.
    Group(usize),
}

/// Semantic action selector. `Ordinal` walks actions depth-first in source
/// order, while the other forms are useful when a caller already tracks the
/// root/group structure.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub enum ActionSelector {
    /// Depth-first action ordinal, zero based.
    Ordinal(usize),
    /// Direct root action ordinal, excluding grouped actions.
    Direct(usize),
    /// Action ordinal inside a direct group ordinal.
    Group { group: usize, action: usize },
}

impl ActionSelector {
    /// Construct a depth-first semantic selector.
    #[must_use]
    pub const fn ordinal(index: usize) -> Self {
        Self::Ordinal(index)
    }

    /// Construct a direct-root selector.
    #[must_use]
    pub const fn direct(index: usize) -> Self {
        Self::Direct(index)
    }

    /// Construct a grouped-action selector.
    #[must_use]
    pub const fn grouped(group: usize, action: usize) -> Self {
        Self::Group { group, action }
    }
}

/// Semantic selector for action-data children.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub enum DataSelector {
    /// Data ordinal among direct `actionData` children of an action.
    Action {
        action: ActionSelector,
        index: usize,
    },
    /// Data ordinal inside an action-data group child.
    Group {
        action: ActionSelector,
        group: usize,
        index: usize,
    },
}

/// Semantic selector for action children.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub enum ChildSelector {
    /// Property ordinal among an action's direct properties.
    Property {
        action: ActionSelector,
        index: usize,
    },
    /// Action-data ordinal among an action's direct data children.
    Data {
        action: ActionSelector,
        index: usize,
    },
    /// Data-group ordinal among an action's direct data-group children.
    DataGroup {
        action: ActionSelector,
        index: usize,
    },
    /// Data ordinal inside an action's data-group child.
    GroupData {
        action: ActionSelector,
        group: usize,
        index: usize,
    },
}

/// A detached bounded action value under construction.
#[derive(Clone, Debug)]
pub struct Draft {
    limits: Limits,
    xml_id: Option<Box<str>>,
    length_unit: LengthUnit,
    time_unit: TimeUnit,
    definitions: Option<OpaquePayload>,
    children: Vec<RootDraft>,
}

#[derive(Clone, Debug)]
enum RootDraft {
    Action(ActionDraft),
    ActionGroup(ActionGroupDraft),
}

impl Draft {
    /// Start a detached action value with caller limits and required units.
    pub fn new(limits: Limits, length_unit: LengthUnit, time_unit: TimeUnit) -> Result<Self> {
        limits.validate()?;
        Ok(Self {
            limits,
            xml_id: None,
            length_unit,
            time_unit,
            definitions: None,
            children: Vec::new(),
        })
    }

    /// Start with default limits.
    pub fn with_units(length_unit: LengthUnit, time_unit: TimeUnit) -> Result<Self> {
        Self::new(Limits::default(), length_unit, time_unit)
    }

    /// Borrow the active budgets.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Set an optional root XML identifier.
    pub fn with_xml_id(mut self, value: impl AsRef<str>) -> Result<Self> {
        self.xml_id = Some(xml_id(value.as_ref())?);
        Ok(self)
    }

    /// Set the optional first InkML definitions payload.
    pub fn definitions(mut self, payload: OpaquePayload) -> Result<Self> {
        if !self.children.is_empty() || self.definitions.is_some() {
            return Err(invalid(
                "ink action definitions must be unique and precede actions",
            ));
        }
        self.definitions = Some(payload);
        Ok(self)
    }

    /// Append a direct action.
    pub fn action(mut self, action: ActionDraft) -> Result<Self> {
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action root allocation failed"))?;
        self.children.push(RootDraft::Action(action));
        Ok(self)
    }

    /// Alias for [`Draft::action`].
    pub fn add_action(self, action: ActionDraft) -> Result<Self> {
        self.action(action)
    }

    /// Append a direct nonempty action group.
    pub fn action_group(mut self, group: ActionGroupDraft) -> Result<Self> {
        self.children
            .try_reserve(1)
            .map_err(|_| invalid("ink action root allocation failed"))?;
        self.children.push(RootDraft::ActionGroup(group));
        Ok(self)
    }

    /// Alias for [`Draft::action_group`].
    pub fn add_action_group(self, group: ActionGroupDraft) -> Result<Self> {
        self.action_group(group)
    }

    /// Validate, preflight, emit, and immediately read back the value.
    pub fn finish(self) -> Result<Prepared> {
        validate_draft(&self)?;
        let output_len = draft_len(&self)?;
        if output_len > self.limits.max_output_bytes {
            return Err(limit(
                "ink action output bytes",
                self.limits.max_output_bytes,
            ));
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_| invalid("ink action output allocation failed"))?;
        emit_draft(&mut output, &self)?;
        debug_assert_eq!(output.len(), output_len);
        let profile = read_profile(&output)?;
        validate_profile_limits(&profile, self.limits)?;
        ensure_unique_profile_ids(&profile)?;
        if !matches_draft(&self, &profile) {
            return Err(invalid("ink action authoring readback differs from draft"));
        }
        Ok(Prepared::from_profile(profile))
    }
}

/// A checked, read-back detached action value.
#[derive(Clone, Debug)]
pub struct Prepared {
    bytes: Arc<[u8]>,
    profile: Profile,
}

impl Prepared {
    fn from_profile(profile: Profile) -> Self {
        Self {
            bytes: profile.source.clone(),
            profile,
        }
    }

    /// Borrow the exact generated bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Alias for [`Prepared::as_bytes`].
    #[must_use]
    pub fn source(&self) -> &[u8] {
        self.as_bytes()
    }

    /// Borrow the strict typed readback.
    #[must_use]
    pub fn profile(&self) -> &Profile {
        &self.profile
    }

    /// Alias for [`Prepared::profile`].
    #[must_use]
    pub fn readback(&self) -> &Profile {
        self.profile()
    }

    /// Return a cloned byte vector for a sink or host-owned buffer.
    pub fn to_vec(&self) -> Result<Vec<u8>> {
        let mut result = Vec::new();
        result
            .try_reserve_exact(self.bytes.len())
            .map_err(|_| invalid("ink action prepared output allocation failed"))?;
        result.extend_from_slice(&self.bytes);
        Ok(result)
    }
}

/// A source-checked reversible byte patch.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    /// Apply this patch only to the exact source from which it was made.
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>> {
        if source != self.before.as_ref() {
            return Err(invalid("ink action patch source is stale"));
        }
        copy_bytes(&self.after)
    }

    /// Reverse the patch with the same source check.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact result bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }
}

/// Result of a source-backed edit, carrying both typed readback and inverse.
#[derive(Clone, Debug)]
pub struct Commit {
    prepared: Prepared,
    patch: Patch,
}

impl Commit {
    /// Borrow the checked edited value.
    #[must_use]
    pub fn prepared(&self) -> &Prepared {
        &self.prepared
    }

    /// Borrow the strict typed readback directly.
    #[must_use]
    pub fn profile(&self) -> &Profile {
        self.prepared.profile()
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow final bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.prepared.as_bytes()
    }

    /// Clone final bytes for a host-owned buffer.
    pub fn to_vec(&self) -> Result<Vec<u8>> {
        self.prepared.to_vec()
    }
}

/// A source-backed transaction over one strict profile.
#[derive(Clone, Debug)]
pub struct Edit {
    profile: Profile,
    limits: Limits,
    operations: Vec<EditOp>,
}

#[derive(Clone, Debug)]
enum EditOp {
    SetActionType(ActionSelector, ActionType),
    SetStartTime(ActionSelector, Box<str>),
    SetPropertyName(ChildSelector, Box<str>),
    SetPropertyValue(ChildSelector, Box<str>),
    SetDataName(DataSelector, Box<str>),
    SetDataReference(DataSelector, Option<Box<str>>),
    SetActionGroupType(usize, ActionType),
    SetActionGroupStartTime(usize, Box<str>),
    SetDataGroupName(ChildSelector, Box<str>),
    Add(ActionParent, ActionDraft),
    AddGroup(ActionGroupDraft),
    AddProperty(ChildParent, PropertyDraft),
    AddData(ActionSelector, ActionDataDraft),
    AddDataGroup(ActionSelector, DataGroupDraft),
    MoveBefore(ActionSelector, ActionSelector),
    MoveGroupBefore(usize, usize),
    MoveChild(ChildSelector, ChildSelector),
    MoveGroupData(DataSelector, DataSelector),
    Clear(ActionSelector),
    RemoveAction(ActionSelector),
    RemoveGroup(usize),
    RemoveProperty(ChildSelector),
    RemoveData(DataSelector),
    RemoveDataGroup(ChildSelector),
}

#[derive(Clone, Copy, Debug)]
enum ChildParent {
    Action(ActionSelector),
}

impl Edit {
    /// Start an edit with default output/resource limits.
    pub fn new(profile: Profile) -> Result<Self> {
        Self::with_limits(profile, Limits::default())
    }

    /// Clone a profile into a source-backed edit.
    pub fn from_profile(profile: &Profile) -> Result<Self> {
        Self::new(profile.clone())
    }

    /// Start an edit with explicit output/resource limits.
    pub fn with_limits(profile: Profile, limits: Limits) -> Result<Self> {
        limits.validate()?;
        validate_profile_limits(&profile, limits)?;
        Ok(Self {
            profile,
            limits,
            operations: Vec::new(),
        })
    }

    /// Borrow the unchanged source snapshot.
    #[must_use]
    pub fn profile(&self) -> &Profile {
        &self.profile
    }

    /// Queue an action-type replacement.
    pub fn set_action_type(&mut self, selector: ActionSelector, value: ActionType) -> Result<()> {
        checked_action_type(value.clone())?;
        bounded_scalar(value.as_str(), "action type", self.limits)?;
        self.push(EditOp::SetActionType(selector, value))
    }

    /// Queue a decimal start-time replacement.
    pub fn set_start_time(
        &mut self,
        selector: ActionSelector,
        value: impl AsRef<str>,
    ) -> Result<()> {
        let value = decimal(value.as_ref())?;
        bounded_scalar(&value, "ink action startTime", self.limits)?;
        self.push(EditOp::SetStartTime(selector, value))
    }

    /// Queue a property-name replacement.
    pub fn set_property_name(
        &mut self,
        selector: ChildSelector,
        value: impl AsRef<str>,
    ) -> Result<()> {
        let value = scalar(value.as_ref(), "property name")?;
        bounded_scalar(&value, "property name", self.limits)?;
        self.push(EditOp::SetPropertyName(selector, value))
    }

    /// Queue a property-value replacement.
    pub fn set_property_value(
        &mut self,
        selector: ChildSelector,
        value: impl AsRef<str>,
    ) -> Result<()> {
        let value = scalar(value.as_ref(), "property value")?;
        bounded_scalar(&value, "property value", self.limits)?;
        self.push(EditOp::SetPropertyValue(selector, value))
    }

    /// Queue an action-data name replacement.
    pub fn set_data_name(&mut self, selector: DataSelector, value: impl AsRef<str>) -> Result<()> {
        let value = scalar(value.as_ref(), "action data name")?;
        bounded_scalar(&value, "action data name", self.limits)?;
        self.push(EditOp::SetDataName(selector, value))
    }

    /// Queue an action-group type replacement.
    pub fn set_action_group_type(&mut self, group: usize, value: ActionType) -> Result<()> {
        checked_action_type(value.clone())?;
        bounded_scalar(value.as_str(), "action group type", self.limits)?;
        self.push(EditOp::SetActionGroupType(group, value))
    }

    /// Queue an action-group decimal start-time replacement.
    pub fn set_action_group_start_time(
        &mut self,
        group: usize,
        value: impl AsRef<str>,
    ) -> Result<()> {
        let value = decimal(value.as_ref())?;
        bounded_scalar(&value, "action group startTime", self.limits)?;
        self.push(EditOp::SetActionGroupStartTime(group, value))
    }

    /// Queue an action-data-group name replacement.
    pub fn set_data_group_name(
        &mut self,
        selector: ChildSelector,
        value: impl AsRef<str>,
    ) -> Result<()> {
        if locate_child(&self.profile, selector)?.kind != ChildKind::DataGroup {
            return Err(invalid("ink action selector is not a data group"));
        }
        let value = scalar(value.as_ref(), "action data-group name")?;
        bounded_scalar(&value, "action data-group name", self.limits)?;
        self.push(EditOp::SetDataGroupName(selector, value))
    }

    /// Alias for [`Edit::set_action_group_type`].
    pub fn set_group_type(&mut self, group: usize, value: ActionType) -> Result<()> {
        self.set_action_group_type(group, value)
    }

    /// Alias for [`Edit::set_action_group_start_time`].
    pub fn set_group_start_time(&mut self, group: usize, value: impl AsRef<str>) -> Result<()> {
        self.set_action_group_start_time(group, value)
    }

    /// Queue an action-data reference replacement.
    pub fn set_data_reference(
        &mut self,
        selector: DataSelector,
        value: Option<impl AsRef<str>>,
    ) -> Result<()> {
        let value = value
            .map(|value| scalar(value.as_ref(), "action data reference"))
            .transpose()?;
        if let Some(value) = &value {
            bounded_scalar(value, "action data reference", self.limits)?;
        }
        self.push(EditOp::SetDataReference(selector, value))
    }

    /// Identity changes are refused until a host supplies a modeled ref
    /// closure.  Replacing an id with itself is a no-op.
    pub fn set_action_id(
        &mut self,
        selector: ActionSelector,
        value: impl AsRef<str>,
    ) -> Result<()> {
        let action = locate_action(&self.profile, selector)?;
        let value = xml_id(value.as_ref())?;
        if action.xml_id.as_deref() == Some(&value) {
            Ok(())
        } else {
            Err(invalid(
                "ink action identity edit requires a modeled reference closure",
            ))
        }
    }

    /// Data identity changes are refused until a host supplies a modeled ref
    /// closure.  Replacing an id with itself is a no-op.
    pub fn set_data_id(&mut self, selector: DataSelector, value: impl AsRef<str>) -> Result<()> {
        let data = locate_data(&self.profile, selector)?;
        let value = xml_id(value.as_ref())?;
        if data.xml_id.as_deref() == Some(&value) {
            Ok(())
        } else {
            Err(invalid(
                "ink action identity edit requires a modeled reference closure",
            ))
        }
    }

    /// Append an authored action under the selected root/group parent.
    pub fn add_action(&mut self, parent: ActionParent, action: ActionDraft) -> Result<()> {
        let mut nodes = 1usize;
        let action_depth = match parent {
            ActionParent::Root => 2,
            ActionParent::Group(_) => 3,
        };
        let mut max_depth = 0usize;
        validate_action_draft(
            &action,
            self.limits,
            &mut nodes,
            action_depth,
            &mut max_depth,
            false,
        )?;
        self.push(EditOp::Add(parent, action))
    }

    /// Append an authored action group at the root.
    pub fn add_action_group(&mut self, group: ActionGroupDraft) -> Result<()> {
        let mut nodes = 1usize;
        let mut max_depth = 0usize;
        validate_group_draft(&group, self.limits, &mut nodes, &mut max_depth, false)?;
        self.push(EditOp::AddGroup(group))
    }

    /// Alias for [`Edit::add_action_group`].
    pub fn add_group(&mut self, group: ActionGroupDraft) -> Result<()> {
        self.add_action_group(group)
    }

    /// Append a property to an existing action, before its data children.
    pub fn add_property(&mut self, action: ActionSelector, property: PropertyDraft) -> Result<()> {
        self.push(EditOp::AddProperty(ChildParent::Action(action), property))
    }

    /// Append direct action data to an existing action.
    pub fn add_data(&mut self, action: ActionSelector, data: ActionDataDraft) -> Result<()> {
        let mut nodes = 1usize;
        let data_depth = action_depth_for_selector(&self.profile, action)?
            .checked_add(1)
            .ok_or_else(|| limit("ink action XML depth", self.limits.max_depth))?;
        let mut max_depth = 0usize;
        validate_data_draft(
            &data,
            self.limits,
            &mut nodes,
            data_depth,
            &mut max_depth,
            false,
        )?;
        self.push(EditOp::AddData(action, data))
    }

    /// Append an action-data group to an existing action.
    pub fn add_data_group(&mut self, action: ActionSelector, group: DataGroupDraft) -> Result<()> {
        let group_depth = action_depth_for_selector(&self.profile, action)?
            .checked_add(1)
            .ok_or_else(|| limit("ink action XML depth", self.limits.max_depth))?;
        let mut max_depth = 0usize;
        validate_data_group_draft(&group, self.limits, group_depth, &mut max_depth, false)?;
        self.push(EditOp::AddDataGroup(action, group))
    }

    /// Move one action before another action in the same root/group sequence.
    pub fn move_before(&mut self, from: ActionSelector, before: ActionSelector) -> Result<()> {
        self.push(EditOp::MoveBefore(from, before))
    }

    /// Move one direct root action group before another group.
    pub fn move_group_before(&mut self, from: usize, before: usize) -> Result<()> {
        self.push(EditOp::MoveGroupBefore(from, before))
    }

    /// Move two direct children of one action while preserving schema order.
    pub fn move_child_before(&mut self, from: ChildSelector, before: ChildSelector) -> Result<()> {
        self.push(EditOp::MoveChild(from, before))
    }

    /// Alias for [`Edit::move_child_before`].
    pub fn move_child(&mut self, from: ChildSelector, before: ChildSelector) -> Result<()> {
        self.move_child_before(from, before)
    }

    /// Move two data children within one action-data group.
    pub fn move_group_data_before(
        &mut self,
        from: DataSelector,
        before: DataSelector,
    ) -> Result<()> {
        self.push(EditOp::MoveGroupData(from, before))
    }

    /// Alias for [`Edit::move_group_data_before`].
    pub fn move_group_data(&mut self, from: DataSelector, before: DataSelector) -> Result<()> {
        self.move_group_data_before(from, before)
    }

    /// Remove all direct children from an action while retaining its start tag.
    pub fn clear(&mut self, selector: ActionSelector) -> Result<()> {
        self.push(EditOp::Clear(selector))
    }

    /// Remove a direct action.
    pub fn remove_action(&mut self, selector: ActionSelector) -> Result<()> {
        self.push(EditOp::RemoveAction(selector))
    }

    /// Remove a direct root action group.
    pub fn remove_action_group(&mut self, index: usize) -> Result<()> {
        self.push(EditOp::RemoveGroup(index))
    }

    /// Alias for [`Edit::remove_action_group`].
    pub fn remove_group(&mut self, index: usize) -> Result<()> {
        self.remove_action_group(index)
    }

    /// Remove a selected property.
    pub fn remove_property(&mut self, selector: ChildSelector) -> Result<()> {
        self.push(EditOp::RemoveProperty(selector))
    }

    /// Remove a selected action-data child.
    pub fn remove_data(&mut self, selector: DataSelector) -> Result<()> {
        self.push(EditOp::RemoveData(selector))
    }

    /// Remove a selected action-data group.
    pub fn remove_data_group(&mut self, selector: ChildSelector) -> Result<()> {
        if !matches!(selector, ChildSelector::DataGroup { .. }) {
            return Err(invalid("ink action selector is not a data group"));
        }
        self.push(EditOp::RemoveDataGroup(selector))
    }

    /// Remove a selected data child inside an action-data group.
    pub fn remove_group_data(&mut self, selector: DataSelector) -> Result<()> {
        if !matches!(selector, DataSelector::Group { .. }) {
            return Err(invalid("ink action selector is not grouped data"));
        }
        self.push(EditOp::RemoveData(selector))
    }

    /// Apply a generic child target removal.
    pub fn remove(&mut self, target: ChildSelector) -> Result<()> {
        match target {
            ChildSelector::Property { .. } => self.remove_property(target),
            ChildSelector::Data { action, index } => {
                self.remove_data(DataSelector::Action { action, index })
            },
            ChildSelector::DataGroup { .. } => self.remove_data_group(target),
            ChildSelector::GroupData {
                action,
                group,
                index,
            } => self.remove_group_data(DataSelector::Group {
                action,
                group,
                index,
            }),
        }
    }

    /// Apply queued edits, returning exact typed readback and a reversible
    /// source-checked patch. Selectors are resolved against the original
    /// retained profile, so a removal or move cannot retarget a later
    /// operation. Scalar intents for one source attribute are coalesced;
    /// conflicting structural ranges are refused. With no queued operations,
    /// the original source allocation is replayed exactly.
    pub fn finish(self) -> Result<Commit> {
        let before = self.profile.source.clone();
        if self.operations.is_empty() {
            let prepared = Prepared::from_profile(self.profile);
            let patch = Patch {
                before: before.clone(),
                after: before,
            };
            return Ok(Commit { prepared, patch });
        }
        let plans = plan_operations(&self.profile, &before, &self.operations, self.limits)?;
        let after = emit_plans(&before, &plans, self.limits)?;
        if after.as_slice() == before.as_ref() {
            let prepared = Prepared::from_profile(self.profile);
            let patch = Patch {
                before: before.clone(),
                after: before,
            };
            return Ok(Commit { prepared, patch });
        }
        let after: Arc<[u8]> = Arc::from(after.into_boxed_slice());
        let profile = read_profile_owned(after.clone())?;
        validate_profile_limits(&profile, self.limits)?;
        ensure_unique_profile_ids(&profile)?;
        let prepared = Prepared {
            bytes: after.clone(),
            profile,
        };
        let patch = Patch { before, after };
        Ok(Commit { prepared, patch })
    }

    /// Alias for [`Edit::finish`].
    pub fn commit(self) -> Result<Commit> {
        self.finish()
    }

    /// Return the checked edited value without discarding the patch.
    pub fn prepare(self) -> Result<Prepared> {
        Ok(self.finish()?.prepared)
    }

    fn push(&mut self, operation: EditOp) -> Result<()> {
        if self.operations.len() >= MAX_EDIT_OPERATIONS {
            return Err(limit("ink action edit operations", MAX_EDIT_OPERATIONS));
        }
        self.operations
            .try_reserve(1)
            .map_err(|_| invalid("ink action edit allocation failed"))?;
        self.operations.push(operation);
        Ok(())
    }
}

fn checked_action_type(value: ActionType) -> Result<ActionType> {
    validate_scalar(value.as_str(), "action type")?;
    Ok(value)
}

fn validate_scalar(value: &str, field: &'static str) -> Result<()> {
    if value.len() > MAX_SCALAR_BYTES || value.bytes().any(|byte| byte == 0) {
        return Err(limit(field, MAX_SCALAR_BYTES));
    }
    super::validate_xml_characters(value, field)?;
    Ok(())
}

fn scalar(value: &str, field: &'static str) -> Result<Box<str>> {
    validate_scalar(value, field)?;
    Ok(value.into())
}

fn xml_id(value: &str) -> Result<Box<str>> {
    let value = scalar(value, "xml:id")?;
    if !is_ncname(&value) {
        return Err(invalid("ink action xml:id is not an NCName"));
    }
    Ok(value)
}

fn decimal(value: &str) -> Result<Box<str>> {
    if value.len() > MAX_SCALAR_BYTES {
        return Err(limit("ink action startTime", MAX_SCALAR_BYTES));
    }
    let collapsed = super::collapse_xml_whitespace(value)?;
    super::validate_decimal(&collapsed)?;
    scalar(&collapsed, "ink action startTime")
}

fn validate_draft(draft: &Draft) -> Result<()> {
    let mut identifiers = Vec::new();
    let mut opaque_identifiers = Vec::new();
    if let Some(id) = draft.xml_id.as_deref() {
        push_id(&mut identifiers, id)?;
    }
    if let Some(payload) = &draft.definitions {
        collect_opaque_ids(payload.as_bytes(), &mut opaque_identifiers)?;
    }
    for child in &draft.children {
        match child {
            RootDraft::Action(action) => {
                collect_draft_ids(action, &mut identifiers)?;
                collect_draft_opaque_ids(action, &mut opaque_identifiers)?;
            },
            RootDraft::ActionGroup(group) => {
                if let Some(id) = group.xml_id.as_deref() {
                    push_id(&mut identifiers, id)?;
                }
                for action in &group.actions {
                    collect_draft_ids(action, &mut identifiers)?;
                    collect_draft_opaque_ids(action, &mut opaque_identifiers)?;
                }
            },
        }
    }
    let mut identifier_set = HashSet::new();
    identifier_set
        .try_reserve(identifiers.len())
        .map_err(|_| invalid("ink action identity allocation failed"))?;
    for id in identifiers {
        if !identifier_set.insert(id) {
            return Err(invalid(
                "ink action draft would create duplicate xml:id values",
            ));
        }
    }
    let mut opaque_identifier_set = HashSet::new();
    opaque_identifier_set
        .try_reserve(opaque_identifiers.len())
        .map_err(|_| invalid("ink action identity allocation failed"))?;
    for id in opaque_identifiers {
        if identifier_set.contains(id.as_ref()) || !opaque_identifier_set.insert(id) {
            return Err(invalid(
                "ink action draft would create duplicate xml:id values",
            ));
        }
    }
    let mut actions = 0usize;
    let mut groups = 0usize;
    let mut nodes = 1usize;
    let mut depth = 1usize;
    if let Some(payload) = &draft.definitions {
        let metrics = check_payload(payload, draft.limits, PayloadRoot::Definitions, true)?;
        nodes = nodes
            .checked_add(metrics.nodes)
            .ok_or_else(|| limit("ink action XML nodes", draft.limits.max_nodes))?;
        depth = depth.max(
            1usize
                .checked_add(metrics.depth)
                .ok_or_else(|| limit("ink action XML depth", draft.limits.max_depth))?,
        );
    }
    if draft.xml_id.is_some() {
        // The identifier has already been checked by `with_xml_id`.
        if draft
            .xml_id
            .as_deref()
            .is_some_and(|value| value.len() > draft.limits.max_scalar_bytes)
        {
            return Err(limit(
                "ink action scalar bytes",
                draft.limits.max_scalar_bytes,
            ));
        }
    }
    for child in &draft.children {
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit("ink action XML nodes", draft.limits.max_nodes))?;
        match child {
            RootDraft::Action(action) => {
                actions = actions
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action records", draft.limits.max_actions))?;
                validate_action_draft(action, draft.limits, &mut nodes, 2, &mut depth, true)?;
            },
            RootDraft::ActionGroup(group) => {
                groups = groups
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action groups", draft.limits.max_action_groups))?;
                bounded_scalar(
                    group.action_type.as_str(),
                    "action group type",
                    draft.limits,
                )?;
                bounded_scalar(&group.start_time, "action group startTime", draft.limits)?;
                super::validate_decimal(&group.start_time)?;
                if let Some(id) = &group.xml_id {
                    bounded_scalar(id, "xml:id", draft.limits)?;
                    if !is_ncname(id) {
                        return Err(invalid("ink action xml:id is not an NCName"));
                    }
                }
                if group.actions.is_empty() {
                    return Err(invalid("ink action group requires an action child"));
                }
                for action in &group.actions {
                    actions = actions
                        .checked_add(1)
                        .ok_or_else(|| limit("ink action records", draft.limits.max_actions))?;
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("ink action XML nodes", draft.limits.max_nodes))?;
                    validate_action_draft(action, draft.limits, &mut nodes, 3, &mut depth, true)?;
                }
            },
        }
    }
    if actions > draft.limits.max_actions {
        return Err(limit("ink action records", draft.limits.max_actions));
    }
    if groups > draft.limits.max_action_groups {
        return Err(limit("ink action groups", draft.limits.max_action_groups));
    }
    if nodes > draft.limits.max_nodes {
        return Err(limit("ink action XML nodes", draft.limits.max_nodes));
    }
    if depth > draft.limits.max_depth {
        return Err(limit("ink action XML depth", draft.limits.max_depth));
    }
    Ok(())
}

fn validate_action_draft(
    action: &ActionDraft,
    limits: Limits,
    nodes: &mut usize,
    action_depth: usize,
    max_depth: &mut usize,
    allow_inherited_payload_namespace: bool,
) -> Result<()> {
    *max_depth = (*max_depth).max(action_depth);
    bounded_scalar(action.action_type.as_str(), "action type", limits)?;
    bounded_scalar(&action.start_time, "ink action startTime", limits)?;
    super::validate_decimal(&action.start_time)?;
    if let Some(id) = &action.xml_id {
        bounded_scalar(id, "xml:id", limits)?;
        if !is_ncname(id) {
            return Err(invalid("ink action xml:id is not an NCName"));
        }
    }
    for child in &action.children {
        *nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
        match child {
            ActionChildDraft::Property(property) => {
                bounded_scalar(&property.name, "property name", limits)?;
                bounded_scalar(&property.value, "property value", limits)?;
                *max_depth = (*max_depth).max(
                    action_depth
                        .checked_add(1)
                        .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?,
                );
            },
            ActionChildDraft::Data(data) => {
                let data_depth = action_depth
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                validate_data_draft(
                    data,
                    limits,
                    nodes,
                    data_depth,
                    max_depth,
                    allow_inherited_payload_namespace,
                )?;
            },
            ActionChildDraft::DataGroup(group) => {
                if group.data.is_empty() {
                    return Err(invalid("ink action data group requires actionData"));
                }
                bounded_scalar(&group.name, "action data-group name", limits)?;
                if let Some(id) = &group.xml_id {
                    bounded_scalar(id, "xml:id", limits)?;
                    if !is_ncname(id) {
                        return Err(invalid("ink action xml:id is not an NCName"));
                    }
                }
                let group_depth = action_depth
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                *max_depth = (*max_depth).max(group_depth);
                let data_depth = group_depth
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                for data in &group.data {
                    *nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                    validate_data_draft(
                        data,
                        limits,
                        nodes,
                        data_depth,
                        max_depth,
                        allow_inherited_payload_namespace,
                    )?;
                }
            },
        }
    }
    Ok(())
}

fn validate_data_draft(
    data: &ActionDataDraft,
    limits: Limits,
    nodes: &mut usize,
    data_depth: usize,
    max_depth: &mut usize,
    allow_inherited_payload_namespace: bool,
) -> Result<()> {
    *max_depth = (*max_depth).max(data_depth);
    if let Some(id) = &data.xml_id {
        bounded_scalar(id, "xml:id", limits)?;
        if !is_ncname(id) {
            return Err(invalid("ink action xml:id is not an NCName"));
        }
    }
    bounded_scalar(&data.name, "action data name", limits)?;
    if let Some(reference) = &data.reference {
        bounded_scalar(reference, "action data reference", limits)?;
    }
    for child in &data.children {
        *nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
        match child {
            DataChildDraft::Transform(payload)
            | DataChildDraft::Trace(payload)
            | DataChildDraft::TraceView(payload) => {
                let root = match child {
                    DataChildDraft::Transform(_) => PayloadRoot::Transform,
                    DataChildDraft::Trace(_) => PayloadRoot::Trace,
                    DataChildDraft::TraceView(_) => PayloadRoot::TraceView,
                };
                let metrics =
                    check_payload(payload, limits, root, allow_inherited_payload_namespace)?;
                *nodes = nodes
                    .checked_add(metrics.nodes.saturating_sub(1))
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                let payload_depth = data_depth
                    .checked_add(metrics.depth)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                *max_depth = (*max_depth).max(payload_depth);
            },
        }
    }
    Ok(())
}

#[derive(Clone, Copy, Debug)]
struct PayloadMetrics {
    nodes: usize,
    depth: usize,
}

#[derive(Clone, Copy, Debug)]
enum PayloadRoot {
    Definitions,
    Transform,
    Trace,
    TraceView,
}

impl PayloadRoot {
    const fn namespace(self) -> &'static [u8] {
        match self {
            Self::Definitions | Self::Trace | Self::TraceView => super::INKML_NAMESPACE.as_bytes(),
            Self::Transform => ACTION_NAMESPACE.as_bytes(),
        }
    }

    const fn local(self) -> &'static [u8] {
        match self {
            Self::Definitions => b"definitions",
            Self::Transform => b"transform",
            Self::Trace => b"trace",
            Self::TraceView => b"traceView",
        }
    }

    const fn inherited_prefix(self) -> &'static [u8] {
        match self {
            Self::Definitions | Self::Trace | Self::TraceView => b"inkml",
            Self::Transform => b"iact",
        }
    }
}

fn check_payload(
    payload: &OpaquePayload,
    limits: Limits,
    expected_root: PayloadRoot,
    allow_inherited_namespace: bool,
) -> Result<PayloadMetrics> {
    if payload.0.len() > limits.max_payload_bytes {
        return Err(limit("ink action payload bytes", limits.max_payload_bytes));
    }
    validate_payload_fragment(
        payload.as_bytes(),
        limits,
        expected_root,
        allow_inherited_namespace,
    )
}

/// Validate an authored opaque value as one complete embedded XML element.
///
/// The strict profile intentionally leaves descendants of admitted payload
/// roots opaque, but an editor still needs a complete element boundary.  A
/// comment-only or whitespace-only payload would otherwise be emitted as
/// direct character content and disappear from the typed readback.
fn validate_payload_fragment(
    payload: &[u8],
    limits: Limits,
    expected_root: PayloadRoot,
    allow_inherited_namespace: bool,
) -> Result<PayloadMetrics> {
    validate_payload_root(payload, expected_root, limits, allow_inherited_namespace)?;
    validate_payload_namespace_grammar(payload, limits, allow_inherited_namespace)?;
    let mut reader = Reader::from_reader(payload);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut depth = 0usize;
    let mut peak_depth = 0usize;
    let mut nodes = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid(
                        "ink action opaque payload has content after its root",
                    ));
                }
                validate_payload_element(&element, &reader, limits)?;
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                peak_depth = peak_depth.max(depth);
                if nodes > limits.max_nodes {
                    return Err(limit("ink action XML nodes", limits.max_nodes));
                }
                if depth > limits.max_depth {
                    return Err(limit("ink action XML depth", limits.max_depth));
                }
                if !root_seen {
                    root_seen = true;
                }
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid(
                        "ink action opaque payload has content after its root",
                    ));
                }
                validate_payload_element(&element, &reader, limits)?;
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                peak_depth = peak_depth.max(element_depth);
                if element_depth > limits.max_depth {
                    return Err(limit("ink action XML depth", limits.max_depth));
                }
                if nodes > limits.max_nodes {
                    return Err(limit("ink action XML nodes", limits.max_nodes));
                }
                if !root_seen {
                    root_seen = true;
                    root_closed = true;
                } else if depth == 0 {
                    return Err(invalid("ink action opaque payload has multiple roots"));
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("ink action opaque payload has an unexpected end"));
                }
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                validate_payload_text(&text, "ink action opaque payload text")?;
                if text.windows(3).any(|window| window == b"]]>") {
                    return Err(invalid(
                        "ink action opaque payload text contains a raw CDATA terminator",
                    ));
                }
                if !root_seen || root_closed {
                    return Err(invalid(
                        "ink action opaque payload must contain one complete root element",
                    ));
                }
            },
            Event::CData(data) => {
                validate_payload_text(&data, "ink action opaque payload CDATA")?;
                if data.windows(3).any(|window| window == b"]]>") {
                    return Err(invalid(
                        "ink action opaque payload CDATA contains a raw CDATA terminator",
                    ));
                }
                if !root_seen || root_closed {
                    return Err(invalid(
                        "ink action opaque payload must contain one complete root element",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                validate_payload_reference(&reference)?;
                if !root_seen || root_closed {
                    return Err(invalid(
                        "ink action opaque payload must contain one complete root element",
                    ));
                }
            },
            Event::Comment(comment) => {
                validate_payload_text(&comment, "ink action opaque payload comment")?;
                if !root_seen || root_closed {
                    return Err(invalid(
                        "ink action opaque payload must contain one complete root element",
                    ));
                }
            },
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "ink action opaque payload cannot contain a declaration, DTD, or PI",
                ));
            },
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err(invalid(
            "ink action opaque payload must contain one complete root element",
        ));
    }
    Ok(PayloadMetrics {
        nodes,
        depth: peak_depth,
    })
}

fn validate_payload_root(
    payload: &[u8],
    expected: PayloadRoot,
    limits: Limits,
    allow_inherited_namespace: bool,
) -> Result<()> {
    let mut reader = NsReader::from_reader(payload);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(super::MAX_NAMESPACE_DECLARATIONS);
    let event = reader
        .read_event()
        .map_err(|error| xml_error(error.to_string()))?;
    let (resolved, event) = reader.resolver().resolve_event(event);
    match event {
        Event::Start(element) | Event::Empty(element) => {
            if element.name().local_name().as_ref() != expected.local() {
                return Err(invalid(
                    "ink action opaque payload has the wrong root element",
                ));
            }
            let namespace = payload_namespace(&resolved, &reader, limits)?;
            let expected_namespace = std::str::from_utf8(expected.namespace())
                .map_err(|error| xml_error(error.to_string()))?;
            let inherited_namespace = allow_inherited_namespace
                && matches!(
                    &resolved,
                    ResolveResult::Unknown(prefix)
                        if prefix.as_slice() == expected.inherited_prefix()
                            && !payload_has_namespace_declaration(
                                &element,
                                expected.inherited_prefix(),
                            )
                );
            if namespace.as_deref() != Some(expected_namespace) && !inherited_namespace {
                return Err(invalid(
                    "ink action opaque payload root has the wrong namespace",
                ));
            }
            Ok(())
        },
        Event::Eof => Err(invalid(
            "ink action opaque payload must contain one complete root element",
        )),
        Event::Decl(_) | Event::DocType(_) | Event::PI(_) => Err(invalid(
            "ink action opaque payload cannot contain a declaration, DTD, or PI",
        )),
        Event::End(_) => Err(invalid(
            "ink action opaque payload has invalid root framing",
        )),
        Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) | Event::Comment(_) => Err(
            invalid("ink action opaque payload must contain one complete root element"),
        ),
    }
}

fn payload_has_namespace_declaration(
    element: &quick_xml::events::BytesStart<'_>,
    prefix: &[u8],
) -> bool {
    first_wins(element).any(|attribute| {
        let Ok(attribute) = attribute else {
            return false;
        };
        let key = attribute.key.as_ref();
        (prefix.is_empty() && key == b"xmlns") || key.strip_prefix(b"xmlns:") == Some(prefix)
    })
}

/// The prefixes `element` declares, the empty prefix standing for `xmlns`:
/// the set of `prefix` values for which
/// [`payload_has_namespace_declaration`] holds, read in one pass so that a
/// caller asking about many attributes of one tag scans it once.
fn payload_declared_prefixes<'a>(
    element: &'a quick_xml::events::BytesStart<'_>,
) -> SeenNames<&'a [u8]> {
    let mut prefixes = SeenNames::new();
    for attribute in first_wins(element).flatten() {
        let key = attribute.key.into_inner();
        if key == b"xmlns" {
            prefixes.insert(&[][..]);
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            prefixes.insert(prefix);
        }
    }
    prefixes
}

#[derive(Debug, Hash, PartialEq, Eq)]
struct PayloadExpandedName {
    namespace: Option<Box<str>>,
    local: Box<[u8]>,
}

fn validate_payload_namespace_grammar(
    payload: &[u8],
    limits: Limits,
    allow_inherited_namespace: bool,
) -> Result<()> {
    let mut reader = NsReader::from_reader(payload);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(super::MAX_NAMESPACE_DECLARATIONS);
    let mut stack: Vec<PayloadExpandedName> = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        let (resolved, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid(
                        "ink action opaque payload has content after its root",
                    ));
                }
                let name = payload_expanded_element_name(
                    &element,
                    &resolved,
                    &reader,
                    limits,
                    allow_inherited_namespace,
                )?;
                validate_payload_attributes_with_namespaces(
                    &element,
                    &reader,
                    limits,
                    allow_inherited_namespace,
                )?;
                if !root_seen {
                    root_seen = true;
                }
                stack
                    .try_reserve(1)
                    .map_err(|_| invalid("ink action payload namespace stack allocation failed"))?;
                stack.push(name);
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid(
                        "ink action opaque payload has content after its root",
                    ));
                }
                let _name = payload_expanded_element_name(
                    &element,
                    &resolved,
                    &reader,
                    limits,
                    allow_inherited_namespace,
                )?;
                validate_payload_attributes_with_namespaces(
                    &element,
                    &reader,
                    limits,
                    allow_inherited_namespace,
                )?;
                if !root_seen {
                    root_seen = true;
                    root_closed = true;
                } else if stack.is_empty() {
                    return Err(invalid("ink action opaque payload has multiple roots"));
                }
            },
            Event::End(element) => {
                let expected = stack
                    .pop()
                    .ok_or_else(|| invalid("ink action opaque payload has an unexpected end"))?;
                let actual = payload_expanded_end_name(
                    &element,
                    &resolved,
                    &reader,
                    limits,
                    allow_inherited_namespace,
                )?;
                if actual != expected {
                    return Err(invalid(
                        "ink action opaque payload closing element has the wrong expanded name",
                    ));
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "ink action opaque payload cannot contain a declaration, DTD, or PI",
                ));
            },
            Event::Text(text) => {
                validate_payload_text(&text, "ink action opaque payload text")?;
            },
            Event::CData(data) => {
                validate_payload_text(&data, "ink action opaque payload CDATA")?;
            },
            Event::Comment(comment) => {
                validate_payload_text(&comment, "ink action opaque payload comment")?;
            },
            Event::GeneralRef(reference) => validate_payload_reference(&reference)?,
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid(
            "ink action opaque payload must contain one complete root element",
        ));
    }
    Ok(())
}

fn payload_expanded_element_name<R: std::io::BufRead>(
    element: &quick_xml::events::BytesStart<'_>,
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
    limits: Limits,
    allow_inherited_namespace: bool,
) -> Result<PayloadExpandedName> {
    let element_name = element.name();
    let raw_name =
        std::str::from_utf8(element_name.as_ref()).map_err(|error| xml_error(error.to_string()))?;
    if !is_qualified_name(raw_name) {
        return Err(invalid("ink action opaque payload element name is invalid"));
    }
    let namespace = payload_resolved_namespace(
        resolved,
        reader,
        Some(element),
        limits,
        allow_inherited_namespace,
    )?;
    let local = element_name.local_name();
    if local.as_ref().len() > limits.max_scalar_bytes {
        return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
    }
    let mut local_copy = Vec::new();
    local_copy
        .try_reserve_exact(local.as_ref().len())
        .map_err(|_| invalid("ink action opaque payload name allocation failed"))?;
    local_copy.extend_from_slice(local.as_ref());
    Ok(PayloadExpandedName {
        namespace,
        local: local_copy.into_boxed_slice(),
    })
}

fn payload_expanded_end_name<R: std::io::BufRead>(
    element: &quick_xml::events::BytesEnd<'_>,
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
    limits: Limits,
    allow_inherited_namespace: bool,
) -> Result<PayloadExpandedName> {
    let element_name = element.name();
    let raw_name =
        std::str::from_utf8(element_name.as_ref()).map_err(|error| xml_error(error.to_string()))?;
    if !is_qualified_name(raw_name) {
        return Err(invalid("ink action opaque payload closing name is invalid"));
    }
    let namespace =
        payload_resolved_namespace(resolved, reader, None, limits, allow_inherited_namespace)?;
    let local = element_name.local_name();
    if local.as_ref().len() > limits.max_scalar_bytes {
        return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
    }
    let mut local_copy = Vec::new();
    local_copy
        .try_reserve_exact(local.as_ref().len())
        .map_err(|_| invalid("ink action opaque payload name allocation failed"))?;
    local_copy.extend_from_slice(local.as_ref());
    Ok(PayloadExpandedName {
        namespace,
        local: local_copy.into_boxed_slice(),
    })
}

fn payload_resolved_namespace<R: std::io::BufRead>(
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
    element: Option<&quick_xml::events::BytesStart<'_>>,
    limits: Limits,
    allow_inherited_namespace: bool,
) -> Result<Option<Box<str>>> {
    match resolved {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(value)) => {
            let value = payload_decode_namespace(value, reader, limits)?;
            Ok(Some(value))
        },
        ResolveResult::Unknown(prefix)
            if allow_inherited_namespace
                && element.is_none_or(|element| {
                    !payload_has_namespace_declaration(element, prefix.as_slice())
                }) =>
        {
            let namespace = match prefix.as_slice() {
                b"iact" => ACTION_NAMESPACE,
                b"inkml" => super::INKML_NAMESPACE,
                _ => {
                    return Err(invalid(
                        "ink action opaque payload uses an undeclared namespace prefix",
                    ));
                },
            };
            Ok(Some(namespace.into()))
        },
        ResolveResult::Unknown(_) => Err(invalid(
            "ink action opaque payload uses an undeclared namespace prefix",
        )),
    }
}

fn payload_decode_namespace<R: std::io::BufRead>(
    value: &[u8],
    reader: &NsReader<R>,
    limits: Limits,
) -> Result<Box<str>> {
    let decoded = reader
        .decoder()
        .decode(value)
        .map_err(|error| xml_error(error.to_string()))?;
    let unescaped =
        quick_xml::escape::unescape(&decoded).map_err(|error| xml_error(error.to_string()))?;
    if unescaped.len() > limits.max_scalar_bytes {
        return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
    }
    super::validate_xml_characters(&unescaped, "ink action namespace")?;
    Ok(unescaped.into_owned().into_boxed_str())
}

fn validate_payload_attributes_with_namespaces<R: std::io::BufRead>(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &NsReader<R>,
    limits: Limits,
    allow_inherited_namespace: bool,
) -> Result<()> {
    let mut expanded = HashSet::new();
    // Whether the tag declares an attribute's unresolved prefix is a question
    // about the whole tag. Read its declarations once, on the first such
    // attribute, rather than scanning the tag again for each of them.
    let mut declared_prefixes = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        let raw_key = attribute.key.as_ref();
        let key = std::str::from_utf8(raw_key).map_err(|error| xml_error(error.to_string()))?;
        if !is_qualified_name(key) {
            return Err(invalid(
                "ink action opaque payload attribute name is invalid",
            ));
        }
        if raw_key.len() > limits.max_scalar_bytes
            || attribute.value.len() > limits.max_scalar_bytes
        {
            return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
        }
        if attribute.value.contains(&b'<') {
            return Err(invalid(
                "ink action opaque attribute contains a raw '<' delimiter",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| xml_error(error.to_string()))?;
        if value.len() > limits.max_scalar_bytes {
            return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
        }
        super::validate_xml_characters(&value, "ink action opaque attribute")?;

        let namespace_declaration =
            raw_key == b"xmlns" || raw_key.strip_prefix(b"xmlns:").is_some();
        if namespace_declaration {
            let prefix = if raw_key == b"xmlns" {
                &[][..]
            } else {
                raw_key.strip_prefix(b"xmlns:").unwrap_or_default()
            };
            if prefix == b"xmlns"
                || value == "http://www.w3.org/2000/xmlns/"
                || (value == crate::ink::XML_NAMESPACE && prefix != b"xml")
                || (prefix == b"xml" && value != crate::ink::XML_NAMESPACE)
                || (!prefix.is_empty() && value.is_empty())
            {
                return Err(invalid(
                    "ink action opaque payload has an invalid namespace binding",
                ));
            }
            continue;
        }

        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match resolved {
            ResolveResult::Unbound => None,
            ResolveResult::Bound(Namespace(value)) => {
                Some(payload_decode_namespace(value, reader, limits)?)
            },
            ResolveResult::Unknown(prefix)
                if allow_inherited_namespace
                    && !declared_prefixes
                        .get_or_insert_with(|| payload_declared_prefixes(element))
                        .contains(&prefix.as_slice()) =>
            {
                let namespace = match prefix.as_slice() {
                    b"iact" => ACTION_NAMESPACE,
                    b"inkml" => super::INKML_NAMESPACE,
                    b"xml" => crate::ink::XML_NAMESPACE,
                    _ => {
                        return Err(invalid(
                            "ink action opaque payload attribute uses an undeclared namespace prefix",
                        ));
                    },
                };
                Some(namespace.into())
            },
            ResolveResult::Unknown(_) => {
                return Err(invalid(
                    "ink action opaque payload attribute uses an undeclared namespace prefix",
                ));
            },
        };
        let local = local.as_ref();
        if local.len() > limits.max_scalar_bytes {
            return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
        }
        let mut local_copy = Vec::new();
        local_copy
            .try_reserve_exact(local.len())
            .map_err(|_| invalid("ink action opaque payload attribute allocation failed"))?;
        local_copy.extend_from_slice(local);
        let name = PayloadExpandedName {
            namespace,
            local: local_copy.into_boxed_slice(),
        };
        if !expanded.insert(name) {
            return Err(invalid(
                "ink action opaque payload has duplicate expanded attributes",
            ));
        }
    }
    Ok(())
}

fn payload_namespace<R: std::io::BufRead>(
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
    limits: Limits,
) -> Result<Option<String>> {
    let ResolveResult::Bound(Namespace(value)) = resolved else {
        return Ok(None);
    };
    let decoded = reader
        .decoder()
        .decode(value)
        .map_err(|error| xml_error(error.to_string()))?;
    let unescaped =
        quick_xml::escape::unescape(&decoded).map_err(|error| xml_error(error.to_string()))?;
    if unescaped.len() > limits.max_scalar_bytes {
        return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
    }
    super::validate_xml_characters(&unescaped, "ink action namespace")?;
    Ok(Some(unescaped.into_owned()))
}

fn validate_payload_text<'a>(value: &'a [u8], what: &str) -> Result<&'a str> {
    let value = std::str::from_utf8(value).map_err(|error| xml_error(error.to_string()))?;
    super::validate_xml_characters(value, what)?;
    Ok(value)
}

fn validate_payload_reference(reference: &BytesRef<'_>) -> Result<()> {
    let value =
        std::str::from_utf8(reference.as_ref()).map_err(|error| xml_error(error.to_string()))?;
    match value {
        "amp" | "lt" | "gt" | "apos" | "quot" => Ok(()),
        value if value.strip_prefix("#x").is_some() => {
            let digits = value.strip_prefix("#x").unwrap_or_default();
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(invalid(
                    "ink action hexadecimal character reference is invalid",
                ));
            }
            let codepoint = u32::from_str_radix(digits, 16)
                .map_err(|_| invalid("ink action hexadecimal character reference is invalid"))?;
            validate_payload_reference_codepoint(codepoint)
        },
        value if value.strip_prefix('#').is_some() => {
            let digits = value.strip_prefix('#').unwrap_or_default();
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid("ink action decimal character reference is invalid"));
            }
            let codepoint = digits
                .parse::<u32>()
                .map_err(|_| invalid("ink action decimal character reference is invalid"))?;
            validate_payload_reference_codepoint(codepoint)
        },
        _ => Err(invalid(
            "ink action general entity references are not supported",
        )),
    }
}

fn validate_payload_reference_codepoint(codepoint: u32) -> Result<()> {
    let character = char::from_u32(codepoint)
        .ok_or_else(|| invalid("ink action character reference is invalid"))?;
    if matches!(
        character,
        '\u{9}'
            | '\u{A}'
            | '\u{D}'
            | '\u{20}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
    ) {
        Ok(())
    } else {
        Err(invalid(
            "ink action character reference is not an XML 1.0 character",
        ))
    }
}

fn validate_payload_element(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
    limits: Limits,
) -> Result<()> {
    if element.name().as_ref().len() > limits.max_scalar_bytes {
        return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
    }
    let mut attributes = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        attributes = attributes.checked_add(1).ok_or_else(|| {
            limit(
                "ink action XML attributes",
                super::MAX_ATTRIBUTES_PER_ELEMENT,
            )
        })?;
        if attributes > super::MAX_ATTRIBUTES_PER_ELEMENT {
            return Err(limit(
                "ink action XML attributes",
                super::MAX_ATTRIBUTES_PER_ELEMENT,
            ));
        }
        if attribute.key.as_ref().len() > limits.max_scalar_bytes
            || attribute.value.len() > limits.max_scalar_bytes
        {
            return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
        }
        if attribute.value.contains(&b'<') {
            return Err(invalid(
                "ink action opaque attribute contains a raw '<' delimiter",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| xml_error(error.to_string()))?;
        if value.len() > limits.max_scalar_bytes {
            return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
        }
        super::validate_xml_characters(&value, "ink action opaque attribute")?;
    }
    Ok(())
}

fn bounded_scalar(value: &str, field: &'static str, limits: Limits) -> Result<()> {
    if value.len() > limits.max_scalar_bytes {
        return Err(limit(field, limits.max_scalar_bytes));
    }
    super::validate_xml_characters(value, field)
}

fn draft_len(draft: &Draft) -> Result<usize> {
    let mut len = 0usize;
    add_len(&mut len, b"<iact:actions xmlns:iact=\"")?;
    add_len(&mut len, ACTION_NAMESPACE.as_bytes())?;
    add_len(&mut len, b"\" xmlns:inkml=\"")?;
    add_len(&mut len, super::INKML_NAMESPACE.as_bytes())?;
    add_len(&mut len, b"\"")?;
    attr_len(&mut len, b"lengthUnit", draft.length_unit.as_str())?;
    attr_len(&mut len, b"timeUnit", draft.time_unit.as_str())?;
    if let Some(id) = &draft.xml_id {
        attr_len(&mut len, b"xml:id", id)?;
    }
    add_len(&mut len, b">")?;
    if let Some(payload) = &draft.definitions {
        add_len(&mut len, &payload.0)?;
    }
    for child in &draft.children {
        match child {
            RootDraft::Action(action) => action_len(&mut len, action)?,
            RootDraft::ActionGroup(group) => group_len(&mut len, group)?,
        }
    }
    add_len(&mut len, b"</iact:actions>")?;
    Ok(len)
}

fn action_len(len: &mut usize, action: &ActionDraft) -> Result<()> {
    add_len(len, b"<iact:action")?;
    if let Some(id) = &action.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"type", action.action_type.as_str())?;
    attr_len(len, b"startTime", &action.start_time)?;
    add_len(len, b">")?;
    for child in &action.children {
        match child {
            ActionChildDraft::Property(property) => property_len(len, property)?,
            ActionChildDraft::Data(data) => data_len(len, data)?,
            ActionChildDraft::DataGroup(group) => data_group_len(len, group)?,
        }
    }
    add_len(len, b"</iact:action>")
}

fn action_len_with_prefix(len: &mut usize, action: &ActionDraft, prefix: &[u8]) -> Result<()> {
    qname_len(len, prefix, b"action", false)?;
    if let Some(id) = &action.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"type", action.action_type.as_str())?;
    attr_len(len, b"startTime", &action.start_time)?;
    add_len(len, b">")?;
    for child in &action.children {
        match child {
            ActionChildDraft::Property(property) => {
                property_len_with_prefix(len, property, prefix)?
            },
            ActionChildDraft::Data(data) => data_len_with_prefix(len, data, prefix)?,
            ActionChildDraft::DataGroup(group) => data_group_len_with_prefix(len, group, prefix)?,
        }
    }
    qname_len(len, prefix, b"action", true)?;
    add_len(len, b">")
}

fn property_len_with_prefix(
    len: &mut usize,
    property: &PropertyDraft,
    prefix: &[u8],
) -> Result<()> {
    qname_len(len, prefix, b"property", false)?;
    attr_len(len, b"name", &property.name)?;
    attr_len(len, b"value", &property.value)?;
    add_len(len, b"/>")
}

fn data_len_with_prefix(len: &mut usize, data: &ActionDataDraft, prefix: &[u8]) -> Result<()> {
    qname_len(len, prefix, b"actionData", false)?;
    if let Some(id) = &data.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"name", &data.name)?;
    if let Some(reference) = &data.reference {
        attr_len(len, b"ref", reference)?;
    }
    add_len(len, b">")?;
    for child in &data.children {
        match child {
            DataChildDraft::Transform(payload)
            | DataChildDraft::Trace(payload)
            | DataChildDraft::TraceView(payload) => add_len(len, &payload.0)?,
        }
    }
    qname_len(len, prefix, b"actionData", true)?;
    add_len(len, b">")
}

fn data_group_len_with_prefix(
    len: &mut usize,
    group: &DataGroupDraft,
    prefix: &[u8],
) -> Result<()> {
    qname_len(len, prefix, b"actionDataGroup", false)?;
    if let Some(id) = &group.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"name", &group.name)?;
    add_len(len, b">")?;
    for data in &group.data {
        data_len_with_prefix(len, data, prefix)?;
    }
    qname_len(len, prefix, b"actionDataGroup", true)?;
    add_len(len, b">")
}

fn qname_len(len: &mut usize, prefix: &[u8], local: &[u8], closing: bool) -> Result<()> {
    if closing {
        add_len(len, b"</")?;
    } else {
        add_len(len, b"<")?;
    }
    if !prefix.is_empty() {
        add_len(len, prefix)?;
        add_len(len, b":")?;
    }
    add_len(len, local)
}

fn group_len(len: &mut usize, group: &ActionGroupDraft) -> Result<()> {
    add_len(len, b"<iact:actionGroup")?;
    if let Some(id) = &group.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"type", group.action_type.as_str())?;
    attr_len(len, b"startTime", &group.start_time)?;
    add_len(len, b">")?;
    for action in &group.actions {
        action_len(len, action)?;
    }
    add_len(len, b"</iact:actionGroup>")
}

fn group_len_with_prefix(len: &mut usize, group: &ActionGroupDraft, prefix: &[u8]) -> Result<()> {
    qname_len(len, prefix, b"actionGroup", false)?;
    if let Some(id) = &group.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"type", group.action_type.as_str())?;
    attr_len(len, b"startTime", &group.start_time)?;
    add_len(len, b">")?;
    for action in &group.actions {
        action_len_with_prefix(len, action, prefix)?;
    }
    qname_len(len, prefix, b"actionGroup", true)?;
    add_len(len, b">")
}

fn property_len(len: &mut usize, property: &PropertyDraft) -> Result<()> {
    add_len(len, b"<iact:property")?;
    attr_len(len, b"name", &property.name)?;
    attr_len(len, b"value", &property.value)?;
    add_len(len, b"/>")
}

fn data_len(len: &mut usize, data: &ActionDataDraft) -> Result<()> {
    add_len(len, b"<iact:actionData")?;
    if let Some(id) = &data.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"name", &data.name)?;
    if let Some(reference) = &data.reference {
        attr_len(len, b"ref", reference)?;
    }
    add_len(len, b">")?;
    for child in &data.children {
        match child {
            DataChildDraft::Transform(payload)
            | DataChildDraft::Trace(payload)
            | DataChildDraft::TraceView(payload) => add_len(len, &payload.0)?,
        }
    }
    add_len(len, b"</iact:actionData>")
}

fn data_group_len(len: &mut usize, group: &DataGroupDraft) -> Result<()> {
    add_len(len, b"<iact:actionDataGroup")?;
    if let Some(id) = &group.xml_id {
        attr_len(len, b"xml:id", id)?;
    }
    attr_len(len, b"name", &group.name)?;
    add_len(len, b">")?;
    for data in &group.data {
        data_len(len, data)?;
    }
    add_len(len, b"</iact:actionDataGroup>")
}

fn emit_draft(output: &mut Vec<u8>, draft: &Draft) -> Result<()> {
    output.extend_from_slice(b"<iact:actions xmlns:iact=\"");
    output.extend_from_slice(ACTION_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\" xmlns:inkml=\"");
    output.extend_from_slice(super::INKML_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\"");
    emit_attr(output, b"lengthUnit", draft.length_unit.as_str());
    emit_attr(output, b"timeUnit", draft.time_unit.as_str());
    if let Some(id) = &draft.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    output.push(b'>');
    if let Some(payload) = &draft.definitions {
        output.extend_from_slice(&payload.0);
    }
    for child in &draft.children {
        match child {
            RootDraft::Action(action) => emit_action(output, action),
            RootDraft::ActionGroup(group) => emit_group(output, group),
        }
    }
    output.extend_from_slice(b"</iact:actions>");
    Ok(())
}

fn emit_action(output: &mut Vec<u8>, action: &ActionDraft) {
    output.extend_from_slice(b"<iact:action");
    if let Some(id) = &action.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"type", action.action_type.as_str());
    emit_attr(output, b"startTime", &action.start_time);
    output.push(b'>');
    for child in &action.children {
        match child {
            ActionChildDraft::Property(property) => {
                output.extend_from_slice(b"<iact:property");
                emit_attr(output, b"name", &property.name);
                emit_attr(output, b"value", &property.value);
                output.extend_from_slice(b"/>");
            },
            ActionChildDraft::Data(data) => emit_data(output, data),
            ActionChildDraft::DataGroup(group) => emit_data_group(output, group),
        }
    }
    output.extend_from_slice(b"</iact:action>");
}

fn emit_action_with_prefix(output: &mut Vec<u8>, action: &ActionDraft, prefix: &[u8]) {
    emit_qname(output, prefix, b"action", false);
    if let Some(id) = &action.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"type", action.action_type.as_str());
    emit_attr(output, b"startTime", &action.start_time);
    output.push(b'>');
    for child in &action.children {
        match child {
            ActionChildDraft::Property(property) => {
                emit_qname(output, prefix, b"property", false);
                emit_attr(output, b"name", &property.name);
                emit_attr(output, b"value", &property.value);
                output.extend_from_slice(b"/>");
            },
            ActionChildDraft::Data(data) => emit_data_with_prefix(output, data, prefix),
            ActionChildDraft::DataGroup(group) => {
                emit_data_group_with_prefix(output, group, prefix);
            },
        }
    }
    emit_qname(output, prefix, b"action", true);
    output.push(b'>');
}

fn emit_data_with_prefix(output: &mut Vec<u8>, data: &ActionDataDraft, prefix: &[u8]) {
    emit_qname(output, prefix, b"actionData", false);
    if let Some(id) = &data.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"name", &data.name);
    if let Some(reference) = &data.reference {
        emit_attr(output, b"ref", reference);
    }
    output.push(b'>');
    for child in &data.children {
        match child {
            DataChildDraft::Transform(payload)
            | DataChildDraft::Trace(payload)
            | DataChildDraft::TraceView(payload) => output.extend_from_slice(&payload.0),
        }
    }
    emit_qname(output, prefix, b"actionData", true);
    output.push(b'>');
}

fn emit_data_group_with_prefix(output: &mut Vec<u8>, group: &DataGroupDraft, prefix: &[u8]) {
    emit_qname(output, prefix, b"actionDataGroup", false);
    if let Some(id) = &group.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"name", &group.name);
    output.push(b'>');
    for data in &group.data {
        emit_data_with_prefix(output, data, prefix);
    }
    emit_qname(output, prefix, b"actionDataGroup", true);
    output.push(b'>');
}

fn emit_qname(output: &mut Vec<u8>, prefix: &[u8], local: &[u8], closing: bool) {
    if closing {
        output.extend_from_slice(b"</");
    } else {
        output.push(b'<');
    }
    if !prefix.is_empty() {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
}

fn emit_group(output: &mut Vec<u8>, group: &ActionGroupDraft) {
    output.extend_from_slice(b"<iact:actionGroup");
    if let Some(id) = &group.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"type", group.action_type.as_str());
    emit_attr(output, b"startTime", &group.start_time);
    output.push(b'>');
    for action in &group.actions {
        emit_action(output, action);
    }
    output.extend_from_slice(b"</iact:actionGroup>");
}

fn emit_group_with_prefix(output: &mut Vec<u8>, group: &ActionGroupDraft, prefix: &[u8]) {
    emit_qname(output, prefix, b"actionGroup", false);
    if let Some(id) = &group.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"type", group.action_type.as_str());
    emit_attr(output, b"startTime", &group.start_time);
    output.push(b'>');
    for action in &group.actions {
        emit_action_with_prefix(output, action, prefix);
    }
    emit_qname(output, prefix, b"actionGroup", true);
    output.push(b'>');
}

fn emit_data(output: &mut Vec<u8>, data: &ActionDataDraft) {
    output.extend_from_slice(b"<iact:actionData");
    if let Some(id) = &data.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"name", &data.name);
    if let Some(reference) = &data.reference {
        emit_attr(output, b"ref", reference);
    }
    output.push(b'>');
    for child in &data.children {
        match child {
            DataChildDraft::Transform(payload)
            | DataChildDraft::Trace(payload)
            | DataChildDraft::TraceView(payload) => output.extend_from_slice(&payload.0),
        }
    }
    output.extend_from_slice(b"</iact:actionData>");
}

fn emit_data_group(output: &mut Vec<u8>, group: &DataGroupDraft) {
    output.extend_from_slice(b"<iact:actionDataGroup");
    if let Some(id) = &group.xml_id {
        emit_attr(output, b"xml:id", id);
    }
    emit_attr(output, b"name", &group.name);
    output.push(b'>');
    for data in &group.data {
        emit_data(output, data);
    }
    output.extend_from_slice(b"</iact:actionDataGroup>");
}

fn add_len(total: &mut usize, bytes: &[u8]) -> Result<()> {
    *total = total
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("ink action output length overflow"))?;
    Ok(())
}

fn attr_len(total: &mut usize, name: &[u8], value: &str) -> Result<()> {
    add_len(total, name)?;
    *total = total
        .checked_add(4)
        .ok_or_else(|| invalid("ink action output length overflow"))?;
    *total = total
        .checked_add(escaped_len(value)?)
        .ok_or_else(|| invalid("ink action output length overflow"))?;
    Ok(())
}

fn escaped_len(value: &str) -> Result<usize> {
    let mut len = 0usize;
    for character in value.chars() {
        let amount = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\'' => 6,
            '\t' => 5,
            '\n' => 5,
            '\r' => 5,
            _ => character.len_utf8(),
        };
        len = len
            .checked_add(amount)
            .ok_or_else(|| invalid("ink action output length overflow"))?;
    }
    Ok(len)
}

fn emit_attr(output: &mut Vec<u8>, name: &[u8], value: &str) {
    output.push(b' ');
    output.extend_from_slice(name);
    output.extend_from_slice(b"=\"");
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\'' => output.extend_from_slice(b"&apos;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut buffer = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
    output.push(b'"');
}

fn matches_draft(draft: &Draft, profile: &Profile) -> bool {
    if profile.xml_id.as_deref() != draft.xml_id.as_deref()
        || profile.length_unit != draft.length_unit
        || profile.time_unit != draft.time_unit
        || profile.definitions.as_ref().map(|span| profile.xml(*span))
            != draft.definitions.as_ref().map(|payload| payload.as_bytes())
        || profile.children.len() != draft.children.len()
    {
        return false;
    }
    draft
        .children
        .iter()
        .zip(&profile.children)
        .all(|(draft, child)| match (draft, child) {
            (RootDraft::Action(left), RootChild::Action(right)) => {
                matches_action_draft(left, right, profile)
            },
            (RootDraft::ActionGroup(left), RootChild::ActionGroup(right)) => {
                matches_group_draft(left, right, profile)
            },
            _ => false,
        })
}

fn matches_action_draft(draft: &ActionDraft, action: &Action, profile: &Profile) -> bool {
    if action.xml_id.as_deref() != draft.xml_id.as_deref()
        || action.action_type != draft.action_type
        || action.start_time.as_ref() != draft.start_time.as_ref()
        || action.children.len() != draft.children.len()
    {
        return false;
    }
    let mut properties = action.properties.iter();
    for (left, right) in draft.children.iter().zip(&action.children) {
        match (left, right) {
            (ActionChildDraft::Property(property), ActionChild::Property(actual)) => {
                if properties.next() != Some(actual)
                    || actual.name != property.name
                    || actual.value != property.value
                {
                    return false;
                }
            },
            (ActionChildDraft::Data(data), ActionChild::Data(actual)) => {
                if !matches_data_draft(data, actual, profile) {
                    return false;
                }
            },
            (ActionChildDraft::DataGroup(group), ActionChild::DataGroup(actual)) => {
                if !matches_group_data_draft(group, actual, profile) {
                    return false;
                }
            },
            _ => return false,
        }
    }
    true
}

fn matches_group_draft(draft: &ActionGroupDraft, group: &ActionGroup, profile: &Profile) -> bool {
    group.xml_id.as_deref() == draft.xml_id.as_deref()
        && group.action_type == draft.action_type
        && group.start_time.as_ref() == draft.start_time.as_ref()
        && group.actions.len() == draft.actions.len()
        && draft
            .actions
            .iter()
            .zip(&group.actions)
            .all(|(left, right)| matches_action_draft(left, right, profile))
}

fn matches_data_draft(draft: &ActionDataDraft, data: &ActionData, profile: &Profile) -> bool {
    data.xml_id.as_deref() == draft.xml_id.as_deref()
        && data.name.as_ref() == draft.name.as_ref()
        && data.reference.as_deref() == draft.reference.as_deref()
        && data.children.len() == draft.children.len()
        && draft
            .children
            .iter()
            .zip(&data.children)
            .all(|(left, right)| {
                let (payload, span) = match (left, right) {
                    (DataChildDraft::Transform(payload), DataChild::Transform(span))
                    | (DataChildDraft::Trace(payload), DataChild::Trace(span))
                    | (DataChildDraft::TraceView(payload), DataChild::TraceView(span)) => {
                        (payload, *span)
                    },
                    _ => return false,
                };
                profile.xml(span) == payload.as_bytes()
            })
}

fn matches_group_data_draft(
    draft: &DataGroupDraft,
    group: &ActionDataGroup,
    profile: &Profile,
) -> bool {
    group.xml_id.as_deref() == draft.xml_id.as_deref()
        && group.name.as_ref() == draft.name.as_ref()
        && group.data.len() == draft.data.len()
        && draft
            .data
            .iter()
            .zip(&group.data)
            .all(|(left, right)| matches_data_draft(left, right, profile))
}

fn locate_action(profile: &Profile, selector: ActionSelector) -> Result<&Action> {
    match selector {
        ActionSelector::Direct(index) => profile
            .children
            .iter()
            .filter_map(|child| match child {
                RootChild::Action(action) => Some(action),
                RootChild::ActionGroup(_) => None,
            })
            .nth(index)
            .ok_or_else(|| invalid("ink action direct selector is out of range")),
        ActionSelector::Group { group, action } => profile
            .children
            .iter()
            .filter_map(|child| match child {
                RootChild::ActionGroup(group) => Some(group),
                RootChild::Action(_) => None,
            })
            .nth(group)
            .and_then(|group| group.actions.get(action))
            .ok_or_else(|| invalid("ink action group selector is out of range")),
        ActionSelector::Ordinal(mut index) => {
            for child in &profile.children {
                match child {
                    RootChild::Action(action) => {
                        if index == 0 {
                            return Ok(action);
                        }
                        index -= 1;
                    },
                    RootChild::ActionGroup(group) => {
                        if let Some(action) = group.actions.get(index) {
                            return Ok(action);
                        }
                        index = index.saturating_sub(group.actions.len());
                    },
                }
            }
            Err(invalid("ink action ordinal selector is out of range"))
        },
    }
}

fn action_span(action: &Action) -> SourceSpan {
    SourceSpan::new(action.source_start, action.source_end)
}

fn locate_group(profile: &Profile, index: usize) -> Result<&ActionGroup> {
    profile
        .children
        .iter()
        .filter_map(|child| match child {
            RootChild::ActionGroup(group) => Some(group),
            RootChild::Action(_) => None,
        })
        .nth(index)
        .ok_or_else(|| invalid("ink action group selector is out of range"))
}

fn locate_property(profile: &Profile, selector: ChildSelector) -> Result<&ActionProperty> {
    let ChildSelector::Property { action, index } = selector else {
        return Err(invalid("ink action selector is not a property"));
    };
    locate_action(profile, action)?
        .properties
        .get(index)
        .ok_or_else(|| invalid("ink action property selector is out of range"))
}

fn locate_data(profile: &Profile, selector: DataSelector) -> Result<&ActionData> {
    match selector {
        DataSelector::Action { action, index } => {
            let action = locate_action(profile, action)?;
            action
                .children
                .iter()
                .filter_map(|child| match child {
                    ActionChild::Data(data) => Some(data),
                    _ => None,
                })
                .nth(index)
                .ok_or_else(|| invalid("ink action data selector is out of range"))
        },
        DataSelector::Group {
            action,
            group,
            index,
        } => {
            let action = locate_action(profile, action)?;
            action
                .children
                .iter()
                .filter_map(|child| match child {
                    ActionChild::DataGroup(group) => Some(group),
                    _ => None,
                })
                .nth(group)
                .and_then(|group| group.data.get(index))
                .ok_or_else(|| invalid("ink action grouped-data selector is out of range"))
        },
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ParentLocation {
    Root,
    Group(usize),
}

#[derive(Clone, Copy)]
struct ActionLocation<'a> {
    action: &'a Action,
    parent: ParentLocation,
}

fn locate_action_with_parent<'a>(
    profile: &'a Profile,
    selector: ActionSelector,
) -> Result<ActionLocation<'a>> {
    match selector {
        ActionSelector::Direct(index) => {
            let mut seen = 0;
            for child in &profile.children {
                if let RootChild::Action(action) = child {
                    if seen == index {
                        return Ok(ActionLocation {
                            action,
                            parent: ParentLocation::Root,
                        });
                    }
                    seen += 1;
                }
            }
        },
        ActionSelector::Group { group, action } => {
            let group_value = locate_group(profile, group)?;
            if let Some(action) = group_value.actions.get(action) {
                return Ok(ActionLocation {
                    action,
                    parent: ParentLocation::Group(group),
                });
            }
        },
        ActionSelector::Ordinal(mut ordinal) => {
            for child in &profile.children {
                match child {
                    RootChild::Action(action) => {
                        if ordinal == 0 {
                            return Ok(ActionLocation {
                                action,
                                parent: ParentLocation::Root,
                            });
                        }
                        ordinal -= 1;
                    },
                    RootChild::ActionGroup(group) => {
                        if let Some(action) = group.actions.get(ordinal) {
                            return Ok(ActionLocation {
                                action,
                                parent: ParentLocation::Group(group_index(profile, group)?),
                            });
                        }
                        ordinal = ordinal.saturating_sub(group.actions.len());
                    },
                }
            }
        },
    }
    Err(invalid("ink action selector is out of range"))
}

fn group_index(profile: &Profile, wanted: &ActionGroup) -> Result<usize> {
    let mut index = 0;
    for child in &profile.children {
        if let RootChild::ActionGroup(group) = child {
            if std::ptr::eq(group, wanted) {
                return Ok(index);
            }
            index += 1;
        }
    }
    Err(invalid("ink action group selector is out of range"))
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
enum ScalarKey {
    ActionType(SourceSpan),
    StartTime(SourceSpan),
    ActionGroupType(SourceSpan),
    ActionGroupStartTime(SourceSpan),
    PropertyName(SourceSpan),
    PropertyValue(SourceSpan),
    DataName(SourceSpan),
    DataReference(SourceSpan),
    DataGroupName(SourceSpan),
}

fn scalar_key(profile: &Profile, operation: &EditOp) -> Result<Option<ScalarKey>> {
    Ok(match operation {
        EditOp::SetActionType(selector, _) => Some(ScalarKey::ActionType(action_span(
            locate_action(profile, *selector)?,
        ))),
        EditOp::SetStartTime(selector, _) => Some(ScalarKey::StartTime(action_span(
            locate_action(profile, *selector)?,
        ))),
        EditOp::SetActionGroupType(group, _) => Some(ScalarKey::ActionGroupType(
            locate_group(profile, *group)?.source,
        )),
        EditOp::SetActionGroupStartTime(group, _) => Some(ScalarKey::ActionGroupStartTime(
            locate_group(profile, *group)?.source,
        )),
        EditOp::SetPropertyName(selector, _) => Some(ScalarKey::PropertyName(
            locate_property(profile, *selector)?.source,
        )),
        EditOp::SetPropertyValue(selector, _) => Some(ScalarKey::PropertyValue(
            locate_property(profile, *selector)?.source,
        )),
        EditOp::SetDataName(selector, _) => {
            Some(ScalarKey::DataName(locate_data(profile, *selector)?.source))
        },
        EditOp::SetDataReference(selector, _) => Some(ScalarKey::DataReference(
            locate_data(profile, *selector)?.source,
        )),
        EditOp::SetDataGroupName(selector, _) => Some(ScalarKey::DataGroupName(
            locate_data_group(profile, *selector)?.source,
        )),
        _ => None,
    })
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ChildKind {
    Property,
    Data,
    DataGroup,
}

#[derive(Clone, Copy)]
struct ChildLocation {
    span: SourceSpan,
    action_span: SourceSpan,
    kind: ChildKind,
}

#[derive(Clone, Copy)]
struct GroupDataLocation {
    span: SourceSpan,
    action_span: SourceSpan,
    group_span: SourceSpan,
}

fn locate_child(profile: &Profile, selector: ChildSelector) -> Result<ChildLocation> {
    match selector {
        ChildSelector::Property { action, index } => {
            let action_value = locate_action(profile, action)?;
            let child = action_value
                .properties
                .get(index)
                .ok_or_else(|| invalid("ink action property selector is out of range"))?;
            Ok(ChildLocation {
                span: child.source,
                action_span: action_span(action_value),
                kind: ChildKind::Property,
            })
        },
        ChildSelector::Data { action, index } => {
            let action_value = locate_action(profile, action)?;
            let child = action_value
                .children
                .iter()
                .filter_map(|child| match child {
                    ActionChild::Data(data) => Some(data),
                    _ => None,
                })
                .nth(index)
                .ok_or_else(|| invalid("ink action data selector is out of range"))?;
            Ok(ChildLocation {
                span: child.source,
                action_span: action_span(action_value),
                kind: ChildKind::Data,
            })
        },
        ChildSelector::DataGroup { action, index } => {
            let action_value = locate_action(profile, action)?;
            let child = action_value
                .children
                .iter()
                .filter_map(|child| match child {
                    ActionChild::DataGroup(group) => Some(group),
                    _ => None,
                })
                .nth(index)
                .ok_or_else(|| invalid("ink action data-group selector is out of range"))?;
            Ok(ChildLocation {
                span: child.source,
                action_span: action_span(action_value),
                kind: ChildKind::DataGroup,
            })
        },
        ChildSelector::GroupData { .. } => Err(invalid(
            "ink action grouped data needs a grouped-data selector",
        )),
    }
}

fn locate_data_group(profile: &Profile, selector: ChildSelector) -> Result<&ActionDataGroup> {
    let ChildSelector::DataGroup { action, index } = selector else {
        return Err(invalid("ink action selector is not a data group"));
    };
    let action = locate_action(profile, action)?;
    action
        .children
        .iter()
        .filter_map(|child| match child {
            ActionChild::DataGroup(group) => Some(group),
            _ => None,
        })
        .nth(index)
        .ok_or_else(|| invalid("ink action data-group selector is out of range"))
}

fn locate_group_data(profile: &Profile, selector: DataSelector) -> Result<GroupDataLocation> {
    let DataSelector::Group {
        action,
        group,
        index,
    } = selector
    else {
        return Err(invalid("ink action selector is not grouped data"));
    };
    let action_value = locate_action(profile, action)?;
    let group_value = action_value
        .children
        .iter()
        .filter_map(|child| match child {
            ActionChild::DataGroup(group) => Some(group),
            _ => None,
        })
        .nth(group)
        .ok_or_else(|| invalid("ink action data-group selector is out of range"))?;
    let data = group_value
        .data
        .get(index)
        .ok_or_else(|| invalid("ink action grouped-data selector is out of range"))?;
    Ok(GroupDataLocation {
        span: data.source,
        action_span: action_span(action_value),
        group_span: group_value.source,
    })
}

fn group_span(profile: &Profile, index: usize) -> Result<SourceSpan> {
    Ok(locate_group(profile, index)?.source)
}

fn insertion_for_action(
    source: &[u8],
    action: &Action,
    kind: ChildKind,
) -> Result<(usize, Option<Box<[u8]>>)> {
    let span = action_span(action);
    let tag = start_tag(source, span)?;
    if tag.self_closing {
        let name = element_name(source, tag)?;
        return Ok((span.start(), Some(name)));
    }
    let (_, close) = inner_range(source, span)?
        .ok_or_else(|| invalid("ink action child insertion range is invalid"))?;
    let position = match kind {
        ChildKind::Property => action
            .children
            .iter()
            .find_map(|child| match child {
                ActionChild::Data(data) => Some(data.source.start()),
                ActionChild::DataGroup(group) => Some(group.source.start()),
                ActionChild::Property(_) => None,
            })
            .unwrap_or(close),
        ChildKind::Data | ChildKind::DataGroup => close,
    };
    Ok((position, None))
}

fn element_name(source: &[u8], tag: TagRange) -> Result<Box<[u8]>> {
    let mut end = tag.start + 1;
    while end < tag.end
        && !source[end].is_ascii_whitespace()
        && source[end] != b'>'
        && source[end] != b'/'
    {
        end += 1;
    }
    let name = source
        .get(tag.start + 1..end)
        .ok_or_else(|| invalid("ink action element name range is invalid"))?;
    if name.is_empty() || name.len() > super::MAX_TOKEN_BYTES {
        return Err(limit(
            "ink action element name bytes",
            super::MAX_TOKEN_BYTES,
        ));
    }
    let mut result = Vec::new();
    result
        .try_reserve_exact(name.len())
        .map_err(|_| invalid("ink action element name allocation failed"))?;
    result.extend_from_slice(name);
    Ok(result.into_boxed_slice())
}

enum InsertKind<'a> {
    Action(&'a ActionDraft, Arc<[u8]>),
    Group(&'a ActionGroupDraft, Arc<[u8]>),
    Property(&'a PropertyDraft, Arc<[u8]>),
    Data(&'a ActionDataDraft, Arc<[u8]>),
    DataGroup(&'a DataGroupDraft, Arc<[u8]>),
    Many(Vec<InsertKind<'a>>),
}

enum PlanKind<'a> {
    Remove,
    Escaped(&'a str),
    Attribute(&'static [u8], &'a str),
    Insert(InsertKind<'a>),
    ExpandSelfClosing {
        name: Box<[u8]>,
        child: InsertKind<'a>,
    },
    Move {
        from_start: usize,
        from_end: usize,
        before_start: usize,
        before_end: usize,
    },
}

struct Plan<'a> {
    start: usize,
    end: usize,
    inserted_len: usize,
    order: usize,
    nested: bool,
    suppressed: bool,
    kind: PlanKind<'a>,
}

fn insert_len(kind: &InsertKind<'_>) -> Result<usize> {
    let mut length = 0usize;
    match kind {
        InsertKind::Action(action, prefix) => {
            action_len_with_prefix(&mut length, action, prefix.as_ref())?
        },
        InsertKind::Group(group, prefix) => {
            group_len_with_prefix(&mut length, group, prefix.as_ref())?
        },
        InsertKind::Property(property, prefix) => {
            property_len_with_prefix(&mut length, property, prefix.as_ref())?
        },
        InsertKind::Data(data, prefix) => data_len_with_prefix(&mut length, data, prefix.as_ref())?,
        InsertKind::DataGroup(group, prefix) => {
            data_group_len_with_prefix(&mut length, group, prefix.as_ref())?
        },
        InsertKind::Many(children) => {
            for child in children {
                length = length
                    .checked_add(insert_len(child)?)
                    .ok_or_else(|| invalid("ink action insertion length overflow"))?;
            }
        },
    }
    Ok(length)
}

fn plan_len(kind: &PlanKind<'_>, _source: &[u8], start: usize, end: usize) -> Result<usize> {
    match kind {
        PlanKind::Remove => Ok(0),
        PlanKind::Escaped(value) => escaped_len(value),
        PlanKind::Attribute(name, value) => attr_len_value(name, value),
        PlanKind::Insert(kind) => insert_len(kind),
        PlanKind::ExpandSelfClosing { name, child } => insert_len(child)?
            .checked_add(name.len())
            .and_then(|length| length.checked_add(3))
            .and_then(|length| length.checked_add(1))
            .and_then(|length| length.checked_add(end.saturating_sub(start + 2)))
            .ok_or_else(|| invalid("ink action replacement length overflow")),
        PlanKind::Move {
            from_start,
            from_end,
            before_start,
            before_end,
        } => from_end
            .checked_sub(*from_start)
            .and_then(|length| length.checked_add(from_start.saturating_sub(*before_end)))
            .and_then(|length| length.checked_add(before_end.saturating_sub(*before_start)))
            .ok_or_else(|| invalid("ink action move length overflow")),
    }
}

fn make_plan<'a>(
    source: &[u8],
    start: usize,
    end: usize,
    order: usize,
    kind: PlanKind<'a>,
) -> Result<Plan<'a>> {
    if start > end || end > source.len() {
        return Err(invalid("ink action replacement range is invalid"));
    }
    let inserted_len = plan_len(&kind, source, start, end)?;
    Ok(Plan {
        start,
        end,
        inserted_len,
        order,
        nested: false,
        suppressed: false,
        kind,
    })
}

fn plan_operations<'a>(
    profile: &Profile,
    source: &'a [u8],
    operations: &'a [EditOp],
    limits: Limits,
) -> Result<Vec<Plan<'a>>> {
    validate_batch_operations(profile, source, operations, limits)?;
    let prefix: Arc<[u8]> = Arc::from(action_prefix(source)?.into_boxed_slice());
    let mut selected = Vec::new();
    selected
        .try_reserve(operations.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    let mut scalar_seen = HashSet::new();
    scalar_seen
        .try_reserve(operations.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for (index, operation) in operations.iter().enumerate().rev() {
        if let Some(key) = scalar_key(profile, operation)? {
            if !scalar_seen.insert(key) {
                continue;
            }
        }
        selected.push((index, operation));
    }
    selected.sort_by_key(|(index, _)| *index);

    let mut plans = Vec::new();
    plans
        .try_reserve(selected.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    let root_range = root_tag(source)?;
    let root_self_closing = root_range.self_closing;
    let root_insert_at = if root_self_closing {
        None
    } else {
        Some(root_close_start(source)?)
    };
    let mut root_insertions: Vec<(usize, InsertKind<'a>)> = Vec::new();
    for (order, operation) in selected {
        let plan = match operation {
            EditOp::SetActionType(selector, value) => {
                let action = locate_action(profile, *selector)?;
                if action.action_type.as_str() == value.as_str() {
                    continue;
                }
                plan_attribute(
                    source,
                    action_span(action),
                    b"type",
                    value.as_str(),
                    order,
                    limits,
                )?
            },
            EditOp::SetStartTime(selector, value) => {
                let action = locate_action(profile, *selector)?;
                if action.start_time.as_ref() == value.as_ref() {
                    continue;
                }
                plan_attribute(
                    source,
                    action_span(action),
                    b"startTime",
                    value,
                    order,
                    limits,
                )?
            },
            EditOp::SetActionGroupType(group, value) => {
                let group = locate_group(profile, *group)?;
                if group.action_type.as_str() == value.as_str() {
                    continue;
                }
                plan_attribute(source, group.source, b"type", value.as_str(), order, limits)?
            },
            EditOp::SetActionGroupStartTime(group, value) => {
                let group = locate_group(profile, *group)?;
                if group.start_time.as_ref() == value.as_ref() {
                    continue;
                }
                plan_attribute(source, group.source, b"startTime", value, order, limits)?
            },
            EditOp::SetPropertyName(selector, value) => {
                let property = locate_property(profile, *selector)?;
                if property.name.as_ref() == value.as_ref() {
                    continue;
                }
                plan_attribute(source, property.source, b"name", value, order, limits)?
            },
            EditOp::SetPropertyValue(selector, value) => {
                let property = locate_property(profile, *selector)?;
                if property.value.as_ref() == value.as_ref() {
                    continue;
                }
                plan_attribute(source, property.source, b"value", value, order, limits)?
            },
            EditOp::SetDataName(selector, value) => {
                let data = locate_data(profile, *selector)?;
                if data.name.as_ref() == value.as_ref() {
                    continue;
                }
                plan_attribute(source, data.source, b"name", value, order, limits)?
            },
            EditOp::SetDataGroupName(selector, value) => {
                let child = locate_child(profile, *selector)?;
                if child.kind != ChildKind::DataGroup {
                    return Err(invalid("ink action selector is not a data group"));
                }
                let group = locate_data_group(profile, *selector)?;
                if group.name.as_ref() == value.as_ref() {
                    continue;
                }
                plan_attribute(source, child.span, b"name", value, order, limits)?
            },
            EditOp::SetDataReference(selector, value) => {
                let data = locate_data(profile, *selector)?;
                if data_reference_semantically_equal(data.reference.as_deref(), value.as_deref())? {
                    continue;
                }
                let Some(plan) =
                    plan_data_reference(source, data.source, value.as_deref(), order, limits)?
                else {
                    continue;
                };
                plan
            },
            EditOp::Add(parent, action) => match parent {
                ActionParent::Root => {
                    if root_self_closing {
                        root_insertions
                            .try_reserve(1)
                            .map_err(|_| invalid("ink action edit allocation failed"))?;
                        root_insertions.push((order, InsertKind::Action(action, prefix.clone())));
                        continue;
                    }
                    plan_root_insert(
                        source,
                        root_range,
                        root_insert_at,
                        InsertKind::Action(action, prefix.clone()),
                        order,
                    )?
                },
                ActionParent::Group(index) => {
                    let at = action_insertion(profile, source, ActionParent::Group(*index))?;
                    make_plan(
                        source,
                        at,
                        at,
                        order,
                        PlanKind::Insert(InsertKind::Action(action, prefix.clone())),
                    )?
                },
            },
            EditOp::AddGroup(group) => {
                if root_self_closing {
                    root_insertions
                        .try_reserve(1)
                        .map_err(|_| invalid("ink action edit allocation failed"))?;
                    root_insertions.push((order, InsertKind::Group(group, prefix.clone())));
                    continue;
                }
                plan_root_insert(
                    source,
                    root_range,
                    root_insert_at,
                    InsertKind::Group(group, prefix.clone()),
                    order,
                )?
            },
            EditOp::AddProperty(ChildParent::Action(selector), property) => {
                let action = locate_action(profile, *selector)?;
                plan_child_insert(
                    source,
                    action,
                    ChildKind::Property,
                    InsertKind::Property(property, prefix.clone()),
                    order,
                )?
            },
            EditOp::AddData(selector, data) => {
                let action = locate_action(profile, *selector)?;
                plan_child_insert(
                    source,
                    action,
                    ChildKind::Data,
                    InsertKind::Data(data, prefix.clone()),
                    order,
                )?
            },
            EditOp::AddDataGroup(selector, group) => {
                let action = locate_action(profile, *selector)?;
                plan_child_insert(
                    source,
                    action,
                    ChildKind::DataGroup,
                    InsertKind::DataGroup(group, prefix.clone()),
                    order,
                )?
            },
            EditOp::MoveBefore(from, before) => {
                let from = locate_action_with_parent(profile, *from)?;
                let before = locate_action_with_parent(profile, *before)?;
                plan_move(
                    source,
                    action_span(from.action),
                    action_span(before.action),
                    from.parent == before.parent,
                    order,
                )?
            },
            EditOp::MoveGroupBefore(from, before) => plan_move(
                source,
                group_span(profile, *from)?,
                group_span(profile, *before)?,
                true,
                order,
            )?,
            EditOp::MoveChild(from, before) => {
                let from = locate_child(profile, *from)?;
                let before = locate_child(profile, *before)?;
                if from.action_span != before.action_span {
                    return Err(invalid("ink action child move requires one parent"));
                }
                if (from.kind == ChildKind::Data || from.kind == ChildKind::DataGroup)
                    && before.kind == ChildKind::Property
                {
                    return Err(invalid("ink action child move violates schema order"));
                }
                plan_move(source, from.span, before.span, true, order)?
            },
            EditOp::MoveGroupData(from, before) => {
                let from = locate_group_data(profile, *from)?;
                let before = locate_group_data(profile, *before)?;
                if from.action_span != before.action_span || from.group_span != before.group_span {
                    return Err(invalid("ink action grouped-data move requires one group"));
                }
                plan_move(source, from.span, before.span, true, order)?
            },
            EditOp::Clear(selector) => {
                let action = locate_action(profile, *selector)?;
                let Some((start, end)) = inner_range(source, action_span(action))? else {
                    continue;
                };
                make_plan(source, start, end, order, PlanKind::Remove)?
            },
            EditOp::RemoveAction(selector) => {
                let action = locate_action(profile, *selector)?;
                let span = action_span(action);
                make_plan(source, span.start(), span.end(), order, PlanKind::Remove)?
            },
            EditOp::RemoveGroup(index) => {
                let group = locate_group(profile, *index)?;
                let span = group.source;
                make_plan(source, span.start(), span.end(), order, PlanKind::Remove)?
            },
            EditOp::RemoveProperty(selector) => {
                let property = locate_property(profile, *selector)?;
                let span = property.source;
                make_plan(source, span.start(), span.end(), order, PlanKind::Remove)?
            },
            EditOp::RemoveData(selector) => {
                let data = locate_data(profile, *selector)?;
                let span = data.source;
                make_plan(source, span.start(), span.end(), order, PlanKind::Remove)?
            },
            EditOp::RemoveDataGroup(selector) => {
                let ChildSelector::DataGroup { action, index } = *selector else {
                    return Err(invalid("ink action selector is not a data group"));
                };
                let action = locate_action(profile, action)?;
                let group = action
                    .children
                    .iter()
                    .filter_map(|child| match child {
                        ActionChild::DataGroup(group) => Some(group),
                        _ => None,
                    })
                    .nth(index)
                    .ok_or_else(|| invalid("ink action data-group selector is out of range"))?;
                let span = group.source;
                make_plan(source, span.start(), span.end(), order, PlanKind::Remove)?
            },
        };
        plans.push(plan);
    }
    if !root_insertions.is_empty() {
        root_insertions.sort_by_key(|(order, _)| *order);
        let first_order = root_insertions[0].0;
        let mut children = Vec::new();
        children
            .try_reserve(root_insertions.len())
            .map_err(|_| invalid("ink action edit allocation failed"))?;
        for (_, child) in root_insertions {
            children.push(child);
        }
        plans.push(plan_root_insert(
            source,
            root_range,
            root_insert_at,
            InsertKind::Many(children),
            first_order,
        )?);
    }
    coalesce_insertions(&mut plans)?;
    coalesce_expansions(source, &mut plans)?;
    plans.sort_by(|left, right| {
        left.start
            .cmp(&right.start)
            .then(left.end.cmp(&right.end))
            .then(left.order.cmp(&right.order))
    });
    mark_nested_plans(&mut plans)?;
    validate_plan_ranges(&plans)?;
    Ok(plans)
}

fn action_insertion(profile: &Profile, source: &[u8], parent: ActionParent) -> Result<usize> {
    match parent {
        ActionParent::Root => root_close_start(source),
        ActionParent::Group(index) => {
            let group = locate_group(profile, index)?;
            inner_range(source, group.source)?
                .map(|(_, close)| close)
                .ok_or_else(|| invalid("ink action group is self-closing"))
        },
    }
}

fn root_tag(source: &[u8]) -> Result<TagRange> {
    let mut reader = Reader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action root offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action root offset exceeds usize"))?;
        match event {
            Event::Start(_) => {
                return Ok(TagRange {
                    start,
                    end,
                    self_closing: false,
                });
            },
            Event::Empty(_) => {
                return Ok(TagRange {
                    start,
                    end,
                    self_closing: true,
                });
            },
            Event::Eof => return Err(invalid("ink action root is absent")),
            _ => {},
        }
    }
}

fn plan_root_insert<'a>(
    source: &[u8],
    tag: TagRange,
    close_start: Option<usize>,
    child: InsertKind<'a>,
    order: usize,
) -> Result<Plan<'a>> {
    if tag.self_closing {
        let name = element_name(source, tag)?;
        make_plan(
            source,
            tag.start,
            tag.end,
            order,
            PlanKind::ExpandSelfClosing { name, child },
        )
    } else {
        let at = close_start.ok_or_else(|| invalid("ink action root closing element is absent"))?;
        make_plan(source, at, at, order, PlanKind::Insert(child))
    }
}

fn plan_attribute<'a>(
    source: &[u8],
    span: SourceSpan,
    name: &'static [u8],
    value: &'a str,
    order: usize,
    limits: Limits,
) -> Result<Plan<'a>> {
    let tag = start_tag(source, span)?;
    if let Some(attribute) = find_attr(source, tag, name)? {
        make_plan(
            source,
            attribute.value_start,
            attribute.value_end,
            order,
            PlanKind::Escaped(value),
        )
    } else {
        let insertion = if tag.self_closing {
            tag.end.saturating_sub(2)
        } else {
            tag.end.saturating_sub(1)
        };
        make_plan(
            source,
            insertion,
            insertion,
            order,
            PlanKind::Attribute(name, value),
        )
    }
    .and_then(|plan| {
        if plan.inserted_len > limits.max_output_bytes {
            Err(limit("ink action output bytes", limits.max_output_bytes))
        } else {
            Ok(plan)
        }
    })
}

fn plan_data_reference<'a>(
    source: &[u8],
    span: SourceSpan,
    value: Option<&'a str>,
    order: usize,
    _limits: Limits,
) -> Result<Option<Plan<'a>>> {
    let tag = start_tag(source, span)?;
    if let Some(attribute) = find_attr(source, tag, b"ref")? {
        match value {
            Some(value) => make_plan(
                source,
                attribute.value_start,
                attribute.value_end,
                order,
                PlanKind::Escaped(value),
            )
            .map(Some),
            None => make_plan(
                source,
                attribute.start,
                attribute.end,
                order,
                PlanKind::Remove,
            )
            .map(Some),
        }
    } else {
        let Some(value) = value else {
            return Ok(None);
        };
        let insertion = if tag.self_closing {
            tag.end.saturating_sub(2)
        } else {
            tag.end.saturating_sub(1)
        };
        make_plan(
            source,
            insertion,
            insertion,
            order,
            PlanKind::Attribute(b"ref", value),
        )
        .map(Some)
    }
}

fn insert_kind_rank(kind: &InsertKind<'_>) -> u8 {
    match kind {
        InsertKind::Property(_, _) => 0,
        InsertKind::Data(_, _) | InsertKind::DataGroup(_, _) => 1,
        InsertKind::Action(_, _) | InsertKind::Group(_, _) | InsertKind::Many(_) => 2,
    }
}

fn collect_insert_items<'a>(
    kind: InsertKind<'a>,
    order: usize,
    items: &mut Vec<(u8, usize, InsertKind<'a>)>,
) -> Result<()> {
    match kind {
        InsertKind::Many(children) => {
            for child in children {
                collect_insert_items(child, order, items)?;
            }
        },
        kind => {
            items
                .try_reserve(1)
                .map_err(|_| invalid("ink action edit allocation failed"))?;
            let rank = insert_kind_rank(&kind);
            items.push((rank, order, kind));
        },
    }
    Ok(())
}

fn combined_insert_kind<'a>(mut items: Vec<(u8, usize, InsertKind<'a>)>) -> Result<InsertKind<'a>> {
    items.sort_by(|left, right| left.0.cmp(&right.0).then(left.1.cmp(&right.1)));
    let mut children = Vec::new();
    children
        .try_reserve(items.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for (_, _, child) in items {
        children.push(child);
    }
    if children.len() == 1 {
        Ok(children
            .pop()
            .ok_or_else(|| invalid("ink action insertion is empty"))?)
    } else {
        Ok(InsertKind::Many(children))
    }
}

fn take_insert_kind<'a>(kind: &mut PlanKind<'a>) -> Option<InsertKind<'a>> {
    let previous = std::mem::replace(kind, PlanKind::Remove);
    match previous {
        PlanKind::Insert(kind) => Some(kind),
        other => {
            *kind = other;
            None
        },
    }
}

fn take_expand_kind<'a>(kind: &mut PlanKind<'a>) -> Option<(Box<[u8]>, InsertKind<'a>)> {
    let previous = std::mem::replace(kind, PlanKind::Remove);
    match previous {
        PlanKind::ExpandSelfClosing { name, child } => Some((name, child)),
        other => {
            *kind = other;
            None
        },
    }
}

/// Merge insertions at one source offset so schema-ordered children do not
/// depend on the order in which the caller queued them.
fn coalesce_insertions(plans: &mut [Plan<'_>]) -> Result<()> {
    let mut group_by_range: HashMap<(usize, usize), usize> = HashMap::new();
    group_by_range
        .try_reserve(plans.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    let mut groups: Vec<(usize, Vec<(u8, usize, InsertKind<'_>)>)> = Vec::new();
    groups
        .try_reserve(plans.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for (index, plan) in plans.iter_mut().enumerate() {
        if plan.suppressed {
            continue;
        }
        let (start, end, order) = (plan.start, plan.end, plan.order);
        let Some(kind) = take_insert_kind(&mut plan.kind) else {
            continue;
        };
        let group_index = if let Some(&group_index) = group_by_range.get(&(start, end)) {
            group_index
        } else {
            let group_index = groups.len();
            group_by_range.insert((start, end), group_index);
            groups.push((index, Vec::new()));
            group_index
        };
        let (_, items) = groups
            .get_mut(group_index)
            .ok_or_else(|| invalid("ink action insertion group is invalid"))?;
        collect_insert_items(kind, order, items)?;
        if groups[group_index].0 != index {
            plan.suppressed = true;
        }
    }
    for (leader, items) in groups {
        let kind = combined_insert_kind(items)?;
        plans[leader].inserted_len = insert_len(&kind)?;
        plans[leader].kind = PlanKind::Insert(kind);
    }
    Ok(())
}

/// Merge child additions that target the same self-closing action.  This is
/// needed for callers that queue both a property and data child before the
/// final source is materialized.
fn coalesce_expansions(source: &[u8], plans: &mut [Plan<'_>]) -> Result<()> {
    let mut group_by_key: HashMap<(usize, usize, Box<[u8]>), usize> = HashMap::new();
    group_by_key
        .try_reserve(plans.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    let mut groups: Vec<(usize, Box<[u8]>, Vec<(u8, usize, InsertKind<'_>)>)> = Vec::new();
    groups
        .try_reserve(plans.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for (index, plan) in plans.iter_mut().enumerate() {
        if plan.suppressed {
            continue;
        }
        let (start, end, order) = (plan.start, plan.end, plan.order);
        let Some((name, kind)) = take_expand_kind(&mut plan.kind) else {
            continue;
        };
        let key = (start, end, name.clone());
        let group_index = if let Some(&group_index) = group_by_key.get(&key) {
            group_index
        } else {
            let group_index = groups.len();
            group_by_key.insert(key, group_index);
            groups.push((index, name, Vec::new()));
            group_index
        };
        let (_, _, items) = groups
            .get_mut(group_index)
            .ok_or_else(|| invalid("ink action expansion group is invalid"))?;
        collect_insert_items(kind, order, items)?;
        if groups[group_index].0 != index {
            plan.suppressed = true;
        }
    }
    for (leader, name, items) in groups {
        let child = combined_insert_kind(items)?;
        let kind = PlanKind::ExpandSelfClosing { name, child };
        plans[leader].inserted_len =
            plan_len(&kind, source, plans[leader].start, plans[leader].end)?;
        plans[leader].kind = kind;
    }
    Ok(())
}

fn plan_child_insert<'a>(
    source: &[u8],
    action: &Action,
    kind: ChildKind,
    child: InsertKind<'a>,
    order: usize,
) -> Result<Plan<'a>> {
    let (at, self_closing_name) = insertion_for_action(source, action, kind)?;
    if let Some(name) = self_closing_name {
        let span = action_span(action);
        make_plan(
            source,
            span.start(),
            span.end(),
            order,
            PlanKind::ExpandSelfClosing { name, child },
        )
    } else {
        make_plan(source, at, at, order, PlanKind::Insert(child))
    }
}

fn plan_move(
    source: &[u8],
    from: SourceSpan,
    before: SourceSpan,
    same_parent: bool,
    order: usize,
) -> Result<Plan<'_>> {
    if !same_parent {
        return Err(invalid("ink action move requires one parent sequence"));
    }
    if from == before || from.end() == before.start() {
        return make_plan(source, 0, 0, order, PlanKind::Remove);
    }
    if (from.start() < before.start() && from.end() > before.start())
        || (before.start() < from.start() && before.end() > from.start())
    {
        return Err(invalid("ink action move spans overlap"));
    }
    let (start, end) = if from.start() < before.start() {
        (from.start(), before.end())
    } else {
        (before.start(), from.end())
    };
    make_plan(
        source,
        start,
        end,
        order,
        PlanKind::Move {
            from_start: from.start(),
            from_end: from.end(),
            before_start: before.start(),
            before_end: before.end(),
        },
    )
}

fn mark_nested_plans(plans: &mut [Plan<'_>]) -> Result<()> {
    let mut moves = Vec::new();
    moves
        .try_reserve(plans.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for (index, plan) in plans.iter().enumerate() {
        if matches!(&plan.kind, PlanKind::Move { .. }) {
            moves.push((plan.start, plan.end, index));
        }
    }
    // A source move is represented as one contiguous interval.  Crossing or
    // nested move intervals cannot be applied without retargeting child
    // selectors, so reject them with one linear sweep over source order.
    for pair in moves.windows(2) {
        if pair[1].0 < pair[0].1 {
            return Err(invalid("ink action move operations overlap"));
        }
    }
    let mut move_index = 0usize;
    for (index, plan) in plans.iter_mut().enumerate() {
        while move_index < moves.len() && moves[move_index].1 <= plan.start {
            move_index += 1;
        }
        let Some(&(move_start, move_end, owner)) = moves.get(move_index) else {
            break;
        };
        if index == owner {
            continue;
        }
        let inside = if plan.start == plan.end {
            plan.start > move_start && plan.start < move_end
        } else {
            plan.start >= move_start && plan.end <= move_end
        };
        if inside {
            plan.nested = true;
        }
    }
    mark_suppressed_plans(plans)?;
    Ok(())
}

fn mark_suppressed_plans(plans: &mut [Plan<'_>]) -> Result<()> {
    let mut previous_remove = None;
    let mut removals = Vec::new();
    removals
        .try_reserve(plans.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    let mut min_start_by_end: HashMap<usize, usize> = HashMap::new();
    min_start_by_end
        .try_reserve(plans.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for plan in plans.iter_mut() {
        if !matches!(&plan.kind, PlanKind::Remove) || plan.start == plan.end {
            continue;
        }
        let range = (plan.start, plan.end);
        if previous_remove == Some(range) {
            plan.suppressed = true;
            continue;
        }
        previous_remove = Some(range);
        removals.push(range);
        min_start_by_end
            .entry(plan.end)
            .and_modify(|start| *start = (*start).min(plan.start))
            .or_insert(plan.start);
    }

    // `plans` is source-sorted before this function runs.  A max-end heap
    // turns the old removal-range cross product into a bounded sweep while
    // still handling crossing ranges conservatively.
    let mut active = BinaryHeap::new();
    let mut next_removal = 0usize;
    for plan in plans {
        while next_removal < removals.len() && removals[next_removal].0 <= plan.start {
            let (start, end) = removals[next_removal];
            active.push((end, start));
            next_removal += 1;
        }
        while active.peek().is_some_and(|(end, _)| *end <= plan.start) {
            active.pop();
        }
        let Some(&(max_end, _)) = active.peek() else {
            continue;
        };
        let min_start = min_start_by_end
            .get(&max_end)
            .copied()
            .unwrap_or(plan.start);
        let contained = if plan.start == plan.end {
            min_start < plan.start && max_end > plan.start
        } else {
            max_end > plan.end || (max_end == plan.end && min_start < plan.start)
        };
        if contained {
            plan.suppressed = true;
        }
    }
    Ok(())
}

fn validate_plan_ranges(plans: &[Plan<'_>]) -> Result<()> {
    let mut previous_end = 0usize;
    let mut seen = 0usize;
    for plan in plans {
        if plan.nested || plan.suppressed {
            continue;
        }
        let index = seen;
        seen += 1;
        if index != 0 && plan.start < previous_end {
            return Err(invalid(
                "ink action edit operations overlap source ranges; coalesce or split the edit",
            ));
        }
        if plan.start != plan.end {
            previous_end = plan.end;
        }
    }
    Ok(())
}

fn emit_plans(source: &[u8], plans: &[Plan<'_>], limits: Limits) -> Result<Vec<u8>> {
    let mut output_len = source.len();
    for plan in plans {
        if plan.nested || plan.suppressed {
            continue;
        }
        let inserted_len = match &plan.kind {
            PlanKind::Move { .. } => move_inserted_len(plan, plans)?,
            _ => plan.inserted_len,
        };
        output_len = output_len
            .checked_sub(plan.end - plan.start)
            .and_then(|length| length.checked_add(inserted_len))
            .ok_or_else(|| invalid("ink action output length overflow"))?;
    }
    if output_len > limits.max_output_bytes || output_len > super::MAX_SOURCE_BYTES {
        return Err(limit("ink action output bytes", limits.max_output_bytes));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("ink action output allocation failed"))?;
    let mut cursor = 0usize;
    for plan in plans {
        if plan.nested || plan.suppressed {
            continue;
        }
        output.extend_from_slice(&source[cursor..plan.start]);
        emit_plan(&mut output, source, plan, plans)?;
        cursor = plan.end;
    }
    output.extend_from_slice(&source[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn emit_plan(
    output: &mut Vec<u8>,
    source: &[u8],
    plan: &Plan<'_>,
    all_plans: &[Plan<'_>],
) -> Result<()> {
    match &plan.kind {
        PlanKind::Remove => {},
        PlanKind::Escaped(value) => emit_escaped(output, value),
        PlanKind::Attribute(name, value) => emit_attr(output, name, value),
        PlanKind::Insert(kind) => emit_insert(output, kind),
        PlanKind::ExpandSelfClosing { name, child } => {
            let open_end = plan.end.saturating_sub(2);
            output.extend_from_slice(&source[plan.start..open_end]);
            output.push(b'>');
            emit_insert(output, child);
            output.extend_from_slice(b"</");
            output.extend_from_slice(name);
            output.push(b'>');
        },
        PlanKind::Move {
            from_start,
            from_end,
            before_start,
            before_end,
        } => {
            let nested = nested_for_move(plan, all_plans)?;
            if from_start < before_start {
                emit_segment(output, source, *from_end, *before_start, &nested)?;
                emit_segment(output, source, *from_start, *from_end, &nested)?;
                emit_segment(output, source, *before_start, *before_end, &nested)?;
            } else {
                emit_segment(output, source, *from_start, *from_end, &nested)?;
                emit_segment(output, source, *before_start, *before_end, &nested)?;
                emit_segment(output, source, *before_end, *from_start, &nested)?;
            }
        },
    }
    Ok(())
}

fn nested_for_move<'a>(move_plan: &Plan<'a>, plans: &'a [Plan<'a>]) -> Result<Vec<&'a Plan<'a>>> {
    // Plans are source-sorted before emission.  Restrict the scan to starts
    // that can fall inside this move; scanning every plan for every move made
    // a batch of moves quadratic in the number of plans.
    let lower = plan_start_lower_bound(plans, move_plan.start);
    let upper = plan_start_upper_bound(plans, move_plan.end);
    let mut nested = Vec::new();
    nested
        .try_reserve(upper.saturating_sub(lower))
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for plan in &plans[lower..upper] {
        let inside = if plan.start == plan.end {
            plan.start > move_plan.start && plan.start < move_plan.end
        } else {
            plan.start >= move_plan.start && plan.end <= move_plan.end
        };
        if plan.nested && !plan.suppressed && inside {
            nested.push(plan);
        }
    }
    Ok(nested)
}

fn plan_start_lower_bound(plans: &[Plan<'_>], start: usize) -> usize {
    let mut low = 0;
    let mut high = plans.len();
    while low < high {
        let middle = low + (high - low) / 2;
        if plans[middle].start < start {
            low = middle + 1;
        } else {
            high = middle;
        }
    }
    low
}

fn plan_start_upper_bound(plans: &[Plan<'_>], start: usize) -> usize {
    let mut low = 0;
    let mut high = plans.len();
    while low < high {
        let middle = low + (high - low) / 2;
        if plans[middle].start <= start {
            low = middle + 1;
        } else {
            high = middle;
        }
    }
    low
}

fn segment_len(start: usize, end: usize, nested: &[&Plan<'_>]) -> Result<usize> {
    let mut length = end
        .checked_sub(start)
        .ok_or_else(|| invalid("ink action source segment is invalid"))?;
    for plan in nested {
        if plan.start >= start && plan.end <= end {
            let inserted_len = match &plan.kind {
                PlanKind::Move { .. } => {
                    return Err(invalid("nested ink action moves are not supported"));
                },
                _ => plan.inserted_len,
            };
            length = length
                .checked_sub(plan.end - plan.start)
                .and_then(|value| value.checked_add(inserted_len))
                .ok_or_else(|| invalid("ink action segment length overflow"))?;
        }
    }
    Ok(length)
}

fn move_inserted_len(move_plan: &Plan<'_>, plans: &[Plan<'_>]) -> Result<usize> {
    let nested = nested_for_move(move_plan, plans)?;
    let PlanKind::Move {
        from_start,
        from_end,
        before_start,
        before_end,
    } = &move_plan.kind
    else {
        return Err(invalid("ink action move plan is invalid"));
    };
    let from_len = segment_len(*from_start, *from_end, &nested)?;
    let (middle_start, middle_end) = if from_start < before_start {
        (*from_end, *before_start)
    } else {
        (*before_end, *from_start)
    };
    let between_len = segment_len(middle_start, middle_end, &nested)?;
    let before_len = segment_len(*before_start, *before_end, &nested)?;
    from_len
        .checked_add(between_len)
        .and_then(|length| length.checked_add(before_len))
        .ok_or_else(|| invalid("ink action move length overflow"))
}

fn emit_segment(
    output: &mut Vec<u8>,
    source: &[u8],
    start: usize,
    end: usize,
    nested: &[&Plan<'_>],
) -> Result<()> {
    let mut plans: Vec<&Plan<'_>> = Vec::new();
    plans
        .try_reserve(nested.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for plan in nested.iter().copied() {
        if plan.start >= start && plan.end <= end {
            plans.push(plan);
        }
    }
    plans.sort_by(|left, right| {
        left.start
            .cmp(&right.start)
            .then(left.end.cmp(&right.end))
            .then(left.order.cmp(&right.order))
    });
    let mut cursor = start;
    for plan in plans {
        if plan.start < cursor {
            return Err(invalid("nested ink action edit ranges overlap"));
        }
        output.extend_from_slice(&source[cursor..plan.start]);
        emit_plan(output, source, plan, &[])?;
        cursor = plan.end;
    }
    output.extend_from_slice(&source[cursor..end]);
    Ok(())
}

fn emit_insert(output: &mut Vec<u8>, kind: &InsertKind<'_>) {
    match kind {
        InsertKind::Action(action, prefix) => {
            emit_action_with_prefix(output, action, prefix.as_ref())
        },
        InsertKind::Group(group, prefix) => emit_group_with_prefix(output, group, prefix.as_ref()),
        InsertKind::Property(property, prefix) => {
            emit_qname(output, prefix.as_ref(), b"property", false);
            emit_attr(output, b"name", &property.name);
            emit_attr(output, b"value", &property.value);
            output.extend_from_slice(b"/>");
        },
        InsertKind::Data(data, prefix) => emit_data_with_prefix(output, data, prefix.as_ref()),
        InsertKind::DataGroup(group, prefix) => {
            emit_data_group_with_prefix(output, group, prefix.as_ref())
        },
        InsertKind::Many(children) => {
            for child in children {
                emit_insert(output, child);
            }
        },
    }
}

fn validate_profile_limits(profile: &Profile, limits: Limits) -> Result<()> {
    if profile.source.len() > limits.max_output_bytes {
        return Err(limit("ink action source bytes", limits.max_output_bytes));
    }
    let mut actions = 0usize;
    let mut groups = 0usize;
    for child in &profile.children {
        match child {
            RootChild::Action(action) => {
                actions = actions
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action records", limits.max_actions))?;
                validate_profile_action(action, profile, limits)?;
            },
            RootChild::ActionGroup(group) => {
                groups = groups
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action groups", limits.max_action_groups))?;
                validate_profile_group(group, profile, limits, &mut actions)?;
            },
        }
    }
    if actions > limits.max_actions {
        return Err(limit("ink action records", limits.max_actions));
    }
    if groups > limits.max_action_groups {
        return Err(limit("ink action groups", limits.max_action_groups));
    }
    if let Some(span) = profile.definitions {
        check_span_payload(profile, span, limits)?;
    }

    let mut reader = Reader::from_reader(profile.source.as_ref());
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                check_profile_attributes(&element, limits, &reader)?;
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                check_profile_attributes(&element, limits, &reader)?;
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => break,
            _ => {},
        }
        if nodes > limits.max_nodes {
            return Err(limit("ink action XML nodes", limits.max_nodes));
        }
        if depth > limits.max_depth {
            return Err(limit("ink action XML depth", limits.max_depth));
        }
    }
    Ok(())
}

fn check_profile_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    limits: Limits,
    reader: &Reader<&[u8]>,
) -> Result<()> {
    if element.name().as_ref().len() > limits.max_scalar_bytes {
        return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
    }
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        if attribute.key.as_ref().len() > limits.max_scalar_bytes
            || attribute.value.as_ref().len() > limits.max_scalar_bytes
        {
            return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| xml_error(error.to_string()))?;
        if value.len() > limits.max_scalar_bytes {
            return Err(limit("ink action scalar bytes", limits.max_scalar_bytes));
        }
        super::validate_xml_characters(&value, "ink action attribute")?;
    }
    Ok(())
}

fn validate_profile_group(
    group: &ActionGroup,
    profile: &Profile,
    limits: Limits,
    actions: &mut usize,
) -> Result<()> {
    check_profile_scalar(group.xml_id.as_deref(), "xml:id", limits)?;
    check_profile_scalar(Some(&group.start_time), "ink action startTime", limits)?;
    for action in &group.actions {
        *actions = actions
            .checked_add(1)
            .ok_or_else(|| limit("ink action records", limits.max_actions))?;
        validate_profile_action(action, profile, limits)?;
    }
    Ok(())
}

fn validate_profile_action(action: &Action, profile: &Profile, limits: Limits) -> Result<()> {
    check_profile_scalar(action.xml_id.as_deref(), "xml:id", limits)?;
    check_profile_scalar(Some(&action.start_time), "ink action startTime", limits)?;
    check_profile_scalar(Some(action.action_type.as_str()), "action type", limits)?;
    for child in &action.children {
        match child {
            ActionChild::Property(property) => {
                check_profile_scalar(Some(&property.name), "property name", limits)?;
                check_profile_scalar(Some(&property.value), "property value", limits)?;
            },
            ActionChild::Data(data) => validate_profile_data(data, profile, limits)?,
            ActionChild::DataGroup(group) => {
                check_profile_scalar(group.xml_id.as_deref(), "xml:id", limits)?;
                check_profile_scalar(Some(&group.name), "action data-group name", limits)?;
                for data in &group.data {
                    validate_profile_data(data, profile, limits)?;
                }
            },
        }
    }
    Ok(())
}

fn validate_profile_data(data: &ActionData, profile: &Profile, limits: Limits) -> Result<()> {
    check_profile_scalar(data.xml_id.as_deref(), "xml:id", limits)?;
    check_profile_scalar(Some(&data.name), "action data name", limits)?;
    check_profile_scalar(data.reference.as_deref(), "action data reference", limits)?;
    for child in &data.children {
        check_span_payload(profile, child.source_span(), limits)?;
    }
    Ok(())
}

fn check_profile_scalar(value: Option<&str>, field: &'static str, limits: Limits) -> Result<()> {
    if value.is_some_and(|value| value.len() > limits.max_scalar_bytes) {
        return Err(limit(field, limits.max_scalar_bytes));
    }
    Ok(())
}

fn check_span_payload(profile: &Profile, span: SourceSpan, limits: Limits) -> Result<()> {
    if span.end().saturating_sub(span.start()) > limits.max_payload_bytes {
        return Err(limit("ink action payload bytes", limits.max_payload_bytes));
    }
    if span.end() > profile.source.len() {
        return Err(invalid("ink action payload source range is invalid"));
    }
    Ok(())
}

fn validate_batch_operations(
    profile: &Profile,
    source: &[u8],
    operations: &[EditOp],
    limits: Limits,
) -> Result<()> {
    let id_index = build_id_index(profile)?;
    let mut removed_spans = Vec::new();
    removed_spans
        .try_reserve(operations.len())
        .map_err(|_| invalid("ink action removal allocation failed"))?;
    for operation in operations {
        match operation {
            EditOp::RemoveAction(selector) => {
                removed_spans.push(action_span(locate_action(profile, *selector)?));
            },
            EditOp::RemoveGroup(index) => removed_spans.push(locate_group(profile, *index)?.source),
            EditOp::RemoveData(selector) => {
                removed_spans.push(locate_data(profile, *selector)?.source)
            },
            EditOp::RemoveDataGroup(selector) => {
                let ChildSelector::DataGroup { action, index } = *selector else {
                    return Err(invalid("ink action selector is not a data group"));
                };
                let action = locate_action(profile, action)?;
                let group = action
                    .children
                    .iter()
                    .filter_map(|child| match child {
                        ActionChild::DataGroup(group) => Some(group),
                        _ => None,
                    })
                    .nth(index)
                    .ok_or_else(|| invalid("ink action data-group selector is out of range"))?;
                removed_spans.push(group.source);
            },
            EditOp::Clear(selector) => {
                let action = locate_action(profile, *selector)?;
                if let Some((start, end)) = inner_range(source, action_span(action))? {
                    removed_spans.push(SourceSpan::new(start, end));
                }
            },
            _ => {},
        }
    }
    normalize_spans(&mut removed_spans);
    let mut new_ids = Vec::new();
    let mut new_opaque_ids: Vec<Box<str>> = Vec::new();
    let existing_actions = profile_action_count(profile);
    let existing_groups = profile_group_count(profile);
    let existing_nodes = source_node_count(source)?;
    let mut added_actions = 0usize;
    let mut added_groups = 0usize;
    let mut added_nodes = 0usize;
    let mut added_depth = 0usize;
    for operation in operations {
        match operation {
            EditOp::Add(parent, action) => {
                let mut nodes = 1usize;
                let action_depth = match parent {
                    ActionParent::Root => 2,
                    ActionParent::Group(_) => 3,
                };
                validate_action_draft(
                    action,
                    limits,
                    &mut nodes,
                    action_depth,
                    &mut added_depth,
                    false,
                )?;
                added_actions = added_actions
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action records", limits.max_actions))?;
                added_nodes = added_nodes
                    .checked_add(nodes)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                collect_draft_ids(action, &mut new_ids)?;
                collect_draft_opaque_ids(action, &mut new_opaque_ids)?;
            },
            EditOp::AddGroup(group) => {
                let mut group_nodes = 1usize;
                validate_group_draft(group, limits, &mut group_nodes, &mut added_depth, false)?;
                added_groups = added_groups
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action groups", limits.max_action_groups))?;
                added_actions = added_actions
                    .checked_add(group.actions.len())
                    .ok_or_else(|| limit("ink action records", limits.max_actions))?;
                added_nodes = added_nodes
                    .checked_add(group_nodes)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                if let Some(id) = &group.xml_id {
                    push_id(&mut new_ids, id)?;
                }
                for action in &group.actions {
                    collect_draft_ids(action, &mut new_ids)?;
                    collect_draft_opaque_ids(action, &mut new_opaque_ids)?;
                }
            },
            EditOp::AddProperty(ChildParent::Action(selector), property) => {
                bounded_scalar(&property.name, "property name", limits)?;
                bounded_scalar(&property.value, "property value", limits)?;
                let action_depth = action_depth_for_selector(profile, *selector)?;
                added_depth = added_depth.max(
                    action_depth
                        .checked_add(1)
                        .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?,
                );
                added_nodes = added_nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
            },
            EditOp::AddData(selector, data) => {
                let mut nodes = 1usize;
                let data_depth = action_depth_for_selector(profile, *selector)?
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                validate_data_draft(
                    data,
                    limits,
                    &mut nodes,
                    data_depth,
                    &mut added_depth,
                    false,
                )?;
                added_nodes = added_nodes
                    .checked_add(nodes)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                collect_draft_data_ids(data, &mut new_ids)?;
                collect_draft_data_opaque_ids(data, &mut new_opaque_ids)?;
            },
            EditOp::AddDataGroup(selector, group) => {
                if group.data.is_empty() {
                    return Err(invalid("ink action data group requires actionData"));
                }
                bounded_scalar(&group.name, "action data-group name", limits)?;
                let group_depth = action_depth_for_selector(profile, *selector)?
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                added_depth = added_depth.max(group_depth);
                if let Some(id) = &group.xml_id {
                    bounded_scalar(id, "xml:id", limits)?;
                    push_id(&mut new_ids, id)?;
                }
                for data in &group.data {
                    let mut nodes = 1usize;
                    let data_depth = group_depth
                        .checked_add(1)
                        .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
                    validate_data_draft(
                        data,
                        limits,
                        &mut nodes,
                        data_depth,
                        &mut added_depth,
                        false,
                    )?;
                    added_nodes = added_nodes
                        .checked_add(nodes)
                        .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
                    collect_draft_data_ids(data, &mut new_ids)?;
                    collect_draft_data_opaque_ids(data, &mut new_opaque_ids)?;
                }
                added_nodes = added_nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
            },
            _ => {},
        }
    }
    let (removed_actions, removed_groups) = removed_semantic_counts(profile, &removed_spans);
    let removed_nodes = count_nodes_in_spans(source, &removed_spans)?;
    let remaining_actions = existing_actions.saturating_sub(removed_actions);
    let remaining_groups = existing_groups.saturating_sub(removed_groups);
    let remaining_nodes = existing_nodes.saturating_sub(removed_nodes);
    if remaining_actions.saturating_add(added_actions) > limits.max_actions {
        return Err(limit("ink action records", limits.max_actions));
    }
    if remaining_groups.saturating_add(added_groups) > limits.max_action_groups {
        return Err(limit("ink action groups", limits.max_action_groups));
    }
    if remaining_nodes.saturating_add(added_nodes) > limits.max_nodes {
        return Err(limit("ink action XML nodes", limits.max_nodes));
    }
    if added_depth > limits.max_depth {
        return Err(limit("ink action XML depth", limits.max_depth));
    }
    let mut new_id_set = HashSet::new();
    new_id_set
        .try_reserve(new_ids.len())
        .map_err(|_| invalid("ink action edit allocation failed"))?;
    for id in &new_ids {
        if id_index.declaration_spans.contains_key(&**id) || !new_id_set.insert(&**id) {
            return Err(invalid(
                "ink action edit would create duplicate xml:id values",
            ));
        }
    }
    let mut new_opaque_id_set: HashSet<Box<str>> = HashSet::new();
    new_opaque_id_set
        .try_reserve(new_opaque_ids.len())
        .map_err(|_| invalid("ink action identity allocation failed"))?;
    for id in new_opaque_ids {
        if id_index.declaration_spans.contains_key(id.as_ref())
            || new_id_set.contains(id.as_ref())
            || !new_opaque_id_set.insert(id)
        {
            return Err(invalid(
                "ink action edit would create duplicate xml:id values",
            ));
        }
    }
    let references = effective_reference_index(&id_index, profile, operations)?;
    check_removed_declarations(&id_index, &references, &removed_spans)?;
    Ok(())
}

/// Sort and collapse nested/crossing removal ranges before they are used for
/// both quota accounting and reference closure.  Adjacent ranges remain
/// separate so a declaration in the untouched gap cannot be mistaken for a
/// removed descendant.
fn normalize_spans(spans: &mut Vec<SourceSpan>) {
    spans.sort_by(|left, right| {
        left.start()
            .cmp(&right.start())
            .then_with(|| right.end().cmp(&left.end()))
    });
    let mut write = 0usize;
    for read in 0..spans.len() {
        let span = spans[read];
        if write == 0 {
            spans[write] = span;
            write += 1;
            continue;
        }
        let previous = spans[write - 1];
        if span.start() < previous.end() {
            if span.end() > previous.end() {
                spans[write - 1] = SourceSpan::new(previous.start(), span.end());
            }
        } else {
            spans[write] = span;
            write += 1;
        }
    }
    spans.truncate(write);
}

fn removed_semantic_counts(profile: &Profile, removed_spans: &[SourceSpan]) -> (usize, usize) {
    let mut actions = 0usize;
    let mut groups = 0usize;
    for child in &profile.children {
        match child {
            RootChild::Action(action) => {
                if spans_contain(action_span(action), removed_spans) {
                    actions = actions.saturating_add(1);
                }
            },
            RootChild::ActionGroup(group) => {
                if spans_contain(group.source, removed_spans) {
                    groups = groups.saturating_add(1);
                }
                for action in &group.actions {
                    if spans_contain(action_span(action), removed_spans) {
                        actions = actions.saturating_add(1);
                    }
                }
            },
        }
    }
    (actions, groups)
}

fn count_nodes_in_spans(source: &[u8], spans: &[SourceSpan]) -> Result<usize> {
    if spans.is_empty() {
        return Ok(0);
    }
    let mut reader = Reader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    let mut nodes = 0usize;
    let mut span_index = 0usize;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action XML node offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action XML node offset exceeds usize"))?;
        match event {
            Event::Start(_) | Event::Empty(_) => {
                while span_index < spans.len() && start >= spans[span_index].end() {
                    span_index += 1;
                }
                if spans
                    .get(span_index)
                    .is_some_and(|span| span.start() <= start && end <= span.end())
                {
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("ink action XML nodes", super::MAX_NODES))?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(nodes)
}

fn profile_action_count(profile: &Profile) -> usize {
    profile
        .children
        .iter()
        .map(|child| match child {
            RootChild::Action(_) => 1,
            RootChild::ActionGroup(group) => group.actions.len(),
        })
        .sum()
}

fn profile_group_count(profile: &Profile) -> usize {
    profile
        .children
        .iter()
        .filter(|child| matches!(child, RootChild::ActionGroup(_)))
        .count()
}

fn action_depth_for_selector(profile: &Profile, selector: ActionSelector) -> Result<usize> {
    let location = locate_action_with_parent(profile, selector)?;
    Ok(match location.parent {
        ParentLocation::Root => 2,
        ParentLocation::Group(_) => 3,
    })
}

fn source_node_count(source: &[u8]) -> Result<usize> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut nodes = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
        {
            Event::Start(_) | Event::Empty(_) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("ink action XML nodes", super::MAX_NODES))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(nodes)
}

fn validate_data_group_draft(
    group: &DataGroupDraft,
    limits: Limits,
    group_depth: usize,
    max_depth: &mut usize,
    allow_inherited_payload_namespace: bool,
) -> Result<()> {
    if group.data.is_empty() {
        return Err(invalid("ink action data group requires actionData"));
    }
    bounded_scalar(&group.name, "action data-group name", limits)?;
    if let Some(id) = &group.xml_id {
        bounded_scalar(id, "xml:id", limits)?;
        if !is_ncname(id) {
            return Err(invalid("ink action xml:id is not an NCName"));
        }
    }
    let data_depth = group_depth
        .checked_add(1)
        .ok_or_else(|| limit("ink action XML depth", limits.max_depth))?;
    for data in &group.data {
        let mut nodes = 1usize;
        validate_data_draft(
            data,
            limits,
            &mut nodes,
            data_depth,
            max_depth,
            allow_inherited_payload_namespace,
        )?;
    }
    Ok(())
}

fn validate_group_draft(
    group: &ActionGroupDraft,
    limits: Limits,
    nodes: &mut usize,
    max_depth: &mut usize,
    allow_inherited_payload_namespace: bool,
) -> Result<()> {
    if group.actions.is_empty() {
        return Err(invalid("ink action group requires an action child"));
    }
    *max_depth = (*max_depth).max(2);
    bounded_scalar(group.action_type.as_str(), "action group type", limits)?;
    bounded_scalar(&group.start_time, "action group startTime", limits)?;
    super::validate_decimal(&group.start_time)?;
    if let Some(id) = &group.xml_id {
        bounded_scalar(id, "xml:id", limits)?;
        if !is_ncname(id) {
            return Err(invalid("ink action xml:id is not an NCName"));
        }
    }
    for action in &group.actions {
        *nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit("ink action XML nodes", limits.max_nodes))?;
        validate_action_draft(
            action,
            limits,
            nodes,
            3,
            max_depth,
            allow_inherited_payload_namespace,
        )?;
    }
    Ok(())
}

#[derive(Debug)]
struct IdIndex {
    declaration_spans: HashMap<Box<str>, Vec<SourceSpan>>,
    references: Vec<IndexedReference>,
}

#[derive(Debug)]
struct IndexedReference {
    target: Box<str>,
    span: Option<SourceSpan>,
    /// True for values found in a payload whose schema is deliberately opaque.
    /// Such values are matched conservatively after XML decoding, never by raw
    /// source substring search.
    opaque: bool,
}

#[derive(Debug)]
struct EffectiveReference {
    span: Option<SourceSpan>,
}

fn build_id_index(profile: &Profile) -> Result<IdIndex> {
    let mut index = IdIndex {
        declaration_spans: HashMap::new(),
        references: Vec::new(),
    };
    index
        .declaration_spans
        .try_reserve(profile.children.len().saturating_add(1))
        .map_err(|_| invalid("ink action identifier index allocation failed"))?;
    index
        .references
        .try_reserve(profile.children.len())
        .map_err(|_| invalid("ink action reference index allocation failed"))?;
    if let Some(id) = profile.xml_id.as_deref() {
        push_declaration(&mut index, id, SourceSpan::new(0, profile.source().len()))?;
    }
    if let Some(span) = profile.definitions {
        index_opaque_payload(profile, span, &mut index)?;
    }
    for child in &profile.children {
        match child {
            RootChild::Action(action) => index_action(profile, action, &mut index)?,
            RootChild::ActionGroup(group) => {
                if let Some(id) = group.xml_id.as_deref() {
                    push_declaration(&mut index, id, group.source)?;
                }
                for action in &group.actions {
                    index_action(profile, action, &mut index)?;
                }
            },
        }
    }
    Ok(index)
}

fn ensure_unique_profile_ids(profile: &Profile) -> Result<()> {
    let index = build_id_index(profile)?;
    if index
        .declaration_spans
        .values()
        .any(|spans| spans.len() > 1)
    {
        return Err(invalid(
            "ink action edit would create duplicate xml:id values",
        ));
    }
    Ok(())
}

fn index_action(profile: &Profile, action: &Action, index: &mut IdIndex) -> Result<()> {
    if let Some(id) = action.xml_id.as_deref() {
        push_declaration(index, id, action_span(action))?;
    }
    for child in &action.children {
        match child {
            ActionChild::Property(_) => {},
            ActionChild::Data(data) => index_data(profile, data, index)?,
            ActionChild::DataGroup(group) => {
                if let Some(id) = group.xml_id.as_deref() {
                    push_declaration(index, id, group.source)?;
                }
                for data in &group.data {
                    index_data(profile, data, index)?;
                }
            },
        }
    }
    Ok(())
}

fn index_data(profile: &Profile, data: &ActionData, index: &mut IdIndex) -> Result<()> {
    if let Some(id) = data.xml_id.as_deref() {
        push_declaration(index, id, data.source)?;
    }
    if let Some(reference) = data.reference.as_deref() {
        push_reference(index, reference, Some(data.source), false)?;
    }
    for child in &data.children {
        index_opaque_payload(profile, child.source_span(), index)?;
    }
    Ok(())
}

fn push_declaration(index: &mut IdIndex, value: &str, span: SourceSpan) -> Result<()> {
    push_declaration_span(index, value, span)
}

fn push_opaque_declaration(index: &mut IdIndex, value: &str, span: SourceSpan) -> Result<()> {
    push_declaration_span(index, value, span)
}

fn push_declaration_span(index: &mut IdIndex, value: &str, span: SourceSpan) -> Result<()> {
    if let Some(spans) = index.declaration_spans.get_mut(value) {
        spans
            .try_reserve(1)
            .map_err(|_| invalid("ink action identifier index allocation failed"))?;
        spans.push(span);
        return Ok(());
    }
    let key = copy_index_string(value)?;
    let mut spans = Vec::new();
    spans
        .try_reserve(1)
        .map_err(|_| invalid("ink action identifier index allocation failed"))?;
    spans.push(span);
    index
        .declaration_spans
        .try_reserve(1)
        .map_err(|_| invalid("ink action identifier index allocation failed"))?;
    index.declaration_spans.insert(key, spans);
    Ok(())
}

fn push_reference(
    index: &mut IdIndex,
    value: &str,
    span: Option<SourceSpan>,
    opaque: bool,
) -> Result<()> {
    let Some(target) = reference_target(value)? else {
        return Ok(());
    };
    index
        .references
        .try_reserve(1)
        .map_err(|_| invalid("ink action reference index allocation failed"))?;
    index.references.push(IndexedReference {
        target,
        span,
        opaque,
    });
    Ok(())
}

fn copy_index_string(value: &str) -> Result<Box<str>> {
    if value.len() > super::MAX_ATTRIBUTE_VALUE_BYTES {
        return Err(limit(
            "ink action reference value bytes",
            super::MAX_ATTRIBUTE_VALUE_BYTES,
        ));
    }
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_| invalid("ink action reference index allocation failed"))?;
    copy.push_str(value);
    Ok(copy.into_boxed_str())
}

fn reference_target(value: &str) -> Result<Option<Box<str>>> {
    let value = super::collapse_xml_whitespace(value)?;
    let value = value.strip_prefix('#').unwrap_or_else(|| {
        value
            .rsplit_once('#')
            .map_or(value.as_str(), |(_, fragment)| fragment)
    });
    if value.is_empty() {
        return Ok(None);
    }
    Ok(Some(normalize_unreserved_percent(value)?))
}

fn data_reference_semantically_equal(left: Option<&str>, right: Option<&str>) -> Result<bool> {
    match (left, right) {
        (None, None) => Ok(true),
        (Some(left), Some(right)) => {
            Ok(super::collapse_xml_whitespace(left)? == super::collapse_xml_whitespace(right)?)
        },
        _ => Ok(false),
    }
}

fn normalize_unreserved_percent(value: &str) -> Result<Box<str>> {
    let mut normalized = String::new();
    normalized
        .try_reserve_exact(value.len())
        .map_err(|_| invalid("ink action reference index allocation failed"))?;
    let bytes = value.as_bytes();
    let mut cursor = 0;
    while cursor < bytes.len() {
        if bytes[cursor] == b'%'
            && cursor + 2 < bytes.len()
            && let (Some(high), Some(low)) =
                (hex_digit(bytes[cursor + 1]), hex_digit(bytes[cursor + 2]))
        {
            let decoded = (high << 4) | low;
            if is_uri_unreserved(decoded) {
                normalized.push(decoded as char);
                cursor += 3;
                continue;
            }
        }
        let character = value[cursor..]
            .chars()
            .next()
            .ok_or_else(|| invalid("ink action reference value is not UTF-8"))?;
        normalized.push(character);
        cursor += character.len_utf8();
    }
    Ok(normalized.into_boxed_str())
}

fn hex_digit(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

fn is_uri_unreserved(value: u8) -> bool {
    value.is_ascii_alphanumeric() || matches!(value, b'-' | b'.' | b'_' | b'~')
}

fn index_opaque_payload(profile: &Profile, span: SourceSpan, index: &mut IdIndex) -> Result<()> {
    let payload = profile.xml(span);
    if payload.is_empty() {
        return Err(invalid("ink action opaque payload source range is empty"));
    }
    let mut reader = Reader::from_reader(payload);
    let origin = ReaderOrigin::of(payload);
    reader.config_mut().trim_text(false);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action opaque payload offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action opaque payload offset exceeds usize"))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let element_span = SourceSpan::new(
                    span.start()
                        .checked_add(start)
                        .ok_or_else(|| invalid("ink action opaque payload offset overflow"))?,
                    span.start()
                        .checked_add(end)
                        .ok_or_else(|| invalid("ink action opaque payload offset overflow"))?,
                );
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" || key.strip_prefix(b"xmlns:").is_some() || key == b"xml:id"
                    {
                        if key == b"xml:id" {
                            let value = decode_opaque_attribute(&attribute, &reader)?;
                            push_opaque_declaration(index, &value, element_span)?;
                        }
                        continue;
                    }
                    let value = decode_opaque_attribute(&attribute, &reader)?;
                    // Keep all decoded opaque attribute values in the index.
                    // Unknown payload schemas may designate a custom attribute
                    // as an IDREF; exact XML-decoded matching is conservative
                    // without treating comments, tag names, or substrings as
                    // references. `opaque` records retain that uncertainty.
                    push_reference(index, &value, Some(element_span), true)?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn decode_opaque_attribute(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    reader: &Reader<&[u8]>,
) -> Result<String> {
    if attribute.value.len() > super::MAX_ATTRIBUTE_VALUE_BYTES {
        return Err(limit(
            "ink action opaque attribute value bytes",
            super::MAX_ATTRIBUTE_VALUE_BYTES,
        ));
    }
    let value = attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
        .map_err(|error| xml_error(error.to_string()))?;
    if value.len() > super::MAX_ATTRIBUTE_VALUE_BYTES {
        return Err(limit(
            "ink action opaque attribute value bytes",
            super::MAX_ATTRIBUTE_VALUE_BYTES,
        ));
    }
    super::validate_xml_characters(&value, "ink action opaque attribute")?;
    Ok(value.into_owned())
}

fn effective_reference_index(
    index: &IdIndex,
    profile: &Profile,
    operations: &[EditOp],
) -> Result<HashMap<Box<str>, Vec<EffectiveReference>>> {
    let mut intents: HashMap<SourceSpan, Option<Box<str>>> = HashMap::new();
    intents
        .try_reserve(operations.len())
        .map_err(|_| invalid("ink action reference intent allocation failed"))?;
    for operation in operations {
        if let EditOp::SetDataReference(selector, value) = operation {
            let data = locate_data(profile, *selector)?;
            let target = value
                .as_deref()
                .map(reference_target)
                .transpose()?
                .flatten();
            intents.insert(data.source, target);
        }
    }
    let mut references: HashMap<Box<str>, Vec<EffectiveReference>> = HashMap::new();
    references
        .try_reserve(index.references.len())
        .map_err(|_| invalid("ink action reference index allocation failed"))?;
    for reference in &index.references {
        if !reference.opaque {
            if let Some(span) = reference.span {
                if let Some(intent) = intents.get(&span) {
                    let Some(target) = intent else {
                        continue;
                    };
                    insert_effective_reference(&mut references, target.clone(), reference.span)?;
                    continue;
                }
            }
        }
        insert_effective_reference(&mut references, reference.target.clone(), reference.span)?;
    }
    for operation in operations {
        match operation {
            EditOp::Add(_, action) => {
                index_draft_action_references(action, &mut references)?;
            },
            EditOp::AddGroup(group) => {
                for action in &group.actions {
                    index_draft_action_references(action, &mut references)?;
                }
            },
            EditOp::AddData(_, data) => {
                index_draft_data_reference(data, &mut references)?;
            },
            EditOp::AddDataGroup(_, group) => {
                for data in &group.data {
                    index_draft_data_reference(data, &mut references)?;
                }
            },
            _ => {},
        }
    }
    Ok(references)
}

fn insert_effective_reference(
    references: &mut HashMap<Box<str>, Vec<EffectiveReference>>,
    target: Box<str>,
    span: Option<SourceSpan>,
) -> Result<()> {
    references
        .try_reserve(1)
        .map_err(|_| invalid("ink action reference index allocation failed"))?;
    let values = references.entry(target).or_default();
    values
        .try_reserve(1)
        .map_err(|_| invalid("ink action reference index allocation failed"))?;
    values.push(EffectiveReference { span });
    Ok(())
}

fn index_draft_action_references(
    action: &ActionDraft,
    references: &mut HashMap<Box<str>, Vec<EffectiveReference>>,
) -> Result<()> {
    for child in &action.children {
        match child {
            ActionChildDraft::Data(data) => index_draft_data_reference(data, references)?,
            ActionChildDraft::DataGroup(group) => {
                for data in &group.data {
                    index_draft_data_reference(data, references)?;
                }
            },
            ActionChildDraft::Property(_) => {},
        }
    }
    Ok(())
}

fn index_draft_data_reference(
    data: &ActionDataDraft,
    references: &mut HashMap<Box<str>, Vec<EffectiveReference>>,
) -> Result<()> {
    if let Some(value) = data.reference.as_deref() {
        if let Some(target) = reference_target(value)? {
            insert_effective_reference(references, target, None)?;
        }
    }
    for child in &data.children {
        let payload = match child {
            DataChildDraft::Transform(payload)
            | DataChildDraft::Trace(payload)
            | DataChildDraft::TraceView(payload) => payload,
        };
        index_draft_opaque_references(payload.as_bytes(), references)?;
    }
    Ok(())
}

fn index_draft_opaque_references(
    payload: &[u8],
    references: &mut HashMap<Box<str>, Vec<EffectiveReference>>,
) -> Result<()> {
    let mut reader = Reader::from_reader(payload);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    if key == b"xml:id" || key == b"xmlns" || key.strip_prefix(b"xmlns:").is_some()
                    {
                        continue;
                    }
                    let value = decode_opaque_attribute(&attribute, &reader)?;
                    if let Some(target) = reference_target(&value)? {
                        insert_effective_reference(references, target, None)?;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn check_removed_declarations(
    index: &IdIndex,
    references: &HashMap<Box<str>, Vec<EffectiveReference>>,
    removed_spans: &[SourceSpan],
) -> Result<()> {
    for (id, spans) in &index.declaration_spans {
        if id.is_empty() || !spans.iter().any(|span| spans_contain(*span, removed_spans)) {
            continue;
        }
        if spans
            .iter()
            .any(|span| !spans_contain(*span, removed_spans))
        {
            return Err(invalid(
                "ink action removal would break an xml:id reference (ambiguous or opaque)",
            ));
        }
        if references.get(id).is_some_and(|values| {
            values.iter().any(|reference| {
                reference
                    .span
                    .is_none_or(|span| !spans_contain(span, removed_spans))
            })
        }) {
            return Err(invalid(
                "ink action removal would break an xml:id reference (ambiguous or opaque)",
            ));
        }
    }
    Ok(())
}

fn spans_contain(span: SourceSpan, outers: &[SourceSpan]) -> bool {
    outers
        .iter()
        .copied()
        .any(|outer| span_contains(outer, span))
}

fn span_contains(outer: SourceSpan, inner: SourceSpan) -> bool {
    outer.start() <= inner.start() && inner.end() <= outer.end()
}

fn collect_draft_ids<'a>(action: &'a ActionDraft, ids: &mut Vec<&'a str>) -> Result<()> {
    if let Some(id) = action.xml_id.as_deref() {
        push_id(ids, id)?;
    }
    for child in &action.children {
        match child {
            ActionChildDraft::Property(_) => {},
            ActionChildDraft::Data(data) => collect_draft_data_ids(data, ids)?,
            ActionChildDraft::DataGroup(group) => {
                if let Some(id) = group.xml_id.as_deref() {
                    push_id(ids, id)?;
                }
                for data in &group.data {
                    collect_draft_data_ids(data, ids)?;
                }
            },
        }
    }
    Ok(())
}

fn collect_draft_data_ids<'a>(data: &'a ActionDataDraft, ids: &mut Vec<&'a str>) -> Result<()> {
    if let Some(id) = data.xml_id.as_deref() {
        push_id(ids, id)?;
    }
    Ok(())
}

fn collect_draft_opaque_ids(action: &ActionDraft, ids: &mut Vec<Box<str>>) -> Result<()> {
    for child in &action.children {
        match child {
            ActionChildDraft::Property(_) => {},
            ActionChildDraft::Data(data) => collect_draft_data_opaque_ids(data, ids)?,
            ActionChildDraft::DataGroup(group) => {
                for data in &group.data {
                    collect_draft_data_opaque_ids(data, ids)?;
                }
            },
        }
    }
    Ok(())
}

fn collect_draft_data_opaque_ids(data: &ActionDataDraft, ids: &mut Vec<Box<str>>) -> Result<()> {
    for child in &data.children {
        let payload = match child {
            DataChildDraft::Transform(payload)
            | DataChildDraft::Trace(payload)
            | DataChildDraft::TraceView(payload) => payload,
        };
        collect_opaque_ids(payload.as_bytes(), ids)?;
    }
    Ok(())
}

fn collect_opaque_ids(payload: &[u8], ids: &mut Vec<Box<str>>) -> Result<()> {
    let mut reader = Reader::from_reader(payload);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    loop {
        match reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
                    if attribute.key.as_ref() != b"xml:id" {
                        continue;
                    }
                    let value = decode_opaque_attribute(&attribute, &reader)?;
                    if value.is_empty() {
                        continue;
                    }
                    ids.try_reserve(1)
                        .map_err(|_| invalid("ink action identity allocation failed"))?;
                    ids.push(value.into_boxed_str());
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn push_id<'a>(ids: &mut Vec<&'a str>, id: &'a str) -> Result<()> {
    ids.try_reserve(1)
        .map_err(|_| invalid("ink action identity allocation failed"))?;
    ids.push(id);
    Ok(())
}

fn attr_len_value(name: &[u8], value: &str) -> Result<usize> {
    let mut len = 0;
    attr_len(&mut len, name, value)?;
    Ok(len)
}

fn emit_escaped(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\'' => output.extend_from_slice(b"&apos;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

#[derive(Clone, Copy)]
struct TagRange {
    start: usize,
    end: usize,
    self_closing: bool,
}
struct AttributeRange {
    start: usize,
    value_start: usize,
    value_end: usize,
    end: usize,
}

fn start_tag(source: &[u8], span: SourceSpan) -> Result<TagRange> {
    let start = span.start();
    let end = span.end();
    if start >= end || end > source.len() || source.get(start) != Some(&b'<') {
        return Err(invalid(
            "ink action source span does not begin with an element",
        ));
    }
    let mut quote = None;
    for index in start + 1..end {
        match (quote, source[index]) {
            (Some(mark), byte) if byte == mark => quote = None,
            (None, b'\'' | b'"') => quote = Some(source[index]),
            (None, b'>') => {
                let self_closing = index > start + 1 && source[index - 1] == b'/';
                return Ok(TagRange {
                    start,
                    end: index + 1,
                    self_closing,
                });
            },
            _ => {},
        }
    }
    Err(invalid("ink action start tag is unterminated"))
}

fn find_attr(source: &[u8], tag: TagRange, wanted: &[u8]) -> Result<Option<AttributeRange>> {
    let mut cursor = tag.start + 1;
    while cursor < tag.end
        && !source[cursor].is_ascii_whitespace()
        && source[cursor] != b'>'
        && source[cursor] != b'/'
    {
        cursor += 1;
    }
    loop {
        while cursor < tag.end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.end || source[cursor] == b'>' || source[cursor] == b'/' {
            return Ok(None);
        }
        let key_start = cursor;
        while cursor < tag.end
            && !source[cursor].is_ascii_whitespace()
            && source[cursor] != b'='
            && source[cursor] != b'>'
            && source[cursor] != b'/'
        {
            cursor += 1;
        }
        let key_end = cursor;
        while cursor < tag.end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if source.get(cursor) != Some(&b'=') {
            return Err(invalid("ink action attribute source range is invalid"));
        }
        cursor += 1;
        while cursor < tag.end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *source
            .get(cursor)
            .ok_or_else(|| invalid("ink action attribute source range is invalid"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid("ink action attribute source range is invalid"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.end && source[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.end {
            return Err(invalid("ink action attribute source range is invalid"));
        }
        let value_end = cursor;
        cursor += 1;
        if &source[key_start..key_end] == wanted {
            return Ok(Some(AttributeRange {
                start: key_start.saturating_sub(1),
                value_start,
                value_end,
                end: cursor,
            }));
        }
    }
}

fn inner_range(source: &[u8], span: SourceSpan) -> Result<Option<(usize, usize)>> {
    let tag = start_tag(source, span)?;
    if tag.self_closing {
        return Ok(None);
    }
    let close_start = element_close_start(source, span)?;
    if close_start < tag.end {
        return Err(invalid("ink action source closing range is invalid"));
    }
    Ok(Some((tag.end, close_start)))
}

fn element_close_start(source: &[u8], span: SourceSpan) -> Result<usize> {
    if span.end() > source.len() || span.start() >= span.end() {
        return Err(invalid("ink action source range is invalid"));
    }
    let fragment = &source[span.start()..span.end()];
    let mut reader = Reader::from_reader(fragment);
    let origin = ReaderOrigin::of(fragment);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action closing offset exceeds usize"))?;
        match event {
            Event::Start(_) => depth = depth.saturating_add(1),
            Event::Empty(_) if depth == 0 => {
                return Err(invalid("ink action element is unexpectedly empty"));
            },
            Event::End(_) if depth == 1 => {
                let close_start = fragment[..end]
                    .iter()
                    .rposition(|byte| *byte == b'<')
                    .ok_or_else(|| invalid("ink action closing range is invalid"))?;
                return span
                    .start()
                    .checked_add(close_start)
                    .ok_or_else(|| invalid("ink action closing offset overflows usize"));
            },
            Event::End(_) if depth > 0 => depth -= 1,
            Event::Eof => break,
            _ => {},
        }
    }
    Err(invalid("ink action closing element is absent"))
}

fn root_close_start(source: &[u8]) -> Result<usize> {
    let mut reader = Reader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("ink action root offset exceeds usize"))?;
        match event {
            Event::Start(_) => depth += 1,
            Event::Empty(_) if depth == 0 => {
                return Err(invalid("ink action root is self-closing"));
            },
            Event::End(_) if depth == 1 => {
                let close_start = source[..end]
                    .iter()
                    .rposition(|byte| *byte == b'<')
                    .ok_or_else(|| invalid("ink action root closing range is invalid"))?;
                return Ok(close_start);
            },
            Event::End(_) if depth > 0 => depth -= 1,
            Event::Eof => break,
            _ => {},
        }
    }
    Err(invalid("ink action root closing element is absent"))
}

fn action_prefix(source: &[u8]) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let name = element.name();
                let name = name.as_ref();
                let prefix_len = name.iter().position(|byte| *byte == b':').unwrap_or(0);
                let prefix = &name[..prefix_len];
                if prefix.len() > super::MAX_TOKEN_BYTES {
                    return Err(limit(
                        "ink action namespace prefix bytes",
                        super::MAX_TOKEN_BYTES,
                    ));
                }
                let mut result = Vec::new();
                result
                    .try_reserve_exact(prefix.len())
                    .map_err(|_| invalid("ink action namespace prefix allocation failed"))?;
                result.extend_from_slice(prefix);
                return Ok(result);
            },
            Event::Eof => return Err(invalid("ink action root is absent")),
            _ => {},
        }
    }
}

fn copy_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_| invalid("ink action source allocation failed"))?;
    output.extend_from_slice(bytes);
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    super::invalid(message)
}
fn limit(resource: &'static str, limit: usize) -> Error {
    super::limit(resource, limit)
}
fn xml_error(message: impl Into<String>) -> Error {
    Error::Xml(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn authored() -> Prepared {
        let property = PropertyDraft::new("dataType", "ink").expect("property");
        let data = ActionDataDraft::new()
            .with_xml_id("stroke1")
            .expect("data id")
            .transform(OpaquePayload::new(b"<iact:transform/>").expect("transform"))
            .expect("transform order")
            .trace(
                OpaquePayload::new(b"<inkml:trace><!-- </iact:action> --></inkml:trace>")
                    .expect("trace"),
            )
            .expect("trace payload");
        let action = ActionDraft::new(ActionType::Add, " 1\n")
            .expect("action")
            .with_xml_id("action1")
            .expect("action id")
            .property(property)
            .expect("property order")
            .data(data)
            .expect("data");
        Draft::with_units(LengthUnit::Centimeter, TimeUnit::Second)
            .expect("draft")
            .action(action)
            .expect("root action")
            .finish()
            .expect("authoring readback")
    }

    #[test]
    fn draft_preflights_and_reads_back_opaque_payloads() {
        let prepared = authored();
        assert_eq!(prepared.profile().children().len(), 1);
        assert!(
            prepared
                .as_bytes()
                .windows(b"<!-- </iact:action> -->".len())
                .any(|window| window == b"<!-- </iact:action> -->")
        );
    }

    #[test]
    fn edit_scalar_add_and_inverse_are_source_checked() {
        let prepared = authored();
        let original = prepared.as_bytes().to_vec();
        let profile = prepared.profile().clone();
        let mut edit = Edit::new(profile).expect("edit");
        edit.set_property_value(
            ChildSelector::Property {
                action: ActionSelector::ordinal(0),
                index: 0,
            },
            "style",
        )
        .expect("set property");
        let added = ActionDraft::new(ActionType::Custom("future".into()), "2")
            .expect("future action")
            .with_xml_id("action2")
            .expect("future id");
        edit.add_action(ActionParent::Root, added)
            .expect("add action");
        let commit = edit.finish().expect("commit");
        assert_ne!(commit.as_bytes(), original.as_slice());
        assert!(
            commit
                .as_bytes()
                .windows(b"value=\"style\"".len())
                .any(|window| window == b"value=\"style\"")
        );
        let inverse = commit
            .patch()
            .inverse()
            .apply(commit.as_bytes())
            .expect("inverse");
        assert_eq!(inverse, original);
        assert!(commit.patch().apply(b"<stale/>").is_err());
    }

    #[test]
    fn no_op_edit_replays_exact_source() {
        let prepared = authored();
        let original = prepared.as_bytes().to_vec();
        let edit = Edit::new(prepared.profile().clone()).expect("edit");
        let commit = edit.finish().expect("no-op");
        assert_eq!(commit.as_bytes(), original.as_slice());
        assert_eq!(commit.patch().before(), original.as_slice());
        assert_eq!(commit.patch().after(), original.as_slice());
    }

    #[test]
    fn group_insertion_and_semantic_reordering_keep_profile_valid() {
        let first = ActionDraft::new(ActionType::Add, "0")
            .expect("first action")
            .with_xml_id("g1")
            .expect("first id");
        let group = ActionGroupDraft::with_metadata(ActionType::Custom("batch".into()), "0", first)
            .expect("group")
            .with_xml_id("group1")
            .expect("group id");
        let prepared = Draft::with_units(LengthUnit::Meter, TimeUnit::Millisecond)
            .expect("draft")
            .action_group(group)
            .expect("root group")
            .finish()
            .expect("group readback");
        let second = ActionDraft::new(ActionType::Remove, "1")
            .expect("second action")
            .with_xml_id("g2")
            .expect("second id");
        let mut edit = Edit::from_profile(prepared.profile()).expect("edit");
        edit.add_action(ActionParent::Group(0), second)
            .expect("group insertion");
        let commit = edit.finish().expect("group commit");
        assert_eq!(commit.prepared().profile().action_groups().count(), 1);
        assert_eq!(
            commit
                .prepared()
                .profile()
                .action_groups()
                .next()
                .unwrap()
                .actions()
                .len(),
            2
        );
        let mut reorder = Edit::from_profile(commit.prepared().profile()).expect("reorder");
        reorder
            .move_before(ActionSelector::grouped(0, 1), ActionSelector::grouped(0, 0))
            .expect("move");
        let reordered = reorder.finish().expect("reordered commit");
        assert_eq!(
            reordered
                .prepared()
                .profile()
                .action_groups()
                .next()
                .unwrap()
                .actions()[0]
                .xml_id(),
            Some("g2")
        );
    }

    #[test]
    fn caller_output_limit_is_checked_before_authoring_buffer() {
        let limits = Limits {
            max_output_bytes: 1,
            ..Limits::default()
        };
        let draft = Draft::new(limits, LengthUnit::Meter, TimeUnit::Second)
            .expect("limits")
            .action(ActionDraft::new(ActionType::Add, "0").expect("action"))
            .expect("root action");
        assert!(draft.finish().is_err());
    }

    fn tag(content: &str) -> quick_xml::events::BytesStart<'static> {
        let name_len = content.find(' ').unwrap_or(content.len());
        quick_xml::events::BytesStart::from_content(content.to_owned(), name_len)
    }

    /// `e`, then `distinct` names, as many repeats of the last, and `tail`.
    fn repeated_names(distinct: usize, tail: &str) -> quick_xml::events::BytesStart<'static> {
        let mut content = String::from("e");
        for index in 0..distinct {
            content.push_str(&format!(" n{index:05}=\"\""));
        }
        let last = format!(" n{:05}=\"\"", distinct - 1);
        for _ in 0..distinct {
            content.push_str(&last);
        }
        content.push_str(tail);
        tag(&content)
    }

    #[test]
    fn payload_declared_prefixes_answer_what_the_per_prefix_scan_answers() {
        for (content, declared) in [
            ("e", &[][..]),
            (
                "e xmlns=\"urn:d\" xmlns:a=\"urn:a\" b:c=\"1\"",
                &[&b""[..], b"a"][..],
            ),
            // A repeated declaration counts once, at its first occurrence.
            ("e xmlns:a=\"urn:1\" xmlns:a=\"urn:2\"", &[&b"a"[..]]),
            // `xmlns:` names the empty prefix, as it did.
            ("e xmlns:=\"urn:x\"", &[&b""[..]]),
            // A declaration after a malformed attribute still counts.
            ("e bad=v xmlns:iact=\"urn:x\"", &[&b"iact"[..]]),
        ] {
            let element = tag(content);
            let prefixes = payload_declared_prefixes(&element);
            for prefix in [&b""[..], b"a", b"b", b"iact", b"inkml", b"xml"] {
                let expected = declared.contains(&prefix);
                assert_eq!(
                    payload_has_namespace_declaration(&element, prefix),
                    expected,
                    "{content}"
                );
                assert_eq!(prefixes.contains(&prefix), expected, "{content}");
            }
        }
    }

    #[test]
    fn namespace_declarations_are_found_on_tags_of_many_repeated_names() {
        // 20,000 names, 20,000 repeats of the last, then one declaration.
        let element = repeated_names(20_000, " xmlns:iact=\"urn:late\"");
        assert!(payload_has_namespace_declaration(&element, b"iact"));
        assert!(!payload_has_namespace_declaration(&element, b"inkml"));
        let prefixes = payload_declared_prefixes(&element);
        assert!(prefixes.contains(&&b"iact"[..]));
        assert!(!prefixes.contains(&&b"inkml"[..]));
    }

    #[test]
    fn inherited_attribute_prefixes_are_resolved_with_one_scan_of_their_tag() {
        // 50,000 distinct attributes whose `iact` prefix the payload inherits:
        // each asks whether its own tag declares `iact`.
        let mut payload = String::from("<iact:transform");
        for index in 0..50_000 {
            payload.push_str(&format!(" iact:a{index:05}=\"\""));
        }
        payload.push_str("/>");
        let mut reader = NsReader::from_reader(payload.as_bytes());
        let Event::Empty(element) = reader.read_event().expect("payload root") else {
            unreachable!("the payload is one empty element");
        };
        validate_payload_attributes_with_namespaces(&element, &reader, Limits::default(), true)
            .expect("inherited prefixes resolve");

        // The payload is refused by the per-element attribute cap, as before.
        let error = validate_payload_fragment(
            payload.as_bytes(),
            Limits::default(),
            PayloadRoot::Transform,
            true,
        )
        .unwrap_err();
        assert!(
            matches!(
                error,
                Error::Limit {
                    resource: "ink action XML attributes",
                    limit,
                } if limit == super::super::MAX_ATTRIBUTES_PER_ELEMENT
            ),
            "{error}"
        );
    }

    #[test]
    fn an_inherited_prefix_the_tag_also_declares_is_refused_as_before() {
        let accepted = b"<iact:transform><child iact:x=\"1\"/></iact:transform>";
        validate_payload_fragment(accepted, Limits::default(), PayloadRoot::Transform, true)
            .expect("an inherited attribute prefix resolves");

        // `bad=v` ends the reader's namespace scope before `xmlns:iact`, so
        // `iact:x` is unresolved although its tag declares `iact`.
        let refused = b"<iact:transform><child iact:x=\"1\" bad=v xmlns:iact=\"urn:other\"/></iact:transform>";
        let error =
            validate_payload_fragment(refused, Limits::default(), PayloadRoot::Transform, true)
                .unwrap_err();
        assert!(
            matches!(
                &error,
                Error::Invalid(message)
                    if message == "ink action opaque payload attribute uses an undeclared namespace prefix"
            ),
            "{error}"
        );
    }
}
