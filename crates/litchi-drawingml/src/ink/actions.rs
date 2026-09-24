//! Source-backed PowerPoint 2014 ink-action metadata.

use std::{fmt, io::Write, sync::Arc};

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml_name::is_qualified_name;
use quick_xml::{
    XmlVersion,
    events::{BytesDecl, BytesRef, BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

use crate::{Error, Result};

use super::{
    ACTION_NAMESPACE, INKML_NAMESPACE, MAX_ATTRIBUTE_VALUE_BYTES, MAX_DEPTH, MAX_NODES,
    MAX_SOURCE_BYTES, MAX_TOKEN_BYTES, SourceSpan,
};

#[path = "actions_edit.rs"]
pub mod edit;

pub use edit::{
    ActionDataDraft, ActionDraft, ActionGroupDraft, ActionParent, ActionSelector, ChildSelector,
    Commit, DataChildDraft, DataGroupDraft, DataSelector, Draft, Edit, Limits, OpaquePayload,
    Patch, Prepared, PropertyDraft,
};

/// Maximum action records retained by one action part.
pub const MAX_ACTIONS: usize = 65_536;
/// Maximum action-group records retained by one action part.
pub const MAX_ACTION_GROUPS: usize = 16_384;

const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_ATTRIBUTES_PER_ELEMENT: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceId {
    Action,
    Other,
    Unbound,
    Unknown,
}

/// Reserved or future PowerPoint ink action type.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ActionType {
    /// Add ink data.
    Add,
    /// Remove ink data.
    Remove,
    /// Transform ink data.
    Transform,
    /// A future or user-defined action type.
    Custom(Box<str>),
}

impl ActionType {
    fn parse(value: &str) -> Result<Self> {
        Self::parse_with_empty(value, false)
    }

    fn parse_profile(value: &str) -> Result<Self> {
        Self::parse_with_empty(value, true)
    }

    fn parse_with_empty(value: &str, allow_empty: bool) -> Result<Self> {
        if (!allow_empty && value.is_empty())
            || value.len() > 256
            || value.bytes().any(|byte| byte == 0)
        {
            return Err(invalid("ink action type is empty or overlong"));
        }
        Ok(match value {
            "add" => Self::Add,
            "remove" => Self::Remove,
            "transform" => Self::Transform,
            _ => Self::Custom(value.into()),
        })
    }
    /// Return the source lexical action type.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Add => "add",
            Self::Remove => "remove",
            Self::Transform => "transform",
            Self::Custom(value) => value,
        }
    }
}

/// Typed metadata for one CT_Action element.
///
/// The compatibility [`read`] path populates the action attributes while
/// retaining child markup opaque; child accessors are populated by
/// [`read_profile`].
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Action {
    xml_id: Option<Box<str>>,
    action_type: ActionType,
    start_time: Box<str>,
    source_start: usize,
    source_end: usize,
    properties: Vec<ActionProperty>,
    children: Vec<ActionChild>,
}

impl Action {
    /// Optional `xml:id` lexical value.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
    /// The action type.
    #[must_use]
    pub const fn action_type(&self) -> &ActionType {
        &self.action_type
    }
    /// The exact decimal start-time lexical value.
    #[must_use]
    pub fn start_time(&self) -> &str {
        &self.start_time
    }
    /// Source span of the complete action element.
    #[must_use]
    pub const fn source_span(&self) -> (usize, usize) {
        (self.source_start, self.source_end)
    }
    /// Borrow the exact action element XML.
    #[must_use]
    pub fn xml<'a>(&self, document: &'a Actions) -> &'a [u8] {
        document
            .source
            .get(self.source_start..self.source_end)
            .unwrap_or_default()
    }
    /// Additional properties in source order.
    #[must_use]
    pub fn properties(&self) -> &[ActionProperty] {
        &self.properties
    }
    /// Action data children in source order.
    #[must_use]
    pub fn children(&self) -> &[ActionChild] {
        &self.children
    }
    /// Borrow this action's source span from an arbitrary retained source.
    #[must_use]
    pub fn xml_from<'a>(&self, source: &'a [u8]) -> &'a [u8] {
        source
            .get(self.source_start..self.source_end)
            .unwrap_or_default()
    }
}

/// Exact InkML length units admitted by the strict PowerPoint ink-action
/// profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LengthUnit {
    /// Metres (`m`).
    Meter,
    /// Centimetres (`cm`).
    Centimeter,
    /// Millimetres (`mm`).
    Millimeter,
    /// Inches (`in`).
    Inch,
    /// Points (`pt`).
    Point,
    /// Picas (`pc`).
    Pica,
    /// Relative em units (`em`).
    Em,
    /// Relative ex units (`ex`).
    Ex,
}

impl LengthUnit {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "m" => Ok(Self::Meter),
            "cm" => Ok(Self::Centimeter),
            "mm" => Ok(Self::Millimeter),
            "in" => Ok(Self::Inch),
            "pt" => Ok(Self::Point),
            "pc" => Ok(Self::Pica),
            "em" => Ok(Self::Em),
            "ex" => Ok(Self::Ex),
            _ => Err(invalid(
                "ink actions lengthUnit is not an InkML standard unit",
            )),
        }
    }
    /// Return the exact schema lexical value.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Meter => "m",
            Self::Centimeter => "cm",
            Self::Millimeter => "mm",
            Self::Inch => "in",
            Self::Point => "pt",
            Self::Pica => "pc",
            Self::Em => "em",
            Self::Ex => "ex",
        }
    }
}

/// Exact InkML time units admitted by the strict PowerPoint ink-action
/// profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TimeUnit {
    /// Seconds (`s`).
    Second,
    /// Milliseconds (`ms`).
    Millisecond,
}

impl TimeUnit {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "s" => Ok(Self::Second),
            "ms" => Ok(Self::Millisecond),
            _ => Err(invalid(
                "ink actions timeUnit is not an InkML standard unit",
            )),
        }
    }
    /// Return the exact schema lexical value.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Second => "s",
            Self::Millisecond => "ms",
        }
    }
}

/// A CT_ActionProperty child retained as typed metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct ActionProperty {
    name: Box<str>,
    value: Box<str>,
    source: SourceSpan,
}

impl ActionProperty {
    /// Required property name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    /// Optional property value, including the schema default (`ink`).
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
    /// Complete source span of the property element.
    #[must_use]
    pub const fn source_span(&self) -> SourceSpan {
        self.source
    }
    /// Borrow the exact property element XML from the retained profile source.
    #[must_use]
    pub fn xml<'a>(&self, profile: &'a Profile) -> &'a [u8] {
        profile.xml(self.source)
    }
}

/// A typed child of CT_ActionData, in source order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DataChild {
    /// The optional CT_Matrix transform.
    Transform(SourceSpan),
    /// An InkML trace.
    Trace(SourceSpan),
    /// An InkML trace view.
    TraceView(SourceSpan),
}

impl DataChild {
    /// Source span of the complete child element.
    #[must_use]
    pub const fn source_span(self) -> SourceSpan {
        match self {
            Self::Transform(span) | Self::Trace(span) | Self::TraceView(span) => span,
        }
    }
}

/// A CT_ActionData element retained as typed metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct ActionData {
    xml_id: Option<Box<str>>,
    name: Box<str>,
    reference: Option<Box<str>>,
    children: Vec<DataChild>,
    source: SourceSpan,
}

impl ActionData {
    /// Optional `xml:id` lexical value.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
    /// Effective data name, including the schema default (`stroke`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    /// Optional `ref` lexical value.
    #[must_use]
    pub fn reference(&self) -> Option<&str> {
        self.reference.as_deref()
    }
    /// Recognized children in source order.
    #[must_use]
    pub fn children(&self) -> &[DataChild] {
        &self.children
    }
    /// Complete source span of the data element.
    #[must_use]
    pub const fn source_span(&self) -> SourceSpan {
        self.source
    }
    /// Borrow the exact data element XML from the retained profile source.
    #[must_use]
    pub fn xml<'a>(&self, profile: &'a Profile) -> &'a [u8] {
        profile.xml(self.source)
    }
}

/// A CT_ActionDataGroup child retained as typed metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct ActionDataGroup {
    xml_id: Option<Box<str>>,
    name: Box<str>,
    data: Vec<ActionData>,
    source: SourceSpan,
}

impl ActionDataGroup {
    /// Optional `xml:id` lexical value.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
    /// Effective data-group name, including the schema default (`stroke`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    /// Action data in source order.
    #[must_use]
    pub fn data(&self) -> &[ActionData] {
        &self.data
    }
    /// Complete source span of the data-group element.
    #[must_use]
    pub const fn source_span(&self) -> SourceSpan {
        self.source
    }
    /// Borrow the exact data-group XML from the retained profile source.
    #[must_use]
    pub fn xml<'a>(&self, profile: &'a Profile) -> &'a [u8] {
        profile.xml(self.source)
    }
}

/// A typed child of CT_Action, in source order.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ActionChild {
    /// A CT_ActionProperty child.
    Property(ActionProperty),
    /// A CT_ActionData child.
    Data(ActionData),
    /// A CT_ActionDataGroup child.
    DataGroup(ActionDataGroup),
}

/// A CT_ActionGroup element retained as typed metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct ActionGroup {
    xml_id: Option<Box<str>>,
    action_type: ActionType,
    start_time: Box<str>,
    actions: Vec<Action>,
    source: SourceSpan,
}

impl ActionGroup {
    /// Optional `xml:id` lexical value.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
    /// Action-group type.
    #[must_use]
    pub const fn action_type(&self) -> &ActionType {
        &self.action_type
    }
    /// XML-whitespace-collapsed decimal start-time value.
    ///
    /// The retained profile source remains available through [`Profile::source`]
    /// and the action span, so collapsing this typed value does not rewrite or
    /// discard the source lexical form.
    #[must_use]
    pub fn start_time(&self) -> &str {
        &self.start_time
    }
    /// Actions in source order.
    #[must_use]
    pub fn actions(&self) -> &[Action] {
        &self.actions
    }
    /// Complete source span of the group element.
    #[must_use]
    pub const fn source_span(&self) -> SourceSpan {
        self.source
    }
    /// Borrow the exact group XML from the retained profile source.
    #[must_use]
    pub fn xml<'a>(&self, profile: &'a Profile) -> &'a [u8] {
        profile.xml(self.source)
    }
}

/// A direct child of CT_Actions, in source order.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum RootChild {
    /// A CT_ActionGroup child.
    ActionGroup(ActionGroup),
    /// A CT_Action child.
    Action(Action),
}

/// Strict, namespace-bound CT_Actions metadata with exact source replay.
///
/// This profile validates the recognized §2.21 structure and preserves opaque
/// descendants only inside admitted InkML `definitions`, `trace`, `traceView`,
/// and `transform` payload boundaries. Those payloads are not a full InkML
/// schema validation pass; their bytes remain available through the retained
/// source. It does not interpret or execute `add`, `remove`, or `transform`
/// actions, and it does not require the optional action-data names described as
/// semantic conventions by §2.21.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Profile {
    source: Arc<[u8]>,
    xml_id: Option<Box<str>>,
    length_unit: LengthUnit,
    time_unit: TimeUnit,
    definitions: Option<SourceSpan>,
    children: Vec<RootChild>,
}

impl Profile {
    /// Borrow exact source bytes.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }
    /// Share the retained exact source allocation with a source-backed owner.
    ///
    /// Readers which need to retain both the parsed profile and the original
    /// bytes should clone this handle instead of copying `source()`.  The
    /// allocation is immutable and remains bounded by the profile reader.
    #[must_use]
    pub fn shared_source(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }
    /// Optional root `xml:id` lexical value.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
    /// Exact length unit.
    #[must_use]
    pub const fn length_unit(&self) -> LengthUnit {
        self.length_unit
    }
    /// Exact time unit.
    #[must_use]
    pub const fn time_unit(&self) -> TimeUnit {
        self.time_unit
    }
    /// Optional InkML definitions element span.
    #[must_use]
    pub const fn definitions_span(&self) -> Option<SourceSpan> {
        self.definitions
    }
    /// Direct action/group children in source order.
    #[must_use]
    pub fn children(&self) -> &[RootChild] {
        &self.children
    }
    /// Borrow direct actions in source order without allocating a projection.
    #[must_use]
    pub fn actions(&self) -> impl Iterator<Item = &Action> {
        self.children.iter().filter_map(|child| match child {
            RootChild::Action(action) => Some(action),
            RootChild::ActionGroup(_) => None,
        })
    }
    /// Borrow direct action groups in source order without allocating a projection.
    #[must_use]
    pub fn action_groups(&self) -> impl Iterator<Item = &ActionGroup> {
        self.children.iter().filter_map(|child| match child {
            RootChild::ActionGroup(group) => Some(group),
            RootChild::Action(_) => None,
        })
    }
    /// Borrow a checked span from the retained profile source.
    #[must_use]
    pub fn xml(&self, span: SourceSpan) -> &[u8] {
        self.source.get(span.range()).unwrap_or_default()
    }
}

/// Immutable, source-backed PowerPoint ink-actions part.
#[derive(Debug, Clone)]
#[must_use]
pub struct Actions {
    source: Arc<[u8]>,
    length_unit: Box<str>,
    time_unit: Box<str>,
    action_groups: usize,
    actions: Vec<Action>,
}

impl PartialEq for Actions {
    fn eq(&self, other: &Self) -> bool {
        self.source.as_ref() == other.source.as_ref()
            && self.length_unit == other.length_unit
            && self.time_unit == other.time_unit
            && self.action_groups == other.action_groups
            && self.actions == other.actions
    }
}
impl Eq for Actions {}

impl Actions {
    /// Borrow exact source bytes.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }
    /// Required InkML length unit.
    #[must_use]
    pub fn length_unit(&self) -> &str {
        &self.length_unit
    }
    /// Required InkML time unit.
    #[must_use]
    pub fn time_unit(&self) -> &str {
        &self.time_unit
    }
    /// Number of action groups.
    #[must_use]
    pub const fn action_group_count(&self) -> usize {
        self.action_groups
    }
    /// Borrow actions in source order.
    #[must_use]
    pub fn actions(&self) -> &[Action] {
        &self.actions
    }
}

/// Read one complete iact:actions element.
///
/// # Errors
///
/// Returns an error for malformed XML, invalid required attributes, or
/// exhausted resource limits.
pub fn read(xml: &[u8]) -> Result<Actions> {
    if xml.len() > MAX_SOURCE_BYTES {
        return Err(limit("ink actions source bytes", MAX_SOURCE_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    let mut stack = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut groups = 0usize;
    let mut actions = Vec::new();
    let mut length_unit = None;
    let mut time_unit = None;
    let mut declaration_seen = false;
    let mut preamble_content_seen = false;
    let mut legacy_fragment = false;

    loop {
        let start = pos(&reader, origin)?;
        let event = reader.read_event().map_err(xml_error)?;
        let (resolved, event) = reader.resolver().resolve_event(event);
        let end = pos(&reader, origin)?;
        let is_declaration = matches!(&event, Event::Decl(_));
        match event {
            Event::Decl(declaration)
                if !root_seen && !declaration_seen && !preamble_content_seen =>
            {
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) if !root_seen => {
                validate_element(&element, &reader, false)?;
                legacy_fragment = require_root(&element, &resolved, &reader)?;
                root_seen = true;
                increment_nodes(&mut nodes)?;
                enforce_depth(1)?;
                length_unit = Some(required_attr(&element, b"lengthUnit", &reader)?);
                time_unit = Some(required_attr(&element, b"timeUnit", &reader)?);
                validate_unit(length_unit.as_deref().unwrap_or_default(), "lengthUnit")?;
                validate_unit(time_unit.as_deref().unwrap_or_default(), "timeUnit")?;
                reserve_one(&mut stack, "ink actions XML stack")?;
                stack.push(Frame {
                    namespace: namespace_id(&resolved, &reader)?,
                    action: None,
                    action_group: false,
                    direct_actions: 0,
                });
            },
            Event::Empty(element) if !root_seen => {
                validate_element(&element, &reader, false)?;
                legacy_fragment = require_root(&element, &resolved, &reader)?;
                root_seen = true;
                root_closed = true;
                increment_nodes(&mut nodes)?;
                enforce_depth(1)?;
                length_unit = Some(required_attr(&element, b"lengthUnit", &reader)?);
                time_unit = Some(required_attr(&element, b"timeUnit", &reader)?);
                validate_unit(length_unit.as_deref().unwrap_or_default(), "lengthUnit")?;
                validate_unit(time_unit.as_deref().unwrap_or_default(), "timeUnit")?;
            },
            Event::Start(element) if root_seen && !root_closed => {
                validate_element(&element, &reader, false)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                let local = element.name().local_name();
                let action_group =
                    is_action_namespace(&resolved, &element, legacy_fragment, &reader)?
                        && local.as_ref() == b"actionGroup";
                let action_element =
                    is_action_namespace(&resolved, &element, legacy_fragment, &reader)?
                        && local.as_ref() == b"action";
                let action = if action_group {
                    increment_groups(&mut groups)?;
                    require_action_attrs(&element, &reader)?;
                    None
                } else if action_element {
                    ensure_action_capacity(&actions)?;
                    record_group_action(&mut stack);
                    let index = actions.len();
                    reserve_one(&mut actions, "ink action records")?;
                    actions.push(parse_action(&element, start, end, &reader)?);
                    Some(index)
                } else {
                    None
                };
                reserve_one(&mut stack, "ink actions XML stack")?;
                stack.push(Frame {
                    namespace: namespace_id(&resolved, &reader)?,
                    action,
                    action_group,
                    direct_actions: 0,
                });
            },
            Event::Empty(element) if root_seen && !root_closed => {
                validate_element(&element, &reader, false)?;
                validate_unknown_prefix(&resolved, &element, legacy_fragment)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                let local = element.name().local_name();
                if is_action_namespace(&resolved, &element, legacy_fragment, &reader)?
                    && local.as_ref() == b"actionGroup"
                {
                    increment_groups(&mut groups)?;
                    require_action_attrs(&element, &reader)?;
                    return Err(invalid("ink actionGroup requires an action child"));
                } else if is_action_namespace(&resolved, &element, legacy_fragment, &reader)?
                    && local.as_ref() == b"action"
                {
                    ensure_action_capacity(&actions)?;
                    record_group_action(&mut stack);
                    reserve_one(&mut actions, "ink action records")?;
                    actions.push(parse_action(&element, start, end, &reader)?);
                }
            },
            Event::End(element) if root_seen && !root_closed => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("ink actions has an unexpected end"))?;
                if frame.action_group && frame.direct_actions == 0 {
                    return Err(invalid("ink actionGroup requires an action child"));
                }
                validate_end_name(&element, &resolved, legacy_fragment)?;
                if frame.namespace != namespace_id(&resolved, &reader)? {
                    return Err(invalid("ink actions has mismatched closing elements"));
                }
                if let Some(index) = frame.action {
                    actions[index].source_end = end;
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) if root_closed => {
                return Err(invalid("ink actions has content after its root"));
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) => {
                return Err(invalid("ink actions has an invalid root transition"));
            },
            Event::Text(text)
                if (!root_seen || root_closed)
                    && !validate_text(&text, "ink actions text")?
                        .bytes()
                        .all(|byte| byte.is_ascii_whitespace()) =>
            {
                return Err(invalid("ink actions has text outside its root"));
            },
            Event::Text(text) => {
                validate_text(&text, "ink actions text")?;
            },
            Event::CData(data) => {
                validate_text(&data, "ink actions CDATA")?;
                if !root_seen || root_closed {
                    return Err(invalid("ink actions CDATA is not allowed outside its root"));
                }
            },
            Event::GeneralRef(reference) => {
                validate_reference(&reference)?;
                if !root_seen || root_closed {
                    return Err(invalid("ink actions has a reference outside its root"));
                }
            },
            Event::Comment(comment) => {
                validate_text(&comment, "ink actions comment")?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "ink actions rejects DTDs and processing instructions",
                ));
            },
            Event::Decl(_) => {
                return Err(invalid(
                    "ink actions has a duplicate or late XML declaration",
                ));
            },
            Event::Eof => break,
        }
        if !is_declaration {
            preamble_content_seen = true;
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("ink actions root is absent or unterminated"));
    }
    let length_unit = length_unit.ok_or_else(|| invalid("ink actions lengthUnit is missing"))?;
    let time_unit = time_unit.ok_or_else(|| invalid("ink actions timeUnit is missing"))?;
    let source = copy_source(xml)?;
    Ok(Actions {
        source,
        length_unit: length_unit.into(),
        time_unit: time_unit.into(),
        action_groups: groups,
        actions,
    })
}

/// Read a namespace-bound CT_Actions value using the strict §2.21 structural
/// profile.
///
/// The existing [`read`] function intentionally accepts the historical
/// unresolved-`iact` fragment form and treats unrecognized descendants as
/// opaque. This entry point is explicit: the root and recognized children must
/// resolve to their normative action namespace, while unknown descendants are
/// still retained byte-for-byte in [`Profile::source`] inside those payload
/// boundaries. Schema-adjacent conventions such as the required
/// `stroke`/`target` data names for reserved action types are deliberately not
/// enforced here; this profile is inert and does not execute actions.
pub fn read_profile(xml: &[u8]) -> Result<Profile> {
    read_profile_with_source(xml, None)
}

/// Read a profile while reusing an already-owned source allocation.
///
/// This is crate-internal so the borrowed public reader keeps its existing
/// source-copy contract.  Source-backed edits already own the emitted bytes;
/// retaining that allocation avoids a second full-source copy during typed
/// readback.
pub(crate) fn read_profile_owned(source: Arc<[u8]>) -> Result<Profile> {
    let xml = Arc::clone(&source);
    read_profile_with_source(&xml, Some(source))
}

fn read_profile_with_source(xml: &[u8], owned_source: Option<Arc<[u8]>>) -> Result<Profile> {
    if xml.len() > MAX_SOURCE_BYTES {
        return Err(limit("ink actions source bytes", MAX_SOURCE_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);

    let mut stack = Vec::new();
    let mut root_closed = false;
    let mut root = None;
    let mut nodes = 0usize;
    let mut declaration_seen = false;
    let mut preamble_content_seen = false;

    loop {
        let start = pos(&reader, origin)?;
        let event = reader.read_event().map_err(xml_error)?;
        let (resolved, event) = reader.resolver().resolve_event(event);
        let end = pos(&reader, origin)?;
        let is_declaration = matches!(&event, Event::Decl(_));
        match event {
            Event::Decl(declaration)
                if root.is_none() && !declaration_seen && !preamble_content_seen =>
            {
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) if root.is_none() => {
                validate_profile_element(&element, &resolved, &reader)?;
                let root_state = profile_root(&element, &resolved, &reader)?;
                root = Some(root_state);
                increment_nodes(&mut nodes)?;
                enforce_depth(1)?;
                reserve_one(&mut stack, "ink actions profile XML stack")?;
                stack.push(ProfileFrame {
                    namespace: NamespaceId::Action,
                    kind: ProfileFrameKind::Root,
                });
            },
            Event::Empty(element) if root.is_none() => {
                validate_profile_element(&element, &resolved, &reader)?;
                let root_state = profile_root(&element, &resolved, &reader)?;
                root = Some(root_state);
                increment_nodes(&mut nodes)?;
                enforce_depth(1)?;
                root_closed = true;
            },
            Event::Start(element) if root.is_some() && !root_closed => {
                validate_profile_element(&element, &resolved, &reader)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                let kind = profile_start_kind(
                    &mut stack,
                    root.as_mut()
                        .ok_or_else(|| invalid("ink actions profile root is missing"))?,
                    &element,
                    &resolved,
                    start,
                    &reader,
                )?;
                let namespace = namespace_id(&resolved, &reader)?;
                reserve_one(&mut stack, "ink actions profile XML stack")?;
                stack.push(ProfileFrame { namespace, kind });
            },
            Event::Empty(element) if root.is_some() && !root_closed => {
                validate_profile_element(&element, &resolved, &reader)?;
                increment_nodes(&mut nodes)?;
                enforce_depth(stack.len().saturating_add(1))?;
                let kind = profile_start_kind(
                    &mut stack,
                    root.as_mut()
                        .ok_or_else(|| invalid("ink actions profile root is missing"))?,
                    &element,
                    &resolved,
                    start,
                    &reader,
                )?;
                let value = profile_finish_empty(kind, start, end)?;
                attach_profile_value(
                    stack.last_mut(),
                    root.as_mut()
                        .ok_or_else(|| invalid("ink actions profile root is missing"))?,
                    value,
                )?;
            },
            Event::End(element) if root.is_some() && !root_closed => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("ink actions profile has an unexpected end"))?;
                validate_end_name(&element, &resolved, false)?;
                if frame.namespace != namespace_id(&resolved, &reader)? {
                    return Err(invalid(
                        "ink actions profile has mismatched closing namespaces",
                    ));
                }
                if stack.is_empty() {
                    if !matches!(frame.kind, ProfileFrameKind::Root) {
                        return Err(invalid("ink actions profile root frame is invalid"));
                    }
                    root_closed = true;
                } else {
                    let value = profile_finish(frame.kind, end)?;
                    attach_profile_value(
                        stack.last_mut(),
                        root.as_mut()
                            .ok_or_else(|| invalid("ink actions profile root is missing"))?,
                        value,
                    )?;
                }
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) if root_closed => {
                return Err(invalid("ink actions profile has content after its root"));
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) => {
                return Err(invalid(
                    "ink actions profile has an invalid root transition",
                ));
            },
            Event::Text(text) => {
                let text = validate_text(&text, "ink actions profile text")?;
                if text.contains("]]>") {
                    return Err(invalid(
                        "ink actions profile text contains a raw CDATA terminator",
                    ));
                }
                if (root.is_none() || root_closed) && !text.bytes().all(is_xml_whitespace) {
                    return Err(invalid("ink actions profile has text outside its root"));
                }
                if root.is_some() && !root_closed && !text.is_empty() {
                    ensure_profile_character_content(&stack, text.bytes().all(is_xml_whitespace))?;
                }
            },
            Event::CData(data) => {
                let data = validate_text(&data, "ink actions profile CDATA")?;
                if root.is_none() || root_closed {
                    return Err(invalid("ink actions profile has CDATA outside its root"));
                }
                if !data.is_empty() {
                    ensure_profile_character_content(&stack, data.bytes().all(is_xml_whitespace))?;
                }
            },
            Event::GeneralRef(reference) => {
                validate_reference(&reference)?;
                if root.is_none() || root_closed {
                    return Err(invalid(
                        "ink actions profile has a reference outside its root",
                    ));
                }
                let whitespace = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .is_some_and(|character| matches!(character, ' ' | '\t' | '\r' | '\n'));
                ensure_profile_character_content(&stack, whitespace)?;
            },
            Event::Comment(comment) => {
                validate_text(&comment, "ink actions profile comment")?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "ink actions profile rejects DTDs and processing instructions",
                ));
            },
            Event::Decl(_) => {
                return Err(invalid(
                    "ink actions profile has a duplicate or late XML declaration",
                ));
            },
            Event::Eof => break,
        }
        if !is_declaration {
            preamble_content_seen = true;
        }
    }

    if root.is_none() || !root_closed || !stack.is_empty() {
        return Err(invalid(
            "ink actions profile root is absent or unterminated",
        ));
    }
    let root = root.ok_or_else(|| invalid("ink actions profile root is missing"))?;
    let source = match owned_source {
        Some(source) => source,
        None => copy_source(xml)?,
    };
    Ok(Profile {
        source,
        xml_id: root.xml_id,
        length_unit: root.length_unit,
        time_unit: root.time_unit,
        definitions: root.definitions,
        children: root.children,
    })
}

/// Write a strict profile using its exact retained source bytes.
pub fn write_profile(profile: &Profile) -> Result<Vec<u8>> {
    if profile.source.len() > MAX_SOURCE_BYTES {
        return Err(limit("ink actions source bytes", MAX_SOURCE_BYTES));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(profile.source.len())
        .map_err(|_| invalid("ink actions profile output allocation failed"))?;
    output.extend_from_slice(profile.source());
    Ok(output)
}

/// Write a strict profile to a caller-provided sink.
pub fn write_profile_to<W: Write>(writer: &mut W, profile: &Profile) -> Result<()> {
    writer.write_all(profile.source())?;
    Ok(())
}

/// Write one action part using exact retained source bytes.
///
/// # Errors
///
/// Returns an error only when the caller-provided sink fails.
pub fn write_to<W: Write>(writer: &mut W, actions: &Actions) -> Result<()> {
    writer.write_all(actions.source())?;
    Ok(())
}

/// Write one action part to a byte vector.
///
/// # Errors
///
/// Returns an error when source validation fails.
pub fn write(actions: &Actions) -> Result<Vec<u8>> {
    if actions.source().len() > MAX_SOURCE_BYTES {
        return Err(limit("ink actions source bytes", MAX_SOURCE_BYTES));
    }
    Ok(actions.source().to_vec())
}

#[derive(Debug)]
struct ProfileRoot {
    xml_id: Option<Box<str>>,
    length_unit: LengthUnit,
    time_unit: TimeUnit,
    definitions: Option<SourceSpan>,
    children: Vec<RootChild>,
    definitions_seen: bool,
    action_seen: bool,
    action_count: usize,
    group_count: usize,
}

#[derive(Debug)]
struct ProfileFrame {
    namespace: NamespaceId,
    kind: ProfileFrameKind,
}

#[derive(Debug, Clone, Copy)]
enum ProfileAttributeSet {
    Root,
    Action,
    Property,
    ActionData,
    ActionDataGroup,
}

fn validate_profile_attributes(
    element: &BytesStart<'_>,
    allowed: ProfileAttributeSet,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.strip_prefix(b"xmlns:").is_some() {
            continue;
        }
        if key == b"xml:id" {
            if !matches!(
                allowed,
                ProfileAttributeSet::Root
                    | ProfileAttributeSet::Action
                    | ProfileAttributeSet::ActionData
                    | ProfileAttributeSet::ActionDataGroup
            ) {
                return Err(invalid(
                    "ink actions profile schema-owned element has an unexpected xml:id",
                ));
            }
            continue;
        }
        if attribute.key.prefix().is_some() {
            return Err(invalid(
                "ink actions profile schema-owned element has an unknown qualified attribute",
            ));
        }
        let is_allowed = match allowed {
            ProfileAttributeSet::Root => matches!(key, b"lengthUnit" | b"timeUnit"),
            ProfileAttributeSet::Action => matches!(key, b"type" | b"startTime"),
            ProfileAttributeSet::Property => matches!(key, b"name" | b"value"),
            ProfileAttributeSet::ActionData => matches!(key, b"name" | b"ref"),
            ProfileAttributeSet::ActionDataGroup => key == b"name",
        };
        if !is_allowed {
            return Err(invalid(
                "ink actions profile schema-owned element has an unknown attribute",
            ));
        }
    }
    Ok(())
}

#[derive(Debug)]
enum ProfileFrameKind {
    Root,
    Definitions {
        source_start: usize,
    },
    ActionGroup {
        xml_id: Option<Box<str>>,
        action_type: ActionType,
        start_time: Box<str>,
        actions: Vec<Action>,
        source_start: usize,
    },
    Action {
        xml_id: Option<Box<str>>,
        action_type: ActionType,
        start_time: Box<str>,
        properties: Vec<ActionProperty>,
        children: Vec<ActionChild>,
        source_start: usize,
        data_seen: bool,
    },
    Property {
        name: Box<str>,
        value: Box<str>,
        source_start: usize,
    },
    DataGroup {
        xml_id: Option<Box<str>>,
        name: Box<str>,
        data: Vec<ActionData>,
        source_start: usize,
    },
    Data {
        xml_id: Option<Box<str>>,
        name: Box<str>,
        reference: Option<Box<str>>,
        children: Vec<DataChild>,
        source_start: usize,
        transform_seen: bool,
        trace_seen: bool,
    },
    DataChild {
        kind: DataChildKind,
        source_start: usize,
    },
    Opaque,
}

#[derive(Debug, Clone, Copy)]
enum DataChildKind {
    Transform,
    Trace,
    TraceView,
}

#[derive(Debug)]
enum ProfileValue {
    Definitions(SourceSpan),
    ActionGroup(ActionGroup),
    Action(Action),
    Property(ActionProperty),
    DataGroup(ActionDataGroup),
    Data(ActionData),
    DataChild(DataChild),
    None,
}

fn profile_root<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
) -> Result<ProfileRoot> {
    if !namespace_matches(resolved, ACTION_NAMESPACE.as_bytes(), reader)?
        || element.name().local_name().as_ref() != b"actions"
    {
        return Err(invalid(
            "ink actions profile root must be bound iact:actions",
        ));
    }
    validate_profile_attributes(element, ProfileAttributeSet::Root)?;
    let length = required_attr(element, b"lengthUnit", reader)?;
    let time = required_attr(element, b"timeUnit", reader)?;
    Ok(ProfileRoot {
        xml_id: optional_xml_id(element, reader)?,
        length_unit: LengthUnit::parse(&length)?,
        time_unit: TimeUnit::parse(&time)?,
        definitions: None,
        children: Vec::new(),
        definitions_seen: false,
        action_seen: false,
        action_count: 0,
        group_count: 0,
    })
}

fn validate_profile_element<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
) -> Result<()> {
    validate_element(element, reader, true)?;
    validate_unknown_prefix(resolved, element, false)
}

fn profile_start_kind<R: std::io::BufRead>(
    stack: &mut [ProfileFrame],
    root: &mut ProfileRoot,
    element: &BytesStart<'_>,
    resolved: &ResolveResult<'_>,
    start: usize,
    reader: &NsReader<R>,
) -> Result<ProfileFrameKind> {
    let local = element.name().local_name();
    let action = namespace_matches(resolved, ACTION_NAMESPACE.as_bytes(), reader)?;
    let inkml = namespace_matches(resolved, INKML_NAMESPACE.as_bytes(), reader)?;
    let local = local.as_ref();
    let known_action = is_known_action_local(local);
    let parent = stack
        .last_mut()
        .ok_or_else(|| invalid("ink actions profile has no parent frame"))?;
    match &mut parent.kind {
        ProfileFrameKind::Root => {
            if inkml && local == b"definitions" {
                if root.definitions_seen || root.action_seen {
                    return Err(invalid(
                        "ink actions profile definitions must precede actions and be unique",
                    ));
                }
                root.definitions_seen = true;
                return Ok(ProfileFrameKind::Definitions {
                    source_start: start,
                });
            }
            if action && local == b"actionGroup" {
                increment_profile_group(root)?;
                validate_profile_attributes(element, ProfileAttributeSet::Action)?;
                root.action_seen = true;
                let (xml_id, action_type, start_time) = profile_action_attributes(element, reader)?;
                return Ok(ProfileFrameKind::ActionGroup {
                    xml_id,
                    action_type,
                    start_time,
                    actions: Vec::new(),
                    source_start: start,
                });
            }
            if action && local == b"action" {
                increment_profile_action(root)?;
                validate_profile_attributes(element, ProfileAttributeSet::Action)?;
                root.action_seen = true;
                let (xml_id, action_type, start_time) = profile_action_attributes(element, reader)?;
                return Ok(ProfileFrameKind::Action {
                    xml_id,
                    action_type,
                    start_time,
                    properties: Vec::new(),
                    children: Vec::new(),
                    source_start: start,
                    data_seen: false,
                });
            }
            if action && known_action {
                return Err(invalid(
                    "ink actions profile recognized child has invalid root ancestry",
                ));
            }
            return Err(invalid(
                "ink actions profile root contains an unknown direct child",
            ));
        },
        ProfileFrameKind::Definitions { .. } => {
            return Ok(ProfileFrameKind::Opaque);
        },
        ProfileFrameKind::ActionGroup { .. } => {
            if action && local == b"action" {
                increment_profile_action(root)?;
                validate_profile_attributes(element, ProfileAttributeSet::Action)?;
                let (xml_id, action_type, start_time) = profile_action_attributes(element, reader)?;
                return Ok(ProfileFrameKind::Action {
                    xml_id,
                    action_type,
                    start_time,
                    properties: Vec::new(),
                    children: Vec::new(),
                    source_start: start,
                    data_seen: false,
                });
            }
            if action && known_action {
                return Err(invalid(
                    "ink actions profile recognized child has invalid action-group ancestry",
                ));
            }
            return Err(invalid(
                "ink actions profile actionGroup contains an unknown direct child",
            ));
        },
        ProfileFrameKind::Action { data_seen, .. } => {
            if action && local == b"property" {
                validate_profile_attributes(element, ProfileAttributeSet::Property)?;
                if *data_seen {
                    return Err(invalid(
                        "ink actions profile property must precede action data",
                    ));
                }
                let name = strict_token(required_attr(element, b"name", reader)?, "property name")?;
                let value = strict_token(
                    optional_attr(element, b"value", reader)?.unwrap_or_else(|| "ink".to_owned()),
                    "property value",
                )?;
                return Ok(ProfileFrameKind::Property {
                    name,
                    value,
                    source_start: start,
                });
            }
            if action && (local == b"actionData" || local == b"actionDataGroup") {
                *data_seen = true;
                if local == b"actionData" {
                    validate_profile_attributes(element, ProfileAttributeSet::ActionData)?;
                    let (xml_id, name, reference) = profile_data_attributes(element, reader)?;
                    return Ok(ProfileFrameKind::Data {
                        xml_id,
                        name,
                        reference,
                        children: Vec::new(),
                        source_start: start,
                        transform_seen: false,
                        trace_seen: false,
                    });
                }
                validate_profile_attributes(element, ProfileAttributeSet::ActionDataGroup)?;
                let xml_id = optional_xml_id(element, reader)?;
                let name = strict_token(
                    optional_attr(element, b"name", reader)?.unwrap_or_else(|| "stroke".to_owned()),
                    "action data-group name",
                )?;
                return Ok(ProfileFrameKind::DataGroup {
                    xml_id,
                    name,
                    data: Vec::new(),
                    source_start: start,
                });
            }
            if action && known_action {
                return Err(invalid(
                    "ink actions profile recognized child has invalid action ancestry",
                ));
            }
            return Err(invalid(
                "ink actions profile action contains an unknown direct child",
            ));
        },
        ProfileFrameKind::DataGroup { .. } => {
            if action && local == b"actionData" {
                validate_profile_attributes(element, ProfileAttributeSet::ActionData)?;
                let (xml_id, name, reference) = profile_data_attributes(element, reader)?;
                return Ok(ProfileFrameKind::Data {
                    xml_id,
                    name,
                    reference,
                    children: Vec::new(),
                    source_start: start,
                    transform_seen: false,
                    trace_seen: false,
                });
            }
            if action && known_action {
                return Err(invalid(
                    "ink actions profile recognized child has invalid data-group ancestry",
                ));
            }
            return Err(invalid(
                "ink actions profile actionDataGroup contains an unknown direct child",
            ));
        },
        ProfileFrameKind::Data {
            transform_seen,
            trace_seen,
            ..
        } => {
            if action && local == b"transform" {
                if *transform_seen || *trace_seen {
                    return Err(invalid(
                        "ink actions profile transform must be first and unique",
                    ));
                }
                *transform_seen = true;
                return Ok(ProfileFrameKind::DataChild {
                    kind: DataChildKind::Transform,
                    source_start: start,
                });
            }
            if inkml && (local == b"trace" || local == b"traceView") {
                *trace_seen = true;
                return Ok(ProfileFrameKind::DataChild {
                    kind: if local == b"trace" {
                        DataChildKind::Trace
                    } else {
                        DataChildKind::TraceView
                    },
                    source_start: start,
                });
            }
            if action && known_action {
                return Err(invalid(
                    "ink actions profile recognized child has invalid data ancestry",
                ));
            }
            return Err(invalid(
                "ink actions profile actionData contains an unknown direct child",
            ));
        },
        ProfileFrameKind::Property { .. } => {
            return Err(invalid(
                "ink actions profile property contains a child element",
            ));
        },
        ProfileFrameKind::DataChild { .. } | ProfileFrameKind::Opaque => {},
    }
    Ok(ProfileFrameKind::Opaque)
}

fn profile_finish_empty(kind: ProfileFrameKind, _start: usize, end: usize) -> Result<ProfileValue> {
    profile_finish(kind, end)
}

fn profile_finish(kind: ProfileFrameKind, end: usize) -> Result<ProfileValue> {
    match kind {
        ProfileFrameKind::Root => Err(invalid("ink actions profile root cannot be nested")),
        ProfileFrameKind::Definitions { source_start } => Ok(ProfileValue::Definitions(
            SourceSpan::new(source_start, end),
        )),
        ProfileFrameKind::ActionGroup {
            xml_id,
            action_type,
            start_time,
            actions,
            source_start,
        } => {
            if actions.is_empty() {
                return Err(invalid(
                    "ink actions profile actionGroup requires an action child",
                ));
            }
            Ok(ProfileValue::ActionGroup(ActionGroup {
                xml_id,
                action_type,
                start_time,
                actions,
                source: SourceSpan::new(source_start, end),
            }))
        },
        ProfileFrameKind::Action {
            xml_id,
            action_type,
            start_time,
            properties,
            children,
            source_start,
            ..
        } => Ok(ProfileValue::Action(Action {
            xml_id,
            action_type,
            start_time,
            source_start,
            source_end: end,
            properties,
            children,
        })),
        ProfileFrameKind::Property {
            name,
            value,
            source_start,
        } => Ok(ProfileValue::Property(ActionProperty {
            name,
            value,
            source: SourceSpan::new(source_start, end),
        })),
        ProfileFrameKind::DataGroup {
            xml_id,
            name,
            data,
            source_start,
        } => {
            if data.is_empty() {
                return Err(invalid(
                    "ink actions profile actionDataGroup requires actionData",
                ));
            }
            Ok(ProfileValue::DataGroup(ActionDataGroup {
                xml_id,
                name,
                data,
                source: SourceSpan::new(source_start, end),
            }))
        },
        ProfileFrameKind::Data {
            xml_id,
            name,
            reference,
            children,
            source_start,
            ..
        } => Ok(ProfileValue::Data(ActionData {
            xml_id,
            name,
            reference,
            children,
            source: SourceSpan::new(source_start, end),
        })),
        ProfileFrameKind::DataChild { kind, source_start } => Ok(ProfileValue::DataChild(
            DataChild::from_kind(kind, SourceSpan::new(source_start, end)),
        )),
        ProfileFrameKind::Opaque => Ok(ProfileValue::None),
    }
}

fn attach_profile_value(
    frame: Option<&mut ProfileFrame>,
    root: &mut ProfileRoot,
    value: ProfileValue,
) -> Result<()> {
    let Some(frame) = frame else {
        return Err(invalid("ink actions profile value has no parent"));
    };
    match (&mut frame.kind, value) {
        (ProfileFrameKind::Root, ProfileValue::Definitions(span)) => {
            root.definitions = Some(span);
        },
        (ProfileFrameKind::Root, ProfileValue::ActionGroup(group)) => {
            root.children
                .try_reserve(1)
                .map_err(|_| invalid("ink actions profile root allocation failed"))?;
            root.children.push(RootChild::ActionGroup(group));
        },
        (ProfileFrameKind::Root, ProfileValue::Action(action)) => {
            root.children
                .try_reserve(1)
                .map_err(|_| invalid("ink actions profile root allocation failed"))?;
            root.children.push(RootChild::Action(action));
        },
        (ProfileFrameKind::ActionGroup { actions, .. }, ProfileValue::Action(action)) => {
            actions
                .try_reserve(1)
                .map_err(|_| invalid("ink actions profile action-group allocation failed"))?;
            actions.push(action);
        },
        (ProfileFrameKind::Action { properties, .. }, ProfileValue::Property(property)) => {
            properties
                .try_reserve(1)
                .map_err(|_| invalid("ink actions profile property allocation failed"))?;
            properties.push(property.clone());
            if let ProfileFrameKind::Action { children, .. } = &mut frame.kind {
                children
                    .try_reserve(1)
                    .map_err(|_| invalid("ink actions profile action allocation failed"))?;
                children.push(ActionChild::Property(property));
            }
        },
        (ProfileFrameKind::Action { children, .. }, ProfileValue::Data(data)) => {
            children
                .try_reserve(1)
                .map_err(|_| invalid("ink actions profile action allocation failed"))?;
            children.push(ActionChild::Data(data));
        },
        (ProfileFrameKind::Action { children, .. }, ProfileValue::DataGroup(group)) => {
            children
                .try_reserve(1)
                .map_err(|_| invalid("ink actions profile action allocation failed"))?;
            children.push(ActionChild::DataGroup(group));
        },
        (ProfileFrameKind::DataGroup { data, .. }, ProfileValue::Data(item)) => {
            data.try_reserve(1)
                .map_err(|_| invalid("ink actions profile data-group allocation failed"))?;
            data.push(item);
        },
        (ProfileFrameKind::Data { children, .. }, ProfileValue::DataChild(child)) => {
            children
                .try_reserve(1)
                .map_err(|_| invalid("ink actions profile data allocation failed"))?;
            children.push(child);
        },
        (_, ProfileValue::None) => {},
        _ => {
            return Err(invalid(
                "ink actions profile recognized child has invalid ancestry",
            ));
        },
    }
    Ok(())
}

impl DataChild {
    fn from_kind(kind: DataChildKind, span: SourceSpan) -> Self {
        match kind {
            DataChildKind::Transform => Self::Transform(span),
            DataChildKind::Trace => Self::Trace(span),
            DataChildKind::TraceView => Self::TraceView(span),
        }
    }
}

fn profile_action_attributes<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<(Option<Box<str>>, ActionType, Box<str>)> {
    let xml_id = optional_xml_id(element, reader)?;
    let action_type = ActionType::parse_profile(&required_attr(element, b"type", reader)?)?;
    let start_time = collapse_xml_whitespace(&required_attr(element, b"startTime", reader)?)?;
    validate_decimal(&start_time)?;
    Ok((xml_id, action_type, start_time.into_boxed_str()))
}

fn collapse_xml_whitespace(value: &str) -> Result<String> {
    let mut result = String::new();
    result
        .try_reserve_exact(value.len())
        .map_err(|_| invalid("ink action startTime allocation failed"))?;
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, ' ' | '\t' | '\r' | '\n') {
            if !result.is_empty() {
                pending_space = true;
            }
        } else {
            if pending_space {
                result.push(' ');
                pending_space = false;
            }
            result.push(character);
        }
    }
    Ok(result)
}

fn profile_data_attributes<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<(Option<Box<str>>, Box<str>, Option<Box<str>>)> {
    let xml_id = optional_xml_id(element, reader)?;
    let name = strict_token(
        optional_attr(element, b"name", reader)?.unwrap_or_else(|| "stroke".to_owned()),
        "action data name",
    )?;
    let reference = optional_attr(element, b"ref", reader)?
        .map(|value| strict_token(value, "action data reference"))
        .transpose()?;
    Ok((xml_id, name, reference))
}

fn strict_token(value: String, field: &'static str) -> Result<Box<str>> {
    if value.len() > MAX_TOKEN_BYTES || value.bytes().any(|byte| byte == 0) {
        return Err(limit(field, MAX_TOKEN_BYTES));
    }
    Ok(value.into_boxed_str())
}

fn increment_profile_action(root: &mut ProfileRoot) -> Result<()> {
    let next = root
        .action_count
        .checked_add(1)
        .ok_or_else(|| limit("ink actions", MAX_ACTIONS))?;
    if next > MAX_ACTIONS {
        return Err(limit("ink actions", MAX_ACTIONS));
    }
    root.action_count = next;
    Ok(())
}

fn increment_profile_group(root: &mut ProfileRoot) -> Result<()> {
    let next = root
        .group_count
        .checked_add(1)
        .ok_or_else(|| limit("ink action groups", MAX_ACTION_GROUPS))?;
    if next > MAX_ACTION_GROUPS {
        return Err(limit("ink action groups", MAX_ACTION_GROUPS));
    }
    root.group_count = next;
    Ok(())
}

fn optional_attr<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    name: &[u8],
    reader: &NsReader<R>,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.prefix().is_some() || attribute.key.local_name().as_ref() != name {
            continue;
        }
        if result.is_some() {
            return Err(invalid(
                "ink actions profile has a duplicate typed attribute",
            ));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "ink actions attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "ink actions profile attribute")?;
        result = Some(value.into_owned());
    }
    Ok(result)
}

fn optional_xml_id<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<Option<Box<str>>> {
    let mut result = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() != b"xml:id" {
            continue;
        }
        if result.is_some() {
            return Err(invalid("ink actions profile has a duplicate xml:id"));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "ink actions attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "ink actions profile xml:id")?;
        result = Some(strict_token(value.into_owned(), "xml:id")?);
    }
    Ok(result)
}

fn is_known_action_local(local: &[u8]) -> bool {
    matches!(
        local,
        b"action"
            | b"actionGroup"
            | b"property"
            | b"actionData"
            | b"actionDataGroup"
            | b"transform"
    )
}

fn namespace_matches<R: std::io::BufRead>(
    resolved: &ResolveResult<'_>,
    expected: &[u8],
    reader: &NsReader<R>,
) -> Result<bool> {
    let ResolveResult::Bound(Namespace(value)) = resolved else {
        return Ok(false);
    };
    if *value == expected {
        return Ok(true);
    }
    Ok(decoded_namespace(value, reader)?.as_bytes() == expected)
}

fn decoded_namespace<R: std::io::BufRead>(value: &[u8], reader: &NsReader<R>) -> Result<String> {
    let decoded = reader.decoder().decode(value).map_err(xml_error)?;
    let unescaped = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
    let mut result = String::new();
    result
        .try_reserve_exact(unescaped.len())
        .map_err(|_| invalid("ink actions namespace allocation failed"))?;
    result.push_str(&unescaped);
    Ok(result)
}

fn ensure_profile_character_content(stack: &[ProfileFrame], whitespace: bool) -> Result<()> {
    if matches!(
        stack.last().map(|frame| &frame.kind),
        Some(ProfileFrameKind::Property { .. })
    ) {
        return Err(invalid("ink action property has character content"));
    }
    if whitespace {
        return Ok(());
    }
    match stack.last().map(|frame| &frame.kind) {
        Some(ProfileFrameKind::Definitions { .. })
        | Some(ProfileFrameKind::DataChild { .. })
        | Some(ProfileFrameKind::Opaque)
        | None => Ok(()),
        _ => Err(invalid(
            "ink actions profile recognized element-only content contains text",
        )),
    }
}

#[derive(Debug)]
struct Frame {
    namespace: NamespaceId,
    action: Option<usize>,
    action_group: bool,
    direct_actions: usize,
}

fn record_group_action(stack: &mut [Frame]) {
    if let Some(frame) = stack.last_mut()
        && frame.action_group
    {
        frame.direct_actions = frame.direct_actions.saturating_add(1);
    }
}

fn parse_action<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
    reader: &NsReader<R>,
) -> Result<Action> {
    let action_type = required_attr(element, b"type", reader)?;
    let start_time = required_attr(element, b"startTime", reader)?;
    validate_decimal(&start_time)?;
    Ok(Action {
        xml_id: optional_xml_id(element, reader)?,
        action_type: ActionType::parse(&action_type)?,
        start_time: start_time.into_boxed_str(),
        source_start: start,
        source_end: end,
        properties: Vec::new(),
        children: Vec::new(),
    })
}

fn require_action_attrs<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
) -> Result<()> {
    let kind = required_attr(element, b"type", reader)?;
    let time = required_attr(element, b"startTime", reader)?;
    ActionType::parse(&kind)?;
    validate_decimal(&time)
}

fn required_attr<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    name: &[u8],
    reader: &NsReader<R>,
) -> Result<String> {
    let mut result = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.prefix().is_some() {
            continue;
        }
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        if result.is_some() {
            return Err(invalid("ink actions has a duplicate typed attribute"));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "ink actions attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "ink actions attribute")?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_| invalid("ink actions attribute allocation failed"))?;
        owned.push_str(&value);
        result = Some(owned);
    }
    result.ok_or_else(|| {
        invalid(format!(
            "ink actions attribute '{}' is missing",
            String::from_utf8_lossy(name)
        ))
    })
}

fn validate_decimal(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > 256 {
        return Err(invalid("ink action decimal is empty or overlong"));
    }
    let mut digits = 0usize;
    let mut dot = false;
    for (index, byte) in value.bytes().enumerate() {
        if (byte == b'+' || byte == b'-') && index == 0 {
            continue;
        }
        if byte == b'.' && !dot {
            dot = true;
            continue;
        }
        if !byte.is_ascii_digit() {
            return Err(invalid("ink action startTime is not xsd:decimal"));
        }
        digits += 1;
    }
    if digits == 0 {
        Err(invalid("ink action startTime has no digits"))
    } else {
        Ok(())
    }
}

fn validate_unit(value: &str, name: &'static str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_TOKEN_BYTES
        || value
            .bytes()
            .any(|byte| byte == 0 || byte.is_ascii_whitespace())
    {
        return Err(invalid(format!(
            "ink actions {name} is empty, overlong, or contains whitespace"
        )));
    }
    Ok(())
}

fn require_root<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
) -> Result<bool> {
    if element.name().local_name().as_ref() != b"actions" {
        return Err(invalid("ink actions root must be iact:actions"));
    }
    if namespace_matches(resolved, ACTION_NAMESPACE.as_bytes(), reader)? {
        return Ok(false);
    }
    match resolved {
        ResolveResult::Unknown(prefix)
            if prefix.as_slice() == b"iact"
                && !has_namespace_declaration(element, prefix.as_slice()) =>
        {
            Ok(true)
        },
        _ => Err(invalid("ink actions root must be iact:actions")),
    }
}

fn is_action_namespace(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    legacy_fragment: bool,
    reader: &NsReader<impl std::io::BufRead>,
) -> Result<bool> {
    if namespace_matches(resolved, ACTION_NAMESPACE.as_bytes(), reader)? {
        return Ok(true);
    }
    Ok(match resolved {
        ResolveResult::Unknown(prefix) => {
            legacy_fragment
                && prefix.as_slice() == b"iact"
                && !has_namespace_declaration(element, prefix.as_slice())
        },
        ResolveResult::Bound(_) | ResolveResult::Unbound => false,
    })
}

fn namespace_id<R: std::io::BufRead>(
    resolved: &ResolveResult<'_>,
    reader: &NsReader<R>,
) -> Result<NamespaceId> {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => {
            let action = namespace_matches(resolved, ACTION_NAMESPACE.as_bytes(), reader)?;
            std::str::from_utf8(value).map_err(xml_error)?;
            Ok(if action {
                NamespaceId::Action
            } else {
                NamespaceId::Other
            })
        },
        ResolveResult::Unknown(prefix) => {
            std::str::from_utf8(prefix).map_err(xml_error)?;
            Ok(NamespaceId::Unknown)
        },
        ResolveResult::Unbound => Ok(NamespaceId::Unbound),
    }
}

fn validate_unknown_prefix(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    legacy_fragment: bool,
) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = resolved {
        std::str::from_utf8(prefix).map_err(xml_error)?;
        if !legacy_fragment
            || has_namespace_declaration(element, prefix.as_slice())
            || prefix.as_slice() != b"iact"
        {
            return Err(invalid(
                "ink actions element uses an undeclared namespace prefix",
            ));
        }
    }
    Ok(())
}

fn validate_end_name(
    element: &quick_xml::events::BytesEnd<'_>,
    resolved: &ResolveResult<'_>,
    legacy_fragment: bool,
) -> Result<()> {
    let element_name = element.name();
    let name = std::str::from_utf8(element_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("ink actions closing element name is invalid"));
    }
    if let ResolveResult::Unknown(prefix) = resolved {
        std::str::from_utf8(prefix).map_err(xml_error)?;
        if !legacy_fragment || prefix.as_slice() != b"iact" {
            return Err(invalid(
                "ink actions closing element uses an undeclared namespace prefix",
            ));
        }
    }
    Ok(())
}

fn validate_element<R: std::io::BufRead>(
    element: &BytesStart<'_>,
    reader: &NsReader<R>,
    strict_profile: bool,
) -> Result<()> {
    let element_name = element.name();
    let name = std::str::from_utf8(element_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("ink actions element name is invalid"));
    }
    let mut attribute_keys = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute_keys.len() >= MAX_ATTRIBUTES_PER_ELEMENT {
            return Err(limit(
                "ink actions attributes per element",
                MAX_ATTRIBUTES_PER_ELEMENT,
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        if !is_qualified_name(name) {
            return Err(invalid("ink actions attribute name is invalid"));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "ink actions attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        if strict_profile && attribute.value.contains(&b'<') {
            return Err(invalid(
                "ink actions profile attribute contains a raw less-than sign",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_characters(&value, "ink actions attribute")?;
        if strict_profile {
            let key = attribute.key.as_ref();
            let declared_prefix = if key == b"xmlns" {
                Some(b"".as_slice())
            } else {
                key.strip_prefix(b"xmlns:")
            };
            if let Some(prefix) = declared_prefix {
                const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
                const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
                if prefix == b"xmlns"
                    || value == XMLNS_NAMESPACE
                    || (value == XML_NAMESPACE && prefix != b"xml")
                    || (prefix == b"xml" && value != XML_NAMESPACE)
                    || (!prefix.is_empty() && value.is_empty())
                {
                    return Err(invalid(
                        "ink actions profile has an invalid namespace binding",
                    ));
                }
            }
        }
        if let Some(prefix) = attribute.key.prefix()
            && !matches!(prefix.as_ref(), b"xml" | b"xmlns")
            && matches!(
                reader.resolver().resolve_attribute(attribute.key).0,
                ResolveResult::Unknown(_)
            )
        {
            return Err(invalid(
                "ink actions attribute uses an undeclared namespace prefix",
            ));
        }
        for key in attribute_keys.iter().copied() {
            if expanded_attribute_names_equal(key, attribute.key, reader)? {
                return Err(invalid(
                    "ink actions element has duplicate expanded attributes",
                ));
            }
        }
        attribute_keys
            .try_reserve(1)
            .map_err(|_| invalid("ink actions attribute-key allocation failed"))?;
        attribute_keys.push(attribute.key);
    }
    Ok(())
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let raw = declaration.as_ref();
    std::str::from_utf8(raw).map_err(xml_error)?;
    let mut cursor = 0;
    skip_decl_whitespace(raw, &mut cursor);
    if !consume_decl_token(raw, &mut cursor, b"xml") || !decl_whitespace(raw.get(cursor).copied()) {
        return Err(invalid("ink actions XML declaration is malformed"));
    }
    skip_decl_whitespace(raw, &mut cursor);
    let (name, value) = parse_decl_attribute(raw, &mut cursor)?;
    if name != b"version" || value != b"1.0" {
        return Err(invalid(
            "ink actions XML declaration must start with version 1.0",
        ));
    }
    let mut previous = b"version".as_slice();
    while {
        skip_decl_whitespace(raw, &mut cursor);
        cursor < raw.len()
    } {
        let (name, value) = parse_decl_attribute(raw, &mut cursor)?;
        let valid_order = match (previous, name) {
            (b"version", b"encoding") => value.eq_ignore_ascii_case(b"utf-8"),
            (b"version" | b"encoding", b"standalone") => matches!(value, b"yes" | b"no"),
            _ => false,
        };
        if !valid_order {
            return Err(invalid(
                "ink actions XML declaration has an invalid or duplicate attribute",
            ));
        }
        previous = name;
    }
    Ok(())
}

fn decl_whitespace(value: Option<u8>) -> bool {
    matches!(value, Some(b' ' | b'\t' | b'\r' | b'\n'))
}

fn skip_decl_whitespace(raw: &[u8], cursor: &mut usize) {
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|byte| decl_whitespace(Some(byte)))
    {
        *cursor += 1;
    }
}

fn consume_decl_token(raw: &[u8], cursor: &mut usize, token: &[u8]) -> bool {
    raw.get(*cursor..)
        .is_some_and(|remaining| remaining.starts_with(token))
        .then(|| *cursor += token.len())
        .is_some()
}

fn parse_decl_attribute<'a>(raw: &'a [u8], cursor: &mut usize) -> Result<(&'a [u8], &'a [u8])> {
    let name_start = *cursor;
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|byte| byte.is_ascii_alphabetic())
    {
        *cursor += 1;
    }
    if *cursor == name_start {
        return Err(invalid(
            "ink actions XML declaration attribute name is missing",
        ));
    }
    let name = &raw[name_start..*cursor];
    skip_decl_whitespace(raw, cursor);
    if raw.get(*cursor) != Some(&b'=') {
        return Err(invalid(
            "ink actions XML declaration attribute equals sign is missing",
        ));
    }
    *cursor += 1;
    skip_decl_whitespace(raw, cursor);
    let quote = raw
        .get(*cursor)
        .copied()
        .filter(|value| matches!(value, b'\'' | b'"'))
        .ok_or_else(|| invalid("ink actions XML declaration attribute quote is missing"))?;
    *cursor += 1;
    let value_start = *cursor;
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|value| value != quote)
    {
        *cursor += 1;
    }
    if raw.get(*cursor) != Some(&quote) {
        return Err(invalid(
            "ink actions XML declaration attribute is unterminated",
        ));
    }
    let value = &raw[value_start..*cursor];
    *cursor += 1;
    Ok((name, value))
}

fn validate_xml_characters(value: &str, what: &str) -> Result<()> {
    if super::xml_characters::valid(value) {
        Ok(())
    } else {
        Err(invalid(format!("{what} contains an invalid XML character")))
    }
}

fn is_xml10_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

fn validate_text<'a>(value: &'a [u8], what: &str) -> Result<&'a str> {
    let value = std::str::from_utf8(value).map_err(xml_error)?;
    validate_xml_characters(value, what)?;
    Ok(value)
}

fn is_xml_whitespace(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

fn validate_reference(reference: &BytesRef<'_>) -> Result<()> {
    let value = std::str::from_utf8(reference.as_ref()).map_err(xml_error)?;
    match value {
        "amp" | "lt" | "gt" | "apos" | "quot" => Ok(()),
        value if value.strip_prefix("#x").is_some() => {
            let digits = value.strip_prefix("#x").unwrap_or_default();
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(invalid(
                    "ink actions hexadecimal character reference is invalid",
                ));
            }
            let codepoint = u32::from_str_radix(digits, 16)
                .map_err(|_| invalid("ink actions hexadecimal character reference is invalid"))?;
            let character = char::from_u32(codepoint)
                .ok_or_else(|| invalid("ink actions character reference is invalid"))?;
            if is_xml10_character(character) {
                Ok(())
            } else {
                Err(invalid(
                    "ink actions character reference is not an XML 1.0 character",
                ))
            }
        },
        value if value.strip_prefix('#').is_some() => {
            let digits = value.strip_prefix('#').unwrap_or_default();
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid(
                    "ink actions decimal character reference is invalid",
                ));
            }
            let codepoint = digits
                .parse::<u32>()
                .map_err(|_| invalid("ink actions decimal character reference is invalid"))?;
            let character = char::from_u32(codepoint)
                .ok_or_else(|| invalid("ink actions character reference is invalid"))?;
            if is_xml10_character(character) {
                Ok(())
            } else {
                Err(invalid(
                    "ink actions character reference is not an XML 1.0 character",
                ))
            }
        },
        _ => Err(invalid(
            "ink actions general entity references are not supported",
        )),
    }
}

fn has_namespace_declaration(element: &BytesStart<'_>, prefix: &[u8]) -> bool {
    element.attributes().any(|attribute| {
        let Ok(attribute) = attribute else {
            return false;
        };
        let key = attribute.key.as_ref();
        (prefix.is_empty() && key == b"xmlns") || (key.strip_prefix(b"xmlns:") == Some(prefix))
    })
}

fn expanded_attribute_names_equal<R: std::io::BufRead>(
    left: QName<'_>,
    right: QName<'_>,
    reader: &NsReader<R>,
) -> Result<bool> {
    let (left_namespace, left_local) = reader.resolver().resolve_attribute(left);
    let (right_namespace, right_local) = reader.resolver().resolve_attribute(right);
    if left_local != right_local {
        return Ok(false);
    }
    match (left_namespace, right_namespace) {
        (ResolveResult::Unbound, ResolveResult::Unbound) => Ok(true),
        (ResolveResult::Bound(Namespace(left)), ResolveResult::Bound(Namespace(right))) => {
            if left == right {
                return Ok(true);
            }
            Ok(decoded_namespace(left, reader)? == decoded_namespace(right, reader)?)
        },
        _ => Ok(false),
    }
}

fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_| invalid(format!("{resource} allocation failed")))
}

fn copy_source(xml: &[u8]) -> Result<Arc<[u8]>> {
    let mut source = Vec::new();
    source
        .try_reserve_exact(xml.len())
        .map_err(|_| invalid("ink actions source allocation failed"))?;
    source.extend_from_slice(xml);
    Ok(Arc::from(source.into_boxed_slice()))
}
fn pos<R: std::io::BufRead>(reader: &NsReader<R>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("ink actions offset exceeds usize"))
}
fn increment_nodes(nodes: &mut usize) -> Result<()> {
    let next = nodes
        .checked_add(1)
        .ok_or_else(|| limit("ink actions XML nodes", MAX_NODES))?;
    enforce_nodes(next)?;
    *nodes = next;
    Ok(())
}
fn enforce_nodes(nodes: usize) -> Result<()> {
    if nodes > MAX_NODES {
        Err(limit("ink actions XML nodes", MAX_NODES))
    } else {
        Ok(())
    }
}
fn enforce_depth(depth: usize) -> Result<()> {
    if depth > MAX_DEPTH {
        Err(limit("ink actions XML depth", MAX_DEPTH))
    } else {
        Ok(())
    }
}
fn increment_groups(groups: &mut usize) -> Result<()> {
    let next = groups
        .checked_add(1)
        .ok_or_else(|| limit("ink action groups", MAX_ACTION_GROUPS))?;
    if next > MAX_ACTION_GROUPS {
        return Err(limit("ink action groups", MAX_ACTION_GROUPS));
    }
    *groups = next;
    Ok(())
}
fn ensure_action_capacity(actions: &[Action]) -> Result<()> {
    if actions.len() >= MAX_ACTIONS {
        Err(limit("ink actions", MAX_ACTIONS))
    } else {
        Ok(())
    }
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}
fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
