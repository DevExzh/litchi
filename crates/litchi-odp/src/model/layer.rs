//! Typed, inert OpenDocument drawing-layer declarations.
//!
//! The declaration model is deliberately separate from shape rendering.  A
//! layer is a named grouping owner; `draw:display` and `draw:protected` are
//! retained as metadata and no visibility, protection, or rendering behavior
//! is executed by this crate.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    events::{BytesStart, Event},
    name::{Namespace, PrefixDeclaration, ResolveResult},
    reader::NsReader,
};

pub(crate) const DRAW_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(crate) const DR3D_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
pub(crate) const ANIM_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
pub(crate) const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(crate) const PRESENTATION_NAMESPACE: &str =
    "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
pub(crate) const STYLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(crate) const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

const MAX_XML_BYTES: usize = 128 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;
const MAX_DEPTH: usize = 512;
const MAX_LAYERS: usize = 65_536;
const MAX_LAYER_SETS: usize = 65_536;
const MAX_NAME_BYTES: usize = 4 * 1024;
const MAX_VALUE_BYTES: usize = 1024 * 1024;

/// The source owner of one `draw:layer-set` declaration.
#[derive(Clone, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum LayerOwner {
    /// A layer set in `office:master-styles` in `styles.xml`.
    MasterStyles,
    /// A layer set directly owned by a presentation page in `content.xml`.
    Page(usize),
    /// A layer set directly owned by a named `style:master-page` in `styles.xml`.
    MasterPage(String),
}

impl LayerOwner {
    /// Select the global layer declarations.
    #[must_use]
    pub const fn master_styles() -> Self {
        Self::MasterStyles
    }

    /// Select a page-local layer declaration set by checked page position.
    #[must_use]
    pub const fn page(index: usize) -> Self {
        Self::Page(index)
    }

    /// Select a master-page-local layer declaration set by exact name.
    ///
    /// The name is checked when it is resolved against a source snapshot.
    #[must_use]
    pub fn master_page(name: impl Into<String>) -> Self {
        Self::MasterPage(name.into())
    }
}

/// One typed `draw:layer` declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Layer {
    name: String,
    display: Option<String>,
    protected: Option<bool>,
}

impl Layer {
    /// Create a detached layer with the required `draw:name`.
    ///
    /// # Errors
    /// Returns an error for an empty, invalid, or oversized name.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let layer = Self {
            name: name.into(),
            display: None,
            protected: None,
        };
        validate_layer(&layer)?;
        Ok(layer)
    }

    /// Set the lexical ODF display policy.  The value is inert and is not
    /// interpreted by the presentation reader.
    ///
    /// # Errors
    /// Returns an error for XML controls or oversized values.
    pub fn with_display(mut self, display: impl Into<String>) -> Result<Self> {
        let display = display.into();
        validate_display(&display)?;
        self.display = Some(display);
        Ok(self)
    }

    /// Set the optional inert protection flag.
    #[must_use]
    pub const fn with_protected(mut self, protected: bool) -> Self {
        self.protected = Some(protected);
        self
    }

    /// Return the required layer name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the source lexical `draw:display` value, if present.
    #[must_use]
    pub fn display(&self) -> Option<&str> {
        self.display.as_deref()
    }

    /// Return the optional source `draw:protected` value.
    #[must_use]
    pub const fn protected(&self) -> Option<bool> {
        self.protected
    }

    pub(crate) fn parsed(
        name: String,
        display: Option<String>,
        protected: Option<bool>,
    ) -> Result<Self> {
        let layer = Self {
            name,
            display,
            protected,
        };
        validate_layer(&layer)?;
        Ok(layer)
    }

    pub(crate) fn to_xml(&self, prefix: &str, declare_draw: bool) -> Result<String> {
        validate_layer(self)?;
        let qualified = |local: &str| {
            if prefix.is_empty() {
                local.to_owned()
            } else {
                format!("{prefix}:{local}")
            }
        };
        let mut output = String::with_capacity(128 + self.name.len());
        output.push('<');
        output.push_str(&qualified("layer"));
        if declare_draw {
            output.push_str(" xmlns:");
            output.push_str(if prefix.is_empty() { "draw" } else { prefix });
            output.push_str("=\"");
            output.push_str(DRAW_NAMESPACE);
            output.push('"');
        }
        output.push(' ');
        output.push_str(&qualified("name"));
        output.push_str("=\"");
        output.push_str(&escape_xml(&self.name));
        output.push('"');
        if let Some(display) = &self.display {
            output.push(' ');
            output.push_str(&qualified("display"));
            output.push_str("=\"");
            output.push_str(&escape_xml(display));
            output.push('"');
        }
        if let Some(protected) = self.protected {
            output.push(' ');
            output.push_str(&qualified("protected"));
            output.push_str("=\"");
            output.push_str(if protected { "true" } else { "false" });
            output.push('"');
        }
        output.push_str("/>");
        Ok(output)
    }
}

/// One source layer set and its declarations in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LayerSet {
    pub(crate) owner: LayerOwner,
    pub(crate) layers: Vec<Layer>,
}

impl LayerSet {
    /// Return the source owner selector.
    #[must_use]
    pub fn owner(&self) -> &LayerOwner {
        &self.owner
    }

    /// Return declarations in source order.
    #[must_use]
    pub fn layers(&self) -> &[Layer] {
        &self.layers
    }

    /// Select one uniquely named layer within this set.
    ///
    /// # Errors
    /// Returns an error when a malformed source contains duplicate names.
    pub fn get(&self, name: &str) -> Result<Option<&Layer>> {
        let mut found = None;
        for layer in &self.layers {
            if layer.name == name {
                if found.is_some() {
                    return invalid("ODP layer name is ambiguous within its layer set");
                }
                found = Some(layer);
            }
        }
        Ok(found)
    }
}

/// Lazy typed inventory of all ODP drawing-layer declaration sets.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct LayerInventory {
    pub(crate) sets: Vec<LayerSet>,
}

impl LayerInventory {
    /// Return all explicit layer sets in package document order.
    #[must_use]
    pub fn sets(&self) -> &[LayerSet] {
        &self.sets
    }

    /// Return the first global layer set, if one is declared.
    #[must_use]
    pub fn master_styles(&self) -> Option<&LayerSet> {
        self.sets
            .iter()
            .find(|set| set.owner == LayerOwner::MasterStyles)
    }

    /// Return a page-local layer set, if one is declared.
    #[must_use]
    pub fn page(&self, index: usize) -> Option<&LayerSet> {
        self.sets
            .iter()
            .find(|set| set.owner == LayerOwner::Page(index))
    }

    /// Return a named master-page-local layer set, if one is declared.
    #[must_use]
    pub fn master_page(&self, name: &str) -> Option<&LayerSet> {
        self.sets.iter().find(
            |set| matches!(&set.owner, LayerOwner::MasterPage(candidate) if candidate == name),
        )
    }

    /// Return whether no explicit layer set is present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.sets.is_empty()
    }
}

#[derive(Clone)]
pub(crate) struct LayerSetLocation {
    pub(crate) owner: LayerOwner,
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) empty: bool,
    pub(crate) end_tag_start: Option<usize>,
    pub(crate) name_prefix: String,
    pub(crate) draw_prefix: Option<String>,
    pub(crate) layers: Vec<LayerLocation>,
}

#[derive(Clone)]
pub(crate) struct LayerLocation {
    pub(crate) layer: Layer,
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) empty: bool,
    pub(crate) name_value: (usize, usize),
}

#[derive(Clone)]
pub(crate) struct ParentLocation {
    pub(crate) start: usize,
    pub(crate) layer_insert_at: usize,
    pub(crate) end: usize,
    pub(crate) empty: bool,
    pub(crate) qualified_name: Vec<u8>,
    pub(crate) draw_prefix: Option<String>,
}

#[derive(Clone)]
pub(crate) struct LayerReference {
    pub(crate) owner: LayerOwner,
    pub(crate) value: String,
    pub(crate) range: (usize, usize),
}

#[derive(Clone)]
pub(crate) struct ParsedLayers {
    pub(crate) sets: Vec<LayerSetLocation>,
    pub(crate) references: Vec<LayerReference>,
    pub(crate) parents: Vec<(LayerOwner, ParentLocation)>,
}

#[derive(Clone)]
struct Frame {
    namespace: String,
    local: Vec<u8>,
    start: usize,
    owner: Option<LayerOwner>,
    has_local_layer_set: bool,
    saw_layer_order_tail: bool,
    master_name: Option<String>,
    draw_prefix: Option<String>,
    qualified_name: Vec<u8>,
}

/// Parse all layer declarations in one ODP XML part and return source spans.
///
/// This is an internal source scanner.  It retains only small ranges and typed
/// values; the source XML remains owned by the immutable package snapshot.
pub(crate) fn scan(xml: &str, path: &str) -> Result<ParsedLayers> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("ODP layer XML exceeds the 128 MiB limit");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut frames = Vec::<Frame>::new();
    let mut sets = Vec::new();
    let mut refs = Vec::new();
    let mut parents = Vec::new();
    let mut page_count = 0usize;
    let mut event_count = 0usize;
    let mut layer_count = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prologue_content_seen = false;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| error("ODP layer XML position overflow"))?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|cause| error(format!("invalid ODP layer XML in {path}: {cause}")))?;
        let event_end = xml_tag_end(xml.as_bytes(), event_start).unwrap_or(event_start);
        if event_count >= MAX_EVENTS {
            return invalid("ODP layer XML exceeds the event limit");
        }
        event_count += 1;
        match event {
            Event::Start(element) => {
                let depth = frames
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| error("ODP layer XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid("ODP layer XML nesting exceeds 512 levels");
                }
                let (namespace, local) = resolved_name(&resolved, element.local_name().as_ref())?;
                if frames.is_empty() && root_closed {
                    return invalid("ODP layer XML has more than one root element");
                }
                root_seen = true;
                if let Some(parent) = frames.last_mut() {
                    note_layer_order_child(parent, &namespace, &local)?;
                }
                let parent = frames.last();
                let draw_prefix = if namespace == DRAW_NAMESPACE {
                    Some(qualified_prefix(element.name().as_ref()))
                } else {
                    current_draw_prefix(&reader)
                };
                let master_name = if namespace == STYLE_NAMESPACE && local == b"master-page" {
                    Some(master_page_name(&reader, &element)?)
                } else {
                    parent.and_then(|frame| frame.master_name.clone())
                };
                let page_index = if namespace == DRAW_NAMESPACE && local == b"page" {
                    let index = page_count;
                    page_count = page_count
                        .checked_add(1)
                        .ok_or_else(|| error("ODP layer page count overflow"))?;
                    Some(index)
                } else {
                    parent.and_then(|frame| {
                        frame.owner.as_ref().and_then(|owner| match owner {
                            LayerOwner::Page(index) => Some(*index),
                            _ => None,
                        })
                    })
                };
                let inherited_owner = parent.and_then(|frame| frame.owner.clone());
                let owner = if namespace == DRAW_NAMESPACE && local == b"page" {
                    page_index.map(LayerOwner::Page)
                } else if namespace == OFFICE_NAMESPACE && local == b"master-styles" {
                    Some(LayerOwner::MasterStyles)
                } else if namespace == STYLE_NAMESPACE && local == b"master-page" {
                    master_name.clone().map(LayerOwner::MasterPage)
                } else {
                    inherited_owner
                };
                let mut frame = Frame {
                    namespace: namespace.clone(),
                    local: local.clone(),
                    start: event_start,
                    owner: owner.clone(),
                    has_local_layer_set: false,
                    saw_layer_order_tail: false,
                    master_name: master_name.clone(),
                    draw_prefix,
                    qualified_name: element.name().as_ref().to_vec(),
                };
                if let Some(owner) = owner_boundary(&namespace, &local, frame.owner.as_ref()) {
                    parents.push((
                        owner,
                        ParentLocation {
                            start: event_start,
                            layer_insert_at: event_end,
                            end: 0,
                            empty: false,
                            qualified_name: frame.qualified_name.clone(),
                            draw_prefix: frame.draw_prefix.clone(),
                        },
                    ));
                }
                if namespace == DRAW_NAMESPACE
                    && local == b"layer-set"
                    && sets.len() >= MAX_LAYER_SETS
                {
                    return invalid("ODP layer sets exceed the configured limit");
                }
                if namespace == DRAW_NAMESPACE && local == b"layer-set" {
                    let Some(parent_frame) = parent else {
                        return invalid("ODP draw:layer-set has no owner");
                    };
                    let set_owner = layer_set_owner(parent_frame)?;
                    if !matches!(
                        parent_frame.local.as_slice(),
                        b"page" | b"master-styles" | b"master-page"
                    ) {
                        return invalid("ODP draw:layer-set is outside an allowed owner");
                    }
                    if sets
                        .iter()
                        .any(|set: &LayerSetLocation| set.owner == set_owner)
                    {
                        return invalid("ODP layer owner contains duplicate layer sets");
                    }
                    let name_prefix = qualified_prefix(element.name().as_ref());
                    let location = LayerSetLocation {
                        owner: set_owner.clone(),
                        start: event_start,
                        end: event_end,
                        empty: false,
                        end_tag_start: None,
                        name_prefix,
                        draw_prefix: parent_frame
                            .draw_prefix
                            .clone()
                            .or_else(|| Some("draw".to_owned())),
                        layers: Vec::new(),
                    };
                    sets.push(location);
                    if let Some(parent_frame) = frames.last_mut() {
                        parent_frame.has_local_layer_set = true;
                    }
                    frame.owner = Some(set_owner);
                } else if namespace == DRAW_NAMESPACE && local == b"layer" {
                    let Some(set_frame) = frames.last() else {
                        return invalid("ODP draw:layer has no layer-set owner");
                    };
                    if set_frame.namespace != DRAW_NAMESPACE || set_frame.local != b"layer-set" {
                        return invalid("ODP draw:layer is outside draw:layer-set");
                    }
                    let set = sets
                        .last_mut()
                        .ok_or_else(|| error("ODP layer-set scanner state is missing"))?;
                    if set.layers.len() >= MAX_LAYERS || layer_count >= MAX_LAYERS {
                        return invalid("ODP layer declarations exceed the configured limit");
                    }
                    let name = attr_value(&reader, &element, DRAW_NAMESPACE, b"name")?
                        .ok_or_else(|| error("ODP draw:layer requires draw:name"))?;
                    let display = attr_value(&reader, &element, DRAW_NAMESPACE, b"display")?;
                    let protected = attr_value(&reader, &element, DRAW_NAMESPACE, b"protected")?
                        .map(|value| parse_bool(&value))
                        .transpose()?;
                    let layer = Layer::parsed(name, display, protected)?;
                    if set
                        .layers
                        .iter()
                        .any(|candidate| candidate.layer.name == layer.name)
                    {
                        return invalid("ODP layer-set contains duplicate draw:name values");
                    }
                    let name_key = attr_raw_key(&reader, &element, DRAW_NAMESPACE, b"name")?
                        .ok_or_else(|| error("ODP draw:layer name attribute is missing"))?;
                    let name_value =
                        attr_raw_value_range(&xml.as_bytes()[event_start..event_end], &name_key)
                            .ok_or_else(|| error("ODP draw:layer name source span is missing"))?;
                    set.layers.push(LayerLocation {
                        layer,
                        start: event_start,
                        end: event_end,
                        empty: true,
                        name_value: (event_start + name_value.0, event_start + name_value.1),
                    });
                    layer_count += 1;
                    set.end = event_end;
                } else {
                    if let Some(key) = attr_raw_key(&reader, &element, DRAW_NAMESPACE, b"layer")? {
                        if let Some(value) =
                            attr_value(&reader, &element, DRAW_NAMESPACE, b"layer")?
                        {
                            if let Some(owner) =
                                effective_owner(&frames, &sets, page_index, master_name.as_deref())
                            {
                                if let Some(range) = attr_raw_value_range(
                                    &xml.as_bytes()[event_start..event_end],
                                    &key,
                                ) {
                                    refs.push(LayerReference {
                                        owner,
                                        value,
                                        range: (event_start + range.0, event_start + range.1),
                                    });
                                }
                            }
                        }
                    }
                }
                frames.push(frame);
            },
            Event::Empty(element) => {
                let depth = frames
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| error("ODP layer XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid("ODP layer XML nesting exceeds 512 levels");
                }
                let (namespace, local) = resolved_name(&resolved, element.local_name().as_ref())?;
                let root_empty = frames.is_empty();
                if root_empty && root_closed {
                    return invalid("ODP layer XML has more than one root element");
                }
                root_seen = true;
                if let Some(parent) = frames.last_mut() {
                    note_layer_order_child(parent, &namespace, &local)?;
                }
                let parent = frames.last();
                if let Some(parent) = parent {
                    note_layer_prefix_child(&mut parents, parent, &namespace, &local, event_end);
                }
                let draw_prefix = if namespace == DRAW_NAMESPACE {
                    Some(qualified_prefix(element.name().as_ref()))
                } else {
                    current_draw_prefix(&reader)
                };
                let master_name = if namespace == STYLE_NAMESPACE && local == b"master-page" {
                    Some(master_page_name(&reader, &element)?)
                } else {
                    parent.and_then(|frame| frame.master_name.clone())
                };
                let page_index = if namespace == DRAW_NAMESPACE && local == b"page" {
                    let index = page_count;
                    page_count = page_count
                        .checked_add(1)
                        .ok_or_else(|| error("ODP layer page count overflow"))?;
                    Some(index)
                } else {
                    parent.and_then(|frame| {
                        frame.owner.as_ref().and_then(|owner| match owner {
                            LayerOwner::Page(index) => Some(*index),
                            _ => None,
                        })
                    })
                };
                let empty_owner = if namespace == DRAW_NAMESPACE && local == b"page" {
                    page_index.map(LayerOwner::Page)
                } else if namespace == OFFICE_NAMESPACE && local == b"master-styles" {
                    Some(LayerOwner::MasterStyles)
                } else if namespace == STYLE_NAMESPACE && local == b"master-page" {
                    master_name.clone().map(LayerOwner::MasterPage)
                } else {
                    None
                };
                if let Some(owner) = empty_owner {
                    parents.push((
                        owner,
                        ParentLocation {
                            start: event_start,
                            layer_insert_at: event_end,
                            end: event_end,
                            empty: true,
                            qualified_name: element.name().as_ref().to_vec(),
                            draw_prefix,
                        },
                    ));
                }
                if namespace == DRAW_NAMESPACE
                    && local == b"layer-set"
                    && sets.len() >= MAX_LAYER_SETS
                {
                    return invalid("ODP layer sets exceed the configured limit");
                }
                if namespace == DRAW_NAMESPACE && local == b"layer-set" {
                    let Some(parent_frame) = parent else {
                        return invalid("ODP draw:layer-set has no owner");
                    };
                    let set_owner = layer_set_owner(parent_frame)?;
                    if !matches!(
                        parent_frame.local.as_slice(),
                        b"page" | b"master-styles" | b"master-page"
                    ) {
                        return invalid("ODP draw:layer-set is outside an allowed owner");
                    }
                    if sets
                        .iter()
                        .any(|set: &LayerSetLocation| set.owner == set_owner)
                    {
                        return invalid("ODP layer owner contains duplicate layer sets");
                    }
                    let name_prefix = qualified_prefix(element.name().as_ref());
                    let location = LayerSetLocation {
                        owner: set_owner,
                        start: event_start,
                        end: event_end,
                        empty: true,
                        end_tag_start: None,
                        name_prefix,
                        draw_prefix: parent_frame
                            .draw_prefix
                            .clone()
                            .or_else(|| Some("draw".to_owned())),
                        layers: Vec::new(),
                    };
                    let name_prefix = location.name_prefix.clone();
                    let _ = name_prefix;
                    sets.push(location);
                    if let Some(parent_frame) = frames.last_mut() {
                        parent_frame.has_local_layer_set = true;
                    }
                } else if namespace == DRAW_NAMESPACE && local == b"layer" {
                    let Some(set) = sets.last_mut() else {
                        return invalid("ODP draw:layer has no layer-set owner");
                    };
                    if set.layers.len() >= MAX_LAYERS || layer_count >= MAX_LAYERS {
                        return invalid("ODP layer declarations exceed the configured limit");
                    }
                    let name = attr_value(&reader, &element, DRAW_NAMESPACE, b"name")?
                        .ok_or_else(|| error("ODP draw:layer requires draw:name"))?;
                    let display = attr_value(&reader, &element, DRAW_NAMESPACE, b"display")?;
                    let protected = attr_value(&reader, &element, DRAW_NAMESPACE, b"protected")?
                        .map(|value| parse_bool(&value))
                        .transpose()?;
                    let layer = Layer::parsed(name, display, protected)?;
                    if set
                        .layers
                        .iter()
                        .any(|candidate| candidate.layer.name == layer.name)
                    {
                        return invalid("ODP layer-set contains duplicate draw:name values");
                    }
                    let name_key = attr_raw_key(&reader, &element, DRAW_NAMESPACE, b"name")?
                        .ok_or_else(|| error("ODP draw:layer name attribute is missing"))?;
                    let name_value =
                        attr_raw_value_range(&xml.as_bytes()[event_start..event_end], &name_key)
                            .ok_or_else(|| error("ODP draw:layer name source span is missing"))?;
                    set.layers.push(LayerLocation {
                        layer,
                        start: event_start,
                        end: event_end,
                        empty: true,
                        name_value: (event_start + name_value.0, event_start + name_value.1),
                    });
                    layer_count += 1;
                    set.end = event_end;
                } else {
                    if let Some(key) = attr_raw_key(&reader, &element, DRAW_NAMESPACE, b"layer")?
                        && let Some(value) =
                            attr_value(&reader, &element, DRAW_NAMESPACE, b"layer")?
                        && let Some(owner) =
                            effective_owner(&frames, &sets, page_index, master_name.as_deref())
                        && let Some(range) =
                            attr_raw_value_range(&xml.as_bytes()[event_start..event_end], &key)
                    {
                        refs.push(LayerReference {
                            owner,
                            value,
                            range: (event_start + range.0, event_start + range.1),
                        });
                    }
                }
                if root_empty {
                    root_closed = true;
                }
            },
            Event::End(element) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| error("ODP layer XML has an unexpected end"))?;
                let (namespace, local) = resolved_name(&resolved, element.local_name().as_ref())?;
                if frame.namespace != namespace || frame.local != local {
                    return invalid("ODP layer XML end element does not match its start");
                }
                if let Some(parent) = frames.last() {
                    note_layer_prefix_child(
                        &mut parents,
                        parent,
                        &frame.namespace,
                        &frame.local,
                        event_end,
                    );
                }
                if namespace == DRAW_NAMESPACE && local == b"layer" {
                    let set = sets
                        .last_mut()
                        .ok_or_else(|| error("ODP layer scanner lost layer-set"))?;
                    let location = set
                        .layers
                        .last_mut()
                        .ok_or_else(|| error("ODP layer scanner lost layer"))?;
                    location.end = event_end;
                    location.empty = false;
                    set.end = event_end;
                    set.end_tag_start = Some(event_start);
                } else if namespace == DRAW_NAMESPACE && local == b"layer-set" {
                    let set = sets
                        .last_mut()
                        .ok_or_else(|| error("ODP layer scanner lost layer-set"))?;
                    set.end = event_end;
                    set.end_tag_start = Some(event_start);
                }
                if let Some((_, parent)) = parents
                    .iter_mut()
                    .rev()
                    .find(|(_, parent)| parent.start == frame.start)
                {
                    parent.end = event_end;
                }
                if frames.is_empty() {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::DocType(_) => return invalid("ODP layer XML doctype is prohibited"),
            Event::Comment(_) | Event::PI(_) => {
                if !root_seen {
                    prologue_content_seen = true;
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if frames.is_empty() => {
                return invalid("ODP layer XML has non-whitespace content outside its root");
            },
            Event::CData(_) | Event::GeneralRef(_) => {},
            Event::Text(text) if frames.is_empty() => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return invalid("ODP layer XML has non-whitespace content outside its root");
                }
                if !root_seen {
                    prologue_content_seen = true;
                }
            },
            Event::Text(_) => {},
            Event::Decl(_) => {
                if declaration_seen || root_seen || prologue_content_seen {
                    return invalid("ODP XML declaration is outside the document prologue");
                }
                declaration_seen = true;
            },
        }
        buffer.clear();
    }
    if !root_seen {
        return invalid("ODP layer XML has no root element");
    }
    if !frames.is_empty() {
        return invalid("ODP layer XML has unterminated elements");
    }
    // A page or master may declare its layer-set after a shape that refers to
    // it. Resolve references only after the complete part reveals whether the
    // local owner exists; otherwise the reference uses global master-styles.
    for reference in &mut refs {
        match &reference.owner {
            LayerOwner::Page(index)
                if !sets.iter().any(|set| set.owner == LayerOwner::Page(*index)) =>
            {
                reference.owner = LayerOwner::MasterStyles;
            },
            LayerOwner::MasterPage(name)
                if !sets
                    .iter()
                    .any(|set| set.owner == LayerOwner::MasterPage(name.clone())) =>
            {
                reference.owner = LayerOwner::MasterStyles;
            },
            _ => {},
        }
    }
    if sets.len() > MAX_LAYER_SETS
        || sets.iter().map(|set| set.layers.len()).sum::<usize>() > MAX_LAYERS
    {
        return invalid("ODP layer inventory exceeds configured limits");
    }
    Ok(ParsedLayers {
        sets,
        references: refs,
        parents,
    })
}

/// Parse a presentation's content and styles layer declarations lazily.
pub(crate) fn inventory(content: &str, styles: Option<&str>) -> Result<LayerInventory> {
    let mut parsed = scan(content, "content.xml")?;
    if let Some(styles) = styles {
        let styles = scan(styles, "styles.xml")?;
        if parsed.sets.iter().any(|candidate| {
            styles
                .sets
                .iter()
                .any(|other| candidate.owner == other.owner)
        }) {
            return invalid("ODP package contains duplicate layer-set owners");
        }
        parsed.sets.extend(styles.sets);
    }
    let sets = parsed
        .sets
        .into_iter()
        .map(|set| LayerSet {
            owner: set.owner,
            layers: set.layers.into_iter().map(|layer| layer.layer).collect(),
        })
        .collect();
    Ok(LayerInventory { sets })
}

fn layer_set_owner(frame: &Frame) -> Result<LayerOwner> {
    match frame.namespace.as_str() {
        DRAW_NAMESPACE if frame.local == b"page" => frame
            .owner
            .clone()
            .ok_or_else(|| error("ODP page layer owner is missing")),
        OFFICE_NAMESPACE if frame.local == b"master-styles" => Ok(LayerOwner::MasterStyles),
        STYLE_NAMESPACE if frame.local == b"master-page" => frame
            .master_name
            .clone()
            .map(LayerOwner::MasterPage)
            .ok_or_else(|| error("ODP master page layer owner has no name")),
        _ => invalid("ODP layer-set has an invalid parent"),
    }
}

fn effective_owner(
    frames: &[Frame],
    _sets: &[LayerSetLocation],
    page_index: Option<usize>,
    master_name: Option<&str>,
) -> Option<LayerOwner> {
    if let Some(page) = page_index {
        return Some(LayerOwner::Page(page));
    }
    if let Some(name) = master_name {
        return Some(LayerOwner::MasterPage(name.to_owned()));
    }
    frames.last().and_then(|frame| frame.owner.clone())
}

fn master_page_name(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    let name = attr_value(reader, element, STYLE_NAMESPACE, b"name")?.ok_or_else(|| {
        error("ODP style:master-page requires a style:name NCName owner identity")
    })?;
    if name.len() > MAX_NAME_BYTES {
        return invalid("ODP style:master-page style:name exceeds the configured limit");
    }
    validate_odf_ncname(&name, "ODP style:master-page style:name")?;
    Ok(name)
}

fn validate_odf_ncname(value: &str, description: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return invalid(format!("{description} cannot be empty"));
    };
    if !is_ncname_start(first) || !chars.all(is_ncname_char) {
        return invalid(format!("invalid {description} '{value}'"));
    }
    Ok(())
}

// XML 1.0 NameStartChar/NameChar ranges, with ':' removed for NCName.
fn is_ncname_start(character: char) -> bool {
    let code = character as u32;
    matches!(
        code,
        0x0041..=0x005A
            | 0x005F
            | 0x0061..=0x007A
            | 0x00C0..=0x00D6
            | 0x00D8..=0x00F6
            | 0x00F8..=0x02FF
            | 0x0370..=0x037D
            | 0x037F..=0x1FFF
            | 0x200C..=0x200D
            | 0x2070..=0x218F
            | 0x2C00..=0x2FEF
            | 0x3001..=0xD7FF
            | 0xF900..=0xFDCF
            | 0xFDF0..=0xFFFD
            | 0x10000..=0xEFFFF
    )
}

fn is_ncname_char(character: char) -> bool {
    let code = character as u32;
    is_ncname_start(character)
        || matches!(code, 0x002D | 0x002E | 0x0030..=0x0039 | 0x00B7 | 0x0300..=0x036F | 0x203F..=0x2040)
}

fn resolved_name(resolved: &ResolveResult<'_>, local: &[u8]) -> Result<(String, Vec<u8>)> {
    let local = local.to_vec();
    match resolved {
        ResolveResult::Bound(Namespace(namespace)) => {
            let namespace = std::str::from_utf8(namespace)
                .map_err(|_| error("ODP layer XML namespace is not UTF-8"))?
                .to_owned();
            Ok((namespace, local))
        },
        ResolveResult::Unbound => Ok((String::new(), local)),
        ResolveResult::Unknown(_) => {
            invalid("ODP layer XML uses an unresolved element namespace prefix")
        },
    }
}

fn note_layer_order_child(
    parent: &mut Frame,
    child_namespace: &str,
    child_local: &[u8],
) -> Result<()> {
    let is_page = parent.namespace == DRAW_NAMESPACE && parent.local == b"page";
    let is_master_page = parent.namespace == STYLE_NAMESPACE && parent.local == b"master-page";
    if !is_page && !is_master_page {
        return Ok(());
    }
    if child_namespace == DRAW_NAMESPACE && child_local == b"layer-set" {
        if parent.saw_layer_order_tail {
            return invalid("ODP draw:layer-set is out of schema order for its owner");
        }
        return Ok(());
    }
    let prefix_child = if is_page {
        child_namespace == SVG_NAMESPACE && matches!(child_local, b"title" | b"desc")
    } else {
        child_namespace == STYLE_NAMESPACE
            && matches!(
                child_local,
                b"header"
                    | b"header-left"
                    | b"header-first"
                    | b"footer"
                    | b"footer-left"
                    | b"footer-first"
            )
    };
    if parent.has_local_layer_set && prefix_child {
        return invalid("ODP layer owner has a schema-prefix child after draw:layer-set");
    }
    if parent.has_local_layer_set {
        return Ok(());
    }
    if prefix_child {
        return Ok(());
    }
    // Namespace-bound ODF children are known schema content and therefore
    // establish the post-prefix portion of the owner. Foreign extensions,
    // including unqualified opaque children, remain source-preserving and do
    // not make a later declaration appear malformed.
    if matches!(
        child_namespace,
        ANIM_NAMESPACE
            | DR3D_NAMESPACE
            | DRAW_NAMESPACE
            | OFFICE_NAMESPACE
            | PRESENTATION_NAMESPACE
            | STYLE_NAMESPACE
            | SVG_NAMESPACE
    ) {
        parent.saw_layer_order_tail = true;
    }
    Ok(())
}

fn qualified_prefix(name: &[u8]) -> String {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return String::new();
    };
    String::from_utf8_lossy(&name[..separator]).into_owned()
}

fn attr_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &[u8],
) -> Result<Option<String>> {
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|cause| error(format!("invalid ODP layer attribute: {cause}")))?;
        let (resolved, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if local_name.as_ref() != local {
            continue;
        }
        let ResolveResult::Bound(Namespace(uri)) = resolved else {
            continue;
        };
        if uri != namespace.as_bytes() {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|cause| error(format!("invalid ODP layer attribute value: {cause}")))?;
        if value.len() > MAX_VALUE_BYTES {
            return invalid("ODP layer attribute value exceeds the configured limit");
        }
        return Ok(Some(value.into_owned()));
    }
    Ok(None)
}

fn attr_raw_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &[u8],
) -> Result<Option<Vec<u8>>> {
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|cause| error(format!("invalid ODP layer attribute: {cause}")))?;
        let (resolved, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if local_name.as_ref() != local {
            continue;
        }
        if let ResolveResult::Bound(Namespace(uri)) = resolved
            && uri == namespace.as_bytes()
        {
            return Ok(Some(attribute.key.as_ref().to_vec()));
        }
    }
    Ok(None)
}

fn current_draw_prefix(reader: &NsReader<&[u8]>) -> Option<String> {
    reader
        .resolver()
        .bindings()
        .find_map(|(prefix, namespace)| {
            if namespace != Namespace(DRAW_NAMESPACE.as_bytes()) {
                return None;
            }
            Some(match prefix {
                PrefixDeclaration::Default => String::new(),
                PrefixDeclaration::Named(prefix) => String::from_utf8_lossy(prefix).into_owned(),
            })
        })
}

fn owner_boundary(namespace: &str, local: &[u8], owner: Option<&LayerOwner>) -> Option<LayerOwner> {
    match (namespace, local) {
        (DRAW_NAMESPACE, b"page") => owner.cloned(),
        (OFFICE_NAMESPACE, b"master-styles") => Some(LayerOwner::MasterStyles),
        (STYLE_NAMESPACE, b"master-page") => owner.cloned(),
        _ => None,
    }
}

fn note_layer_prefix_child(
    parents: &mut [(LayerOwner, ParentLocation)],
    parent: &Frame,
    child_namespace: &str,
    child_local: &[u8],
    child_end: usize,
) {
    let allowed = match (parent.namespace.as_str(), parent.local.as_slice()) {
        (DRAW_NAMESPACE, b"page") => {
            child_namespace == SVG_NAMESPACE && matches!(child_local, b"title" | b"desc")
        },
        (STYLE_NAMESPACE, b"master-page") => {
            child_namespace == STYLE_NAMESPACE
                && matches!(
                    child_local,
                    b"header"
                        | b"header-left"
                        | b"header-first"
                        | b"footer"
                        | b"footer-left"
                        | b"footer-first"
                )
        },
        _ => false,
    };
    if !allowed {
        return;
    }
    if let Some((_, location)) = parents
        .iter_mut()
        .rev()
        .find(|(_, location)| location.start == parent.start)
    {
        location.layer_insert_at = child_end;
    }
}

fn attr_raw_value_range(tag: &[u8], key: &[u8]) -> Option<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && !matches!(tag[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len()
            || tag[cursor] == b'>'
            || (tag[cursor] == b'/' && tag.get(cursor + 1) == Some(&b'>'))
        {
            return None;
        }
        let key_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let actual_key = tag.get(key_start..cursor)?;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return None;
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag.get(cursor)?;
        if quote != b'\'' && quote != b'"' {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return None;
        }
        if actual_key == key {
            return Some((value_start, cursor));
        }
        cursor += 1;
    }
    None
}

fn xml_tag_end(bytes: &[u8], start: usize) -> Option<usize> {
    if bytes.get(start) != Some(&b'<') {
        return None;
    }
    let mut quote = None;
    for (offset, byte) in bytes.get(start + 1..)?.iter().enumerate() {
        match quote {
            Some(current) if *byte == current => quote = None,
            Some(_) => {},
            None if *byte == b'\'' || *byte == b'"' => quote = Some(*byte),
            None if *byte == b'>' => return Some(start + offset + 2),
            None => {},
        }
    }
    None
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid("ODP draw:protected must be true, false, 1, or 0"),
    }
}

fn validate_layer(layer: &Layer) -> Result<()> {
    validate_name(&layer.name, "ODP layer name")?;
    if let Some(display) = &layer.display {
        validate_display(display)?;
    }
    Ok(())
}

pub(crate) fn validate_name(name: &str, label: &str) -> Result<()> {
    if name.is_empty() || name.len() > MAX_NAME_BYTES || name.chars().any(char::is_control) {
        return invalid(format!(
            "{label} is empty, contains controls, or exceeds {MAX_NAME_BYTES} bytes"
        ));
    }
    Ok(())
}

fn validate_text(value: &str, label: &str) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES || value.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        return invalid(format!("{label} is invalid or exceeds the configured size limit"));
    }
    Ok(())
}

fn validate_display(value: &str) -> Result<()> {
    if !matches!(value, "always" | "screen" | "printer" | "none") {
        return invalid("ODP draw:display must be always, screen, printer, or none");
    }
    validate_text(value, "ODP layer display")
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "scanner boundary assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn scanner_rejects_missing_root() {
        assert!(scan("", "content.xml").is_err());
    }

    #[test]
    fn scanner_rejects_multiple_roots() {
        assert!(scan("<first/><second/>", "content.xml").is_err());
    }

    #[test]
    fn scanner_rejects_non_whitespace_outside_root() {
        assert!(scan("text<root/>", "content.xml").is_err());
        assert!(scan("<root/>text", "content.xml").is_err());
        assert!(scan("<![CDATA[text]]><root/>", "content.xml").is_err());
        assert!(scan("<root/>&amp;", "content.xml").is_err());
    }

    #[test]
    fn scanner_retains_xml_permitted_misc_around_root() {
        assert!(
            scan(
                r#"<?xml version="1.0"?><!--before--><root/><!--after--><?after?>"#,
                "content.xml"
            )
            .is_ok()
        );
        assert!(
            scan(
                r#"<!--before--><?xml version="1.0"?><root/>"#,
                "content.xml"
            )
            .is_err()
        );
    }

    #[test]
    fn ncname_uses_xml_code_point_ranges() {
        assert!(validate_odf_ncname("À.layer_1", "test").is_ok());
        assert!(validate_odf_ncname("µ", "test").is_err());
        assert!(validate_odf_ncname("ª", "test").is_err());
        assert!(validate_odf_ncname("º", "test").is_err());
        assert!(validate_odf_ncname("name·1", "test").is_ok());
    }
}
