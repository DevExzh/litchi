//! Inert typed ODF XForms model declarations.
//!
//! ODF embeds W3C XForms models below `office:forms`.  This module models the
//! declaration layer (`xforms:model`, `instance`, `bind`, and `submission`)
//! without implementing an XForms processor: XPath, expressions, actions,
//! submissions, network access, and instance evaluation remain strings.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::doc_markdown,
    clippy::format_push_string,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::missing_errors_doc,
    clippy::must_use_candidate,
    clippy::needless_pass_by_value,
    clippy::similar_names,
    reason = "the XForms model is a bounded inert XML declaration projection"
)]

use crate::namespace::{FORMNS, OFFICENS, XFORMSNS, XMLNS};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 512;
const MAX_MODELS: usize = 16_384;
const MAX_CHILDREN: usize = 65_536;
const MAX_ATTRIBUTES: usize = 256;
const MAX_STRING_BYTES: usize = 16 * 1024 * 1024;

/// A namespace-resolved attribute retained on a model or known child.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Attribute {
    /// Namespace URI, or the empty string for an unqualified attribute.
    pub namespace_uri: String,
    /// Local attribute name.
    pub local_name: String,
    /// Lexical value.
    pub value: String,
    /// Source prefix when the attribute was qualified.
    pub prefix: Option<String>,
}

/// One namespace binding retained from the source context of an XForms model.
/// The bindings are inert: they only provide the prefix context needed to
/// interpret XPath/CURIE strings or re-emit opaque XML fragments.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NamespaceBinding {
    /// Prefix, or `None` for the default namespace.
    pub prefix: Option<String>,
    /// Namespace URI bound to the prefix.
    pub namespace_uri: String,
}

/// An unknown direct model child retained as an inert XML fragment.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Extension {
    /// Namespace URI.
    pub namespace_uri: String,
    /// Local name.
    pub local_name: String,
    /// Original bounded XML fragment.
    pub xml: String,
}

/// A typed XForms `xforms:instance` declaration.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Instance {
    /// Optional XForms id.
    pub id: Option<String>,
    /// External instance source URI.
    pub src: Option<String>,
    /// Resource URI.
    pub resource: Option<String>,
    /// Media type.
    pub mediatype: Option<String>,
    /// Opaque inline instance XML, excluding the `xforms:instance` wrapper.
    pub content_xml: Option<String>,
    /// Unknown attributes retained by namespace and local name.
    pub attributes: Vec<Attribute>,
    /// Namespace declarations made on the instance wrapper.  This is kept
    /// private because it is serialization context, but it is required when
    /// an inline instance body uses a prefix declared only on that wrapper.
    namespace_declarations: Vec<String>,
}

/// A typed XForms `xforms:bind` declaration.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Bind {
    /// Optional bind id.
    pub id: Option<String>,
    /// `nodeset` expression.
    pub nodeset: Option<String>,
    /// `ref` expression.
    pub reference: Option<String>,
    /// `context` expression.
    pub context: Option<String>,
    /// Referenced model id.
    pub model: Option<String>,
    /// Schema datatype name.
    pub datatype: Option<String>,
    /// `readonly` expression.
    pub readonly: Option<String>,
    /// `relevant` expression.
    pub relevant: Option<String>,
    /// `required` expression.
    pub required: Option<String>,
    /// `constraint` expression.
    pub constraint: Option<String>,
    /// `calculate` expression.
    pub calculate: Option<String>,
    /// Unknown attributes retained by namespace and local name.
    pub attributes: Vec<Attribute>,
}

/// XForms submission method vocabulary.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum SubmissionMethod {
    Get,
    Put,
    Post,
    Delete,
    Other(String),
}

impl SubmissionMethod {
    fn parse(value: &str) -> Self {
        match value {
            "get" => Self::Get,
            "put" => Self::Put,
            "post" => Self::Post,
            "delete" => Self::Delete,
            value => Self::Other(value.to_owned()),
        }
    }

    fn as_str(&self) -> &str {
        match self {
            Self::Get => "get",
            Self::Put => "put",
            Self::Post => "post",
            Self::Delete => "delete",
            Self::Other(value) => value,
        }
    }
}

/// XForms submission replacement policy.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum SubmissionReplace {
    All,
    Instance,
    None,
    Other(String),
}

impl SubmissionReplace {
    fn parse(value: &str) -> Self {
        match value {
            "all" => Self::All,
            "instance" => Self::Instance,
            "none" => Self::None,
            value => Self::Other(value.to_owned()),
        }
    }

    fn as_str(&self) -> &str {
        match self {
            Self::All => "all",
            Self::Instance => "instance",
            Self::None => "none",
            Self::Other(value) => value,
        }
    }
}

/// A typed XForms `xforms:submission` declaration.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Submission {
    /// Optional submission id.
    pub id: Option<String>,
    /// Action URI.
    pub action: Option<String>,
    /// Submission method.
    pub method: Option<SubmissionMethod>,
    /// XML serialization version string.
    pub version: Option<String>,
    /// Serialization media type.
    pub mediatype: Option<String>,
    /// Serialization encoding.
    pub encoding: Option<String>,
    /// Replacement policy.
    pub replace: Option<SubmissionReplace>,
    /// Instance id.
    pub instance: Option<String>,
    /// Validation expression.
    pub validate: Option<String>,
    /// Relevance expression.
    pub relevant: Option<String>,
    /// Unknown attributes retained by namespace and local name.
    pub attributes: Vec<Attribute>,
}

/// Ordered direct child of an XForms model.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum ModelChild {
    /// Instance declaration.
    Instance(Instance),
    /// Bind declaration.
    Bind(Bind),
    /// Submission declaration.
    Submission(Submission),
    /// Unknown or extension child.
    Extension(Extension),
}

/// A typed, inert `xforms:model` declaration.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Model {
    /// Optional model id.
    pub id: Option<String>,
    /// Unknown model attributes retained by namespace and local name.
    pub attributes: Vec<Attribute>,
    /// Ordered direct model children.
    pub children: Vec<ModelChild>,
    /// Namespace declarations inherited by the source model, retained so
    /// opaque child fragments remain namespace-well-formed after replacement.
    namespace_declarations: Vec<String>,
}

#[derive(Debug)]
struct Frame {
    namespace: String,
    local: String,
    start: usize,
    changes: Vec<(String, Option<String>)>,
    model_declarations: Option<Vec<String>>,
}

impl Model {
    /// Return model instances in source order.
    pub fn instances(&self) -> impl Iterator<Item = &Instance> {
        self.children.iter().filter_map(|child| match child {
            ModelChild::Instance(value) => Some(value),
            _ => None,
        })
    }

    /// Return model binds in source order.
    pub fn binds(&self) -> impl Iterator<Item = &Bind> {
        self.children.iter().filter_map(|child| match child {
            ModelChild::Bind(value) => Some(value),
            _ => None,
        })
    }

    /// Return model submissions in source order.
    pub fn submissions(&self) -> impl Iterator<Item = &Submission> {
        self.children.iter().filter_map(|child| match child {
            ModelChild::Submission(value) => Some(value),
            _ => None,
        })
    }

    /// Return the complete in-scope namespace context retained for this model.
    pub fn namespace_bindings(&self) -> Vec<NamespaceBinding> {
        public_namespace_bindings(&self.namespace_declarations)
    }

    /// Validate bounded model structure and all retained lexical values.
    pub fn validate(&self) -> Result<()> {
        // Check the aggregate retained/serialized footprint before walking
        // opaque instance and extension fragments.  This keeps a caller-built
        // model from parsing or validating unbounded child payloads only to
        // reject the complete edit later.
        let output_len = model_output_upper_bound(self)?;
        bounded_edit_length(output_len, "XForms model")?;
        validate_optional("xforms:model id", self.id.as_deref())?;
        validate_reserved_namespace_bindings(&self.namespace_declarations, "xforms:model")?;
        validate_attributes(&self.attributes)?;
        if self.children.len() > MAX_CHILDREN {
            return Err(Error::InvalidFormat(format!(
                "xforms:model exceeds {MAX_CHILDREN} direct children"
            )));
        }
        for child in &self.children {
            match child {
                ModelChild::Instance(value) => value.validate()?,
                ModelChild::Bind(value) => value.validate()?,
                ModelChild::Submission(value) => value.validate()?,
                ModelChild::Extension(value) => {
                    validate_extension(value, &self.namespace_declarations)?;
                },
            }
        }
        Ok(())
    }

    /// Serialize a model for insertion or replacement.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let output_len = model_output_upper_bound(self)?;
        bounded_edit_length(output_len, "XForms model serialization")?;
        let mut output = String::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|source| Error::Allocation {
                resource: "ODT XForms model serialization",
                source,
            })?;
        output.push_str("<xforms:model xmlns:xforms=\"");
        output.push_str(XFORMSNS);
        output.push_str("\" xmlns:office=\"");
        output.push_str(OFFICENS);
        output.push_str("\" xmlns:form=\"");
        output.push_str(FORMNS);
        output.push('"');
        let mut inherited = vec![
            ("xmlns:xforms".to_string(), XFORMSNS.to_string()),
            ("xmlns:office".to_string(), OFFICENS.to_string()),
            ("xmlns:form".to_string(), FORMNS.to_string()),
        ];
        push_namespace_declarations(&mut output, &self.namespace_declarations, &mut inherited);
        push_optional_attr(&mut output, "id", self.id.as_deref());
        push_unknown_attributes(&mut output, &self.attributes, &mut inherited);
        if self.children.is_empty() {
            output.push_str("/>");
        } else {
            output.push('>');
            for child in &self.children {
                write_child(&mut output, child, &inherited)?;
            }
            output.push_str("</xforms:model>");
        }
        Ok(output)
    }
}

impl Instance {
    /// Return the in-scope namespace context retained for this instance
    /// wrapper and its inline content.
    pub fn namespace_bindings(&self) -> Vec<NamespaceBinding> {
        public_namespace_bindings(&self.namespace_declarations)
    }

    /// Validate lexical values and bounded inline content.
    pub fn validate(&self) -> Result<()> {
        validate_reserved_namespace_bindings(&self.namespace_declarations, "xforms:instance")?;
        for (name, value) in [
            ("xforms:instance id", self.id.as_deref()),
            ("xforms:instance src", self.src.as_deref()),
            ("xforms:instance resource", self.resource.as_deref()),
            ("xforms:instance mediatype", self.mediatype.as_deref()),
            ("xforms:instance content", self.content_xml.as_deref()),
        ] {
            validate_optional(name, value)?;
        }
        if let Some(content) = &self.content_xml {
            validate_fragment_with_context(
                content,
                "XForms instance",
                true,
                &self.namespace_declarations,
            )?;
        }
        validate_attributes(&self.attributes)
    }
}

impl Bind {
    /// Validate lexical values and bounded attributes.
    pub fn validate(&self) -> Result<()> {
        for (name, value) in [
            ("xforms:bind id", self.id.as_deref()),
            ("xforms:bind nodeset", self.nodeset.as_deref()),
            ("xforms:bind ref", self.reference.as_deref()),
            ("xforms:bind context", self.context.as_deref()),
            ("xforms:bind model", self.model.as_deref()),
            ("xforms:bind type", self.datatype.as_deref()),
            ("xforms:bind readonly", self.readonly.as_deref()),
            ("xforms:bind relevant", self.relevant.as_deref()),
            ("xforms:bind required", self.required.as_deref()),
            ("xforms:bind constraint", self.constraint.as_deref()),
            ("xforms:bind calculate", self.calculate.as_deref()),
        ] {
            validate_optional(name, value)?;
        }
        validate_attributes(&self.attributes)
    }
}

impl Submission {
    /// Validate lexical values and bounded attributes.
    pub fn validate(&self) -> Result<()> {
        for (name, value) in [
            ("xforms:submission id", self.id.as_deref()),
            ("xforms:submission action", self.action.as_deref()),
            ("xforms:submission version", self.version.as_deref()),
            ("xforms:submission mediatype", self.mediatype.as_deref()),
            ("xforms:submission encoding", self.encoding.as_deref()),
            ("xforms:submission instance", self.instance.as_deref()),
            ("xforms:submission validate", self.validate.as_deref()),
            ("xforms:submission relevant", self.relevant.as_deref()),
        ] {
            validate_optional(name, value)?;
        }
        if let Some(method) = &self.method {
            validate_xml_value("xforms:submission method", method.as_str())?;
        }
        if let Some(replace) = &self.replace {
            validate_xml_value("xforms:submission replace", replace.as_str())?;
        }
        validate_attributes(&self.attributes)
    }
}

/// Parse all `xforms:model` elements directly below `office:forms`.
pub(crate) fn parse_models(xml: &str) -> Result<Vec<Model>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} XForms model limit"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Frame> = Vec::new();
    let mut spans: Vec<(usize, usize, Vec<String>)> = Vec::new();
    let mut namespace_scope = Vec::new();
    let mut depth = 0usize;
    loop {
        let event_position = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms XML: {error}")))?;
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref source) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("XForms XML depth overflow".to_string()))?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "XForms XML exceeds {MAX_DEPTH} levels"
                    )));
                }
                let local = utf8(source.local_name().as_ref(), "XForms element name")?;
                let declarations = namespace_declarations(source)?;
                let changes = apply_namespace_declarations(&mut namespace_scope, &declarations)?;
                let start = event_position;
                let model_declarations = if namespace == XFORMSNS
                    && local == "model"
                    && stack
                        .last()
                        .is_some_and(|frame| frame.namespace == OFFICENS && frame.local == "forms")
                {
                    Some(namespace_scope_to_raw(&namespace_scope))
                } else {
                    None
                };
                stack.push(Frame {
                    namespace,
                    local,
                    start,
                    changes,
                    model_declarations,
                });
            },
            Event::Empty(ref source) => {
                let local = utf8(source.local_name().as_ref(), "XForms element name")?;
                let declarations = namespace_declarations(source)?;
                let changes = apply_namespace_declarations(&mut namespace_scope, &declarations)?;
                if namespace == XFORMSNS
                    && local == "model"
                    && stack
                        .last()
                        .is_some_and(|frame| frame.namespace == OFFICENS && frame.local == "forms")
                {
                    let start = event_position;
                    if spans.len() >= MAX_MODELS {
                        return Err(Error::InvalidFormat(format!(
                            "ODF contains more than {MAX_MODELS} XForms models"
                        )));
                    }
                    spans.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT XForms model spans",
                        source,
                    })?;
                    spans.push((start, event_end, namespace_scope_to_raw(&namespace_scope)));
                }
                restore_namespace_declarations(&mut namespace_scope, changes);
            },
            Event::End(_) => {
                let end = event_end;
                let frame = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("XForms XML stack underflow".to_string())
                })?;
                if frame.namespace == XFORMSNS
                    && frame.local == "model"
                    && stack.last().is_some_and(|parent| {
                        parent.namespace == OFFICENS && parent.local == "forms"
                    })
                {
                    if spans.len() >= MAX_MODELS {
                        return Err(Error::InvalidFormat(format!(
                            "ODF contains more than {MAX_MODELS} XForms models"
                        )));
                    }
                    spans.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT XForms model spans",
                        source,
                    })?;
                    spans.push((
                        frame.start,
                        end,
                        frame
                            .model_declarations
                            .unwrap_or_else(|| namespace_scope_to_raw(&namespace_scope)),
                    ));
                }
                restore_namespace_declarations(&mut namespace_scope, frame.changes);
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XForms XML".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || !namespace_scope.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete XForms document while locating models".to_string(),
        ));
    }
    spans.sort_by_key(|(start, _, _)| *start);
    if spans.len() > MAX_MODELS {
        return Err(Error::InvalidFormat(format!(
            "ODF contains more than {MAX_MODELS} XForms models"
        )));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(spans.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms model projection",
            source,
        })?;
    for (start, end, declarations) in spans {
        let raw = xml
            .get(start..end)
            .ok_or_else(|| Error::InvalidFormat("invalid XForms model span".to_string()))?;
        let owned = inject_namespace_declarations(raw, &declarations)?;
        output.push(parse_model(&owned, &declarations)?);
    }
    validate_model_set(&output)?;
    Ok(output)
}

fn validate_model_set(models: &[Model]) -> Result<()> {
    let mut ids = HashSet::new();
    let mut model_ids = HashSet::new();
    let mut instance_ids = HashSet::new();
    for model in models {
        if let Some(id) = &model.id {
            insert_model_id(&mut ids, &mut model_ids, id, "xforms:model")?;
        }
        for child in &model.children {
            match child {
                ModelChild::Instance(instance) => {
                    if let Some(id) = &instance.id {
                        insert_model_id(&mut ids, &mut instance_ids, id, "xforms:instance")?;
                    }
                },
                ModelChild::Bind(bind) => {
                    if let Some(id) = &bind.id {
                        insert_unique_id(&mut ids, id, "xforms:bind")?;
                    }
                },
                ModelChild::Submission(submission) => {
                    if let Some(id) = &submission.id {
                        insert_unique_id(&mut ids, id, "xforms:submission")?;
                    }
                },
                ModelChild::Extension(_) => {},
            }
        }
    }
    for model in models {
        for child in &model.children {
            match child {
                ModelChild::Bind(bind) => {
                    if let Some(reference) = &bind.model {
                        if !model_ids.contains(reference.as_str()) {
                            return Err(Error::InvalidFormat(format!(
                                "xforms:bind model reference '{reference}' is dangling"
                            )));
                        }
                    }
                },
                ModelChild::Submission(submission) => {
                    if let Some(reference) = &submission.instance {
                        if !instance_ids.contains(reference.as_str()) {
                            return Err(Error::InvalidFormat(format!(
                                "xforms:submission instance reference '{reference}' is dangling"
                            )));
                        }
                    }
                },
                ModelChild::Instance(_) | ModelChild::Extension(_) => {},
            }
        }
    }
    Ok(())
}

fn insert_model_id<'a>(
    ids: &mut HashSet<&'a str>,
    family: &mut HashSet<&'a str>,
    id: &'a str,
    element: &str,
) -> Result<()> {
    if !ids.insert(id) {
        return Err(Error::InvalidFormat(format!(
            "duplicate XForms id '{id}' on {element}"
        )));
    }
    family.insert(id);
    Ok(())
}

fn insert_unique_id<'a>(ids: &mut HashSet<&'a str>, id: &'a str, element: &str) -> Result<()> {
    if !ids.insert(id) {
        return Err(Error::InvalidFormat(format!(
            "duplicate XForms id '{id}' on {element}"
        )));
    }
    Ok(())
}

fn parse_model(xml: &str, inherited_declarations: &[String]) -> Result<Model> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack: Vec<(String, String, usize, Vec<Attribute>)> = Vec::new();
    let mut model = Model {
        namespace_declarations: inherited_declarations.to_vec(),
        ..Model::default()
    };
    let mut children = Vec::new();
    loop {
        let event_position = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms model XML: {error}")))?;
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref source) => {
                depth += 1;
                if depth == 1 {
                    parse_model_attributes(&reader, source, &mut model)?;
                    model.namespace_declarations = merge_namespace_declarations(
                        inherited_declarations,
                        &namespace_declarations(source)?,
                    );
                }
                let ns = namespace.clone();
                let local = utf8(source.local_name().as_ref(), "XForms model element name")?;
                let start = event_position;
                let attrs = attributes(&reader, source)?;
                stack.push((ns, local, start, attrs));
            },
            Event::Empty(ref source) => {
                let ns = namespace.clone();
                let local = utf8(source.local_name().as_ref(), "XForms child name")?;
                let start = event_position;
                if depth == 0 {
                    if ns != XFORMSNS || local != "model" {
                        return Err(Error::InvalidFormat(
                            "XForms model fragment has the wrong root element".to_string(),
                        ));
                    }
                    parse_model_attributes(&reader, source, &mut model)?;
                    model.namespace_declarations = merge_namespace_declarations(
                        inherited_declarations,
                        &namespace_declarations(source)?,
                    );
                } else if depth == 1 {
                    let raw = xml.get(start..event_end).ok_or_else(|| {
                        Error::InvalidFormat("invalid XForms child span".to_string())
                    })?;
                    children.push(parse_child(
                        &ns,
                        &local,
                        raw,
                        &reader,
                        source,
                        &model.namespace_declarations,
                    )?);
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("XForms model stack underflow".to_string())
                })?;
                if depth == 2 {
                    let raw = xml.get(frame.2..event_end).ok_or_else(|| {
                        Error::InvalidFormat("invalid XForms child span".to_string())
                    })?;
                    // The start element's attributes are parsed again against
                    // the child fragment; this preserves the complete opaque
                    // body while exposing the common declaration attributes.
                    children.push(parse_child_values(
                        &frame.0,
                        &frame.1,
                        raw,
                        frame.3,
                        &model.namespace_declarations,
                    )?);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms model depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XForms model XML".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete XForms model fragment".to_string(),
        ));
    }
    model.children = children;
    model.validate()?;
    Ok(model)
}

fn parse_model_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    model: &mut Model,
) -> Result<()> {
    let attrs = attributes(reader, source)?;
    reject_duplicate_known_attributes(&attrs, XFORMSNS, &["id"], "xforms:model")?;
    for attr in attrs {
        if (attr.namespace_uri.is_empty() || attr.namespace_uri == XFORMSNS)
            && attr.local_name == "id"
        {
            model.id = Some(attr.value);
        } else {
            model.attributes.push(attr);
        }
    }
    Ok(())
}

fn parse_child_values(
    namespace: &str,
    local: &str,
    raw: &str,
    attrs: Vec<Attribute>,
    inherited_declarations: &[String],
) -> Result<ModelChild> {
    match (namespace, local) {
        (XFORMSNS, "instance") => Ok(ModelChild::Instance(parse_instance(
            raw,
            attrs,
            inherited_declarations,
        )?)),
        (XFORMSNS, "bind") => Ok(ModelChild::Bind(parse_bind(attrs)?)),
        (XFORMSNS, "submission") => Ok(ModelChild::Submission(parse_submission(attrs)?)),
        _ => {
            bounded_fragment_length(raw, "XForms extension")?;
            Ok(ModelChild::Extension(Extension {
                namespace_uri: namespace.to_owned(),
                local_name: local.to_owned(),
                xml: raw.to_owned(),
            }))
        },
    }
}

fn parse_child(
    namespace: &str,
    local: &str,
    raw: &str,
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    inherited_declarations: &[String],
) -> Result<ModelChild> {
    let attrs = attributes(reader, source)?;
    match (namespace, local) {
        (XFORMSNS, "instance") => Ok(ModelChild::Instance(parse_instance(
            raw,
            attrs,
            inherited_declarations,
        )?)),
        (XFORMSNS, "bind") => Ok(ModelChild::Bind(parse_bind(attrs)?)),
        (XFORMSNS, "submission") => Ok(ModelChild::Submission(parse_submission(attrs)?)),
        _ => {
            bounded_fragment_length(raw, "XForms extension")?;
            Ok(ModelChild::Extension(Extension {
                namespace_uri: namespace.to_owned(),
                local_name: local.to_owned(),
                xml: raw.to_owned(),
            }))
        },
    }
}

fn parse_instance(
    raw: &str,
    attrs: Vec<Attribute>,
    inherited_declarations: &[String],
) -> Result<Instance> {
    let local_declarations = first_namespace_declarations(raw)?;
    reject_duplicate_known_attributes(
        &attrs,
        XFORMSNS,
        &["id", "src", "resource", "mediatype"],
        "xforms:instance",
    )?;
    let mut value = Instance {
        id: take_attr(&attrs, XFORMSNS, "id"),
        src: take_attr(&attrs, XFORMSNS, "src"),
        resource: take_attr(&attrs, XFORMSNS, "resource"),
        mediatype: take_attr(&attrs, XFORMSNS, "mediatype"),
        content_xml: None,
        attributes: unknown_attrs(attrs, XFORMSNS, &["id", "src", "resource", "mediatype"]),
        namespace_declarations: merge_namespace_declarations(
            inherited_declarations,
            &local_declarations,
        ),
    };
    value.content_xml = inner_xml(raw)?;
    Ok(value)
}

fn parse_bind(attrs: Vec<Attribute>) -> Result<Bind> {
    reject_duplicate_known_attributes(
        &attrs,
        XFORMSNS,
        &[
            "id",
            "nodeset",
            "ref",
            "context",
            "model",
            "type",
            "readonly",
            "relevant",
            "required",
            "constraint",
            "calculate",
        ],
        "xforms:bind",
    )?;
    Ok(Bind {
        id: take_attr(&attrs, XFORMSNS, "id"),
        nodeset: take_attr(&attrs, XFORMSNS, "nodeset"),
        reference: take_attr(&attrs, XFORMSNS, "ref"),
        context: take_attr(&attrs, XFORMSNS, "context"),
        model: take_attr(&attrs, XFORMSNS, "model"),
        datatype: take_attr(&attrs, XFORMSNS, "type"),
        readonly: take_attr(&attrs, XFORMSNS, "readonly"),
        relevant: take_attr(&attrs, XFORMSNS, "relevant"),
        required: take_attr(&attrs, XFORMSNS, "required"),
        constraint: take_attr(&attrs, XFORMSNS, "constraint"),
        calculate: take_attr(&attrs, XFORMSNS, "calculate"),
        attributes: unknown_attrs(
            attrs,
            XFORMSNS,
            &[
                "id",
                "nodeset",
                "ref",
                "context",
                "model",
                "type",
                "readonly",
                "relevant",
                "required",
                "constraint",
                "calculate",
            ],
        ),
    })
}

fn parse_submission(attrs: Vec<Attribute>) -> Result<Submission> {
    reject_duplicate_known_attributes(
        &attrs,
        XFORMSNS,
        &[
            "id",
            "action",
            "method",
            "version",
            "mediatype",
            "encoding",
            "replace",
            "instance",
            "validate",
            "relevant",
        ],
        "xforms:submission",
    )?;
    let method = take_attr(&attrs, XFORMSNS, "method").map(|value| SubmissionMethod::parse(&value));
    let replace =
        take_attr(&attrs, XFORMSNS, "replace").map(|value| SubmissionReplace::parse(&value));
    Ok(Submission {
        id: take_attr(&attrs, XFORMSNS, "id"),
        action: take_attr(&attrs, XFORMSNS, "action"),
        method,
        version: take_attr(&attrs, XFORMSNS, "version"),
        mediatype: take_attr(&attrs, XFORMSNS, "mediatype"),
        encoding: take_attr(&attrs, XFORMSNS, "encoding"),
        replace,
        instance: take_attr(&attrs, XFORMSNS, "instance"),
        validate: take_attr(&attrs, XFORMSNS, "validate"),
        relevant: take_attr(&attrs, XFORMSNS, "relevant"),
        attributes: unknown_attrs(
            attrs,
            XFORMSNS,
            &[
                "id",
                "action",
                "method",
                "version",
                "mediatype",
                "encoding",
                "replace",
                "instance",
                "validate",
                "relevant",
            ],
        ),
    })
}

fn write_child(
    output: &mut String,
    child: &ModelChild,
    inherited: &[(String, String)],
) -> Result<()> {
    match child {
        ModelChild::Extension(value) => output.push_str(&value.xml),
        ModelChild::Instance(value) => {
            value.validate()?;
            output.push_str("<xforms:instance");
            let mut child_namespace = inherited.to_vec();
            push_namespace_declarations(
                output,
                &value.namespace_declarations,
                &mut child_namespace,
            );
            push_optional_attr(output, "id", value.id.as_deref());
            push_optional_attr(output, "src", value.src.as_deref());
            push_optional_attr(output, "resource", value.resource.as_deref());
            push_optional_attr(output, "mediatype", value.mediatype.as_deref());
            push_unknown_attributes(output, &value.attributes, &mut child_namespace);
            if let Some(content) = &value.content_xml {
                output.push('>');
                output.push_str(content);
                output.push_str("</xforms:instance>");
            } else {
                output.push_str("/>");
            }
        },
        ModelChild::Bind(value) => {
            value.validate()?;
            output.push_str("<xforms:bind");
            push_optional_attr(output, "id", value.id.as_deref());
            push_optional_attr(output, "nodeset", value.nodeset.as_deref());
            push_optional_attr(output, "ref", value.reference.as_deref());
            push_optional_attr(output, "context", value.context.as_deref());
            push_optional_attr(output, "model", value.model.as_deref());
            push_optional_attr(output, "type", value.datatype.as_deref());
            push_optional_attr(output, "readonly", value.readonly.as_deref());
            push_optional_attr(output, "relevant", value.relevant.as_deref());
            push_optional_attr(output, "required", value.required.as_deref());
            push_optional_attr(output, "constraint", value.constraint.as_deref());
            push_optional_attr(output, "calculate", value.calculate.as_deref());
            let mut child_namespace = inherited.to_vec();
            push_unknown_attributes(output, &value.attributes, &mut child_namespace);
            output.push_str("/>");
        },
        ModelChild::Submission(value) => {
            value.validate()?;
            output.push_str("<xforms:submission");
            push_optional_attr(output, "id", value.id.as_deref());
            push_optional_attr(output, "action", value.action.as_deref());
            push_optional_attr(
                output,
                "method",
                value.method.as_ref().map(SubmissionMethod::as_str),
            );
            push_optional_attr(output, "version", value.version.as_deref());
            push_optional_attr(output, "mediatype", value.mediatype.as_deref());
            push_optional_attr(output, "encoding", value.encoding.as_deref());
            push_optional_attr(
                output,
                "replace",
                value.replace.as_ref().map(SubmissionReplace::as_str),
            );
            push_optional_attr(output, "instance", value.instance.as_deref());
            push_optional_attr(output, "validate", value.validate.as_deref());
            push_optional_attr(output, "relevant", value.relevant.as_deref());
            let mut child_namespace = inherited.to_vec();
            push_unknown_attributes(output, &value.attributes, &mut child_namespace);
            output.push_str("/>");
        },
    }
    Ok(())
}

fn model_output_upper_bound(model: &Model) -> Result<usize> {
    let mut length = 512usize;
    add_xml_budget(&mut length, model.id.as_deref())?;
    for declaration in &model.namespace_declarations {
        add_raw_budget(&mut length, declaration)?;
    }
    add_attribute_budget(&mut length, &model.attributes)?;
    for child in &model.children {
        add_size_budget(&mut length, 256)?;
        match child {
            ModelChild::Instance(value) => {
                add_xml_budget(&mut length, value.id.as_deref())?;
                add_xml_budget(&mut length, value.src.as_deref())?;
                add_xml_budget(&mut length, value.resource.as_deref())?;
                add_xml_budget(&mut length, value.mediatype.as_deref())?;
                add_raw_budget_option(&mut length, value.content_xml.as_deref())?;
                for declaration in &value.namespace_declarations {
                    add_raw_budget(&mut length, declaration)?;
                }
                add_attribute_budget(&mut length, &value.attributes)?;
            },
            ModelChild::Bind(value) => {
                for item in [
                    value.id.as_deref(),
                    value.nodeset.as_deref(),
                    value.reference.as_deref(),
                    value.context.as_deref(),
                    value.model.as_deref(),
                    value.datatype.as_deref(),
                    value.readonly.as_deref(),
                    value.relevant.as_deref(),
                    value.required.as_deref(),
                    value.constraint.as_deref(),
                    value.calculate.as_deref(),
                ] {
                    add_xml_budget(&mut length, item)?;
                }
                add_attribute_budget(&mut length, &value.attributes)?;
            },
            ModelChild::Submission(value) => {
                for item in [
                    value.id.as_deref(),
                    value.action.as_deref(),
                    value.method.as_ref().map(SubmissionMethod::as_str),
                    value.version.as_deref(),
                    value.mediatype.as_deref(),
                    value.encoding.as_deref(),
                    value.replace.as_ref().map(SubmissionReplace::as_str),
                    value.instance.as_deref(),
                    value.validate.as_deref(),
                    value.relevant.as_deref(),
                ] {
                    add_xml_budget(&mut length, item)?;
                }
                add_attribute_budget(&mut length, &value.attributes)?;
            },
            ModelChild::Extension(value) => add_raw_budget(&mut length, &value.xml)?,
        }
    }
    Ok(length)
}

fn add_attribute_budget(total: &mut usize, attributes: &[Attribute]) -> Result<()> {
    for attribute in attributes {
        add_size_budget(total, 256)?;
        add_xml_budget(total, Some(&attribute.namespace_uri))?;
        add_xml_budget(total, Some(&attribute.local_name))?;
        add_xml_budget(total, Some(&attribute.value))?;
        add_xml_budget(total, attribute.prefix.as_deref())?;
    }
    Ok(())
}

fn add_xml_budget(total: &mut usize, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        let amount = escaped_xml_len(value)?;
        add_size_budget(total, amount)?;
    }
    Ok(())
}

fn add_raw_budget_option(total: &mut usize, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        add_raw_budget(total, value)?;
    }
    Ok(())
}

fn add_raw_budget(total: &mut usize, value: &str) -> Result<()> {
    add_size_budget(total, value.len())
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    let mut length = 0usize;
    for byte in value.bytes() {
        let extra = match byte {
            b'&' => 4,
            b'<' | b'>' => 3,
            b'"' | b'\'' => 5,
            _ => 0,
        };
        add_size_budget(&mut length, 1 + extra)?;
    }
    Ok(length)
}

fn add_size_budget(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::InvalidFormat("XForms serialization size overflow".to_string()))?;
    Ok(())
}

fn attributes(reader: &NsReader<&[u8]>, source: &BytesStart<'_>) -> Result<Vec<Attribute>> {
    let mut output = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms attribute: {error}")))?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if output.len() >= MAX_ATTRIBUTES {
            return Err(Error::InvalidFormat(format!(
                "XForms element exceeds {MAX_ATTRIBUTES} attributes"
            )));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace(&namespace)?.unwrap_or_default();
        let prefix = attribute
            .key
            .as_ref()
            .split(|byte| *byte == b':')
            .next()
            .filter(|candidate| *candidate != attribute.key.as_ref())
            .map(|candidate| String::from_utf8_lossy(candidate).into_owned());
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid XForms attribute value: {error}"))
            })?;
        output.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODT XForms attribute projection",
            source,
        })?;
        output.push(Attribute {
            namespace_uri,
            local_name: utf8(local.as_ref(), "XForms attribute name")?,
            value: value.into_owned(),
            prefix,
        });
    }
    if output.len() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "XForms element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    Ok(output)
}

fn namespace_declarations(source: &BytesStart<'_>) -> Result<Vec<String>> {
    let mut output = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            if output.len() >= MAX_ATTRIBUTES {
                return Err(Error::InvalidFormat(format!(
                    "XForms element exceeds {MAX_ATTRIBUTES} namespace declarations"
                )));
            }
            let key = std::str::from_utf8(raw).map_err(|_| {
                Error::InvalidFormat("invalid namespace declaration name".to_string())
            })?;
            let value = std::str::from_utf8(attribute.value.as_ref()).map_err(|_| {
                Error::InvalidFormat("invalid namespace declaration value".to_string())
            })?;
            output.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "ODT XForms namespace declarations",
                source,
            })?;
            output.push(format!(" {key}=\"{value}\""));
        }
    }
    Ok(output)
}

fn merge_namespace_declarations(base: &[String], local: &[String]) -> Vec<String> {
    let mut scope = base
        .iter()
        .filter_map(|declaration| declaration_parts(declaration))
        .map(|(name, value)| (name.to_string(), value.to_string()))
        .collect::<Vec<_>>();
    for declaration in local {
        if let Some((name, value)) = declaration_parts(declaration) {
            replace_namespace(&mut scope, name, value);
        }
    }
    namespace_scope_to_raw(&scope)
}

fn apply_namespace_declarations(
    scope: &mut Vec<(String, String)>,
    declarations: &[String],
) -> Result<Vec<(String, Option<String>)>> {
    let mut changes = Vec::new();
    for declaration in declarations {
        let (name, value) = declaration_parts(declaration).ok_or_else(|| {
            Error::InvalidFormat("invalid namespace declaration span".to_string())
        })?;
        let previous = scope
            .iter()
            .find(|(current, _)| current == name)
            .map(|(_, value)| value.clone());
        changes.push((name.to_string(), previous));
        replace_namespace(scope, name, value);
    }
    Ok(changes)
}

fn restore_namespace_declarations(
    scope: &mut Vec<(String, String)>,
    changes: Vec<(String, Option<String>)>,
) {
    for (name, previous) in changes.into_iter().rev() {
        match previous {
            Some(value) => replace_namespace(scope, &name, &value),
            None => scope.retain(|(current, _)| current != &name),
        }
    }
}

fn replace_namespace(scope: &mut Vec<(String, String)>, name: &str, value: &str) {
    if let Some((_, current)) = scope.iter_mut().find(|(current, _)| current == name) {
        *current = value.to_string();
    } else {
        scope.push((name.to_string(), value.to_string()));
    }
}

fn namespace_scope_to_raw(scope: &[(String, String)]) -> Vec<String> {
    scope
        .iter()
        .map(|(name, value)| format!(" {name}=\"{value}\""))
        .collect()
}

fn push_namespace_declarations(
    output: &mut String,
    declarations: &[String],
    scope: &mut Vec<(String, String)>,
) {
    for declaration in declarations {
        let Some((name, value)) = declaration_parts(declaration) else {
            continue;
        };
        if scope
            .iter()
            .any(|(current, current_value)| current == name && current_value == value)
        {
            continue;
        }
        output.push_str(declaration);
        replace_namespace(scope, name, value);
    }
}

fn declaration_parts(declaration: &str) -> Option<(&str, &str)> {
    let declaration = declaration.trim();
    let (name, value) = declaration.split_once('=')?;
    let value = value.trim();
    let value = value.strip_prefix('"')?.strip_suffix('"')?;
    Some((name.trim(), value))
}

fn public_namespace_bindings(declarations: &[String]) -> Vec<NamespaceBinding> {
    declarations
        .iter()
        .filter_map(|declaration| {
            let (name, value) = declaration_parts(declaration)?;
            let prefix = name.strip_prefix("xmlns:").map(str::to_owned);
            Some(NamespaceBinding {
                prefix,
                namespace_uri: value.to_owned(),
            })
        })
        .collect()
}

fn first_namespace_declarations(raw: &str) -> Result<Vec<String>> {
    let mut reader = NsReader::from_str(raw);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms element: {error}")))?;
        match event {
            Event::Start(source) | Event::Empty(source) => {
                return namespace_declarations(&source);
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XForms element".to_string(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "missing XForms element root".to_string(),
                ));
            },
            _ => buffer.clear(),
        }
    }
}

fn inject_namespace_declarations(raw: &str, declarations: &[String]) -> Result<String> {
    let (open_end, empty, present) = first_tag_span(raw)?;
    let mut insert_len = 0usize;
    for declaration in declarations {
        let Some((name, _)) = declaration_parts(declaration) else {
            continue;
        };
        if !present.contains(name) {
            insert_len = insert_len.checked_add(declaration.len()).ok_or_else(|| {
                Error::InvalidFormat("XForms namespace insertion size overflow".to_string())
            })?;
        }
    }
    if insert_len == 0 {
        return Ok(raw.to_owned());
    }
    let offset = if empty {
        open_end
            .checked_sub(2)
            .ok_or_else(|| Error::InvalidFormat("invalid empty XForms element span".to_string()))?
    } else {
        open_end
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("invalid XForms element span".to_string()))?
    };
    let output_len = raw.len().checked_add(insert_len).ok_or_else(|| {
        Error::InvalidFormat("XForms namespace insertion size overflow".to_string())
    })?;
    bounded_edit_length(output_len, "XForms namespace insertion")?;
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms namespace insertion",
            source,
        })?;
    output.push_str(&raw[..offset]);
    for declaration in declarations {
        let Some((name, _)) = declaration_parts(declaration) else {
            continue;
        };
        if !present.contains(name) {
            output.push_str(declaration);
        }
    }
    output.push_str(&raw[offset..]);
    Ok(output)
}

fn first_tag_span(raw: &str) -> Result<(usize, bool, HashSet<String>)> {
    let mut reader = NsReader::from_str(raw);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XML element: {error}")))?;
        match event {
            Event::Start(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    false,
                    present_namespace_names(&source)?,
                ));
            },
            Event::Empty(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    true,
                    present_namespace_names(&source)?,
                ));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XML element".to_string(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat("missing XML element root".to_string()));
            },
            _ => buffer.clear(),
        }
    }
}

fn present_namespace_names(source: &BytesStart<'_>) -> Result<HashSet<String>> {
    let mut present = HashSet::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid XML namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            present.insert(
                std::str::from_utf8(raw)
                    .map_err(|_| {
                        Error::InvalidFormat("invalid XML namespace declaration name".to_string())
                    })?
                    .to_owned(),
            );
        }
    }
    Ok(present)
}

fn unknown_attrs(attrs: Vec<Attribute>, namespace: &str, known: &[&str]) -> Vec<Attribute> {
    attrs
        .into_iter()
        .filter(|attr| {
            !(known.contains(&attr.local_name.as_str())
                && (attr.namespace_uri == namespace
                    || (namespace == XFORMSNS && attr.namespace_uri.is_empty())))
        })
        .collect()
}

fn take_attr(attrs: &[Attribute], namespace: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|attr| {
            attr.local_name == local
                && (attr.namespace_uri == namespace
                    || (namespace == XFORMSNS && attr.namespace_uri.is_empty()))
        })
        .map(|attr| attr.value.clone())
}

fn inner_xml(raw: &str) -> Result<Option<String>> {
    let mut reader = NsReader::from_str(raw);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut content_start = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid XForms instance XML: {error}"))
        })?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms instance depth overflow".to_string())
                })?;
                if depth == 1 {
                    content_start = Some(after);
                }
            },
            Event::Empty(_) if depth == 0 => return Ok(None),
            Event::End(_) if depth == 1 => {
                let close = before;
                let content = raw
                    .get(content_start.unwrap_or(close)..close)
                    .ok_or_else(|| {
                        Error::InvalidFormat("invalid XForms instance content span".to_string())
                    })?;
                bounded_fragment_length(content, "XForms instance content")?;
                return raw
                    .get(content_start.unwrap_or(close)..close)
                    .map(|value| Some(value.to_owned()))
                    .ok_or_else(|| {
                        Error::InvalidFormat("invalid XForms instance content span".to_string())
                    });
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms instance depth underflow".to_string())
                })?;
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated XForms instance".to_string(),
                ));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XForms instance".to_string(),
                ));
            },
            _ => {},
        }
        buffer.clear();
    }
}

fn push_optional_attr(output: &mut String, local: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push(' ');
        output.push_str(local);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
    }
}

fn push_unknown_attributes(
    output: &mut String,
    attrs: &[Attribute],
    namespace_scope: &mut Vec<(String, String)>,
) {
    let mut dynamic_prefixes: Vec<(String, String)> = Vec::new();
    for attr in attrs {
        let prefix = if attr.namespace_uri.is_empty() {
            String::new()
        } else if let Some(prefix) = canonical_prefix(&attr.namespace_uri).filter(|prefix| {
            namespace_scope.iter().all(|(name, namespace)| {
                name != &format!("xmlns:{prefix}") || namespace == &attr.namespace_uri
            })
        }) {
            prefix.to_string()
        } else if let Some(prefix) = attr.prefix.as_deref().filter(|prefix| {
            is_valid_prefix(prefix)
                && namespace_scope.iter().all(|(name, namespace)| {
                    name != &format!("xmlns:{prefix}") || namespace == &attr.namespace_uri
                })
        }) {
            let prefix = prefix.to_string();
            if !namespace_scope.iter().any(|(name, namespace)| {
                name == &format!("xmlns:{prefix}") && namespace == &attr.namespace_uri
            }) {
                output.push(' ');
                output.push_str("xmlns:");
                output.push_str(&prefix);
                output.push_str("=\"");
                output.push_str(&escape_xml(&attr.namespace_uri));
                output.push('"');
                namespace_scope.push((format!("xmlns:{prefix}"), attr.namespace_uri.clone()));
            }
            prefix
        } else if let Some((prefix, _)) = dynamic_prefixes
            .iter()
            .find(|(_, namespace)| namespace == &attr.namespace_uri)
        {
            prefix.clone()
        } else if let Some((prefix, _)) = namespace_scope.iter().find(|(name, namespace)| {
            namespace == &attr.namespace_uri && name.starts_with("xmlns:")
        }) {
            prefix.trim_start_matches("xmlns:").to_string()
        } else {
            let mut suffix = dynamic_prefixes.len();
            let prefix = loop {
                let candidate = format!("ext{suffix}");
                if namespace_scope
                    .iter()
                    .all(|(name, _)| name != &format!("xmlns:{candidate}"))
                {
                    break candidate;
                }
                suffix = suffix.saturating_add(1);
            };
            output.push(' ');
            output.push_str("xmlns:");
            output.push_str(&prefix);
            output.push_str("=\"");
            output.push_str(&escape_xml(&attr.namespace_uri));
            output.push('"');
            dynamic_prefixes.push((prefix.clone(), attr.namespace_uri.clone()));
            namespace_scope.push((format!("xmlns:{prefix}"), attr.namespace_uri.clone()));
            prefix
        };
        output.push(' ');
        if !prefix.is_empty() {
            output.push_str(&prefix);
            output.push(':');
        }
        output.push_str(&attr.local_name);
        output.push_str("=\"");
        output.push_str(&escape_xml(&attr.value));
        output.push('"');
    }
}

fn is_valid_prefix(prefix: &str) -> bool {
    let Some(first) = prefix.as_bytes().first().copied() else {
        return false;
    };
    (first.is_ascii_alphabetic() || first == b'_')
        && prefix != "xml"
        && prefix != "xmlns"
        && prefix
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.' | b'-'))
}

fn canonical_prefix(namespace: &str) -> Option<&str> {
    match namespace {
        XFORMSNS => Some("xforms"),
        XMLNS => Some("xml"),
        FORMNS => Some("form"),
        OFFICENS => Some("office"),
        _ => None,
    }
}

fn validate_attributes(attrs: &[Attribute]) -> Result<()> {
    if attrs.len() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "XForms element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    let mut seen = HashSet::new();
    for attr in attrs {
        if attr.local_name.is_empty() || attr.local_name.contains(':') || attr.local_name == "xmlns"
        {
            return Err(Error::InvalidFormat(
                "XForms attribute local name is not serializable".to_string(),
            ));
        }
        validate_xml_value("XForms attribute namespace", &attr.namespace_uri)?;
        validate_xml_name(&attr.local_name, "XForms attribute name")?;
        validate_xml_value("XForms attribute value", &attr.value)?;
        if !seen.insert((&attr.namespace_uri, &attr.local_name)) {
            return Err(Error::InvalidFormat(
                "duplicate namespace-resolved XForms attribute".to_string(),
            ));
        }
        if let Some(prefix) = &attr.prefix {
            if !is_valid_prefix(prefix) {
                return Err(Error::InvalidFormat(
                    "XForms attribute prefix is not serializable".to_string(),
                ));
            }
        }
    }
    Ok(())
}

fn reject_duplicate_known_attributes(
    attrs: &[Attribute],
    namespace: &str,
    known: &[&str],
    resource: &str,
) -> Result<()> {
    let mut seen = HashSet::new();
    for attr in attrs {
        if known.contains(&attr.local_name.as_str())
            && (attr.namespace_uri == namespace
                || (namespace == XFORMSNS && attr.namespace_uri.is_empty()))
            && !seen.insert(attr.local_name.as_str())
        {
            return Err(Error::InvalidFormat(format!(
                "duplicate {resource} attribute '{}'",
                attr.local_name
            )));
        }
    }
    Ok(())
}

fn validate_optional(name: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        validate_xml_value(name, value)?;
    }
    Ok(())
}

fn validate_xml_value(name: &str, value: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES || !value.chars().all(is_xml_1_0_char) {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds its bounded XML lexical value"
        )));
    }
    Ok(())
}

fn validate_xml_name(value: &str, name: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    };
    if !(first == '_' || first.is_alphabetic())
        || !chars.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
    {
        return Err(Error::InvalidFormat(format!("invalid XML name '{value}'")));
    }
    validate_xml_value(name, value)
}

fn validate_fragment_with_context(
    xml: &str,
    resource: &str,
    allow_empty: bool,
    inherited_declarations: &[String],
) -> Result<()> {
    if xml.len() > MAX_STRING_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{resource} XML exceeds its bounded size"
        )));
    }
    if xml.trim().is_empty() {
        if allow_empty {
            return Ok(());
        }
        return Err(Error::InvalidFormat(format!(
            "{resource} XML must contain one root element"
        )));
    }
    let owned = if inherited_declarations.is_empty() {
        None
    } else {
        Some(inject_namespace_declarations(xml, inherited_declarations)?)
    };
    let source = owned.as_deref().unwrap_or(xml);
    if source.len() > MAX_STRING_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{resource} XML exceeds its bounded size after namespace context injection"
        )));
    }
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut closed_root = false;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid {resource} XML: {error}")))?;
        resolved_namespace(&resolved)?;
        match event {
            Event::Start(ref value) => {
                validate_fragment_attributes(&reader, value)?;
                if depth == 0 {
                    if closed_root {
                        return Err(Error::InvalidFormat(format!(
                            "{resource} XML contains more than one root element"
                        )));
                    }
                    roots = roots.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat(format!("{resource} root count overflow"))
                    })?;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat(format!("{resource} depth overflow")))?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "{resource} exceeds {MAX_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref value) => {
                validate_fragment_attributes(&reader, value)?;
                if depth == 0 {
                    if closed_root {
                        return Err(Error::InvalidFormat(format!(
                            "{resource} XML contains more than one root element"
                        )));
                    }
                    roots = roots.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat(format!("{resource} root count overflow"))
                    })?;
                    closed_root = true;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat(format!("{resource} stack underflow")))?;
                if depth == 0 {
                    closed_root = true;
                }
            },
            Event::Text(value) if depth == 0 => {
                if !value
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return Err(Error::InvalidFormat(format!(
                        "{resource} XML has non-whitespace text outside its root"
                    )));
                }
            },
            Event::CData(value) if depth == 0 => {
                if !value
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return Err(Error::InvalidFormat(format!(
                        "{resource} XML has non-whitespace text outside its root"
                    )));
                }
            },
            Event::GeneralRef(value) if depth == 0 => {
                if !value
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return Err(Error::InvalidFormat(format!(
                        "{resource} XML has a reference outside its root"
                    )));
                }
            },
            Event::GeneralRef(value) => {
                let name = std::str::from_utf8(value.as_ref()).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "{resource} XML contains an invalid entity reference"
                    ))
                })?;
                if quick_xml::escape::resolve_xml_entity(name).is_none() {
                    return Err(Error::InvalidFormat(format!(
                        "{resource} XML contains an undeclared entity reference"
                    )));
                }
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(format!(
                    "DOCTYPE is not permitted in {resource} XML"
                )));
            },
            Event::Decl(_) => {
                return Err(Error::InvalidFormat(format!(
                    "XML declarations are not permitted inside {resource} fragments"
                )));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || roots != 1 {
        return Err(Error::InvalidFormat(format!(
            "{resource} XML is not one complete fragment"
        )));
    }
    Ok(())
}

fn bounded_fragment_length(value: &str, resource: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{resource} XML exceeds its bounded size"
        )));
    }
    Ok(())
}

fn validate_extension(value: &Extension, inherited_declarations: &[String]) -> Result<()> {
    validate_fragment_with_context(
        &value.xml,
        "XForms extension",
        false,
        inherited_declarations,
    )?;
    let source = if inherited_declarations.is_empty() {
        None
    } else {
        Some(inject_namespace_declarations(
            &value.xml,
            inherited_declarations,
        )?)
    };
    let root = extension_root_name(source.as_deref().unwrap_or(&value.xml))?;
    if root.0 != value.namespace_uri || root.1 != value.local_name {
        return Err(Error::InvalidFormat(
            "XForms extension expanded root does not match its retained identity".to_string(),
        ));
    }
    Ok(())
}

fn extension_root_name(xml: &str) -> Result<(String, String)> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid XForms extension XML: {error}"))
            })?;
        match event {
            Event::Start(source) | Event::Empty(source) => {
                let namespace = resolved_namespace(&resolved)?.unwrap_or_default();
                let local = utf8(source.local_name().as_ref(), "XForms extension root name")?;
                return Ok((namespace, local));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XForms extension XML".to_string(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "missing XForms extension root element".to_string(),
                ));
            },
            _ => buffer.clear(),
        }
    }
}

fn validate_fragment_attributes(reader: &NsReader<&[u8]>, source: &BytesStart<'_>) -> Result<()> {
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid XML fragment attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        resolved_namespace(&namespace)?;
    }
    Ok(())
}

fn validate_reserved_namespace_bindings(declarations: &[String], resource: &str) -> Result<()> {
    for declaration in declarations {
        let Some((name, value)) = declaration_parts(declaration) else {
            return Err(Error::InvalidFormat(format!(
                "{resource} has an invalid namespace declaration"
            )));
        };
        let expected = match name {
            "xmlns:xforms" => Some(XFORMSNS),
            "xmlns:office" => Some(OFFICENS),
            "xmlns:form" => Some(FORMNS),
            "xmlns:xml" => Some(XMLNS),
            _ => None,
        };
        if let Some(expected) = expected {
            if value != expected {
                return Err(Error::InvalidFormat(format!(
                    "{resource} shadows reserved {name} namespace"
                )));
            }
        }
    }
    Ok(())
}

const fn is_xml_1_0_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || (value as u32 >= 0x20 && value as u32 <= 0xD7FF)
        || (value as u32 >= 0xE000 && value as u32 <= 0xFFFD)
        || (value as u32 >= 0x10000 && value as u32 <= 0x10FFFF)
}

fn resolved_namespace(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(value) => Ok(Some(utf8(value.as_ref(), "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn utf8(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 {description}")))
}

/// Find direct model spans for the mutable XML editor.
pub(crate) fn model_spans(xml: &str) -> Result<Vec<(usize, usize)>> {
    // Reuse the parser's validated model discovery and then independently
    // rescan the same bounded source for exact byte spans.
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(String, String, usize)> = Vec::new();
    let mut output = Vec::new();
    let mut depth = 0usize;
    loop {
        let event_position = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms XML: {error}")))?;
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref source) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("XForms XML depth overflow".to_string()))?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "XForms XML exceeds {MAX_DEPTH} levels"
                    )));
                }
                let ns = namespace.clone();
                let local = utf8(source.local_name().as_ref(), "XForms element name")?;
                let start = event_position;
                stack.push((ns, local, start));
            },
            Event::Empty(ref source) => {
                let ns = namespace.clone();
                let local = utf8(source.local_name().as_ref(), "XForms element name")?;
                if ns == XFORMSNS
                    && local == "model"
                    && stack.last().is_some_and(|(parent_ns, parent_local, _)| {
                        parent_ns == OFFICENS && parent_local == "forms"
                    })
                {
                    let start = event_position;
                    if output.len() >= MAX_MODELS {
                        return Err(Error::InvalidFormat(format!(
                            "ODF contains more than {MAX_MODELS} XForms models"
                        )));
                    }
                    output.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT XForms model edit spans",
                        source,
                    })?;
                    output.push((start, event_end));
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("XForms XML stack underflow".to_string())
                })?;
                if frame.0 == XFORMSNS
                    && frame.1 == "model"
                    && stack.last().is_some_and(|(parent_ns, parent_local, _)| {
                        parent_ns == OFFICENS && parent_local == "forms"
                    })
                {
                    if output.len() >= MAX_MODELS {
                        return Err(Error::InvalidFormat(format!(
                            "ODF contains more than {MAX_MODELS} XForms models"
                        )));
                    }
                    output.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODT XForms model edit spans",
                        source,
                    })?;
                    output.push((frame.2, event_end));
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms XML depth underflow".to_string())
                })?;
            },
            Event::Eof => break,
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XForms XML".to_string(),
                ));
            },
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete XForms document while locating model spans".to_string(),
        ));
    }
    output.sort_by_key(|(start, _)| *start);
    Ok(output)
}

pub(crate) fn replace_model(xml: &str, position: usize, model: &Model) -> Result<String> {
    let current = parse_models(xml)?;
    let existing = current.get(position).ok_or_else(|| {
        Error::InvalidFormat(format!("XForms model position {position} is out of range"))
    })?;
    model.validate()?;
    if existing == model {
        return Ok(xml.to_owned());
    }
    let spans = model_spans(xml)?;
    let (start, end) = spans.get(position).copied().ok_or_else(|| {
        Error::InvalidFormat(format!("XForms model position {position} is out of range"))
    })?;
    let source = xml
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("invalid XForms model span".to_string()))?;
    if model_has_direct_lexical_material(source)? {
        return Err(Error::InvalidFormat(
            "changed XForms model would discard source comments, processing instructions, or inter-child whitespace"
                .to_string(),
        ));
    }
    let candidate = splice_replace(xml, start, end, &model.to_xml()?)?;
    parse_models(&candidate)?;
    Ok(candidate)
}

pub(crate) fn remove_model(xml: &str, position: usize) -> Result<String> {
    let current = parse_models(xml)?;
    if current.get(position).is_none() {
        return Err(Error::InvalidFormat(format!(
            "XForms model position {position} is out of range"
        )));
    }
    let spans = model_spans(xml)?;
    let (start, end) = spans.get(position).copied().ok_or_else(|| {
        Error::InvalidFormat(format!("XForms model position {position} is out of range"))
    })?;
    let candidate = splice_replace(xml, start, end, "")?;
    parse_models(&candidate)?;
    Ok(candidate)
}

pub(crate) fn insert_model(xml: &str, model: &Model) -> Result<String> {
    parse_models(xml)?;
    model.validate()?;
    let forms = container_span(xml, OFFICENS, "forms")?.ok_or_else(|| {
        Error::InvalidFormat("document has no office:forms container".to_string())
    })?;
    let fragment = model.to_xml()?;
    let source = xml
        .get(forms.0..forms.1)
        .ok_or_else(|| Error::InvalidFormat("invalid office:forms span".to_string()))?;
    if source.trim_end().ends_with("/>") {
        let close = source.len().checked_sub(2).ok_or_else(|| {
            Error::InvalidFormat("invalid empty office:forms closing delimiter".to_string())
        })?;
        let qname_end = source[1..]
            .find(|ch: char| ch.is_ascii_whitespace() || ch == '/')
            .map_or(close - 1, |index| index + 1);
        let replacement_len = source
            .len()
            .checked_add(fragment.len())
            .and_then(|length| length.checked_add(20))
            .ok_or_else(|| Error::InvalidFormat("XForms insertion size overflow".to_string()))?;
        bounded_edit_length(replacement_len, "XForms insertion")?;
        let mut replacement = String::new();
        replacement
            .try_reserve_exact(replacement_len)
            .map_err(|source| Error::Allocation {
                resource: "ODT XForms insertion",
                source,
            })?;
        replacement.push_str(&source[..close]);
        replacement.push('>');
        replacement.push_str(&fragment);
        replacement.push_str("</");
        replacement.push_str(&source[1..qname_end]);
        replacement.push('>');
        let candidate = splice_replace(xml, forms.0, forms.1, &replacement)?;
        parse_models(&candidate)?;
        return Ok(candidate);
    }
    let candidate = splice_insert(xml, forms.1, &fragment)?;
    parse_models(&candidate)?;
    Ok(candidate)
}

fn model_has_direct_lexical_material(xml: &str) -> Result<bool> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms model XML: {error}")))?;
        match event {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms model depth overflow".to_string())
                })?
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms model depth underflow".to_string())
                })?;
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_)
            | Event::Comment(_)
            | Event::PI(_)
                if depth == 1 =>
            {
                return Ok(true);
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in XForms model XML".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::InvalidFormat(
            "incomplete XForms model XML".to_string(),
        ));
    }
    Ok(false)
}

fn container_span(xml: &str, namespace: &str, local: &str) -> Result<Option<(usize, usize)>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(String, String, usize)> = Vec::new();
    let mut found = None;
    let mut depth = 0usize;
    loop {
        let event_position = reader.buffer_position() as usize;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF XML: {error}")))?;
        let resolved = resolved_namespace(&resolved)?.unwrap_or_default();
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref source) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODF XML exceeds {MAX_DEPTH} levels"
                    )));
                }
                let ns = resolved.clone();
                let name = utf8(source.local_name().as_ref(), "container name")?;
                stack.push((ns, name, event_position));
            },
            Event::Empty(ref source) => {
                let ns = resolved.clone();
                let name = utf8(source.local_name().as_ref(), "container name")?;
                if ns == namespace && name == local {
                    let start = event_position;
                    found = Some((start, end));
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| Error::InvalidFormat("XML stack underflow".to_string()))?;
                if found.is_none() && frame.0 == namespace && frame.1 == local {
                    found = Some((frame.2, event_position));
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth underflow".to_string()))?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete XML while locating forms container".to_string(),
        ));
    }
    Ok(found)
}

fn splice_insert(xml: &str, offset: usize, value: &str) -> Result<String> {
    if offset > xml.len() {
        return Err(Error::InvalidFormat(
            "invalid XML insertion offset".to_string(),
        ));
    }
    let output_len = xml
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("XForms insertion size overflow".to_string()))?;
    bounded_edit_length(output_len, "XForms insertion")?;
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms insertion",
            source,
        })?;
    output.push_str(&xml[..offset]);
    output.push_str(value);
    output.push_str(&xml[offset..]);
    Ok(output)
}

fn splice_replace(xml: &str, start: usize, end: usize, value: &str) -> Result<String> {
    if start > end || end > xml.len() {
        return Err(Error::InvalidFormat(
            "invalid XForms replacement span".to_string(),
        ));
    }
    let removed = end
        .checked_sub(start)
        .ok_or_else(|| Error::InvalidFormat("XForms replacement span underflow".to_string()))?;
    let output_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(value.len()))
        .ok_or_else(|| Error::InvalidFormat("XForms replacement size overflow".to_string()))?;
    bounded_edit_length(output_len, "XForms replacement")?;
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms replacement",
            source,
        })?;
    output.push_str(&xml[..start]);
    output.push_str(value);
    output.push_str(&xml[end..]);
    Ok(output)
}

fn bounded_edit_length(length: usize, operation: &str) -> Result<()> {
    if length > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{operation} exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const XML: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:ext="urn:example:extension"><office:body><office:text><office:forms><xforms:model id="m"><xforms:instance id="i"><data xmlns="urn:example:data"><value>one</value></data></xforms:instance><xforms:bind id="b" nodeset="/data/value" type="xsd:string"/><ext:hook ext:value="keep"/></xforms:model></office:forms></office:text></office:body></office:document-content>"#;

    #[test]
    fn parses_known_xforms_children_and_keeps_extensions() {
        let models = parse_models(XML).expect("models parse");
        assert_eq!(models.len(), 1);
        assert_eq!(models[0].id.as_deref(), Some("m"));
        assert_eq!(models[0].instances().count(), 1);
        assert_eq!(
            models[0].binds().next().unwrap().nodeset.as_deref(),
            Some("/data/value")
        );
        assert!(matches!(models[0].children[2], ModelChild::Extension(_)));
        let round_trip = models[0].to_xml().expect("model writes");
        assert!(round_trip.contains("ext:hook"));
        assert!(round_trip.contains("<value>one</value>"));
        assert!(parse_models(&format!("<office:document-content xmlns:office=\"{OFFICENS}\"><office:forms>{round_trip}</office:forms></office:document-content>")).is_ok());
    }

    #[test]
    fn model_replace_uses_exact_source_span() {
        let mut replacement = parse_models(XML).unwrap().remove(0);
        replacement.id = Some("changed".to_string());
        let updated = replace_model(XML, 0, &replacement).expect("replace");
        assert!(updated.contains(" id=\"changed\""));
        assert!(updated.contains("ext:hook"));
    }

    #[test]
    fn equivalent_model_replacement_is_an_exact_noop_and_changed_lexical_source_refuses() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xforms="http://www.w3.org/2002/xforms"><office:body><office:text><office:forms><xforms:model id="m">
<!-- retain this comment -->
<xforms:instance id="i"/>
<?retain yes?>
</xforms:model></office:forms></office:text></office:body></office:document-content>"#;
        let model = parse_models(xml).unwrap().remove(0);
        assert_eq!(replace_model(xml, 0, &model).unwrap(), xml);

        let mut changed = model;
        changed.id = Some("changed".to_string());
        assert!(replace_model(xml, 0, &changed).is_err());
    }

    #[test]
    fn model_ids_and_references_form_a_closed_document_graph() {
        let dangling = XML.replace(
            "<xforms:bind id=\"b\"",
            "<xforms:bind id=\"b\" model=\"missing\"",
        );
        assert!(parse_models(&dangling).is_err());

        let duplicate = XML.replace("<xforms:instance id=\"i\"", "<xforms:instance id=\"m\"");
        assert!(parse_models(&duplicate).is_err());

        let refs = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xforms="http://www.w3.org/2002/xforms"><office:body><office:text><office:forms><xforms:model id="m1"><xforms:bind id="b1" model="m2"/><xforms:submission id="s1" instance="i2"/></xforms:model><xforms:model id="m2"><xforms:instance id="i2"/></xforms:model></office:forms></office:text></office:body></office:document-content>"#;
        assert!(parse_models(refs).is_ok());
        assert!(remove_model(refs, 1).is_err());
    }

    #[test]
    fn accepts_normative_unqualified_xforms_attributes_and_inherited_context() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xforms="http://www.w3.org/2002/xforms"><office:body><office:text><office:forms xmlns:e="urn:example:forms"><xforms:model id="m" note="a &gt; b"><xforms:instance id="i" xmlns:d="urn:example:data"><d:data><d:value>one</d:value></d:data></xforms:instance><xforms:bind id="b" nodeset="/data/value" type="xsd:string"/><e:hook e:value="keep"/></xforms:model></office:forms></office:text></office:body></office:document-content>"#;
        let model = parse_models(xml).unwrap().remove(0);
        assert_eq!(model.id.as_deref(), Some("m"));
        assert!(model.namespace_bindings().iter().any(|binding| {
            binding.prefix.as_deref() == Some("e") && binding.namespace_uri == "urn:example:forms"
        }));
        assert_eq!(
            model.binds().next().unwrap().datatype.as_deref(),
            Some("xsd:string")
        );
        let instance = model.instances().next().unwrap();
        assert_eq!(
            instance.content_xml.as_deref(),
            Some("<d:data><d:value>one</d:value></d:data>")
        );
        assert!(instance.namespace_bindings().iter().any(|binding| {
            binding.prefix.as_deref() == Some("d") && binding.namespace_uri == "urn:example:data"
        }));
        let round_trip = model.to_xml().unwrap();
        assert!(round_trip.contains("xmlns:e=\"urn:example:forms\""));
        assert!(round_trip.contains("xmlns:d=\"urn:example:data\""));
        assert!(!round_trip.contains("xforms:nodeset="));
        assert!(parse_models(&format!(
            "<office:document-content xmlns:office=\"{OFFICENS}\"><office:forms>{round_trip}</office:forms></office:document-content>"
        ))
        .is_ok());
    }

    #[test]
    fn unknown_attribute_namespaces_are_bound_without_empty_namespace_aliases() {
        let mut model = parse_models(XML).unwrap().remove(0);
        model.attributes.push(Attribute {
            namespace_uri: "urn:example:other".to_string(),
            local_name: "flag".to_string(),
            value: "yes".to_string(),
            prefix: None,
        });
        model.attributes.push(Attribute {
            namespace_uri: String::new(),
            local_name: "plain".to_string(),
            value: "yes".to_string(),
            prefix: None,
        });
        let xml = model.to_xml().unwrap();
        assert!(xml.contains("xmlns:ext0=\"urn:example:other\""));
        assert!(xml.contains("ext0:flag=\"yes\""));
        assert!(xml.contains(" plain=\"yes\""));
        assert!(parse_models(&format!(
            "<office:document-content xmlns:office=\"{OFFICENS}\"><office:forms>{xml}</office:forms></office:document-content>"
        ))
        .is_ok());
    }

    #[test]
    fn rejects_incomplete_or_doctype_opaque_fragments() {
        let mut model = Model::default();
        model.children.push(ModelChild::Extension(Extension {
            namespace_uri: "urn:example:ext".to_string(),
            local_name: "hook".to_string(),
            xml: "<e:hook xmlns:e=\"urn:example:ext\">tail".to_string(),
        }));
        assert!(model.to_xml().is_err());

        let mut model = Model::default();
        model.children.push(ModelChild::Instance(Instance {
            content_xml: Some("<!DOCTYPE data><data/>".to_string()),
            ..Instance::default()
        }));
        assert!(model.to_xml().is_err());

        let mut model = Model::default();
        model.children.push(ModelChild::Extension(Extension {
            namespace_uri: "urn:example:ext".to_string(),
            local_name: "hook".to_string(),
            xml: "<e:hook xmlns:e=\"urn:example:ext\">&custom;</e:hook>".to_string(),
        }));
        assert!(model.to_xml().is_err());

        let mut model = Model::default();
        model.children.push(ModelChild::Extension(Extension {
            namespace_uri: "urn:example:ext".to_string(),
            local_name: "hook".to_string(),
            xml: "<e:hook/>".to_string(),
        }));
        assert!(model.to_xml().is_err());

        let mut model = Model::default();
        model.children.push(ModelChild::Instance(Instance {
            content_xml: Some("<d:data/>".to_string()),
            ..Instance::default()
        }));
        assert!(model.to_xml().is_err());
    }

    #[test]
    fn extension_identity_is_checked_against_retained_xml_root() {
        let mut model = Model::default();
        model.children.push(ModelChild::Extension(Extension {
            namespace_uri: "urn:example:ext".to_string(),
            local_name: "hook".to_string(),
            xml: "<e:other xmlns:e=\"urn:example:ext\"/>".to_string(),
        }));
        assert!(model.to_xml().is_err());
    }

    #[test]
    fn rejects_reserved_namespace_shadowing_before_serialization() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xforms="urn:example:wrong"><office:body><office:text><office:forms xmlns:xf="http://www.w3.org/2002/xforms"><xf:model/></office:forms></office:text></office:body></office:document-content>"#;
        assert!(parse_models(xml).is_err());
    }
}
