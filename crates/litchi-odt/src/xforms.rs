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

use crate::core::ResolvedReader;
use crate::generic::{ChargedXml, FlatMutationBudget, MemoryLease, allocate_xml};
use crate::namespace::{FORMNS, OFFICENS, XFORMSNS, XMLNS};
use litchi_core::{Error, Resource, ResourceLimit, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::Reader;
use std::collections::HashSet;
use std::mem::size_of;

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
        self.validate_with_budget(None)
    }

    fn validate_with_budget(&self, budget: Option<&FlatMutationBudget>) -> Result<()> {
        if let Some(budget) = budget {
            budget.check()?;
            budget.consume_objects(1)?;
        }
        let output_len = model_output_upper_bound(self)?;
        bounded_edit_length(output_len, "XForms model")?;
        validate_optional("xforms:model id", self.id.as_deref())?;
        validate_reserved_namespace_bindings(&self.namespace_declarations, "xforms:model", budget)?;
        validate_attributes_with_budget(&self.attributes, budget)?;
        if self.children.len() > MAX_CHILDREN {
            return Err(Error::InvalidFormat(format!(
                "xforms:model exceeds {MAX_CHILDREN} direct children"
            )));
        }
        for child in &self.children {
            if let Some(budget) = budget {
                budget.check()?;
                budget.consume_objects(1)?;
            }
            match child {
                ModelChild::Instance(value) => value.validate_with_budget(budget)?,
                ModelChild::Bind(value) => value.validate_with_budget(budget)?,
                ModelChild::Submission(value) => value.validate_with_budget(budget)?,
                ModelChild::Extension(value) => {
                    validate_extension_with_budget(value, &self.namespace_declarations, budget)?;
                },
            }
        }
        Ok(())
    }

    /// Serialize a model for insertion or replacement.
    pub fn to_xml(&self) -> Result<String> {
        self.to_xml_with_limit(MAX_XML_BYTES)
    }

    /// Serialize a model after charging its exact encoded length to a caller limit.
    pub(crate) fn to_xml_with_limit(&self, maximum: usize) -> Result<String> {
        self.to_xml_with_limit_and_budget(maximum, None)
            .map(ChargedXml::into_string)
    }

    pub(crate) fn to_xml_with_limit_and_budget(
        &self,
        maximum: usize,
        budget: Option<&FlatMutationBudget>,
    ) -> Result<ChargedXml> {
        // Charge every temporary namespace scope and prefix projection before
        // the length writer runs.  `LengthOutput` follows the same writer
        // path and therefore allocates its temporary scope vectors too.
        let _scratch_memory = budget
            .map(|budget| {
                model_serialization_scratch_bytes(self, Some(budget)).and_then(|amount| {
                    budget.reserve_bytes(amount, "ODT XForms model serialization scratch")
                })
            })
            .transpose()?;
        // Measure before validation and before constructing the output.  The
        // measured form follows the same namespace-prefix choices as the
        // writer below, so a caller cap is charged against the exact result,
        // rather than the older conservative upper bound.
        let output_len = model_output_len(self, maximum, budget)?;
        self.validate_with_budget(budget)?;
        bounded_edit_length_with_limit(output_len, "XForms model serialization", maximum)?;
        let (mut output, memory) =
            allocate_xml(budget, output_len, "ODT XForms model serialization")?;
        if let Some(budget) = budget {
            let mut checked = BudgetedXmlOutput {
                output: &mut output,
                budget,
            };
            write_model(&mut checked, self)?;
        } else {
            write_model(&mut output, self)?;
        }
        debug_assert_eq!(output.len(), output_len);
        Ok(ChargedXml {
            xml: output,
            memory,
        })
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
        self.validate_with_budget(None)
    }

    fn validate_with_budget(&self, budget: Option<&FlatMutationBudget>) -> Result<()> {
        validate_reserved_namespace_bindings(
            &self.namespace_declarations,
            "xforms:instance",
            budget,
        )?;
        for (name, value) in [
            ("xforms:instance id", self.id.as_deref()),
            ("xforms:instance src", self.src.as_deref()),
            ("xforms:instance resource", self.resource.as_deref()),
            ("xforms:instance mediatype", self.mediatype.as_deref()),
            ("xforms:instance content", self.content_xml.as_deref()),
        ] {
            if let Some(budget) = budget {
                budget.check()?;
            }
            validate_optional(name, value)?;
        }
        if let Some(content) = &self.content_xml {
            validate_fragment_with_context_with_budget(
                content,
                "XForms instance",
                true,
                &self.namespace_declarations,
                budget,
            )?;
        }
        validate_attributes_with_budget(&self.attributes, budget)
    }
}

impl Bind {
    /// Validate lexical values and bounded attributes.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_budget(None)
    }

    fn validate_with_budget(&self, budget: Option<&FlatMutationBudget>) -> Result<()> {
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
            if let Some(budget) = budget {
                budget.check()?;
            }
            validate_optional(name, value)?;
        }
        validate_attributes_with_budget(&self.attributes, budget)
    }
}

impl Submission {
    /// Validate lexical values and bounded attributes.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_budget(None)
    }

    fn validate_with_budget(&self, budget: Option<&FlatMutationBudget>) -> Result<()> {
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
            if let Some(budget) = budget {
                budget.check()?;
            }
            validate_optional(name, value)?;
        }
        if let Some(budget) = budget {
            budget.check()?;
            budget.consume_objects(1)?;
        }
        if let Some(method) = &self.method {
            validate_xml_value("xforms:submission method", method.as_str())?;
        }
        if let Some(budget) = budget {
            budget.check()?;
        }
        if let Some(replace) = &self.replace {
            validate_xml_value("xforms:submission replace", replace.as_str())?;
        }
        validate_attributes_with_budget(&self.attributes, budget)
    }
}

/// Parse all `xforms:model` elements directly below `office:forms`.
pub(crate) fn parse_models(xml: &str) -> Result<Vec<Model>> {
    parse_models_inner(xml, None)
}

fn parse_models_inner(xml: &str, budget: Option<&FlatMutationBudget>) -> Result<Vec<Model>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} XForms model limit"
        )));
    }
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Frame> = Vec::new();
    let mut spans: Vec<(usize, usize, Vec<String>)> = Vec::new();
    let mut namespace_scope = Vec::new();
    let mut depth = 0usize;
    loop {
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
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
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
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
                    Some(namespace_scope_to_raw(&namespace_scope)?)
                } else {
                    None
                };
                stack
                    .try_reserve_exact(1)
                    .map_err(|source| Error::Allocation {
                        resource: "ODT XForms parser frame stack",
                        source,
                    })?;
                stack.push(Frame {
                    namespace,
                    local,
                    start,
                    changes,
                    model_declarations,
                });
            },
            Event::Empty(ref source) => {
                if let Some(budget) = budget {
                    budget.observe_depth(depth.saturating_add(1))?;
                }
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
                    spans
                        .try_reserve_exact(1)
                        .map_err(|source| Error::Allocation {
                            resource: "ODT XForms model spans",
                            source,
                        })?;
                    spans.push((start, event_end, namespace_scope_to_raw(&namespace_scope)?));
                }
                restore_namespace_declarations(&mut namespace_scope, changes)?;
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
                    spans
                        .try_reserve_exact(1)
                        .map_err(|source| Error::Allocation {
                            resource: "ODT XForms model spans",
                            source,
                        })?;
                    spans.push((
                        frame.start,
                        end,
                        frame
                            .model_declarations
                            .unwrap_or(namespace_scope_to_raw(&namespace_scope)?),
                    ));
                }
                restore_namespace_declarations(&mut namespace_scope, frame.changes)?;
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
        if let Some(budget) = budget {
            budget.check()?;
        }
        let owned = match budget {
            Some(budget) => inject_namespace_declarations_with_budget(raw, &declarations, budget)?,
            None => ChargedXml {
                xml: inject_namespace_declarations(raw, &declarations)?,
                memory: None,
            },
        };
        output.push(parse_model(&owned.xml, &declarations, budget)?);
    }
    validate_model_set(&output, budget)?;
    Ok(output)
}

/// Parse models while charging the parser's owned projection before it starts.
///
/// The ordinary parser deliberately owns strings and vectors because the
/// public model outlives the XML event stream.  A flat mutation therefore
/// cannot charge `xml.len()` and call that parser memory: the projection also
/// owns attribute values, namespace scopes, opaque instance/extension spans,
/// vector slots, and validation sets.  This plan walks the same borrowed
/// events without allocating and counts each of those categories.  The
/// reservation is retained until the returned projection is dropped by the
/// caller.  It is an upper bound for allocator capacities, rather than a
/// document-size or output-byte limit.
pub(crate) fn parse_models_with_budget(
    xml: &str,
    budget: &FlatMutationBudget,
) -> Result<(Vec<Model>, MemoryLease)> {
    budget.check()?;
    let amount = xforms_parser_memory_plan(xml, Some(budget))?;
    let reservation = budget.reserve_bytes(amount, "ODT XForms model parser projection")?;
    match parse_models_inner(xml, Some(budget)) {
        Ok(models) => Ok((models, MemoryLease::new(reservation))),
        Err(error) => {
            drop(reservation);
            Err(error)
        },
    }
}

#[derive(Default)]
struct XFormsMemoryPlan {
    event_count: usize,
    max_event_bytes: usize,
    element_count: usize,
    attribute_count: usize,
    attribute_bytes: usize,
    namespace_count: usize,
    namespace_bytes: usize,
    model_count: usize,
    model_bytes: usize,
    child_count: usize,
    id_count: usize,
}

impl XFormsMemoryPlan {
    fn total(self) -> Result<usize> {
        // Every parser vector grows with the previous allocation live until
        // the replacement allocation succeeds.  The codecs use
        // `try_reserve_exact`, so charging both the retained destination and
        // its predecessor gives an explicit old-plus-new peak for each
        // independent vector.  This is parser scratch accounting and remains
        // separate from the final XML byte limit.
        let span_slots = self
            .model_count
            .checked_mul(size_of::<(usize, usize, Vec<String>)>())
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| Error::InvalidFormat("XForms span vector size overflow".to_string()))?;
        let frame_slots = self
            .element_count
            .checked_mul(size_of::<Frame>())
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| Error::InvalidFormat("XForms frame vector size overflow".to_string()))?;
        let attribute_slots = self
            .attribute_count
            .checked_mul(size_of::<Attribute>())
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| {
                Error::InvalidFormat("XForms attribute vector size overflow".to_string())
            })?;
        let model_slots = self
            .model_count
            .checked_mul(size_of::<Model>())
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| Error::InvalidFormat("XForms model vector size overflow".to_string()))?;
        let child_slots = self
            .child_count
            .checked_mul(size_of::<ModelChild>())
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| Error::InvalidFormat("XForms child vector size overflow".to_string()))?;
        let vector_slots = frame_slots
            .checked_add(attribute_slots)
            .and_then(|value| value.checked_add(model_slots))
            .and_then(|value| value.checked_add(child_slots))
            .and_then(|value| value.checked_add(span_slots))
            .ok_or_else(|| {
                Error::InvalidFormat("XForms parser vector size overflow".to_string())
            })?;
        let event_buffer = self.max_event_bytes.checked_mul(2).ok_or_else(|| {
            Error::InvalidFormat("XForms parser buffer size overflow".to_string())
        })?;
        let repeated_scope_count = self
            .model_count
            .checked_add(self.child_count)
            .and_then(|value| value.checked_add(2))
            .ok_or_else(|| Error::InvalidFormat("XForms scope repetition overflow".to_string()))?;
        let strings = self
            .attribute_bytes
            .checked_mul(2)
            // Namespace declarations inherited from an ancestor are copied
            // into every direct model's retained namespace scope and can be
            // injected into every model fragment.  Count that repeated
            // ownership explicitly; the source XML contains each declaration
            // only once, so a single XML-length reservation would undercharge
            // namespace-heavy model sets.
            .and_then(|value| {
                self.namespace_bytes
                    .checked_mul(repeated_scope_count)
                    .and_then(|amount| value.checked_add(amount))
            })
            // Each model span is copied into an injected parser fragment
            // before its owned instance/extension strings are retained. Keep
            // both the retained projection and that current fragment in the
            // peak rather than charging only the final model bytes.
            .and_then(|value| {
                self.model_bytes
                    .checked_mul(2)
                    .and_then(|amount| value.checked_add(amount))
            })
            .ok_or_else(|| {
                Error::InvalidFormat("XForms parser string size overflow".to_string())
            })?;
        // HashSet uses bucket storage in addition to the retained strings.
        // Charge a complete bucket-sized slot per possible id and attribute
        // key.  This is tied to observed entries and does not scale with the
        // document cap when the XML is small.
        let validation_sets = self
            .id_count
            .checked_mul(3 * (size_of::<usize>() * 4 + size_of::<&str>()))
            .and_then(|value| {
                value.checked_add(self.attribute_count.checked_mul(size_of::<usize>() * 4)?)
            })
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| Error::InvalidFormat("XForms parser set size overflow".to_string()))?;
        // Namespace scopes are copied into each retained model/child and
        // into temporary injected fragments. Count both the String slots and
        // the scope-change tuples; charging only their byte payloads misses
        // the nested vectors themselves.
        let namespace_scope_slots = self
            .namespace_count
            .checked_mul(repeated_scope_count)
            .and_then(|count| count.checked_mul(size_of::<String>()))
            .and_then(|value| {
                self.namespace_count
                    .checked_mul(size_of::<(String, Option<String>)>())
                    .and_then(|amount| value.checked_add(amount))
            })
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| Error::InvalidFormat("XForms namespace slot overflow".to_string()))?;
        // Namespace-presence and prefix-selection helpers use hash sets while
        // injecting each model. Their bucket storage is separate from the
        // retained XML strings and must be admitted before those helpers run.
        let namespace_helper_sets = self
            .namespace_count
            .checked_mul(self.model_count.saturating_add(1))
            .and_then(|count| count.checked_mul(size_of::<String>() * 4))
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| Error::InvalidFormat("XForms namespace set overflow".to_string()))?;
        vector_slots
            .checked_add(event_buffer)
            .and_then(|value| value.checked_add(strings))
            .and_then(|value| value.checked_add(namespace_scope_slots))
            .and_then(|value| value.checked_add(namespace_helper_sets))
            .and_then(|value| value.checked_add(validation_sets))
            .and_then(|value| {
                self.event_count
                    .checked_mul(size_of::<usize>())
                    .and_then(|value| value.checked_mul(2))
                    .and_then(|amount| value.checked_add(amount))
            })
            .and_then(|value| {
                self.namespace_count
                    .checked_mul(size_of::<(String, String)>())
                    .and_then(|amount| amount.checked_mul(2))
                    .and_then(|amount| value.checked_add(amount))
            })
            .ok_or_else(|| Error::InvalidFormat("XForms parser memory plan overflow".to_string()))
    }
}

fn xforms_parser_memory_plan(xml: &str, budget: Option<&FlatMutationBudget>) -> Result<usize> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} XForms model limit"
        )));
    }
    let mut plan = XFormsMemoryPlan::default();
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut forms_depth = None;
    let mut model_depth = None;
    let mut model_start = None;
    loop {
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
        let event_start = reader.buffer_position() as usize;
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms XML: {error}")))?;
        let (namespace_len, is_office, is_xforms) = match &resolved {
            ResolveResult::Bound(value) => (
                value.as_ref().len(),
                value.as_ref() == OFFICENS.as_bytes(),
                value.as_ref() == XFORMSNS.as_bytes(),
            ),
            ResolveResult::Unbound => (0, false, false),
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unbound namespace prefix '{}'",
                    String::from_utf8_lossy(prefix)
                )));
            },
        };
        drop(resolved);
        let event_end = reader.buffer_position() as usize;
        plan.event_count = plan
            .event_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("XForms event count overflow".to_string()))?;
        plan.max_event_bytes = plan
            .max_event_bytes
            .max(event_end.saturating_sub(event_start));
        match event {
            Event::Start(source) => {
                let local = source.local_name();
                plan.element_count = plan.element_count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms element count overflow".to_string())
                })?;
                plan.attribute_bytes = plan
                    .attribute_bytes
                    .checked_add(namespace_len)
                    .and_then(|value| value.checked_add(local.as_ref().len()))
                    .ok_or_else(|| Error::InvalidFormat("XForms name size overflow".to_string()))?;
                count_xforms_attributes(&mut plan, &source)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("XForms XML depth overflow".to_string()))?;
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
                let is_forms = is_office && local.as_ref() == b"forms";
                if is_forms {
                    forms_depth = Some(depth);
                }
                let is_model = is_xforms
                    && local.as_ref() == b"model"
                    && forms_depth == Some(depth.saturating_sub(1));
                if is_model {
                    model_depth = Some(depth);
                    model_start = Some(event_start);
                    plan.model_count = plan.model_count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("XForms model count overflow".to_string())
                    })?;
                }
            },
            Event::Empty(source) => {
                if let Some(budget) = budget {
                    budget.observe_depth(depth.saturating_add(1))?;
                }
                let local = source.local_name();
                plan.element_count = plan.element_count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms element count overflow".to_string())
                })?;
                plan.attribute_bytes = plan
                    .attribute_bytes
                    .checked_add(namespace_len)
                    .and_then(|value| value.checked_add(local.as_ref().len()))
                    .ok_or_else(|| Error::InvalidFormat("XForms name size overflow".to_string()))?;
                count_xforms_attributes(&mut plan, &source)?;
                if is_xforms && local.as_ref() == b"model" && forms_depth.is_some() {
                    plan.model_count = plan.model_count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("XForms model count overflow".to_string())
                    })?;
                    plan.model_bytes = plan
                        .model_bytes
                        .checked_add(event_end.saturating_sub(event_start))
                        .ok_or_else(|| {
                            Error::InvalidFormat("XForms model span overflow".to_string())
                        })?;
                }
            },
            Event::End(_) => {
                if model_depth == Some(depth) {
                    plan.model_bytes = plan
                        .model_bytes
                        .checked_add(
                            event_end
                                .checked_sub(model_start.unwrap_or(event_end))
                                .unwrap_or_default(),
                        )
                        .ok_or_else(|| {
                            Error::InvalidFormat("XForms model span overflow".to_string())
                        })?;
                    model_depth = None;
                    model_start = None;
                }
                if forms_depth == Some(depth) {
                    forms_depth = None;
                }
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
    }
    plan.child_count = plan.element_count.saturating_sub(plan.model_count);
    plan.id_count = plan.attribute_count;
    plan.total()
}

fn count_xforms_attributes(plan: &mut XFormsMemoryPlan, source: &BytesStart<'_>) -> Result<()> {
    for attribute in source.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms attribute: {error}")))?;
        let key = attribute.key.as_ref();
        plan.attribute_count = plan
            .attribute_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("XForms attribute count overflow".to_string()))?;
        plan.attribute_bytes = plan
            .attribute_bytes
            .checked_add(key.len())
            .and_then(|value| value.checked_add(attribute.value.len()))
            .ok_or_else(|| Error::InvalidFormat("XForms attribute size overflow".to_string()))?;
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            plan.namespace_count = plan.namespace_count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("XForms namespace count overflow".to_string())
            })?;
            plan.namespace_bytes = plan
                .namespace_bytes
                .checked_add(key.len())
                .and_then(|value| value.checked_add(attribute.value.len()))
                .and_then(|value| value.checked_add(4))
                .ok_or_else(|| {
                    Error::InvalidFormat("XForms namespace size overflow".to_string())
                })?;
        }
    }
    Ok(())
}

fn validate_model_set(models: &[Model], budget: Option<&FlatMutationBudget>) -> Result<()> {
    let mut ids = HashSet::new();
    let mut model_ids = HashSet::new();
    let mut instance_ids = HashSet::new();
    for model in models {
        if let Some(budget) = budget {
            budget.check()?;
            budget.consume_objects(1)?;
        }
        if let Some(id) = &model.id {
            insert_model_id(&mut ids, &mut model_ids, id, "xforms:model")?;
        }
        for child in &model.children {
            if let Some(budget) = budget {
                budget.check()?;
                budget.consume_objects(1)?;
            }
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
        if let Some(budget) = budget {
            budget.check()?;
        }
        for child in &model.children {
            if let Some(budget) = budget {
                budget.check()?;
                budget.consume_objects(1)?;
            }
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
    ids.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODT XForms validation ids",
        source,
    })?;
    if !ids.insert(id) {
        return Err(Error::InvalidFormat(format!(
            "duplicate XForms id '{id}' on {element}"
        )));
    }
    family.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODT XForms validation model ids",
        source,
    })?;
    family.insert(id);
    Ok(())
}

fn insert_unique_id<'a>(ids: &mut HashSet<&'a str>, id: &'a str, element: &str) -> Result<()> {
    ids.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODT XForms validation ids",
        source,
    })?;
    if !ids.insert(id) {
        return Err(Error::InvalidFormat(format!(
            "duplicate XForms id '{id}' on {element}"
        )));
    }
    Ok(())
}

fn parse_model(
    xml: &str,
    inherited_declarations: &[String],
    budget: Option<&FlatMutationBudget>,
) -> Result<Model> {
    if let Some(budget) = budget {
        budget.check()?;
    }
    let mut reader = ResolvedReader::from_xml(xml);
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
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
        let event_position = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms model XML: {error}")))?;
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref source) => {
                depth += 1;
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
                if depth == 1 {
                    parse_model_attributes(&reader, source, &mut model)?;
                    model.namespace_declarations = merge_namespace_declarations(
                        inherited_declarations,
                        &namespace_declarations(source)?,
                    )?;
                }
                let ns = namespace.clone();
                let local = utf8(source.local_name().as_ref(), "XForms model element name")?;
                let start = event_position;
                let attrs = attributes(&reader, source)?;
                stack
                    .try_reserve_exact(1)
                    .map_err(|source| Error::Allocation {
                        resource: "ODT XForms model frame stack",
                        source,
                    })?;
                stack.push((ns, local, start, attrs));
            },
            Event::Empty(ref source) => {
                if let Some(budget) = budget {
                    budget.observe_depth(depth.saturating_add(1))?;
                }
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
                    )?;
                } else if depth == 1 {
                    let raw = xml.get(start..event_end).ok_or_else(|| {
                        Error::InvalidFormat("invalid XForms child span".to_string())
                    })?;
                    children
                        .try_reserve_exact(1)
                        .map_err(|source| Error::Allocation {
                            resource: "ODT XForms model children",
                            source,
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
                    children
                        .try_reserve_exact(1)
                        .map_err(|source| Error::Allocation {
                            resource: "ODT XForms model children",
                            source,
                        })?;
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
    model.validate_with_budget(budget)?;
    Ok(model)
}

fn parse_model_attributes(
    reader: &ResolvedReader<'_>,
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
    reader: &ResolvedReader<'_>,
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
        )?,
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

trait XmlOutput {
    fn push_str(&mut self, value: &str) -> Result<()>;

    fn push_char(&mut self, value: char) -> Result<()>;
}

impl XmlOutput for String {
    fn push_str(&mut self, value: &str) -> Result<()> {
        String::push_str(self, value);
        Ok(())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        String::push(self, value);
        Ok(())
    }
}

struct BudgetedXmlOutput<'a> {
    output: &'a mut String,
    budget: &'a FlatMutationBudget,
}

impl XmlOutput for BudgetedXmlOutput<'_> {
    fn push_str(&mut self, value: &str) -> Result<()> {
        self.budget.check()?;
        append_checked(self.output, value, Some(self.budget))
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        self.budget.check()?;
        self.output.push(value);
        Ok(())
    }
}

struct LengthOutput<'a> {
    length: usize,
    maximum: usize,
    budget: Option<&'a FlatMutationBudget>,
}

impl<'a> LengthOutput<'a> {
    fn new(maximum: usize, budget: Option<&'a FlatMutationBudget>) -> Self {
        Self {
            length: 0,
            maximum,
            budget,
        }
    }

    fn add(&mut self, amount: usize) -> Result<()> {
        if let Some(budget) = self.budget {
            budget.check()?;
        }
        let length = self.length.checked_add(amount).ok_or_else(|| {
            Error::InvalidFormat("XForms serialization size overflow".to_string())
        })?;
        if length > self.maximum {
            return Err(bounded_edit_length_with_limit(
                length,
                "XForms model serialization",
                self.maximum,
            )
            .expect_err("length above maximum must return an error"));
        }
        self.length = length;
        Ok(())
    }
}

impl XmlOutput for LengthOutput<'_> {
    fn push_str(&mut self, value: &str) -> Result<()> {
        self.add(value.len())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        self.add(value.len_utf8())
    }
}

fn write_model<O: XmlOutput>(output: &mut O, model: &Model) -> Result<()> {
    output.push_str("<xforms:model xmlns:xforms=\"")?;
    output.push_str(XFORMSNS)?;
    output.push_str("\" xmlns:office=\"")?;
    output.push_str(OFFICENS)?;
    output.push_str("\" xmlns:form=\"")?;
    output.push_str(FORMNS)?;
    output.push_char('"')?;
    let mut inherited = Vec::new();
    inherited
        .try_reserve_exact(3)
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms model namespace scope",
            source,
        })?;
    inherited.push(("xmlns:xforms".to_string(), XFORMSNS.to_string()));
    inherited.push(("xmlns:office".to_string(), OFFICENS.to_string()));
    inherited.push(("xmlns:form".to_string(), FORMNS.to_string()));
    push_namespace_declarations(output, &model.namespace_declarations, &mut inherited)?;
    push_optional_attr(output, "id", model.id.as_deref())?;
    push_unknown_attributes(output, &model.attributes, &mut inherited)?;
    if model.children.is_empty() {
        output.push_str("/>")?;
    } else {
        output.push_char('>')?;
        for child in &model.children {
            write_child(output, child, &inherited)?;
        }
        output.push_str("</xforms:model>")?;
    }
    Ok(())
}

fn clone_namespace_scope(scope: &[(String, String)]) -> Result<Vec<(String, String)>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(scope.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms model namespace scope",
            source,
        })?;
    output.extend(scope.iter().cloned());
    Ok(output)
}

fn write_child<O: XmlOutput>(
    output: &mut O,
    child: &ModelChild,
    inherited: &[(String, String)],
) -> Result<()> {
    match child {
        ModelChild::Extension(value) => output.push_str(&value.xml)?,
        ModelChild::Instance(value) => {
            output.push_str("<xforms:instance")?;
            let mut child_namespace = clone_namespace_scope(inherited)?;
            push_namespace_declarations(
                output,
                &value.namespace_declarations,
                &mut child_namespace,
            )?;
            push_optional_attr(output, "id", value.id.as_deref())?;
            push_optional_attr(output, "src", value.src.as_deref())?;
            push_optional_attr(output, "resource", value.resource.as_deref())?;
            push_optional_attr(output, "mediatype", value.mediatype.as_deref())?;
            push_unknown_attributes(output, &value.attributes, &mut child_namespace)?;
            if let Some(content) = &value.content_xml {
                output.push_char('>')?;
                output.push_str(content)?;
                output.push_str("</xforms:instance>")?;
            } else {
                output.push_str("/>")?;
            }
        },
        ModelChild::Bind(value) => {
            output.push_str("<xforms:bind")?;
            push_optional_attr(output, "id", value.id.as_deref())?;
            push_optional_attr(output, "nodeset", value.nodeset.as_deref())?;
            push_optional_attr(output, "ref", value.reference.as_deref())?;
            push_optional_attr(output, "context", value.context.as_deref())?;
            push_optional_attr(output, "model", value.model.as_deref())?;
            push_optional_attr(output, "type", value.datatype.as_deref())?;
            push_optional_attr(output, "readonly", value.readonly.as_deref())?;
            push_optional_attr(output, "relevant", value.relevant.as_deref())?;
            push_optional_attr(output, "required", value.required.as_deref())?;
            push_optional_attr(output, "constraint", value.constraint.as_deref())?;
            push_optional_attr(output, "calculate", value.calculate.as_deref())?;
            let mut child_namespace = clone_namespace_scope(inherited)?;
            push_unknown_attributes(output, &value.attributes, &mut child_namespace)?;
            output.push_str("/>")?;
        },
        ModelChild::Submission(value) => {
            output.push_str("<xforms:submission")?;
            push_optional_attr(output, "id", value.id.as_deref())?;
            push_optional_attr(output, "action", value.action.as_deref())?;
            push_optional_attr(
                output,
                "method",
                value.method.as_ref().map(SubmissionMethod::as_str),
            )?;
            push_optional_attr(output, "version", value.version.as_deref())?;
            push_optional_attr(output, "mediatype", value.mediatype.as_deref())?;
            push_optional_attr(output, "encoding", value.encoding.as_deref())?;
            push_optional_attr(
                output,
                "replace",
                value.replace.as_ref().map(SubmissionReplace::as_str),
            )?;
            push_optional_attr(output, "instance", value.instance.as_deref())?;
            push_optional_attr(output, "validate", value.validate.as_deref())?;
            push_optional_attr(output, "relevant", value.relevant.as_deref())?;
            let mut child_namespace = clone_namespace_scope(inherited)?;
            push_unknown_attributes(output, &value.attributes, &mut child_namespace)?;
            output.push_str("/>")?;
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

fn model_output_len(
    model: &Model,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<usize> {
    let mut output = LengthOutput::new(maximum, budget);
    write_model(&mut output, model)?;
    Ok(output.length)
}

fn model_serialization_scratch_bytes(
    model: &Model,
    budget: Option<&FlatMutationBudget>,
) -> Result<usize> {
    let mut scope_entries = 3usize;
    let mut scope_strings = XFORMSNS
        .len()
        .checked_add(OFFICENS.len())
        .and_then(|value| value.checked_add(FORMNS.len()))
        .and_then(|value| value.checked_add("xmlns:xforms".len()))
        .and_then(|value| value.checked_add("xmlns:office".len()))
        .and_then(|value| value.checked_add("xmlns:form".len()))
        .ok_or_else(|| Error::InvalidFormat("XForms serialization scratch overflow".to_string()))?;
    for declaration in &model.namespace_declarations {
        if let Some(budget) = budget {
            budget.check()?;
        }
        scope_entries = scope_entries
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("XForms scope entry overflow".to_string()))?;
        scope_strings = scope_strings
            .checked_add(declaration.len())
            .ok_or_else(|| Error::InvalidFormat("XForms scope string overflow".to_string()))?;
    }
    let mut attribute_count = model.attributes.len();
    for attribute in &model.attributes {
        if let Some(budget) = budget {
            budget.check()?;
        }
        scope_strings = scope_strings
            .checked_add(attribute.namespace_uri.len())
            .and_then(|value| value.checked_add(attribute.prefix.as_deref().map_or(0, str::len)))
            .and_then(|value| value.checked_add(16))
            .ok_or_else(|| Error::InvalidFormat("XForms attribute scratch overflow".to_string()))?;
    }
    let mut child_scope_count = 0usize;
    for child in &model.children {
        if let Some(budget) = budget {
            budget.check()?;
        }
        let attributes = match child {
            ModelChild::Instance(value) => &value.attributes,
            ModelChild::Bind(value) => &value.attributes,
            ModelChild::Submission(value) => &value.attributes,
            ModelChild::Extension(_) => continue,
        };
        child_scope_count = child_scope_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("XForms child scope overflow".to_string()))?;
        attribute_count = attribute_count
            .checked_add(attributes.len())
            .ok_or_else(|| Error::InvalidFormat("XForms attribute count overflow".to_string()))?;
        for attribute in attributes {
            if let Some(budget) = budget {
                budget.check()?;
            }
            scope_strings = scope_strings
                .checked_add(
                    attribute
                        .namespace_uri
                        .len()
                        .checked_mul(2)
                        .and_then(|value| {
                            value.checked_add(attribute.prefix.as_deref().map_or(0, str::len))
                        })
                        .and_then(|value| value.checked_add(16))
                        .ok_or_else(|| {
                            Error::InvalidFormat("XForms child scratch overflow".to_string())
                        })?,
                )
                .ok_or_else(|| Error::InvalidFormat("XForms child scratch overflow".to_string()))?;
        }
    }
    scope_entries = scope_entries
        .checked_add(attribute_count)
        .ok_or_else(|| Error::InvalidFormat("XForms scope entry overflow".to_string()))?;
    let scope_slot_bytes = scope_entries
        .checked_mul(size_of::<(String, String)>())
        .ok_or_else(|| Error::InvalidFormat("XForms scope slot overflow".to_string()))?;
    let child_scope_bytes = child_scope_count
        .checked_mul(scope_slot_bytes)
        .and_then(|value| value.checked_add(child_scope_count.checked_mul(scope_strings)?))
        .ok_or_else(|| Error::InvalidFormat("XForms child scope overflow".to_string()))?;
    let dynamic_prefix_bytes = attribute_count
        .checked_mul(size_of::<(String, String)>() + 16)
        .ok_or_else(|| Error::InvalidFormat("XForms dynamic prefix overflow".to_string()))?;
    // Validation of inline instances and extension fragments injects the
    // retained model namespace context before parsing the fragment.  That
    // path owns a temporary root-namespace set, a copied fragment, and a
    // parser buffer; admit those allocations before validation begins.
    let mut validation_scratch = 0usize;
    for child in &model.children {
        let (raw, declarations) = match child {
            ModelChild::Instance(value) => match value.content_xml.as_deref() {
                Some(raw) => (raw, &value.namespace_declarations),
                None => continue,
            },
            ModelChild::Extension(value) => (value.xml.as_str(), &model.namespace_declarations),
            ModelChild::Bind(_) | ModelChild::Submission(_) => continue,
        };
        let declaration_bytes = declarations.iter().try_fold(0usize, |total, value| {
            total.checked_add(value.len()).ok_or_else(|| {
                Error::InvalidFormat("XForms validation scratch overflow".to_string())
            })
        })?;
        let fragment = raw.len().checked_add(declaration_bytes).ok_or_else(|| {
            Error::InvalidFormat("XForms validation scratch overflow".to_string())
        })?;
        let namespace_set = declarations
            .len()
            .checked_mul(size_of::<String>() * 4)
            .ok_or_else(|| {
                Error::InvalidFormat("XForms validation namespace set overflow".to_string())
            })?;
        validation_scratch = validation_scratch
            .checked_add(
                fragment
                    .checked_mul(2)
                    .and_then(|value| value.checked_add(namespace_set))
                    .ok_or_else(|| {
                        Error::InvalidFormat("XForms validation scratch overflow".to_string())
                    })?,
            )
            .ok_or_else(|| {
                Error::InvalidFormat("XForms validation scratch overflow".to_string())
            })?;
    }
    scope_slot_bytes
        .checked_add(scope_strings)
        .and_then(|value| value.checked_add(child_scope_bytes))
        .and_then(|value| value.checked_add(dynamic_prefix_bytes))
        .and_then(|value| value.checked_add(validation_scratch))
        .ok_or_else(|| Error::InvalidFormat("XForms serialization scratch overflow".to_string()))
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

fn attributes(reader: &ResolvedReader<'_>, source: &BytesStart<'_>) -> Result<Vec<Attribute>> {
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
        output
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
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
            output
                .try_reserve_exact(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODT XForms namespace declarations",
                    source,
                })?;
            output.push(format!(" {key}=\"{value}\""));
        }
    }
    Ok(output)
}

fn merge_namespace_declarations(base: &[String], local: &[String]) -> Result<Vec<String>> {
    let mut scope = Vec::new();
    scope
        .try_reserve_exact(base.len().saturating_add(local.len()))
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms namespace merge scope",
            source,
        })?;
    for declaration in base {
        if let Some((name, value)) = declaration_parts(declaration) {
            scope.push((name.to_string(), value.to_string()));
        }
    }
    for declaration in local {
        if let Some((name, value)) = declaration_parts(declaration) {
            replace_namespace(&mut scope, name, value)?;
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
        changes
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT XForms namespace changes",
                source,
            })?;
        changes.push((name.to_string(), previous));
        replace_namespace(scope, name, value)?;
    }
    Ok(changes)
}

fn restore_namespace_declarations(
    scope: &mut Vec<(String, String)>,
    changes: Vec<(String, Option<String>)>,
) -> Result<()> {
    for (name, previous) in changes.into_iter().rev() {
        match previous {
            Some(value) => replace_namespace(scope, &name, &value)?,
            None => scope.retain(|(current, _)| current != &name),
        }
    }
    Ok(())
}

fn replace_namespace(scope: &mut Vec<(String, String)>, name: &str, value: &str) -> Result<()> {
    if let Some((_, current)) = scope.iter_mut().find(|(current, _)| current == name) {
        *current = value.to_string();
    } else {
        scope
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT XForms namespace scope",
                source,
            })?;
        scope.push((name.to_string(), value.to_string()));
    }
    Ok(())
}

fn namespace_scope_to_raw(scope: &[(String, String)]) -> Result<Vec<String>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(scope.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT XForms namespace declarations",
            source,
        })?;
    for (name, value) in scope {
        let length = 1usize
            .checked_add(name.len())
            .and_then(|value_length| value_length.checked_add(3))
            .and_then(|value_length| value_length.checked_add(value.len()))
            .ok_or_else(|| Error::InvalidFormat("XForms namespace size overflow".to_string()))?;
        let mut declaration = String::new();
        declaration
            .try_reserve_exact(length)
            .map_err(|source| Error::Allocation {
                resource: "ODT XForms namespace declaration",
                source,
            })?;
        declaration.push(' ');
        declaration.push_str(name);
        declaration.push_str("=\"");
        declaration.push_str(value);
        declaration.push('"');
        output.push(declaration);
    }
    Ok(output)
}

fn push_namespace_declarations<O: XmlOutput>(
    output: &mut O,
    declarations: &[String],
    scope: &mut Vec<(String, String)>,
) -> Result<()> {
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
        output.push_str(declaration)?;
        replace_namespace_checked(scope, name, value)?;
    }
    Ok(())
}

fn replace_namespace_checked(
    scope: &mut Vec<(String, String)>,
    name: &str,
    value: &str,
) -> Result<()> {
    if let Some((_, current)) = scope.iter_mut().find(|(current, _)| current == name) {
        *current = value.to_string();
    } else {
        scope
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT XForms model namespace scope",
                source,
            })?;
        scope.push((name.to_string(), value.to_string()));
    }
    Ok(())
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
    let mut reader = Reader::from_str(raw);
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
    let (open_end, empty, present) = first_tag_span(raw, None)?;
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

/// Namespace injection used while a budgeted model is being validated.
///
/// The root scan owns a temporary namespace-name set and the returned string
/// owns the complete injected fragment. Charge the scan envelope before the
/// scan starts, then charge the exact output allocation separately.
fn inject_namespace_declarations_with_budget(
    raw: &str,
    declarations: &[String],
    budget: &FlatMutationBudget,
) -> Result<ChargedXml> {
    let declaration_bytes = declarations.iter().try_fold(0usize, |total, value| {
        total
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("XForms namespace scan overflow".to_string()))
    })?;
    let scan_bytes = raw
        .len()
        .checked_add(declaration_bytes)
        .and_then(|value| value.checked_add(declarations.len().checked_mul(size_of::<String>())?))
        .ok_or_else(|| Error::InvalidFormat("XForms namespace scan overflow".to_string()))?;
    let scan_memory = budget.reserve_bytes(scan_bytes, "ODT XForms namespace scan")?;
    let (open_end, empty, present) = first_tag_span(raw, Some(budget))?;
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
    let output_len = raw.len().checked_add(insert_len).ok_or_else(|| {
        Error::InvalidFormat("XForms namespace insertion size overflow".to_string())
    })?;
    bounded_edit_length(output_len, "XForms namespace insertion")?;
    let (mut output, memory) =
        allocate_xml(Some(budget), output_len, "ODT XForms namespace insertion")?;
    if insert_len == 0 {
        output.push_str(raw);
        drop(scan_memory);
        return Ok(ChargedXml {
            xml: output,
            memory,
        });
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
    drop(scan_memory);
    Ok(ChargedXml {
        xml: output,
        memory,
    })
}

fn first_tag_span(
    raw: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<(usize, bool, HashSet<String>)> {
    let mut reader = Reader::from_str(raw);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        if let Some(budget) = budget {
            budget.check()?;
        }
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XML element: {error}")))?;
        match event {
            Event::Start(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    false,
                    present_namespace_names(&source, budget)?,
                ));
            },
            Event::Empty(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    true,
                    present_namespace_names(&source, budget)?,
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

fn present_namespace_names(
    source: &BytesStart<'_>,
    budget: Option<&FlatMutationBudget>,
) -> Result<HashSet<String>> {
    let mut present = HashSet::new();
    for attribute in source.attributes() {
        if let Some(budget) = budget {
            budget.check()?;
        }
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid XML namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            present.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "ODT XForms namespace-name set",
                source,
            })?;
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
    let mut reader = Reader::from_str(raw);
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

fn push_optional_attr<O: XmlOutput>(
    output: &mut O,
    local: &str,
    value: Option<&str>,
) -> Result<()> {
    if let Some(value) = value {
        output.push_char(' ')?;
        output.push_str(local)?;
        output.push_str("=\"")?;
        push_escaped_xml(output, value)?;
        output.push_char('"')?;
    }
    Ok(())
}

fn push_unknown_attributes<O: XmlOutput>(
    output: &mut O,
    attrs: &[Attribute],
    namespace_scope: &mut Vec<(String, String)>,
) -> Result<()> {
    let mut dynamic_prefixes: Vec<(String, String)> = Vec::new();
    for attr in attrs {
        let prefix = if attr.namespace_uri.is_empty() {
            String::new()
        } else if let Some(prefix) = canonical_prefix(&attr.namespace_uri).filter(|prefix| {
            namespace_scope.iter().all(|(name, namespace)| {
                !namespace_binding_name(name, prefix) || namespace == &attr.namespace_uri
            })
        }) {
            prefix.to_string()
        } else if let Some(prefix) = attr.prefix.as_deref().filter(|prefix| {
            is_valid_prefix(prefix)
                && namespace_scope.iter().all(|(name, namespace)| {
                    !namespace_binding_name(name, prefix) || namespace == &attr.namespace_uri
                })
        }) {
            let prefix = prefix.to_string();
            if !namespace_scope.iter().any(|(name, namespace)| {
                namespace_binding_name(name, &prefix) && namespace == &attr.namespace_uri
            }) {
                output.push_char(' ')?;
                output.push_str("xmlns:")?;
                output.push_str(&prefix)?;
                output.push_str("=\"")?;
                push_escaped_xml(output, &attr.namespace_uri)?;
                output.push_char('"')?;
                namespace_scope
                    .try_reserve_exact(1)
                    .map_err(|source| Error::Allocation {
                        resource: "ODT XForms model namespace scope",
                        source,
                    })?;
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
            output.push_char(' ')?;
            output.push_str("xmlns:")?;
            output.push_str(&prefix)?;
            output.push_str("=\"")?;
            push_escaped_xml(output, &attr.namespace_uri)?;
            output.push_char('"')?;
            dynamic_prefixes
                .try_reserve_exact(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODT XForms dynamic namespace prefixes",
                    source,
                })?;
            dynamic_prefixes.push((prefix.clone(), attr.namespace_uri.clone()));
            namespace_scope
                .try_reserve_exact(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODT XForms model namespace scope",
                    source,
                })?;
            namespace_scope.push((format!("xmlns:{prefix}"), attr.namespace_uri.clone()));
            prefix
        };
        output.push_char(' ')?;
        if !prefix.is_empty() {
            output.push_str(&prefix)?;
            output.push_char(':')?;
        }
        output.push_str(&attr.local_name)?;
        output.push_str("=\"")?;
        push_escaped_xml(output, &attr.value)?;
        output.push_char('"')?;
    }
    Ok(())
}

fn namespace_binding_name(name: &str, prefix: &str) -> bool {
    name.strip_prefix("xmlns:") == Some(prefix)
}

fn push_escaped_xml<O: XmlOutput>(output: &mut O, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;")?,
            '<' => output.push_str("&lt;")?,
            '>' => output.push_str("&gt;")?,
            '"' => output.push_str("&quot;")?,
            '\'' => output.push_str("&apos;")?,
            _ => output.push_char(character)?,
        }
    }
    Ok(())
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

fn validate_attributes_with_budget(
    attrs: &[Attribute],
    budget: Option<&FlatMutationBudget>,
) -> Result<()> {
    if attrs.len() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "XForms element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    let mut seen = HashSet::new();
    for attr in attrs {
        if let Some(budget) = budget {
            budget.check()?;
        }
        seen.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODT XForms attribute validation set",
            source,
        })?;
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

fn validate_fragment_with_context_with_budget(
    xml: &str,
    resource: &str,
    allow_empty: bool,
    inherited_declarations: &[String],
    budget: Option<&FlatMutationBudget>,
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
    } else if let Some(budget) = budget {
        Some(inject_namespace_declarations_with_budget(
            xml,
            inherited_declarations,
            budget,
        )?)
    } else {
        Some(ChargedXml {
            xml: inject_namespace_declarations(xml, inherited_declarations)?,
            memory: None,
        })
    };
    let source = owned.as_ref().map_or(xml, |value| value.xml.as_str());
    if source.len() > MAX_STRING_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{resource} XML exceeds its bounded size after namespace context injection"
        )));
    }
    let mut reader = ResolvedReader::from_xml(source);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut parser_memory = MemoryLease::default();
    if let Some(budget) = budget {
        parser_memory.reserve_vec(
            budget,
            &mut buffer,
            source.len(),
            "ODT XForms fragment parser buffer",
        )?;
    }
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut closed_root = false;
    loop {
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid {resource} XML: {error}")))?;
        resolved_namespace(&resolved)?;
        match event {
            Event::Start(ref value) => {
                validate_fragment_attributes(&reader, value, budget)?;
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
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "{resource} exceeds {MAX_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref value) => {
                validate_fragment_attributes(&reader, value, budget)?;
                if let Some(budget) = budget {
                    budget.observe_depth(depth.saturating_add(1))?;
                }
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

fn validate_extension_with_budget(
    value: &Extension,
    inherited_declarations: &[String],
    budget: Option<&FlatMutationBudget>,
) -> Result<()> {
    validate_fragment_with_context_with_budget(
        &value.xml,
        "XForms extension",
        false,
        inherited_declarations,
        budget,
    )?;
    let owned = if inherited_declarations.is_empty() {
        None
    } else if let Some(budget) = budget {
        Some(inject_namespace_declarations_with_budget(
            &value.xml,
            inherited_declarations,
            budget,
        )?)
    } else {
        Some(ChargedXml {
            xml: inject_namespace_declarations(&value.xml, inherited_declarations)?,
            memory: None,
        })
    };
    let source = owned
        .as_ref()
        .map_or(value.xml.as_str(), |fragment| fragment.xml.as_str());
    let root = extension_root_name(source, budget)?;
    if root.0 != value.namespace_uri || root.1 != value.local_name {
        return Err(Error::InvalidFormat(
            "XForms extension expanded root does not match its retained identity".to_string(),
        ));
    }
    Ok(())
}

fn extension_root_name(xml: &str, budget: Option<&FlatMutationBudget>) -> Result<(String, String)> {
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut parser_memory = MemoryLease::default();
    if let Some(budget) = budget {
        parser_memory.reserve_vec(
            budget,
            &mut buffer,
            xml.len(),
            "ODT XForms extension root parser",
        )?;
    }
    let mut depth = 0usize;
    loop {
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid XForms extension XML: {error}"))
            })?;
        match event {
            Event::Start(source) => {
                let namespace = resolved_namespace(&resolved)?.unwrap_or_default();
                let local = utf8(source.local_name().as_ref(), "XForms extension root name")?;
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms extension depth overflow".to_string())
                })?;
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
                return Ok((namespace, local));
            },
            Event::Empty(source) => {
                let namespace = resolved_namespace(&resolved)?.unwrap_or_default();
                let local = utf8(source.local_name().as_ref(), "XForms extension root name")?;
                if let Some(budget) = budget {
                    budget.observe_depth(1)?;
                }
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

fn validate_fragment_attributes(
    reader: &ResolvedReader<'_>,
    source: &BytesStart<'_>,
    budget: Option<&FlatMutationBudget>,
) -> Result<()> {
    for attribute in source.attributes() {
        if let Some(budget) = budget {
            budget.check()?;
        }
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

fn validate_reserved_namespace_bindings(
    declarations: &[String],
    resource: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<()> {
    for declaration in declarations {
        if let Some(budget) = budget {
            budget.check()?;
        }
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
            let semantic = quick_xml::escape::unescape(value).map_err(|error| {
                Error::InvalidFormat(format!("{resource} has an invalid namespace URI: {error}"))
            })?;
            if semantic != expected {
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
    crate::elements::xml::normalized_namespace_uri(namespace, "XForms")?
        .map(|uri| utf8(uri, "XForms namespace URI"))
        .transpose()
}

fn utf8(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 {description}")))
}

/// Find direct model spans for the mutable XML editor.
pub(crate) fn model_spans(xml: &str) -> Result<Vec<(usize, usize)>> {
    model_spans_with_optional_budget(xml, None)
}

fn model_spans_with_optional_budget(
    xml: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<Vec<(usize, usize)>> {
    // Reuse the parser's validated model discovery and then independently
    // rescan the same bounded source for exact byte spans.
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(String, String, usize)> = Vec::new();
    let mut output = Vec::new();
    let mut depth = 0usize;
    loop {
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
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
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "XForms XML exceeds {MAX_DEPTH} levels"
                    )));
                }
                let ns = namespace.clone();
                let local = utf8(source.local_name().as_ref(), "XForms element name")?;
                let start = event_position;
                stack
                    .try_reserve_exact(1)
                    .map_err(|source| Error::Allocation {
                        resource: "ODT XForms span frame stack",
                        source,
                    })?;
                stack.push((ns, local, start));
            },
            Event::Empty(ref source) => {
                if let Some(budget) = budget {
                    budget.observe_depth(depth.saturating_add(1))?;
                }
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
                    output
                        .try_reserve_exact(1)
                        .map_err(|source| Error::Allocation {
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
                    output
                        .try_reserve_exact(1)
                        .map_err(|source| Error::Allocation {
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

fn model_spans_with_budget(
    xml: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<(Vec<(usize, usize)>, Option<MemoryLease>)> {
    let memory = budget
        .map(|budget| {
            xforms_parser_memory_plan(xml, Some(budget)).and_then(|amount| {
                budget
                    .reserve_bytes(amount, "ODT XForms model edit span scan")
                    .map(MemoryLease::new)
            })
        })
        .transpose()?;
    let spans = match budget {
        Some(budget) => model_spans_with_optional_budget(xml, Some(budget)),
        None => model_spans(xml),
    };
    match spans {
        Ok(spans) => Ok((spans, memory)),
        Err(error) => {
            drop(memory);
            Err(error)
        },
    }
}

fn parse_models_optional_budget(
    xml: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<(Vec<Model>, Option<MemoryLease>)> {
    match budget {
        Some(budget) => {
            parse_models_with_budget(xml, budget).map(|(models, memory)| (models, Some(memory)))
        },
        None => parse_models(xml).map(|models| (models, None)),
    }
}

fn validate_models_optional_budget(xml: &str, budget: Option<&FlatMutationBudget>) -> Result<()> {
    let (models, memory) = parse_models_optional_budget(xml, budget)?;
    drop(models);
    drop(memory);
    Ok(())
}

pub(crate) fn validate_models_with_budget(xml: &str, budget: &FlatMutationBudget) -> Result<()> {
    validate_models_optional_budget(xml, Some(budget))
}

pub(crate) fn replace_model(xml: &str, position: usize, model: &Model) -> Result<String> {
    replace_model_with_limit(xml, position, model, MAX_XML_BYTES)
}

pub(crate) fn replace_model_with_limit(
    xml: &str,
    position: usize,
    model: &Model,
    maximum: usize,
) -> Result<String> {
    replace_model_with_limit_and_budget(xml, position, model, maximum, None)
        .map(ChargedXml::into_string)
}

pub(crate) fn replace_model_with_limit_and_budget(
    xml: &str,
    position: usize,
    model: &Model,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    let (current, parse_memory) = parse_models_optional_budget(xml, budget)?;
    let existing = current.get(position).ok_or_else(|| {
        Error::InvalidFormat(format!("XForms model position {position} is out of range"))
    })?;
    if existing == model {
        drop(current);
        drop(parse_memory);
        bounded_edit_length_with_limit(xml.len(), "XForms exact no-op", maximum)?;
        let (mut output, memory) = allocate_xml(budget, xml.len(), "ODT XForms exact no-op")?;
        output.push_str(xml);
        return Ok(ChargedXml {
            xml: output,
            memory,
        });
    }
    drop(current);
    drop(parse_memory);
    let replacement = model.to_xml_with_limit_and_budget(maximum, budget)?;
    let (spans, span_memory) = model_spans_with_budget(xml, budget)?;
    let (start, end) = spans.get(position).copied().ok_or_else(|| {
        Error::InvalidFormat(format!("XForms model position {position} is out of range"))
    })?;
    let source = xml
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("invalid XForms model span".to_string()))?;
    if model_has_direct_lexical_material(source, budget)? {
        return Err(Error::InvalidFormat(
            "changed XForms model would discard source comments, processing instructions, or inter-child whitespace"
                .to_string(),
        ));
    }
    let candidate =
        splice_replace_with_limit_and_budget(xml, start, end, &replacement.xml, maximum, budget)?;
    drop(spans);
    drop(span_memory);
    validate_models_optional_budget(&candidate.xml, budget)?;
    Ok(candidate)
}

pub(crate) fn remove_model(xml: &str, position: usize) -> Result<String> {
    remove_model_with_limit(xml, position, MAX_XML_BYTES)
}

pub(crate) fn remove_model_with_limit(
    xml: &str,
    position: usize,
    maximum: usize,
) -> Result<String> {
    remove_model_with_limit_and_budget(xml, position, maximum, None).map(ChargedXml::into_string)
}

pub(crate) fn remove_model_with_limit_and_budget(
    xml: &str,
    position: usize,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    let (current, parse_memory) = parse_models_optional_budget(xml, budget)?;
    if current.get(position).is_none() {
        return Err(Error::InvalidFormat(format!(
            "XForms model position {position} is out of range"
        )));
    }
    drop(current);
    drop(parse_memory);
    let (spans, span_memory) = model_spans_with_budget(xml, budget)?;
    let (start, end) = spans.get(position).copied().ok_or_else(|| {
        Error::InvalidFormat(format!("XForms model position {position} is out of range"))
    })?;
    let candidate = splice_replace_with_limit_and_budget(xml, start, end, "", maximum, budget)?;
    drop(spans);
    drop(span_memory);
    validate_models_optional_budget(&candidate.xml, budget)?;
    Ok(candidate)
}

pub(crate) fn insert_model(xml: &str, model: &Model) -> Result<String> {
    insert_model_with_limit(xml, model, MAX_XML_BYTES)
}

pub(crate) fn insert_model_with_limit(xml: &str, model: &Model, maximum: usize) -> Result<String> {
    insert_model_with_limit_and_budget(xml, model, maximum, None).map(ChargedXml::into_string)
}

pub(crate) fn insert_model_with_limit_and_budget(
    xml: &str,
    model: &Model,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    let (models, parse_memory) = parse_models_optional_budget(xml, budget)?;
    drop(models);
    drop(parse_memory);
    let fragment = model.to_xml_with_limit_and_budget(maximum, budget)?;
    let container_memory = budget
        .map(|budget| {
            xforms_parser_memory_plan(xml, Some(budget)).and_then(|amount| {
                budget
                    .reserve_bytes(amount, "ODT XForms forms span scan")
                    .map(MemoryLease::new)
            })
        })
        .transpose()?;
    let forms = container_span(xml, OFFICENS, "forms", budget)?.ok_or_else(|| {
        Error::InvalidFormat("document has no office:forms container".to_string())
    })?;
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
            .checked_add(fragment.xml.len())
            .and_then(|length| length.checked_add(20))
            .ok_or_else(|| Error::InvalidFormat("XForms insertion size overflow".to_string()))?;
        bounded_edit_length_with_limit(replacement_len, "XForms insertion", maximum)?;
        let (mut replacement, replacement_memory) =
            allocate_xml(budget, replacement_len, "ODT XForms insertion")?;
        append_checked(&mut replacement, &source[..close], budget)?;
        if let Some(budget) = budget {
            budget.check()?;
        }
        replacement.push('>');
        append_checked(&mut replacement, &fragment.xml, budget)?;
        append_checked(&mut replacement, "</", budget)?;
        append_checked(&mut replacement, &source[1..qname_end], budget)?;
        if let Some(budget) = budget {
            budget.check()?;
        }
        replacement.push('>');
        let candidate = splice_replace_with_limit_and_budget(
            xml,
            forms.0,
            forms.1,
            &replacement,
            maximum,
            budget,
        )?;
        drop(replacement_memory);
        drop(container_memory);
        validate_models_optional_budget(&candidate.xml, budget)?;
        return Ok(candidate);
    }
    let candidate =
        splice_insert_with_limit_and_budget(xml, forms.1, &fragment.xml, maximum, budget)?;
    drop(container_memory);
    validate_models_optional_budget(&candidate.xml, budget)?;
    Ok(candidate)
}

fn model_has_direct_lexical_material(
    xml: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<bool> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid XForms model XML: {error}")))?;
        match event {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms model depth overflow".to_string())
                })?;
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("XForms model depth underflow".to_string())
                })?;
            },
            Event::Empty(_) => {
                if let Some(budget) = budget {
                    budget.observe_depth(depth.saturating_add(1))?;
                }
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

fn container_span(
    xml: &str,
    namespace: &str,
    local: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<Option<(usize, usize)>> {
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(String, String, usize)> = Vec::new();
    let mut found = None;
    let mut depth = 0usize;
    loop {
        if let Some(budget) = budget {
            budget.event(depth)?;
        }
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
                if let Some(budget) = budget {
                    budget.observe_depth(depth)?;
                }
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODF XML exceeds {MAX_DEPTH} levels"
                    )));
                }
                let ns = resolved.clone();
                let name = utf8(source.local_name().as_ref(), "container name")?;
                stack
                    .try_reserve_exact(1)
                    .map_err(|source| Error::Allocation {
                        resource: "ODT XForms forms span stack",
                        source,
                    })?;
                stack.push((ns, name, event_position));
            },
            Event::Empty(ref source) => {
                if let Some(budget) = budget {
                    budget.observe_depth(depth.saturating_add(1))?;
                }
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

fn splice_insert_with_limit_and_budget(
    xml: &str,
    offset: usize,
    value: &str,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    if offset > xml.len() {
        return Err(Error::InvalidFormat(
            "invalid XML insertion offset".to_string(),
        ));
    }
    let output_len = xml
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("XForms insertion size overflow".to_string()))?;
    bounded_edit_length_with_limit(output_len, "XForms insertion", maximum)?;
    let (mut output, memory) = allocate_xml(budget, output_len, "ODT XForms insertion")?;
    append_checked(&mut output, &xml[..offset], budget)?;
    append_checked(&mut output, value, budget)?;
    append_checked(&mut output, &xml[offset..], budget)?;
    Ok(ChargedXml {
        xml: output,
        memory,
    })
}

fn splice_replace_with_limit_and_budget(
    xml: &str,
    start: usize,
    end: usize,
    value: &str,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
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
    bounded_edit_length_with_limit(output_len, "XForms replacement", maximum)?;
    let (mut output, memory) = allocate_xml(budget, output_len, "ODT XForms replacement")?;
    append_checked(&mut output, &xml[..start], budget)?;
    append_checked(&mut output, value, budget)?;
    append_checked(&mut output, &xml[end..], budget)?;
    Ok(ChargedXml {
        xml: output,
        memory,
    })
}

fn append_checked(
    output: &mut String,
    value: &str,
    budget: Option<&FlatMutationBudget>,
) -> Result<()> {
    let Some(budget) = budget else {
        output.push_str(value);
        return Ok(());
    };
    let mut start = 0usize;
    for (index, _) in value.char_indices() {
        if index.saturating_sub(start) >= 8 * 1024 {
            budget.check()?;
            output.push_str(&value[start..index]);
            start = index;
        }
    }
    if start < value.len() {
        budget.check()?;
        output.push_str(&value[start..]);
    } else {
        budget.check()?;
    }
    Ok(())
}

fn bounded_edit_length(length: usize, operation: &str) -> Result<()> {
    bounded_edit_length_with_limit(length, operation, MAX_XML_BYTES)
}

fn bounded_edit_length_with_limit(length: usize, operation: &str, maximum: usize) -> Result<()> {
    if length > maximum {
        if maximum < MAX_XML_BYTES {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::OutputBytes,
                observed: u64::try_from(length).unwrap_or(u64::MAX),
                limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                scope: operation.into(),
            }));
        }
        return Err(Error::InvalidFormat(format!(
            "{operation} exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
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
    fn parses_entity_escaped_office_and_xforms_namespace_uris() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:&#x31;.0" xmlns:xforms="http://www.w3.org/&#x32;002/xforms"><office:body><office:text><office:forms><xforms:model id="m"/></office:forms></office:text></office:body></office:document-content>"#;
        let models = parse_models(xml).expect("escaped namespace URIs are semantic");
        assert_eq!(models.len(), 1);
        assert_eq!(models[0].id.as_deref(), Some("m"));
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
    fn model_limit_accepts_exact_output_and_rejects_one_byte_under() {
        let model = parse_models(XML).unwrap().remove(0);
        let expected = model.to_xml().unwrap();
        assert_eq!(model.to_xml_with_limit(expected.len()).unwrap(), expected);
        assert!(matches!(
            model.to_xml_with_limit(expected.len() - 1),
            Err(Error::ResourceLimit(_))
        ));
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
