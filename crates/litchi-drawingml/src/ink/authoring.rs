//! Detached, bounded InkML authoring.
//!
//! This module owns the canonical writer for newly authored InkML.  Parsed
//! [`super::Document`] values remain source backed and are written byte for
//! byte by the existing codec; a [`Draft`] is a separate consuming builder
//! that emits only the small typed subset represented here. Trace data uses
//! the canonical finite X/Y integer-pair profile emitted by this writer; it
//! is never treated as caller supplied XML.

use std::{borrow::Cow, sync::Arc};

use litchi_ooxml_common::xml_name::is_ncname;
use quick_xml::{
    Reader, XmlVersion,
    events::{BytesStart, Event},
};

use crate::{Error, Result};

use super::{
    BrushPropertyName, ContextKind, Document, Guid, INKML_NAMESPACE, InkEffect,
    MAX_ATTRIBUTE_VALUE_BYTES, MAX_BRUSH_PROPERTIES, MAX_CONTEXTS, MAX_DEPTH, MAX_NODES,
    MAX_SOURCE_BYTES, MAX_TOKEN_BYTES, MAX_TRACES, Metadata, NAMESPACE, SemanticType, SourceSpan,
    read_metadata, read_shared,
};

const MAX_AUTHORING_BRUSHES: usize = MAX_BRUSH_PROPERTIES;
const MAX_AUTHORING_ID_BYTES: usize = MAX_TOKEN_BYTES;
const MAX_AUTHORING_REFERENCE_BYTES: usize = 4_096;
const MAX_AUTHORING_ATTRIBUTES: usize = 256;
const EMMA_NAMESPACE: &str = "http://www.w3.org/2003/04/emma";

/// Finite budgets for one detached InkML authoring operation.
///
/// Every field is bounded again against the shared InkML reader's hard
/// ceiling during [`Draft::finish`].  This keeps a caller supplied budget
/// from asking the writer to produce bytes the shared reader cannot reopen.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct AuthoringLimits {
    /// Maximum complete output bytes accepted by the shared InkML reader.
    pub max_source_bytes: usize,
    /// Maximum bytes reserved and emitted by the canonical writer.
    pub max_output_bytes: usize,
    /// Maximum generated XML element nodes, including the root and wrappers.
    pub max_nodes: usize,
    /// Maximum generated element depth.
    pub max_depth: usize,
    /// Maximum generated attributes on one element.
    pub max_attributes_per_element: usize,
    /// Maximum decoded value bytes for one generated attribute.
    pub max_attribute_value_bytes: usize,
    /// Maximum authored `msink:context` records.
    pub max_contexts: usize,
    /// Maximum authored `inkml:brush` records.
    pub max_brushes: usize,
    /// Maximum authored `inkml:brushProperty` records.
    pub max_brush_properties: usize,
    /// Maximum authored `inkml:trace` records.
    pub max_traces: usize,
    /// Maximum bytes in one local authored identifier.
    pub max_id_bytes: usize,
    /// Maximum bytes in one local `#...` reference.
    pub max_reference_bytes: usize,
    /// Maximum bytes in one canonical trace lexical value.
    pub max_trace_text_bytes: usize,
}

impl Default for AuthoringLimits {
    fn default() -> Self {
        Self {
            max_source_bytes: MAX_SOURCE_BYTES,
            max_output_bytes: MAX_SOURCE_BYTES,
            max_nodes: MAX_NODES,
            max_depth: MAX_DEPTH,
            max_attributes_per_element: MAX_AUTHORING_ATTRIBUTES,
            max_attribute_value_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
            max_contexts: MAX_CONTEXTS,
            max_brushes: MAX_AUTHORING_BRUSHES,
            max_brush_properties: MAX_BRUSH_PROPERTIES,
            max_traces: MAX_TRACES,
            max_id_bytes: MAX_AUTHORING_ID_BYTES,
            max_reference_bytes: MAX_AUTHORING_REFERENCE_BYTES,
            max_trace_text_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
        }
    }
}

impl AuthoringLimits {
    /// Validate that all budgets are finite, nonzero, and compatible with the
    /// shared bounded reader.
    ///
    /// # Errors
    ///
    /// Returns an error if a budget is zero or exceeds the shared hard cap.
    pub fn validate(self) -> Result<()> {
        for (resource, value, maximum) in [
            ("source bytes", self.max_source_bytes, MAX_SOURCE_BYTES),
            ("output bytes", self.max_output_bytes, MAX_SOURCE_BYTES),
            ("XML nodes", self.max_nodes, MAX_NODES),
            ("XML depth", self.max_depth, MAX_DEPTH),
            (
                "attributes per element",
                self.max_attributes_per_element,
                MAX_AUTHORING_ATTRIBUTES,
            ),
            (
                "attribute value bytes",
                self.max_attribute_value_bytes,
                MAX_ATTRIBUTE_VALUE_BYTES,
            ),
            ("contexts", self.max_contexts, MAX_CONTEXTS),
            ("brushes", self.max_brushes, MAX_AUTHORING_BRUSHES),
            (
                "brush properties",
                self.max_brush_properties,
                MAX_BRUSH_PROPERTIES,
            ),
            ("traces", self.max_traces, MAX_TRACES),
            (
                "identifier bytes",
                self.max_id_bytes,
                MAX_AUTHORING_ID_BYTES,
            ),
            (
                "reference bytes",
                self.max_reference_bytes,
                MAX_AUTHORING_REFERENCE_BYTES,
            ),
            (
                "trace text bytes",
                self.max_trace_text_bytes,
                MAX_ATTRIBUTE_VALUE_BYTES,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(invalid(format!(
                    "InkML authoring {resource} budget must be in 1..={maximum}"
                )));
            }
        }
        Ok(())
    }
}

/// A detached bounded InkML value builder.
#[derive(Debug, Default)]
pub struct Draft {
    limits: AuthoringLimits,
    contexts: Vec<ContextDraft>,
    brushes: Vec<BrushDraft>,
    traces: Vec<TraceDraft>,
    property_count: usize,
    retained_input_bytes: usize,
}

impl Draft {
    /// Start a detached authoring operation.
    ///
    /// # Errors
    ///
    /// Returns an error for a zero or overlarge budget.
    pub fn new(limits: AuthoringLimits) -> Result<Self> {
        limits.validate()?;
        Ok(Self {
            limits,
            contexts: Vec::new(),
            brushes: Vec::new(),
            traces: Vec::new(),
            property_count: 0,
            retained_input_bytes: 0,
        })
    }

    /// Borrow the budgets governing this draft.
    #[must_use]
    pub const fn limits(&self) -> AuthoringLimits {
        self.limits
    }

    /// Add one context record, consuming the previous draft value.
    ///
    /// # Errors
    ///
    /// Returns a limit or allocation error before the record is appended.
    pub fn context(mut self, context: ContextDraft) -> Result<Self> {
        self.limits.validate()?;
        validate_context(&context, self.limits)?;
        if self.contexts.len() >= self.limits.max_contexts {
            return Err(limit("InkML context records", self.limits.max_contexts));
        }
        let text_bytes = context_text_bytes(&context)?;
        ensure_retained_input(self.retained_input_bytes, text_bytes, self.limits)?;
        try_reserve_one(&mut self.contexts, "InkML context records")?;
        self.contexts.push(context);
        self.retained_input_bytes = self
            .retained_input_bytes
            .checked_add(text_bytes)
            .ok_or_else(|| {
                limit(
                    "InkML retained authoring text",
                    retained_text_limit(self.limits),
                )
            })?;
        Ok(self)
    }

    /// Add one brush record, consuming the previous draft value.
    ///
    /// # Errors
    ///
    /// Returns a limit or allocation error before the record is appended.
    pub fn brush(mut self, brush: BrushDraft) -> Result<Self> {
        self.limits.validate()?;
        validate_brush(&brush, self.limits)?;
        if self.brushes.len() >= self.limits.max_brushes {
            return Err(limit("InkML brush records", self.limits.max_brushes));
        }
        let next_properties = self
            .property_count
            .checked_add(brush.properties.len())
            .ok_or_else(|| limit("InkML brush properties", self.limits.max_brush_properties))?;
        if next_properties > self.limits.max_brush_properties {
            return Err(limit(
                "InkML brush properties",
                self.limits.max_brush_properties,
            ));
        }
        let text_bytes = brush_text_bytes(&brush);
        ensure_retained_input(self.retained_input_bytes, text_bytes, self.limits)?;
        try_reserve_one(&mut self.brushes, "InkML brush records")?;
        self.brushes.push(brush);
        self.property_count = next_properties;
        self.retained_input_bytes = self
            .retained_input_bytes
            .checked_add(text_bytes)
            .ok_or_else(|| {
                limit(
                    "InkML retained authoring text",
                    retained_text_limit(self.limits),
                )
            })?;
        Ok(self)
    }

    /// Add one trace record, consuming the previous draft value.
    ///
    /// # Errors
    ///
    /// Returns a limit or allocation error before the record is appended.
    pub fn trace(mut self, trace: TraceDraft) -> Result<Self> {
        self.limits.validate()?;
        validate_trace(&trace, self.limits)?;
        if self.traces.len() >= self.limits.max_traces {
            return Err(limit("InkML trace records", self.limits.max_traces));
        }
        let text_bytes = trace_text_bytes(&trace)?;
        ensure_retained_input(self.retained_input_bytes, text_bytes, self.limits)?;
        try_reserve_one(&mut self.traces, "InkML trace records")?;
        self.traces.push(trace);
        self.retained_input_bytes = self
            .retained_input_bytes
            .checked_add(text_bytes)
            .ok_or_else(|| {
                limit(
                    "InkML retained authoring text",
                    retained_text_limit(self.limits),
                )
            })?;
        Ok(self)
    }

    /// Finish the draft into canonical namespace-bound InkML.
    ///
    /// The generated bytes are reopened through the shared bounded metadata
    /// reader before publication.  Only typed values represented by this
    /// module are emitted; no external relationship or opaque XML plan is
    /// attached to the result.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid limits, duplicate or unresolved local
    /// identifiers, invalid references, resource exhaustion, XML allocation
    /// failure, or failed canonical readback.
    pub fn finish(self) -> Result<Prepared> {
        self.limits.validate()?;
        let source_ids = self.validate_records()?;
        if self.limits.max_attributes_per_element < 2 {
            return Err(limit(
                "InkML attributes per element",
                self.limits.max_attributes_per_element,
            ));
        }
        validate_attribute_length(INKML_NAMESPACE, self.limits)?;
        validate_attribute_length(NAMESPACE, self.limits)?;
        validate_attribute_length(EMMA_NAMESPACE, self.limits)?;
        validate_attribute_length("1.0", self.limits)?;
        validate_attribute_length("ink", self.limits)?;
        validate_attribute_length("X", self.limits)?;
        validate_attribute_length("Y", self.limits)?;
        validate_attribute_length("integer", self.limits)?;
        let node_count = self.node_count()?;
        if node_count > self.limits.max_nodes {
            return Err(limit("InkML XML nodes", self.limits.max_nodes));
        }
        let depth = self.required_depth();
        if depth > self.limits.max_depth {
            return Err(limit("InkML XML depth", self.limits.max_depth));
        }

        let mut length = LengthSink::new(self.limits.max_output_bytes);
        emit_document(&self, &source_ids, &mut length)?;
        let output_len = length.len;
        if output_len > self.limits.max_source_bytes {
            return Err(limit("InkML source bytes", self.limits.max_source_bytes));
        }
        let mut writer = XmlSink::new(output_len, self.limits.max_output_bytes)?;
        emit_document(&self, &source_ids, &mut writer)?;
        let bytes = Arc::new(writer.finish());
        let document = read_shared(Arc::clone(&bytes))?;
        self.compare_readback(document.metadata())?;
        self.compare_trace_bytes(&document)?;
        let metadata = document.metadata().clone();
        Ok(Prepared { bytes, metadata })
    }

    fn validate_records(&self) -> Result<Vec<Option<Box<str>>>> {
        let definition_contexts = self
            .contexts
            .iter()
            .filter(|context| context.xml_id.is_some())
            .count();
        let local_identifier_count = definition_contexts
            .checked_mul(2)
            .and_then(|count| count.checked_add(self.brushes.len()))
            .ok_or_else(|| invalid("InkML identifier count overflow"))?;
        let mut local_ids = Vec::new();
        local_ids
            .try_reserve(local_identifier_count)
            .map_err(|_| invalid("InkML local identifier allocation failed"))?;
        let mut context_ids = Vec::new();
        context_ids
            .try_reserve(definition_contexts)
            .map_err(|_| invalid("InkML context identifier allocation failed"))?;
        let mut brush_ids = Vec::new();
        brush_ids
            .try_reserve(self.brushes.len())
            .map_err(|_| invalid("InkML brush identifier allocation failed"))?;
        let mut semantic_ids = Vec::new();
        semantic_ids
            .try_reserve(self.contexts.len())
            .map_err(|_| invalid("InkML context GUID allocation failed"))?;
        for context in &self.contexts {
            validate_context(context, self.limits)?;
            if let Some(id) = context.id.as_ref() {
                semantic_ids.push(id.as_str());
            }
            if let Some(id) = context.xml_id.as_deref() {
                local_ids.push(id);
                context_ids.push(id);
            }
        }
        for brush in &self.brushes {
            validate_brush(brush, self.limits)?;
            let id = brush.id.as_ref();
            local_ids.push(id);
            brush_ids.push(id);
        }
        local_ids.sort_unstable();
        let source_ids = self.generated_source_ids(&local_ids)?;
        for id in source_ids.iter().flatten() {
            local_ids.push(id);
        }
        local_ids.sort_unstable();
        context_ids.sort_unstable();
        brush_ids.sort_unstable();
        semantic_ids.sort_unstable();
        if has_duplicate(&local_ids) || has_duplicate(&semantic_ids) {
            return Err(invalid("InkML authored identifiers must be unique"));
        }
        for trace in &self.traces {
            validate_trace(trace, self.limits)?;
            let context_reference = trace
                .context_ref
                .as_deref()
                .ok_or_else(|| invalid("InkML trace contextRef is required for local closure"))?;
            let context_target = reference_target(context_reference)?;
            if context_ids
                .binary_search_by(|candidate| (*candidate).cmp(context_target))
                .is_err()
            {
                return Err(invalid(
                    "InkML trace contextRef has no local context target",
                ));
            }
            let brush_reference = trace
                .brush_ref
                .as_deref()
                .ok_or_else(|| invalid("InkML trace brushRef is required for local closure"))?;
            let brush_target = reference_target(brush_reference)?;
            if brush_ids
                .binary_search_by(|candidate| (*candidate).cmp(brush_target))
                .is_err()
            {
                return Err(invalid("InkML trace brushRef has no local brush target"));
            }
        }
        Ok(source_ids)
    }

    fn generated_source_ids(&self, used_ids: &[&str]) -> Result<Vec<Option<Box<str>>>> {
        let mut source_ids = Vec::new();
        source_ids
            .try_reserve(self.contexts.len())
            .map_err(|_| invalid("InkML generated source identifier allocation failed"))?;
        let mut next_candidate = 0usize;
        for context in &self.contexts {
            if context.xml_id.is_none() {
                source_ids.push(None);
                continue;
            }
            loop {
                let candidate = format!("inkSrc{next_candidate}");
                next_candidate = next_candidate
                    .checked_add(1)
                    .ok_or_else(|| invalid("InkML generated source identifier counter overflow"))?;
                if candidate.len() > self.limits.max_id_bytes {
                    return Err(limit("InkML identifier bytes", self.limits.max_id_bytes));
                }
                validate_attribute_length(&candidate, self.limits)?;
                if used_ids
                    .binary_search_by(|used| (*used).cmp(candidate.as_str()))
                    .is_ok()
                {
                    continue;
                }
                source_ids.push(Some(own_identifier(
                    candidate,
                    "generated inkSource identifier",
                    self.limits.max_id_bytes,
                )?));
                break;
            }
        }
        Ok(source_ids)
    }

    fn node_count(&self) -> Result<usize> {
        let mut count = 1usize; // inkml:ink
        let definition_contexts = self
            .contexts
            .iter()
            .filter(|context| context.xml_id.is_some())
            .count();
        if definition_contexts != 0 || !self.brushes.is_empty() {
            let context_definition_nodes = definition_contexts
                .checked_mul(5)
                .ok_or_else(|| limit("InkML XML nodes", self.limits.max_nodes))?;
            count = count
                .checked_add(1)
                .and_then(|value| value.checked_add(context_definition_nodes))
                .and_then(|value| value.checked_add(self.brushes.len()))
                .ok_or_else(|| limit("InkML XML nodes", self.limits.max_nodes))?;
            count = count
                .checked_add(self.property_count)
                .ok_or_else(|| limit("InkML XML nodes", self.limits.max_nodes))?;
        }
        if !self.contexts.is_empty() || !self.traces.is_empty() {
            let annotation_nodes = self
                .contexts
                .len()
                .checked_mul(4)
                .ok_or_else(|| limit("InkML XML nodes", self.limits.max_nodes))?;
            count = count
                .checked_add(1)
                .and_then(|value| value.checked_add(annotation_nodes))
                .and_then(|value| value.checked_add(self.traces.len()))
                .ok_or_else(|| limit("InkML XML nodes", self.limits.max_nodes))?;
        }
        Ok(count)
    }

    fn required_depth(&self) -> usize {
        let mut depth = 1; // inkml:ink
        if self.contexts.iter().any(|context| context.xml_id.is_some()) || !self.brushes.is_empty()
        {
            depth = depth.max(3); // definitions/brush or context
            if self
                .brushes
                .iter()
                .any(|brush| !brush.properties.is_empty())
            {
                depth = depth.max(4); // definitions/brush/brushProperty
            }
        }
        if !self.contexts.is_empty() || !self.traces.is_empty() {
            depth = depth.max(2); // traceGroup
        }
        if !self.contexts.is_empty() {
            depth = depth.max(6); // annotationXML/emma/interpretation/context
        }
        depth
    }

    fn compare_readback(&self, metadata: &Metadata) -> Result<()> {
        if metadata.context_count() != self.contexts.len()
            || metadata.trace_count() != self.traces.len()
            || metadata.brush_property_count() != self.property_count
        {
            return Err(invalid("InkML canonical metadata count readback mismatch"));
        }
        for (expected, actual) in self.contexts.iter().zip(metadata.contexts()) {
            if actual.kind() != &expected.kind
                || actual.id() != expected.id.as_ref()
                || actual.semantic_type() != expected.semantic_type.as_ref()
                || actual.alignment_level() != expected.alignment_level
                || actual.content_type() != expected.content_type
                || actual.rotation_angle() != expected.rotation_angle
            {
                return Err(invalid("InkML context metadata readback mismatch"));
            }
        }
        for (expected, actual) in self
            .brushes
            .iter()
            .flat_map(|brush| brush.properties.iter())
            .zip(metadata.brush_properties())
        {
            if actual.name() != &expected.name
                || actual.units() != expected.units.as_deref()
                || !authored_brush_value_matches(
                    &expected.name,
                    expected.value.as_ref(),
                    actual.value(),
                )
            {
                return Err(invalid("InkML brush metadata readback mismatch"));
            }
        }
        for (expected, actual) in self.traces.iter().zip(metadata.traces()) {
            if actual.context_ref() != expected.context_ref.as_deref()
                || actual.brush_ref() != expected.brush_ref.as_deref()
            {
                return Err(invalid("InkML trace metadata readback mismatch"));
            }
        }
        Ok(())
    }

    fn compare_trace_bytes(&self, document: &Document) -> Result<()> {
        for (expected, actual) in self.traces.iter().zip(document.traces()) {
            if actual.data(document) != expected.data.as_bytes() {
                return Err(invalid("InkML trace lexical data readback mismatch"));
            }
        }
        Ok(())
    }
}

fn authored_brush_value_matches(name: &BrushPropertyName, expected: &str, actual: &str) -> bool {
    match name {
        BrushPropertyName::Width
        | BrushPropertyName::Height
        | BrushPropertyName::AnchorX
        | BrushPropertyName::AnchorY
        | BrushPropertyName::ScaleFactor
        | BrushPropertyName::Transparency
        | BrushPropertyName::AntiAliased
        | BrushPropertyName::FitToCurve
        | BrushPropertyName::IgnorePressure => {
            collapse_xsd_whitespace(expected).as_ref() == collapse_xsd_whitespace(actual).as_ref()
        },
        BrushPropertyName::Color
        | BrushPropertyName::Tip
        | BrushPropertyName::RasterOp
        | BrushPropertyName::InkEffects
        | BrushPropertyName::Custom(_) => expected == actual,
    }
}

/// Immutable canonical InkML bytes and their typed readback projection.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Prepared {
    bytes: Arc<Vec<u8>>,
    metadata: Metadata,
}

impl Prepared {
    /// Admit bytes only when they are byte-for-byte output of this canonical
    /// authoring profile.
    ///
    /// This is deliberately stricter than [`super::read`]. General or legacy
    /// InkML remains available through the source-backed reader; it is not
    /// silently normalized into a detached [`Prepared`] value.
    ///
    /// # Errors
    ///
    /// Returns an error when the input is malformed, exceeds the supplied
    /// limits, uses an unsupported canonical shape, or differs from the
    /// bytes emitted after typed reconstruction.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, AuthoringLimits::default())
    }

    /// Admit bytes using explicit finite authoring limits.
    ///
    /// The input is scanned and reconstructed while borrowed. No caller input
    /// allocation is retained; the bounded generated allocation is published
    /// only after it compares byte-for-byte equal to the input.
    ///
    /// # Errors
    ///
    /// Returns an error when the input is malformed, exceeds the supplied
    /// limits, uses an unsupported canonical shape, or differs from the
    /// bytes emitted after typed reconstruction.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: AuthoringLimits) -> Result<Self> {
        limits.validate()?;
        if bytes.len() > limits.max_source_bytes {
            return Err(limit("InkML source bytes", limits.max_source_bytes));
        }
        if bytes.len() > limits.max_output_bytes {
            return Err(limit("InkML output bytes", limits.max_output_bytes));
        }
        preflight_import_limits(bytes, limits)?;
        let metadata = read_metadata(bytes)?;
        validate_import_metadata_limits(&metadata, limits)?;
        let draft = import_canonical_draft(bytes, &metadata, limits)?;
        let prepared = draft.finish()?;
        if prepared.as_bytes() != bytes {
            return Err(invalid(
                "InkML bytes are valid XML but are not canonical detached authoring output",
            ));
        }
        Ok(prepared)
    }

    /// Borrow the canonical bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Borrow the typed metadata read back from the canonical bytes.
    #[must_use]
    pub const fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Clone the immutable source allocation for an existing shared reader or
    /// an OPC part owner.
    #[must_use]
    pub fn shared_source(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.bytes)
    }

    /// Reopen the canonical bytes through the source-backed shared reader.
    ///
    /// # Errors
    ///
    /// Returns an error if the immutable prepared bytes no longer satisfy the
    /// shared bounded reader.  The bytes cannot be changed through this type.
    pub fn readback(&self) -> Result<Document> {
        read_shared(Arc::clone(&self.bytes))
    }
}

#[derive(Debug)]
struct ImportedBrush {
    id: Box<str>,
    property_count: usize,
}

#[derive(Debug)]
struct CanonicalParts {
    context_ids: Vec<Box<str>>,
    brushes: Vec<ImportedBrush>,
}

const CANONICAL_ELEMENT_NAMES: &[&[u8]] = &[
    b"inkml:ink",
    b"inkml:definitions",
    b"inkml:context",
    b"inkml:inkSource",
    b"inkml:traceFormat",
    b"inkml:channel",
    b"inkml:brush",
    b"inkml:brushProperty",
    b"inkml:traceGroup",
    b"inkml:annotationXML",
    b"emma:emma",
    b"emma:interpretation",
    b"msink:context",
    b"inkml:trace",
];

/// Count caller-bounded records before the shared metadata projection owns
/// any decoded scalar values. This is intentionally a structural preflight;
/// the namespace-aware shared reader remains the authority for XML validity.
fn preflight_import_limits(bytes: &[u8], limits: AuthoringLimits) -> Result<()> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut contexts = 0usize;
    let mut brushes = 0usize;
    let mut properties = 0usize;
    let mut traces = 0usize;
    let mut in_trace = false;
    let mut trace_text_seen = false;
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("InkML import preflight failed: {error}")))?
        {
            Event::Start(element) => {
                preflight_element(
                    &element,
                    false,
                    &mut depth,
                    &mut nodes,
                    &mut contexts,
                    &mut brushes,
                    &mut properties,
                    &mut traces,
                    limits,
                )?;
                if element.name().as_ref() == b"inkml:trace" {
                    if in_trace {
                        return Err(invalid("InkML import preflight found nested traces"));
                    }
                    in_trace = true;
                    trace_text_seen = false;
                }
            },
            Event::Empty(element) => {
                preflight_element(
                    &element,
                    true,
                    &mut depth,
                    &mut nodes,
                    &mut contexts,
                    &mut brushes,
                    &mut properties,
                    &mut traces,
                    limits,
                )?;
            },
            Event::Text(text) if in_trace => {
                if trace_text_seen {
                    return Err(invalid(
                        "InkML import preflight found multiple trace text events",
                    ));
                }
                trace_text_seen = true;
                if text.as_ref().len() > limits.max_trace_text_bytes {
                    return Err(limit("InkML trace text bytes", limits.max_trace_text_bytes));
                }
            },
            Event::End(element) => {
                if !is_canonical_element_name(element.name().as_ref()) {
                    return Err(invalid(
                        "InkML canonical output contains an unsupported element QName",
                    ));
                }
                if element.name().as_ref() == b"inkml:trace" {
                    if !in_trace {
                        return Err(invalid(
                            "InkML import preflight found an unexpected trace end",
                        ));
                    }
                    if !trace_text_seen {
                        return Err(invalid("InkML import preflight found a trace without text"));
                    }
                    in_trace = false;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("InkML import preflight found an unexpected end"))?;
            },
            Event::Eof => {
                if depth != 0 || in_trace {
                    return Err(invalid(
                        "InkML import preflight found an unterminated element",
                    ));
                }
                return Ok(());
            },
            _ => {},
        }
    }
}

fn preflight_element(
    element: &BytesStart<'_>,
    empty: bool,
    depth: &mut usize,
    nodes: &mut usize,
    contexts: &mut usize,
    brushes: &mut usize,
    properties: &mut usize,
    traces: &mut usize,
    limits: AuthoringLimits,
) -> Result<()> {
    let qualified_name = element.name();
    let name = qualified_name.as_ref();
    if !is_canonical_element_name(name) {
        return Err(invalid(
            "InkML canonical output contains an unsupported element QName",
        ));
    }
    preflight_attributes(element, limits)?;
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("InkML XML nodes", limits.max_nodes))?;
    if *nodes > limits.max_nodes {
        return Err(limit("InkML XML nodes", limits.max_nodes));
    }
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| limit("InkML XML depth", limits.max_depth))?;
    if next_depth > limits.max_depth {
        return Err(limit("InkML XML depth", limits.max_depth));
    }
    if !empty {
        *depth = next_depth;
    }
    match name {
        b"msink:context" => {
            increment_preflight(contexts, limits.max_contexts, "InkML context records")?
        },
        b"inkml:brush" => increment_preflight(brushes, limits.max_brushes, "InkML brush records")?,
        b"inkml:brushProperty" => increment_preflight(
            properties,
            limits.max_brush_properties,
            "InkML brush properties",
        )?,
        b"inkml:trace" => increment_preflight(traces, limits.max_traces, "InkML trace records")?,
        _ => {},
    }
    Ok(())
}

fn increment_preflight(value: &mut usize, maximum: usize, resource: &'static str) -> Result<()> {
    *value = value
        .checked_add(1)
        .ok_or_else(|| limit(resource, maximum))?;
    if *value > maximum {
        return Err(limit(resource, maximum));
    }
    Ok(())
}

fn preflight_attributes(element: &BytesStart<'_>, limits: AuthoringLimits) -> Result<()> {
    let mut count = 0usize;
    let qualified_name = element.name();
    let element_name = qualified_name.as_ref();
    let mut has_inkml_namespace = false;
    let mut has_msink_namespace = false;
    let mut has_emma_namespace = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("InkML import attribute scan failed: {error}")))?;
        count = count.checked_add(1).ok_or_else(|| {
            limit(
                "InkML attributes per element",
                limits.max_attributes_per_element,
            )
        })?;
        if count > limits.max_attributes_per_element {
            return Err(limit(
                "InkML attributes per element",
                limits.max_attributes_per_element,
            ));
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "InkML attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let value = attribute
            .normalized_value(XmlVersion::Explicit1_0)
            .map_err(|error| {
                invalid(format!("InkML import attribute value is invalid: {error}"))
            })?;
        if value.len() > limits.max_attribute_value_bytes {
            return Err(limit(
                "InkML attribute value bytes",
                limits.max_attribute_value_bytes,
            ));
        }
        match attribute.key.as_ref() {
            b"xml:id" if value.len() > limits.max_id_bytes => {
                return Err(limit("InkML identifier bytes", limits.max_id_bytes));
            },
            b"id" if element_name == b"msink:context" && value.len() > limits.max_id_bytes => {
                return Err(limit("InkML identifier bytes", limits.max_id_bytes));
            },
            b"contextRef" | b"brushRef" if value.len() > limits.max_reference_bytes => {
                return Err(limit("InkML reference bytes", limits.max_reference_bytes));
            },
            b"xmlns:inkml" if element_name == b"inkml:ink" => {
                if has_inkml_namespace || value.as_ref() != INKML_NAMESPACE {
                    return Err(invalid(
                        "InkML import has an invalid inkml namespace binding",
                    ));
                }
                has_inkml_namespace = true;
            },
            b"xmlns:msink" if element_name == b"inkml:ink" => {
                if has_msink_namespace || value.as_ref() != NAMESPACE {
                    return Err(invalid(
                        "InkML import has an invalid msink namespace binding",
                    ));
                }
                has_msink_namespace = true;
            },
            b"xmlns:emma" if element_name == b"emma:emma" => {
                if has_emma_namespace || value.as_ref() != EMMA_NAMESPACE {
                    return Err(invalid(
                        "InkML import has an invalid emma namespace binding",
                    ));
                }
                has_emma_namespace = true;
            },
            name if name == b"xmlns" || name.starts_with(b"xmlns:") => {
                return Err(invalid(
                    "InkML import has an unexpected namespace declaration",
                ));
            },
            _ => {},
        }
    }
    if element_name == b"inkml:ink" && (!has_inkml_namespace || !has_msink_namespace) {
        return Err(invalid(
            "InkML import root namespace bindings are incomplete",
        ));
    }
    if element_name == b"emma:emma" && !has_emma_namespace {
        return Err(invalid("InkML import emma namespace binding is missing"));
    }
    Ok(())
}

fn is_canonical_element_name(name: &[u8]) -> bool {
    CANONICAL_ELEMENT_NAMES.contains(&name)
}

fn validate_import_metadata_limits(metadata: &Metadata, limits: AuthoringLimits) -> Result<()> {
    if metadata.context_count() > limits.max_contexts {
        return Err(limit("InkML context records", limits.max_contexts));
    }
    if metadata.trace_count() > limits.max_traces {
        return Err(limit("InkML trace records", limits.max_traces));
    }
    if metadata.brush_property_count() > limits.max_brush_properties {
        return Err(limit("InkML brush properties", limits.max_brush_properties));
    }
    Ok(())
}

fn import_canonical_draft(
    bytes: &[u8],
    metadata: &Metadata,
    limits: AuthoringLimits,
) -> Result<Draft> {
    let parts = parse_canonical_structure(bytes, metadata, limits)?;
    if parts.context_ids.len() > metadata.context_count() {
        return Err(invalid(
            "InkML canonical definitions contain more contexts than annotations",
        ));
    }

    let mut draft = Draft::new(limits)?;
    for (index, context) in metadata.contexts().iter().enumerate() {
        let mut value = ContextDraft::new(context.kind.clone());
        value.id = context.id.clone();
        value.semantic_type = context.semantic_type.clone();
        value.alignment_level = context.alignment_level;
        value.content_type = context.content_type;
        value.rotation_angle = context.rotation_angle;
        // The canonical writer emits definition contexts as a filtered list;
        // the annotation form carries no local XML id. Relative definition
        // order is therefore sufficient to reproduce the exact bytes.
        if let Some(id) = parts.context_ids.get(index) {
            value.xml_id = Some(id.clone());
        }
        draft = draft.context(value)?;
    }

    let mut property_index = 0usize;
    for brush in parts.brushes {
        let property_count = brush.property_count;
        let mut value = BrushDraft {
            property_text_bytes: brush.id.len(),
            id: brush.id,
            properties: Vec::new(),
        };
        for _ in 0..property_count {
            let property = metadata
                .brush_properties()
                .get(property_index)
                .ok_or_else(|| invalid("InkML canonical brush-property count mismatch"))?;
            property_index = property_index
                .checked_add(1)
                .ok_or_else(|| invalid("InkML brush-property index overflow"))?;
            let mut value_draft = BrushPropertyDraft::new(property.name.clone(), property.value())?;
            if let Some(units) = property.units() {
                value_draft = value_draft.with_units(units)?;
            }
            value = value.property(value_draft)?;
        }
        draft = draft.brush(value)?;
    }
    if property_index != metadata.brush_property_count() {
        return Err(invalid("InkML canonical brush-property count mismatch"));
    }

    for trace in metadata.traces() {
        let data = source_text(bytes, trace.data, "trace data")?;
        let mut value = TraceDraft::new(data)?;
        if let Some(reference) = trace.context_ref() {
            value = value.with_context_ref(reference)?;
        }
        if let Some(reference) = trace.brush_ref() {
            value = value.with_brush_ref(reference)?;
        }
        draft = draft.trace(value)?;
    }
    Ok(draft)
}

fn source_text<'a>(source: &'a [u8], span: SourceSpan, field: &'static str) -> Result<&'a str> {
    let bytes = source
        .get(span.range())
        .ok_or_else(|| invalid(format!("InkML {field} source span is invalid")))?;
    std::str::from_utf8(bytes).map_err(|_| invalid(format!("InkML {field} is not valid UTF-8")))
}

struct CanonicalReader<'a> {
    reader: Reader<&'a [u8]>,
}

impl<'a> CanonicalReader<'a> {
    fn new(source: &'a [u8]) -> Self {
        let mut reader = Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        Self { reader }
    }

    fn next(&mut self) -> Result<Event<'a>> {
        self.reader
            .read_event()
            .map_err(|error| invalid(format!("InkML canonical XML scan failed: {error}")))
    }

    fn expect_decl(&mut self) -> Result<()> {
        match self.next()? {
            Event::Decl(_) => Ok(()),
            _ => Err(invalid(
                "InkML canonical output must begin with an XML declaration",
            )),
        }
    }

    fn expect_start(&mut self, name: &'static [u8]) -> Result<BytesStart<'a>> {
        match self.next()? {
            Event::Start(element) if element.name().as_ref() == name => Ok(element),
            _ => Err(invalid(format!(
                "InkML canonical output is missing start element {}",
                String::from_utf8_lossy(name)
            ))),
        }
    }

    fn expect_empty(&mut self, name: &'static [u8]) -> Result<BytesStart<'a>> {
        match self.next()? {
            Event::Empty(element) if element.name().as_ref() == name => Ok(element),
            _ => Err(invalid(format!(
                "InkML canonical output is missing empty element {}",
                String::from_utf8_lossy(name)
            ))),
        }
    }

    fn expect_end(&mut self, name: &'static [u8]) -> Result<()> {
        match self.next()? {
            Event::End(element) if element.name().as_ref() == name => Ok(()),
            _ => Err(invalid(format!(
                "InkML canonical output is missing end element {}",
                String::from_utf8_lossy(name)
            ))),
        }
    }

    fn expect_text(&mut self) -> Result<()> {
        match self.next()? {
            Event::Text(_) => Ok(()),
            _ => Err(invalid("InkML canonical trace must contain one text event")),
        }
    }

    fn expect_eof(&mut self) -> Result<()> {
        match self.next()? {
            Event::Eof => Ok(()),
            _ => Err(invalid("InkML canonical output has trailing content")),
        }
    }
}

fn parse_canonical_structure(
    bytes: &[u8],
    metadata: &Metadata,
    limits: AuthoringLimits,
) -> Result<CanonicalParts> {
    let mut parser = CanonicalReader::new(bytes);
    parser.expect_decl()?;
    let root = parser.expect_start(b"inkml:ink")?;
    drop(root);
    let mut parts = CanonicalParts {
        context_ids: Vec::new(),
        brushes: Vec::new(),
    };
    parts
        .context_ids
        .try_reserve(metadata.context_count())
        .map_err(|_| invalid("InkML canonical context identifier allocation failed"))?;

    let first = parser.next()?;
    let after_optional_definitions = match first {
        Event::Start(element) if element.name().as_ref() == b"inkml:definitions" => {
            drop(element);
            parse_definitions(&mut parser, &mut parts, metadata.context_count(), limits)?;
            parser.next()?
        },
        event => event,
    };

    match after_optional_definitions {
        Event::Start(element) if element.name().as_ref() == b"inkml:traceGroup" => {
            drop(element);
            parse_trace_group(&mut parser, metadata)?;
            parser.expect_end(b"inkml:ink")?;
        },
        Event::End(element) if element.name().as_ref() == b"inkml:ink" => {},
        _ => {
            return Err(invalid(
                "InkML canonical output has an unexpected root child",
            ));
        },
    }
    parser.expect_eof()?;
    if parts.context_ids.len() > metadata.context_count() {
        return Err(invalid(
            "InkML canonical definitions contain more contexts than annotations",
        ));
    }
    if metadata.brush_property_count()
        != parts
            .brushes
            .iter()
            .try_fold(0usize, |count, brush| {
                count.checked_add(brush.property_count)
            })
            .ok_or_else(|| invalid("InkML canonical brush-property count overflow"))?
    {
        return Err(invalid("InkML canonical brush-property count mismatch"));
    }
    Ok(parts)
}

fn parse_definitions(
    parser: &mut CanonicalReader<'_>,
    parts: &mut CanonicalParts,
    expected_contexts: usize,
    limits: AuthoringLimits,
) -> Result<()> {
    let mut brush_seen = false;
    loop {
        match parser.next()? {
            Event::Start(element) if element.name().as_ref() == b"inkml:context" && !brush_seen => {
                if parts.context_ids.len() >= expected_contexts {
                    return Err(invalid(
                        "InkML canonical definitions contain more contexts than annotations",
                    ));
                }
                let id = required_attribute(&element, b"xml:id", "context XML identifier")?;
                let id = own_identifier(id, "context XML identifier", limits.max_id_bytes)?;
                drop(element);
                parse_context_definition(parser, limits)?;
                parts
                    .context_ids
                    .try_reserve(1)
                    .map_err(|_| invalid("InkML canonical context identifier allocation failed"))?;
                parts.context_ids.push(id);
            },
            Event::Start(element) if element.name().as_ref() == b"inkml:brush" => {
                brush_seen = true;
                if parts.brushes.len() >= limits.max_brushes {
                    return Err(limit("InkML brush records", limits.max_brushes));
                }
                let id = required_attribute(&element, b"xml:id", "brush identifier")?;
                let id = own_identifier(id, "brush identifier", limits.max_id_bytes)?;
                drop(element);
                let property_count = parse_brush_definition(parser, limits)?;
                parts
                    .brushes
                    .try_reserve(1)
                    .map_err(|_| invalid("InkML canonical brush allocation failed"))?;
                parts.brushes.push(ImportedBrush { id, property_count });
            },
            Event::End(element) if element.name().as_ref() == b"inkml:definitions" => return Ok(()),
            _ => {
                return Err(invalid(
                    "InkML canonical definitions contain an unexpected child",
                ));
            },
        }
    }
}

fn parse_context_definition(
    parser: &mut CanonicalReader<'_>,
    limits: AuthoringLimits,
) -> Result<()> {
    let source = parser.expect_start(b"inkml:inkSource")?;
    let source_id = required_attribute(&source, b"xml:id", "inkSource XML identifier")?;
    validate_import_identifier(&source_id, "inkSource XML identifier", limits.max_id_bytes)?;
    drop(source);
    parser.expect_start(b"inkml:traceFormat")?;
    parser.expect_empty(b"inkml:channel")?;
    parser.expect_empty(b"inkml:channel")?;
    parser.expect_end(b"inkml:traceFormat")?;
    parser.expect_end(b"inkml:inkSource")?;
    parser.expect_end(b"inkml:context")
}

fn parse_brush_definition(
    parser: &mut CanonicalReader<'_>,
    limits: AuthoringLimits,
) -> Result<usize> {
    let mut property_count = 0usize;
    loop {
        match parser.next()? {
            Event::Empty(element) if element.name().as_ref() == b"inkml:brushProperty" => {
                if property_count >= limits.max_brush_properties {
                    return Err(limit("InkML brush properties", limits.max_brush_properties));
                }
                let _ = required_attribute(&element, b"name", "brush property name")?;
                let _ = required_attribute(&element, b"value", "brush property value")?;
                property_count = property_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("InkML brush-property count overflow"))?;
            },
            Event::End(element) if element.name().as_ref() == b"inkml:brush" => {
                return Ok(property_count);
            },
            _ => {
                return Err(invalid(
                    "InkML canonical brush contains an unexpected child",
                ));
            },
        }
    }
}

fn parse_trace_group(parser: &mut CanonicalReader<'_>, metadata: &Metadata) -> Result<()> {
    for _ in 0..metadata.context_count() {
        parser.expect_start(b"inkml:annotationXML")?;
        parser.expect_start(b"emma:emma")?;
        parser.expect_start(b"emma:interpretation")?;
        parser.expect_empty(b"msink:context")?;
        parser.expect_end(b"emma:interpretation")?;
        parser.expect_end(b"emma:emma")?;
        parser.expect_end(b"inkml:annotationXML")?;
    }
    for _ in 0..metadata.trace_count() {
        let trace = parser.expect_start(b"inkml:trace")?;
        let _ = required_attribute(&trace, b"contextRef", "trace contextRef")?;
        let _ = required_attribute(&trace, b"brushRef", "trace brushRef")?;
        drop(trace);
        parser.expect_text()?;
        parser.expect_end(b"inkml:trace")?;
    }
    parser.expect_end(b"inkml:traceGroup")
}

fn required_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    field: &'static str,
) -> Result<String> {
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("InkML canonical attribute scan failed: {error}")))?;
        if attribute.key.as_ref() == name {
            if value.is_some() {
                return Err(invalid(format!("InkML canonical {field} is duplicated")));
            }
            value = Some(
                attribute
                    .normalized_value(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        invalid(format!(
                            "InkML canonical {field} is not valid XML text: {error}"
                        ))
                    })?
                    .into_owned(),
            );
        }
    }
    value.ok_or_else(|| invalid(format!("InkML canonical {field} is missing")))
}

fn validate_import_identifier(value: &str, field: &'static str, maximum: usize) -> Result<()> {
    if value.len() > maximum {
        return Err(limit("InkML identifier bytes", maximum));
    }
    if !is_ncname(value) {
        return Err(invalid(format!("InkML {field} is not an XML NCName")));
    }
    Ok(())
}

/// Typed context metadata accepted by the detached authoring writer.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ContextDraft {
    kind: ContextKind,
    id: Option<Guid>,
    xml_id: Option<Box<str>>,
    semantic_type: Option<SemanticType>,
    alignment_level: Option<i32>,
    content_type: Option<i32>,
    rotation_angle: Option<i32>,
}

impl ContextDraft {
    /// Create a context with the required InkML context kind.
    #[must_use]
    pub const fn new(kind: ContextKind) -> Self {
        Self {
            kind,
            id: None,
            xml_id: None,
            semantic_type: None,
            alignment_level: None,
            content_type: None,
            rotation_angle: None,
        }
    }

    /// Set the optional Microsoft context GUID `id` attribute.
    #[must_use]
    pub fn with_id(mut self, id: Guid) -> Self {
        self.id = Some(id);
        self
    }

    /// Set the local `xml:id` used by trace `contextRef` values.
    ///
    /// # Errors
    ///
    /// Returns an error unless `id` is a bounded XML `NCName`.
    pub fn with_xml_id(mut self, id: impl AsRef<str>) -> Result<Self> {
        self.xml_id = Some(own_identifier(
            id,
            "context XML identifier",
            MAX_AUTHORING_ID_BYTES,
        )?);
        Ok(self)
    }

    /// Set the optional semantic classification.
    #[must_use]
    pub fn with_semantic_type(mut self, value: SemanticType) -> Self {
        self.semantic_type = Some(value);
        self
    }

    /// Set the optional paragraph alignment level.
    #[must_use]
    pub const fn with_alignment_level(mut self, value: i32) -> Self {
        self.alignment_level = Some(value);
        self
    }

    /// Set the optional paragraph content type.
    #[must_use]
    pub const fn with_content_type(mut self, value: i32) -> Self {
        self.content_type = Some(value);
        self
    }

    /// Set the optional drawing rotation angle.
    #[must_use]
    pub const fn with_rotation_angle(mut self, value: i32) -> Self {
        self.rotation_angle = Some(value);
        self
    }

    /// Borrow the context kind.
    #[must_use]
    pub const fn kind(&self) -> &ContextKind {
        &self.kind
    }

    /// Borrow the optional context identifier.
    #[must_use]
    pub const fn identifier(&self) -> Option<&Guid> {
        self.id.as_ref()
    }

    /// Borrow the optional local `xml:id`.
    #[must_use]
    pub fn xml_identifier(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
}

/// Typed brush-property metadata accepted by the detached authoring writer.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BrushPropertyDraft {
    name: BrushPropertyName,
    value: Box<str>,
    units: Option<Box<str>>,
}

impl BrushPropertyDraft {
    /// Create a property with a checked name and profile-compatible value.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid custom name, invalid XML characters,
    /// an overlong value, or a known-property type mismatch. Width and height
    /// units can be supplied by [`Self::with_units`] before the property is
    /// appended to a brush.
    pub fn new(name: BrushPropertyName, value: impl AsRef<str>) -> Result<Self> {
        validate_property_name(&name)?;
        let value = value.as_ref();
        validate_brush_property_value(&name, value, None, false)?;
        let value = own_attribute_text(value, "brush property value", MAX_ATTRIBUTE_VALUE_BYTES)?;
        Ok(Self {
            name,
            value,
            units: None,
        })
    }

    /// Set the optional units attribute and validate it against the property.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid XML characters or an overlong value.
    pub fn with_units(mut self, units: impl AsRef<str>) -> Result<Self> {
        let units = units.as_ref();
        validate_brush_property_value(&self.name, &self.value, Some(units), false)?;
        self.units = Some(own_attribute_text(
            units,
            "brush property units",
            MAX_ATTRIBUTE_VALUE_BYTES,
        )?);
        Ok(self)
    }

    /// Create an `inkEffects` property from the typed effect vocabulary.
    ///
    /// # Errors
    ///
    /// Returns an error only if the effect's lexical value violates the
    /// shared attribute bound.
    pub fn ink_effect(effect: InkEffect) -> Result<Self> {
        Self::new(BrushPropertyName::InkEffects, effect.as_str())
    }

    /// Borrow the property name.
    #[must_use]
    pub const fn name(&self) -> &BrushPropertyName {
        &self.name
    }

    /// Borrow the exact property value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Borrow the optional units.
    #[must_use]
    pub fn units(&self) -> Option<&str> {
        self.units.as_deref()
    }
}

/// Typed brush record containing local properties.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BrushDraft {
    id: Box<str>,
    properties: Vec<BrushPropertyDraft>,
    property_text_bytes: usize,
}

impl BrushDraft {
    /// Create a brush with a local XML `id`.
    ///
    /// # Errors
    ///
    /// Returns an error unless `id` is a bounded XML `NCName`.
    pub fn new(id: impl AsRef<str>) -> Result<Self> {
        let id = own_identifier(id, "brush identifier", MAX_AUTHORING_ID_BYTES)?;
        let property_text_bytes = id.len();
        Ok(Self {
            id,
            properties: Vec::new(),
            property_text_bytes,
        })
    }

    /// Add a property, consuming the prior brush value.
    ///
    /// # Errors
    ///
    /// Returns an error when the shared brush-property hard bound is reached
    /// or allocation fails.
    pub fn property(mut self, property: BrushPropertyDraft) -> Result<Self> {
        if self.properties.len() >= MAX_BRUSH_PROPERTIES {
            return Err(limit("InkML brush properties", MAX_BRUSH_PROPERTIES));
        }
        let text_bytes = brush_property_text_bytes(&property)?;
        let next_text_bytes = self
            .property_text_bytes
            .checked_add(text_bytes)
            .ok_or_else(|| limit("InkML brush text bytes", MAX_SOURCE_BYTES))?;
        if next_text_bytes > MAX_SOURCE_BYTES {
            return Err(limit("InkML brush text bytes", MAX_SOURCE_BYTES));
        }
        try_reserve_one(&mut self.properties, "InkML brush properties")?;
        self.properties.push(property);
        self.property_text_bytes = next_text_bytes;
        Ok(self)
    }

    /// Borrow the local brush identifier.
    #[must_use]
    pub fn identifier(&self) -> &str {
        &self.id
    }

    /// Borrow authored properties in source order.
    #[must_use]
    pub fn properties(&self) -> &[BrushPropertyDraft] {
        &self.properties
    }
}

/// Checked InkML trace lexical data in the canonical X/Y integer-pair profile
/// and its local references.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TraceDraft {
    data: Box<str>,
    context_ref: Option<Box<str>>,
    brush_ref: Option<Box<str>>,
}

impl TraceDraft {
    /// Create a trace from comma-separated finite X/Y integer pairs with XML
    /// whitespace between coordinates, for example `"10 20, -2 4"`.
    ///
    /// The canonical writer emits a two-channel `integer` trace format, so
    /// delta, wildcard, quoted, and additional-channel forms are intentionally
    /// rejected until they have typed models.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid XML characters, non-finite or incomplete
    /// coordinate pairs, raw markup/entity syntax, or a value over the shared
    /// hard trace bound.
    pub fn new(data: impl AsRef<str>) -> Result<Self> {
        let data = own_trace_text(data, MAX_ATTRIBUTE_VALUE_BYTES)?;
        Ok(Self {
            data,
            context_ref: None,
            brush_ref: None,
        })
    }

    /// Set a local context reference such as `#ctx0`.
    ///
    /// # Errors
    ///
    /// Returns an error for external, empty, or overlong references.
    pub fn with_context_ref(mut self, reference: impl AsRef<str>) -> Result<Self> {
        self.context_ref = Some(own_reference(reference)?);
        Ok(self)
    }

    /// Set a local brush reference such as `#brush0`.
    ///
    /// # Errors
    ///
    /// Returns an error for external, empty, or overlong references.
    pub fn with_brush_ref(mut self, reference: impl AsRef<str>) -> Result<Self> {
        self.brush_ref = Some(own_reference(reference)?);
        Ok(self)
    }

    /// Borrow the checked lexical trace data.
    #[must_use]
    pub fn data(&self) -> &str {
        &self.data
    }

    /// Borrow the optional local context reference.
    #[must_use]
    pub fn context_reference(&self) -> Option<&str> {
        self.context_ref.as_deref()
    }

    /// Borrow the optional local brush reference.
    #[must_use]
    pub fn brush_reference(&self) -> Option<&str> {
        self.brush_ref.as_deref()
    }
}

trait Sink {
    fn raw(&mut self, bytes: &[u8]) -> Result<()>;

    fn escaped(&mut self, value: &str, attribute: bool) -> Result<()> {
        self.escaped_with_options(value, attribute, false)
    }

    fn escaped_preserving_xml_whitespace(&mut self, value: &str) -> Result<()> {
        self.escaped_with_options(value, true, true)
    }

    fn escaped_with_options(
        &mut self,
        value: &str,
        attribute: bool,
        preserve_xml_whitespace: bool,
    ) -> Result<()> {
        let mut segment = 0usize;
        for (index, character) in value.char_indices() {
            let replacement = match character {
                '\t' if preserve_xml_whitespace => Some(b"&#x9;".as_slice()),
                '\n' if preserve_xml_whitespace => Some(b"&#xA;".as_slice()),
                '\r' if preserve_xml_whitespace => Some(b"&#xD;".as_slice()),
                '&' => Some(b"&amp;".as_slice()),
                '<' => Some(b"&lt;".as_slice()),
                '>' => Some(b"&gt;".as_slice()),
                '"' if attribute => Some(b"&quot;".as_slice()),
                '\'' if attribute => Some(b"&apos;".as_slice()),
                _ => None,
            };
            if let Some(replacement) = replacement {
                self.raw(&value.as_bytes()[segment..index])?;
                self.raw(replacement)?;
                segment = index + character.len_utf8();
            }
        }
        self.raw(&value.as_bytes()[segment..])
    }

    fn attr(&mut self, name: &str, value: &str) -> Result<()> {
        self.attr_with_options(name, value, false)
    }

    fn attr_preserving_xml_whitespace(&mut self, name: &str, value: &str) -> Result<()> {
        self.attr_with_options(name, value, true)
    }

    fn attr_with_options(
        &mut self,
        name: &str,
        value: &str,
        preserve_xml_whitespace: bool,
    ) -> Result<()> {
        self.raw(b" ")?;
        self.raw(name.as_bytes())?;
        self.raw(b"=\"")?;
        if preserve_xml_whitespace {
            self.escaped_preserving_xml_whitespace(value)?;
        } else {
            self.escaped(value, true)?;
        }
        self.raw(b"\"")
    }
}

struct LengthSink {
    len: usize,
    max: usize,
}

impl LengthSink {
    const fn new(max: usize) -> Self {
        Self { len: 0, max }
    }
}

impl Sink for LengthSink {
    fn raw(&mut self, bytes: &[u8]) -> Result<()> {
        self.len = self
            .len
            .checked_add(bytes.len())
            .ok_or_else(|| limit("InkML output bytes", self.max))?;
        if self.len > self.max {
            return Err(limit("InkML output bytes", self.max));
        }
        Ok(())
    }
}

struct XmlSink {
    bytes: Vec<u8>,
    max: usize,
}

impl XmlSink {
    fn new(expected: usize, max: usize) -> Result<Self> {
        if expected > max {
            return Err(limit("InkML output bytes", max));
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(expected)
            .map_err(|_| invalid("InkML output allocation failed"))?;
        Ok(Self { bytes, max })
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

impl Sink for XmlSink {
    fn raw(&mut self, bytes: &[u8]) -> Result<()> {
        let next = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| limit("InkML output bytes", self.max))?;
        if next > self.max {
            return Err(limit("InkML output bytes", self.max));
        }
        if self.bytes.capacity().saturating_sub(self.bytes.len()) < bytes.len() {
            self.bytes
                .try_reserve_exact(bytes.len())
                .map_err(|_| invalid("InkML output allocation failed"))?;
        }
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }
}

fn emit_document<S: Sink>(
    draft: &Draft,
    source_ids: &[Option<Box<str>>],
    sink: &mut S,
) -> Result<()> {
    sink.raw(
        br#"<?xml version="1.0" encoding="UTF-8"?><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:msink="http://schemas.microsoft.com/ink/2010/main">"#,
    )?;
    if draft
        .contexts
        .iter()
        .any(|context| context.xml_id.is_some())
        || !draft.brushes.is_empty()
    {
        sink.raw(b"<inkml:definitions>")?;
        for (index, context) in draft.contexts.iter().enumerate() {
            if let Some(id) = context.xml_id.as_deref() {
                let source_id = source_ids
                    .get(index)
                    .and_then(Option::as_deref)
                    .ok_or_else(|| invalid("InkML generated source identifier is missing"))?;
                sink.raw(b"<inkml:context")?;
                sink.attr("xml:id", id)?;
                sink.raw(b"><inkml:inkSource")?;
                sink.attr("xml:id", source_id)?;
                sink.raw(b"><inkml:traceFormat>")?;
                sink.raw(br#"<inkml:channel name="X" type="integer"/><inkml:channel name="Y" type="integer"/>"#)?;
                sink.raw(b"</inkml:traceFormat></inkml:inkSource></inkml:context>")?;
            }
        }
        for brush in &draft.brushes {
            sink.raw(b"<inkml:brush")?;
            sink.attr("xml:id", &brush.id)?;
            sink.raw(b">")?;
            for property in &brush.properties {
                let preserve_custom_lexicals =
                    matches!(property.name(), BrushPropertyName::Custom(_));
                sink.raw(b"<inkml:brushProperty")?;
                if preserve_custom_lexicals {
                    sink.attr_preserving_xml_whitespace("name", property.name.as_str())?;
                    sink.attr_preserving_xml_whitespace("value", &property.value)?;
                } else {
                    sink.attr("name", property.name.as_str())?;
                    sink.attr("value", &property.value)?;
                }
                if let Some(units) = property.units.as_deref() {
                    if preserve_custom_lexicals {
                        sink.attr_preserving_xml_whitespace("units", units)?;
                    } else {
                        sink.attr("units", units)?;
                    }
                }
                sink.raw(b"/>")?;
            }
            sink.raw(b"</inkml:brush>")?;
        }
        sink.raw(b"</inkml:definitions>")?;
    }
    if !draft.contexts.is_empty() || !draft.traces.is_empty() {
        sink.raw(b"<inkml:traceGroup>")?;
        for context in &draft.contexts {
            sink.raw(b"<inkml:annotationXML>")?;
            sink.raw(b"<emma:emma")?;
            sink.attr("xmlns:emma", EMMA_NAMESPACE)?;
            sink.attr("version", "1.0")?;
            sink.raw(b"><emma:interpretation")?;
            sink.attr("emma:mode", "ink")?;
            sink.raw(b"><msink:context")?;
            if let Some(id) = context.id.as_ref() {
                sink.attr("id", id.as_str())?;
            }
            sink.attr("type", context.kind.as_str())?;
            if let Some(value) = context.semantic_type.as_ref() {
                let value = value.as_string();
                sink.attr("semanticType", &value)?;
            }
            if let Some(value) = context.alignment_level {
                let value = value.to_string();
                sink.attr("alignmentLevel", &value)?;
            }
            if let Some(value) = context.content_type {
                let value = value.to_string();
                sink.attr("contentType", &value)?;
            }
            if let Some(value) = context.rotation_angle {
                let value = value.to_string();
                sink.attr("rotationAngle", &value)?;
            }
            sink.raw(b"/></emma:interpretation></emma:emma>")?;
            sink.raw(b"</inkml:annotationXML>")?;
        }
        for trace in &draft.traces {
            sink.raw(b"<inkml:trace")?;
            if let Some(reference) = trace.context_ref.as_deref() {
                sink.attr("contextRef", reference)?;
            }
            if let Some(reference) = trace.brush_ref.as_deref() {
                sink.attr("brushRef", reference)?;
            }
            sink.raw(b">")?;
            sink.escaped(&trace.data, false)?;
            sink.raw(b"</inkml:trace>")?;
        }
        sink.raw(b"</inkml:traceGroup>")?;
    }
    sink.raw(b"</inkml:ink>")
}

fn validate_context(context: &ContextDraft, limits: AuthoringLimits) -> Result<()> {
    if let Some(id) = context.id.as_ref()
        && id.as_str().len() > limits.max_id_bytes
    {
        return Err(limit("InkML identifier bytes", limits.max_id_bytes));
    }
    if let Some(id) = context.id.as_ref() {
        validate_attribute_length(id.as_str(), limits)?;
    }
    if let Some(id) = context.xml_id.as_deref() {
        if id.len() > limits.max_id_bytes {
            return Err(limit("InkML identifier bytes", limits.max_id_bytes));
        }
        validate_attribute_length(id, limits)?;
    }
    validate_attribute_length(context.kind.as_str(), limits)?;
    if let Some(value) = context.semantic_type.as_ref() {
        let value = value.as_string();
        validate_attribute_length(&value, limits)?;
    }
    if (context.alignment_level.is_some() || context.content_type.is_some())
        && context.kind.as_str() != "paragraph"
    {
        return Err(invalid(
            "InkML alignmentLevel/contentType require a paragraph context",
        ));
    }
    if context.rotation_angle.is_some()
        && !matches!(
            context.kind.as_str(),
            "inkDrawing" | "nonInkDrawing" | "mixedDrawing"
        )
    {
        return Err(invalid("InkML rotationAngle requires a drawing context"));
    }
    if context.semantic_type.is_some()
        && !matches!(
            context.kind.as_str(),
            "writingRegion" | "inkDrawing" | "nonInkDrawing" | "mixedDrawing"
        )
    {
        return Err(invalid(
            "InkML semanticType requires a writing or drawing context",
        ));
    }
    for value in [
        context.alignment_level,
        context.content_type,
        context.rotation_angle,
    ]
    .into_iter()
    .flatten()
    {
        let value = value.to_string();
        validate_attribute_length(&value, limits)?;
    }
    let mut attributes = 1; // type
    attributes += usize::from(context.id.is_some());
    attributes += usize::from(context.semantic_type.is_some());
    attributes += usize::from(context.alignment_level.is_some());
    attributes += usize::from(context.content_type.is_some());
    attributes += usize::from(context.rotation_angle.is_some());
    if attributes > limits.max_attributes_per_element {
        return Err(limit(
            "InkML attributes per element",
            limits.max_attributes_per_element,
        ));
    }
    Ok(())
}

fn validate_brush(brush: &BrushDraft, limits: AuthoringLimits) -> Result<()> {
    if brush.id.len() > limits.max_id_bytes {
        return Err(limit("InkML identifier bytes", limits.max_id_bytes));
    }
    validate_attribute_length(&brush.id, limits)?;
    if brush.properties.len() > limits.max_brush_properties {
        return Err(limit("InkML brush properties", limits.max_brush_properties));
    }
    if brush.properties.len() > limits.max_attributes_per_element {
        return Err(limit(
            "InkML brush properties",
            limits.max_attributes_per_element,
        ));
    }
    for property in &brush.properties {
        validate_brush_property(property)?;
        validate_property_name(&property.name)?;
        validate_attribute_length(property.name.as_str(), limits)?;
        if property.value.len() > limits.max_attribute_value_bytes
            || property
                .units
                .as_ref()
                .is_some_and(|units| units.len() > limits.max_attribute_value_bytes)
        {
            return Err(limit(
                "InkML attribute value bytes",
                limits.max_attribute_value_bytes,
            ));
        }
        validate_attribute_length(&property.value, limits)?;
        if let Some(units) = property.units.as_ref() {
            validate_attribute_length(units, limits)?;
        }
        let attributes = 2 + usize::from(property.units.is_some());
        if attributes > limits.max_attributes_per_element {
            return Err(limit(
                "InkML attributes per element",
                limits.max_attributes_per_element,
            ));
        }
    }
    Ok(())
}

fn retained_text_limit(limits: AuthoringLimits) -> usize {
    limits.max_source_bytes.min(limits.max_output_bytes)
}

fn ensure_retained_input(current: usize, incoming: usize, limits: AuthoringLimits) -> Result<()> {
    let retained_limit = retained_text_limit(limits);
    let next = current
        .checked_add(incoming)
        .ok_or_else(|| limit("InkML retained authoring text", retained_limit))?;
    if next > retained_limit {
        return Err(limit("InkML retained authoring text", retained_limit));
    }
    Ok(())
}

fn context_text_bytes(context: &ContextDraft) -> Result<usize> {
    let mut bytes = context.kind.as_str().len();
    if let Some(id) = context.id.as_ref() {
        bytes = checked_text_add(bytes, id.as_str().len())?;
    }
    if let Some(id) = context.xml_id.as_deref() {
        bytes = checked_text_add(bytes, id.len())?;
    }
    if let Some(semantic_type) = context.semantic_type.as_ref() {
        bytes = checked_text_add(bytes, semantic_type_text_bytes(semantic_type))?;
    }
    for value in [
        context.alignment_level,
        context.content_type,
        context.rotation_angle,
    ]
    .into_iter()
    .flatten()
    {
        bytes = checked_text_add(bytes, value.to_string().len())?;
    }
    Ok(bytes)
}

fn brush_text_bytes(brush: &BrushDraft) -> usize {
    brush.property_text_bytes
}

fn brush_property_text_bytes(property: &BrushPropertyDraft) -> Result<usize> {
    let mut bytes = property.name.as_str().len();
    bytes = checked_text_add(bytes, property.value.len())?;
    if let Some(units) = property.units.as_deref() {
        bytes = checked_text_add(bytes, units.len())?;
    }
    Ok(bytes)
}

fn trace_text_bytes(trace: &TraceDraft) -> Result<usize> {
    let mut bytes = trace.data.len();
    if let Some(reference) = trace.context_ref.as_deref() {
        bytes = checked_text_add(bytes, reference.len())?;
    }
    if let Some(reference) = trace.brush_ref.as_deref() {
        bytes = checked_text_add(bytes, reference.len())?;
    }
    Ok(bytes)
}

fn checked_text_add(current: usize, incoming: usize) -> Result<usize> {
    current
        .checked_add(incoming)
        .ok_or_else(|| invalid("InkML retained authoring text length overflow"))
}

fn semantic_type_text_bytes(value: &SemanticType) -> usize {
    match value {
        SemanticType::None => 4,
        SemanticType::Underline => 9,
        SemanticType::Strikethrough => 13,
        SemanticType::Highlight => 9,
        SemanticType::ScratchOut => 9,
        SemanticType::VerticalRange => 13,
        SemanticType::Callout => 7,
        SemanticType::Enclosure => 9,
        SemanticType::Comment => 7,
        SemanticType::Container => 9,
        SemanticType::Connector => 9,
        SemanticType::Custom(value) => decimal_digit_count(*value),
    }
}

fn decimal_digit_count(value: u32) -> usize {
    let mut remaining = value;
    let mut digits = 1;
    while remaining >= 10 {
        remaining /= 10;
        digits += 1;
    }
    digits
}

fn validate_trace(trace: &TraceDraft, limits: AuthoringLimits) -> Result<()> {
    if trace.data.len() > limits.max_trace_text_bytes {
        return Err(limit("InkML trace text bytes", limits.max_trace_text_bytes));
    }
    for reference in [&trace.context_ref, &trace.brush_ref].into_iter().flatten() {
        if reference.len() > limits.max_reference_bytes {
            return Err(limit("InkML reference bytes", limits.max_reference_bytes));
        }
        validate_attribute_length(reference, limits)?;
    }
    let attributes =
        usize::from(trace.context_ref.is_some()) + usize::from(trace.brush_ref.is_some());
    if attributes > limits.max_attributes_per_element {
        return Err(limit(
            "InkML attributes per element",
            limits.max_attributes_per_element,
        ));
    }
    Ok(())
}

fn validate_attribute_length(value: &str, limits: AuthoringLimits) -> Result<()> {
    if value.len() > limits.max_attribute_value_bytes {
        return Err(limit(
            "InkML attribute value bytes",
            limits.max_attribute_value_bytes,
        ));
    }
    Ok(())
}

fn validate_property_name(name: &BrushPropertyName) -> Result<()> {
    if name.as_str().len() > MAX_TOKEN_BYTES || !name.as_str().chars().all(is_xml10_character) {
        return Err(invalid("InkML brush property name is not valid XML text"));
    }
    Ok(())
}

fn validate_brush_property(property: &BrushPropertyDraft) -> Result<()> {
    validate_brush_property_value(
        &property.name,
        &property.value,
        property.units.as_deref(),
        true,
    )
}

fn validate_brush_property_value(
    name: &BrushPropertyName,
    value: &str,
    units: Option<&str>,
    require_units: bool,
) -> Result<()> {
    let invalid_value = || invalid("InkML brush property value is not profile-compatible");
    match name {
        BrushPropertyName::Width | BrushPropertyName::Height => {
            if !is_xsd_decimal(value)
                || (require_units && units.is_none())
                || units.is_some_and(|unit| !is_length_unit(unit))
            {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::Color => {
            if !is_rgb_hex(value) || units.is_some() {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::Transparency => {
            if !is_bounded_integer(value, 0, 255) || units.is_some() {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::Tip => {
            if !matches!(value, "ellipse" | "rectangle") || units.is_some() {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::RasterOp => {
            if !matches!(
                value,
                "black"
                    | "copyPen"
                    | "maskNotPen"
                    | "maskPenNot"
                    | "maskPen"
                    | "mergeNotPen"
                    | "mergePen"
                    | "mergePenNot"
                    | "noOperation"
                    | "not"
                    | "notCopyPen"
                    | "notMaskPen"
                    | "notMergePen"
                    | "notXOrPen"
                    | "white"
                    | "xOrPen"
            ) || units.is_some()
            {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::AntiAliased
        | BrushPropertyName::FitToCurve
        | BrushPropertyName::IgnorePressure => {
            if !is_xsd_boolean(value) || units.is_some() {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::InkEffects => {
            if !matches!(
                value,
                "none"
                    | "pencil"
                    | "rainbow"
                    | "galaxy"
                    | "gold"
                    | "silver"
                    | "lava"
                    | "ocean"
                    | "rosegold"
                    | "bronze"
            ) || units.is_some()
            {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::AnchorX
        | BrushPropertyName::AnchorY
        | BrushPropertyName::ScaleFactor => {
            if !is_xsd_decimal(value) || units.is_some() {
                return Err(invalid_value());
            }
        },
        BrushPropertyName::Custom(_) => {
            // The profile permits arbitrary names but ignores unknown
            // properties. Preserve their bounded lexical value without
            // pretending to know a future property's scalar type.
            if units.is_some_and(|unit| !unit.chars().all(is_xml10_character)) {
                return Err(invalid_value());
            }
        },
    }
    Ok(())
}

fn is_xsd_decimal(value: &str) -> bool {
    let collapsed = collapse_xsd_whitespace(value);
    let value = collapsed.as_ref();
    let value = value.strip_prefix(['+', '-']).unwrap_or(value);
    if value.is_empty() {
        return false;
    }
    let mut parts = value.split('.');
    let integer = parts.next().unwrap_or_default();
    let fraction = parts.next();
    if parts.next().is_some() {
        return false;
    }
    (integer.chars().all(|character| character.is_ascii_digit())
        && (fraction.is_none()
            || fraction
                .is_some_and(|value| value.chars().all(|character| character.is_ascii_digit()))))
        && (fraction.is_some() || !integer.is_empty())
        && (fraction.is_none_or(|value| !value.is_empty()) || !integer.is_empty())
}

fn is_xsd_boolean(value: &str) -> bool {
    let collapsed = collapse_xsd_whitespace(value);
    matches!(collapsed.as_ref(), "true" | "false" | "1" | "0")
}

fn is_bounded_integer(value: &str, minimum: i32, maximum: i32) -> bool {
    let collapsed = collapse_xsd_whitespace(value);
    let value = collapsed.as_ref();
    let digits = value.strip_prefix(['+', '-']).unwrap_or(value);
    !digits.is_empty()
        && digits.chars().all(|character| character.is_ascii_digit())
        && value
            .parse::<i32>()
            .is_ok_and(|value| (minimum..=maximum).contains(&value))
}

fn is_rgb_hex(value: &str) -> bool {
    value.len() == 7
        && value.starts_with('#')
        && value[1..]
            .chars()
            .all(|character| character.is_ascii_hexdigit())
}

fn is_length_unit(value: &str) -> bool {
    matches!(value, "m" | "cm" | "mm" | "in" | "pt" | "pc" | "em" | "ex")
}

fn collapse_xsd_whitespace(value: &str) -> Cow<'_, str> {
    if !value.chars().any(is_xsd_whitespace) {
        return Cow::Borrowed(value);
    }

    let mut collapsed = String::with_capacity(value.len());
    let mut pending_space = false;
    for character in value.chars() {
        if is_xsd_whitespace(character) {
            pending_space = true;
            continue;
        }
        if pending_space && !collapsed.is_empty() {
            collapsed.push(' ');
        }
        collapsed.push(character);
        pending_space = false;
    }
    Cow::Owned(collapsed)
}

const fn is_xsd_whitespace(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\r' | '\n')
}

fn own_identifier(value: impl AsRef<str>, field: &'static str, maximum: usize) -> Result<Box<str>> {
    let value = value.as_ref();
    if value.len() > maximum || !is_ncname(value) {
        return Err(invalid(format!(
            "InkML {field} is not a bounded XML NCName"
        )));
    }
    own_boxed(value, field)
}

fn own_attribute_text(
    value: impl AsRef<str>,
    field: &'static str,
    maximum: usize,
) -> Result<Box<str>> {
    let value = value.as_ref();
    if value.len() > maximum || !value.chars().all(is_xml10_character) {
        return Err(invalid(format!("InkML {field} is not valid XML text")));
    }
    own_boxed(value, field)
}

fn own_trace_text(value: impl AsRef<str>, maximum: usize) -> Result<Box<str>> {
    let value = value.as_ref();
    if value.len() > maximum
        || !value.chars().all(is_xml10_character)
        || value.contains('<')
        || value.contains('&')
    {
        return Err(invalid(
            "InkML trace data must be bounded XML lexical text without markup or entities",
        ));
    }
    validate_trace_points(value)?;
    own_boxed(value, "trace data")
}

fn validate_trace_points(value: &str) -> Result<()> {
    let mut points = 0usize;
    for point in value.split(',') {
        let mut coordinates = point.split([' ', '\t', '\n', '\r']);
        let x = coordinates
            .find(|token| !token.is_empty())
            .ok_or_else(|| invalid("InkML trace data must not contain empty points"))?;
        let y = coordinates
            .find(|token| !token.is_empty())
            .ok_or_else(|| invalid("InkML trace data must contain complete X/Y pairs"))?;
        if coordinates.any(|token| !token.is_empty()) {
            return Err(invalid(
                "InkML trace data must contain exactly two channels per point",
            ));
        }
        if !is_finite_coordinate(x) || !is_finite_coordinate(y) {
            return Err(invalid(
                "InkML trace data must contain signed decimal integer coordinates",
            ));
        }
        points = points
            .checked_add(1)
            .ok_or_else(|| invalid("InkML trace point count overflow"))?;
    }
    if points == 0 {
        return Err(invalid(
            "InkML trace data must contain at least one comma-delimited X/Y point",
        ));
    }
    Ok(())
}

fn is_finite_coordinate(value: &str) -> bool {
    let digits = value.strip_prefix('-').unwrap_or(value);
    !digits.is_empty()
        && digits.chars().all(|character| character.is_ascii_digit())
        && value.parse::<i64>().is_ok()
}

fn own_reference(value: impl AsRef<str>) -> Result<Box<str>> {
    let value = value.as_ref();
    if value.len() > MAX_AUTHORING_REFERENCE_BYTES {
        return Err(limit(
            "InkML reference bytes",
            MAX_AUTHORING_REFERENCE_BYTES,
        ));
    }
    if !value.starts_with('#') || value.len() == 1 {
        return Err(invalid(
            "InkML authoring references must target a local # identifier",
        ));
    }
    let target = &value[1..];
    if !is_ncname(target) {
        return Err(invalid(
            "InkML local reference target is not a valid identifier",
        ));
    }
    own_boxed(value, "local reference")
}

fn own_boxed(value: &str, field: &str) -> Result<Box<str>> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_| invalid(format!("InkML {field} allocation failed")))?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn reference_target(reference: &str) -> Result<&str> {
    reference
        .strip_prefix('#')
        .filter(|target| !target.is_empty())
        .ok_or_else(|| invalid("InkML local reference target is empty"))
}

fn has_duplicate(values: &[&str]) -> bool {
    values.windows(2).any(|pair| pair[0] == pair[1])
}

fn try_reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_| invalid(format!("{resource} allocation failed")))
}

fn is_xml10_character(value: char) -> bool {
    matches!(
        value,
        '\u{9}'
            | '\u{A}'
            | '\u{D}'
            | '\u{20}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
    )
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

#[cfg(test)]
mod tests {
    use super::*;

    const CONTEXT_ID: &str = "{8646EB18-6E67-4FFA-8739-E20C3C1A0F80}";

    fn context() -> ContextDraft {
        ContextDraft::new(ContextKind::WritingRegion)
            .with_id(Guid::new(CONTEXT_ID).expect("test GUID"))
            .with_xml_id("ctx0")
            .expect("test context XML identifier")
            .with_semantic_type(SemanticType::Comment)
    }

    fn brush() -> BrushDraft {
        BrushDraft::new("brush0")
            .expect("test brush")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Width, "0.06667")
                    .expect("test property")
                    .with_units("cm")
                    .expect("test units"),
            )
            .expect("test brush property")
            .property(
                BrushPropertyDraft::ink_effect(InkEffect::Pencil).expect("test effect property"),
            )
            .expect("test brush property")
    }

    fn trace() -> TraceDraft {
        TraceDraft::new("1 2, 3 4")
            .expect("test trace")
            .with_context_ref("#ctx0")
            .expect("test context reference")
            .with_brush_ref("#brush0")
            .expect("test brush reference")
    }

    fn all_typed_fields_draft() -> Draft {
        let paragraph = ContextDraft::new(ContextKind::Paragraph)
            .with_id(Guid::new(CONTEXT_ID).expect("test GUID"))
            .with_xml_id("paragraph0")
            .expect("paragraph identifier")
            .with_alignment_level(2)
            .with_content_type(3);
        let drawing = ContextDraft::new(ContextKind::InkDrawing)
            .with_xml_id("drawing0")
            .expect("drawing identifier")
            .with_semantic_type(SemanticType::Underline)
            .with_rotation_angle(45);
        let brush = BrushDraft::new("allBrush")
            .expect("brush")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Width, "0.1")
                    .expect("width")
                    .with_units("cm")
                    .expect("width units"),
            )
            .expect("width property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Height, "0.2")
                    .expect("height")
                    .with_units("mm")
                    .expect("height units"),
            )
            .expect("height property")
            .property(BrushPropertyDraft::new(BrushPropertyName::Color, "#123456").expect("color"))
            .expect("color property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Transparency, "128")
                    .expect("transparency"),
            )
            .expect("transparency property")
            .property(BrushPropertyDraft::new(BrushPropertyName::Tip, "ellipse").expect("tip"))
            .expect("tip property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::RasterOp, "copyPen").expect("raster op"),
            )
            .expect("raster op property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::AntiAliased, "true")
                    .expect("anti-aliased"),
            )
            .expect("anti-aliased property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::FitToCurve, "false")
                    .expect("fit to curve"),
            )
            .expect("fit to curve property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::IgnorePressure, "1")
                    .expect("ignore pressure"),
            )
            .expect("ignore pressure property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::InkEffects, "pencil")
                    .expect("ink effect"),
            )
            .expect("ink effect property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::AnchorX, "1.25").expect("anchor x"),
            )
            .expect("anchor x property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::AnchorY, "-2.5").expect("anchor y"),
            )
            .expect("anchor y property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::ScaleFactor, "0.75").expect("scale"),
            )
            .expect("scale property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Custom("future".into()), "opaque")
                    .expect("custom property"),
            )
            .expect("custom property");
        Draft::new(AuthoringLimits::default())
            .expect("limits")
            .context(paragraph)
            .expect("paragraph")
            .context(drawing)
            .expect("drawing")
            .brush(brush)
            .expect("all properties")
            .trace(
                TraceDraft::new("1 2, -3 4")
                    .expect("trace")
                    .with_context_ref("#paragraph0")
                    .expect("context reference")
                    .with_brush_ref("#allBrush")
                    .expect("brush reference"),
            )
            .expect("trace")
    }

    #[test]
    fn canonical_output_reopens_and_shares_bytes() {
        let draft = Draft::new(AuthoringLimits::default())
            .expect("limits")
            .context(context())
            .expect("context")
            .brush(brush())
            .expect("brush")
            .trace(trace())
            .expect("trace");
        let prepared = draft.finish().expect("canonical InkML");
        assert_eq!(prepared.metadata().context_count(), 1);
        assert_eq!(prepared.metadata().trace_count(), 1);
        assert_eq!(prepared.metadata().brush_property_count(), 2);
        assert!(
            prepared
                .as_bytes()
                .windows(b"xmlns:inkml=\"http://www.w3.org/2003/InkML\"".len())
                .any(|window| window == b"xmlns:inkml=\"http://www.w3.org/2003/InkML\"")
        );
        assert!(
            prepared
                .as_bytes()
                .windows(b"xml:id=\"ctx0\"".len())
                .any(|window| { window == b"xml:id=\"ctx0\"" })
        );
        assert!(
            prepared
                .as_bytes()
                .windows(b"emma:mode=\"ink\"".len())
                .any(|window| { window == b"emma:mode=\"ink\"" })
        );
        assert!(
            prepared
                .as_bytes()
                .windows(b"<inkml:channel name=\"X\" type=\"integer\"/>".len())
                .any(|window| window == b"<inkml:channel name=\"X\" type=\"integer\"/>")
        );
        assert!(
            prepared
                .as_bytes()
                .windows(b"<inkml:inkSource xml:id=\"inkSrc0\">".len())
                .any(|window| window == b"<inkml:inkSource xml:id=\"inkSrc0\">")
        );
        assert!(
            prepared
                .as_bytes()
                .windows(b"xml:id=\"brush0\"".len())
                .any(|window| { window == b"xml:id=\"brush0\"" })
        );

        let document = prepared.readback().expect("prepared readback");
        assert_eq!(document.metadata(), prepared.metadata());
        assert_eq!(document.traces()[0].data(&document), b"1 2, 3 4");
        let shared = prepared.shared_source();
        assert!(Arc::ptr_eq(&shared, &document.source));
    }

    #[test]
    fn finish_reopens_schema_whitespace_values_without_rewriting_source() {
        let raw_decimal = " \t1.2\r\n";
        let raw_integer = "\n42\t";
        let raw_boolean = " \t1\r\n";
        let brush = BrushDraft::new("whitespace")
            .expect("brush")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Width, raw_decimal)
                    .expect("decimal")
                    .with_units("m")
                    .expect("metre unit"),
            )
            .expect("width property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Transparency, raw_integer)
                    .expect("integer"),
            )
            .expect("transparency property")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::AntiAliased, raw_boolean)
                    .expect("boolean"),
            )
            .expect("boolean property");
        let prepared = Draft::default()
            .brush(brush)
            .expect("brush")
            .finish()
            .expect("schema whitespace is semantically equivalent on readback");
        let document = prepared.readback().expect("prepared readback");

        assert_eq!(document.source(), prepared.as_bytes());
        for raw in [raw_decimal, raw_integer, raw_boolean] {
            let attribute = format!("value=\"{raw}\"");
            assert!(
                prepared
                    .as_bytes()
                    .windows(attribute.len())
                    .any(|window| window == attribute.as_bytes()),
                "emitted source retained raw value {raw:?}"
            );
        }

        let width = document.brush_properties()[0]
            .effective()
            .expect("effective width");
        assert_eq!(width.value(), "1.2");
        assert_eq!(width.units(), Some("m"));
        assert!(!width.defaulted());
        assert_eq!(
            document.brush_properties()[1]
                .effective()
                .expect("effective transparency")
                .value(),
            "42"
        );
        assert_eq!(
            document.brush_properties()[2]
                .effective()
                .expect("effective anti-aliased")
                .value(),
            "true"
        );
    }

    #[test]
    fn canonical_bytes_import_with_all_typed_fields() {
        let prepared = all_typed_fields_draft().finish().expect("canonical InkML");
        let imported = Prepared::from_bytes(prepared.as_bytes()).expect("canonical import");
        assert_eq!(imported.as_bytes(), prepared.as_bytes());
        assert_eq!(imported.metadata(), prepared.metadata());
        assert_eq!(imported.metadata().context_count(), 2);
        assert_eq!(imported.metadata().trace_count(), 1);
        assert_eq!(imported.metadata().brush_property_count(), 14);
        let document = imported.readback().expect("readback");
        assert_eq!(document.traces()[0].data(&document), b"1 2, -3 4");

        let empty = Draft::default().finish().expect("empty canonical InkML");
        let empty_import = Prepared::from_bytes(empty.as_bytes()).expect("empty import");
        assert_eq!(empty_import.as_bytes(), empty.as_bytes());
    }

    #[test]
    fn canonical_import_rejects_malformed_noncanonical_and_limits() {
        let prepared = Draft::new(AuthoringLimits::default())
            .expect("limits")
            .context(context())
            .expect("context")
            .brush(brush())
            .expect("brush")
            .trace(trace())
            .expect("trace")
            .finish()
            .expect("canonical InkML");
        assert!(
            Prepared::from_bytes(&prepared.as_bytes()[..prepared.as_bytes().len() - 1]).is_err()
        );

        let mut noncanonical = prepared.as_bytes().to_vec();
        let channel = b"name=\"X\"";
        let position = noncanonical
            .windows(channel.len())
            .position(|window| window == channel)
            .expect("canonical channel");
        noncanonical[position + channel.len() - 2] = b'Q';
        assert!(Prepared::from_bytes(&noncanonical).is_err());

        let exact_limits = AuthoringLimits {
            max_source_bytes: prepared.as_bytes().len(),
            max_output_bytes: prepared.as_bytes().len(),
            ..AuthoringLimits::default()
        };
        assert!(Prepared::from_bytes_with_limits(prepared.as_bytes(), exact_limits).is_ok());
        let below_source = AuthoringLimits {
            max_source_bytes: prepared.as_bytes().len() - 1,
            ..exact_limits
        };
        assert!(Prepared::from_bytes_with_limits(prepared.as_bytes(), below_source).is_err());
        let below_trace = AuthoringLimits {
            max_trace_text_bytes: 7,
            ..exact_limits
        };
        assert!(Prepared::from_bytes_with_limits(prepared.as_bytes(), below_trace).is_err());
    }

    #[test]
    fn canonical_import_preflight_rejects_aliases_depth_and_scalar_quotas() {
        let alias = br#"<m:ink xmlns:m="http://www.w3.org/2003/InkML"></m:ink>"#;
        assert!(preflight_import_limits(alias, AuthoringLimits::default()).is_err());

        let rebound = br#"<inkml:ink xmlns:inkml="http://schemas.microsoft.com/ink/2010/main" xmlns:msink="http://schemas.microsoft.com/ink/2010/main"></inkml:ink>"#;
        assert!(preflight_import_limits(rebound, AuthoringLimits::default()).is_err());

        let empty_depth = b"<inkml:ink><inkml:channel/></inkml:ink>";
        let depth_limits = AuthoringLimits {
            max_depth: 1,
            ..AuthoringLimits::default()
        };
        assert!(preflight_import_limits(empty_depth, depth_limits).is_err());

        let oversized_attribute = br#"<inkml:ink data="123"></inkml:ink>"#;
        let attribute_limits = AuthoringLimits {
            max_attribute_value_bytes: 2,
            ..AuthoringLimits::default()
        };
        assert!(preflight_import_limits(oversized_attribute, attribute_limits).is_err());

        let oversized_identifier = br#"<inkml:ink><inkml:context xml:id="long"/></inkml:ink>"#;
        let identifier_limits = AuthoringLimits {
            max_id_bytes: 3,
            ..AuthoringLimits::default()
        };
        assert!(preflight_import_limits(oversized_identifier, identifier_limits).is_err());

        let oversized_context_guid = br#"<inkml:ink><msink:context id="long"/></inkml:ink>"#;
        assert!(preflight_import_limits(oversized_context_guid, identifier_limits).is_err());

        let oversized_reference = br##"<inkml:ink><inkml:trace contextRef="#long" brushRef="#b">1 2</inkml:trace></inkml:ink>"##;
        let reference_limits = AuthoringLimits {
            max_reference_bytes: 3,
            ..AuthoringLimits::default()
        };
        assert!(preflight_import_limits(oversized_reference, reference_limits).is_err());

        let oversized_trace = b"<inkml:ink><inkml:trace>123</inkml:trace></inkml:ink>";
        let trace_limits = AuthoringLimits {
            max_trace_text_bytes: 2,
            ..AuthoringLimits::default()
        };
        assert!(preflight_import_limits(oversized_trace, trace_limits).is_err());
    }

    #[test]
    fn canonical_import_uses_decoded_attribute_quota_for_escaped_values() {
        let escaped = "&".repeat(64);
        let limits = AuthoringLimits {
            max_attribute_value_bytes: 64,
            ..AuthoringLimits::default()
        };
        let brush = BrushDraft::new("escapedBrush")
            .expect("brush")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Custom("future".into()), &escaped)
                    .expect("escaped custom value"),
            )
            .expect("property");
        let prepared = Draft::new(limits)
            .expect("limits")
            .brush(brush)
            .expect("brush")
            .finish()
            .expect("canonical escaped attribute value");
        assert!(
            prepared
                .as_bytes()
                .windows(b"&amp;".len())
                .any(|window| { window == b"&amp;" })
        );
        assert!(Prepared::from_bytes_with_limits(prepared.as_bytes(), limits).is_ok());
    }

    #[test]
    fn custom_brush_property_xml_escaping_preserves_attribute_whitespace() {
        let name = "future\tname\r";
        let value = "\tcustom\r\nvalue";
        let units = "u\nit\t";
        let brush = BrushDraft::new("customWhitespace")
            .expect("brush")
            .property(
                BrushPropertyDraft::new(BrushPropertyName::Custom(name.into()), value)
                    .expect("custom property")
                    .with_units(units)
                    .expect("custom units"),
            )
            .expect("property");
        let prepared = Draft::default()
            .brush(brush)
            .expect("brush")
            .finish()
            .expect("custom XML whitespace is escaped as character references");
        let document = prepared.readback().expect("prepared readback");
        let property = &document.brush_properties()[0];

        assert_eq!(property.name().as_str(), name);
        assert_eq!(property.value(), value);
        assert_eq!(property.units(), Some(units));
        assert!(
            prepared
                .as_bytes()
                .windows(b"&#x9;".len())
                .any(|window| window == b"&#x9;")
        );
        assert!(
            prepared
                .as_bytes()
                .windows(b"&#xA;".len())
                .any(|window| window == b"&#xA;")
        );
        assert!(
            prepared
                .as_bytes()
                .windows(b"&#xD;".len())
                .any(|window| window == b"&#xD;")
        );

        let imported = Prepared::from_bytes(prepared.as_bytes()).expect("custom readback import");
        assert_eq!(imported.as_bytes(), prepared.as_bytes());
    }

    #[test]
    fn rejects_invalid_trace_lexical_text_and_external_refs() {
        assert!(TraceDraft::new("<inkml:evil/>").is_err());
        assert!(TraceDraft::new("&amp;").is_err());
        assert!(TraceDraft::new("garbage").is_err());
        assert!(TraceDraft::new("1 2 3 4").is_err());
        assert!(TraceDraft::new("1 2 3").is_err());
        assert!(TraceDraft::new("1 2,3 4").is_ok());
        assert!(TraceDraft::new("-1 -2,\t3\n4").is_ok());
        assert!(TraceDraft::new("1 2,3 4 5").is_err());
        assert!(TraceDraft::new("1 2,").is_err());
        assert!(TraceDraft::new(",1 2").is_err());
        assert!(TraceDraft::new("1 2,,3 4").is_err());
        assert!(TraceDraft::new("1 2, * *").is_err());
        assert!(TraceDraft::new("1 2, '3' '4'").is_err());
        assert!(TraceDraft::new("1 +2").is_err());
        assert!(TraceDraft::new("1 NaN").is_err());
        assert!(
            TraceDraft::new("1 2")
                .unwrap()
                .with_context_ref("https://example.test")
                .is_err()
        );
        assert!(TraceDraft::new("1 2").unwrap().with_brush_ref("#").is_err());
    }

    #[test]
    fn rejects_duplicate_and_unresolved_local_ids() {
        let first = context();
        let duplicate = context();
        let result = Draft::new(AuthoringLimits::default())
            .expect("limits")
            .context(first)
            .expect("first context")
            .context(duplicate)
            .expect("second context")
            .finish();
        assert!(result.is_err());

        let unresolved = TraceDraft::new("1 2")
            .expect("trace")
            .with_context_ref("#missing")
            .expect("reference");
        let result = Draft::new(AuthoringLimits::default())
            .expect("limits")
            .trace(unresolved)
            .expect("trace insertion")
            .finish();
        assert!(result.is_err());
    }

    #[test]
    fn rejects_incompatible_context_attributes_and_missing_trace_refs() {
        let incompatible = ContextDraft::new(ContextKind::WritingRegion).with_alignment_level(1);
        assert!(
            Draft::new(AuthoringLimits::default())
                .expect("limits")
                .context(incompatible)
                .is_err()
        );

        let missing_refs = Draft::new(AuthoringLimits::default())
            .expect("limits")
            .trace(TraceDraft::new("1 2").expect("trace"))
            .expect("trace insertion")
            .finish();
        assert!(missing_refs.is_err());
    }

    #[test]
    fn rejects_invalid_known_brush_property_values_and_units() {
        assert!(BrushPropertyDraft::new(BrushPropertyName::Color, "red").is_err());
        assert!(
            BrushPropertyDraft::new(BrushPropertyName::Color, "#123456")
                .expect("color")
                .with_units("cm")
                .is_err()
        );
        assert!(BrushPropertyDraft::new(BrushPropertyName::Transparency, "256").is_err());
        assert!(BrushPropertyDraft::new(BrushPropertyName::FitToCurve, "yes").is_err());
        assert!(BrushPropertyDraft::new(BrushPropertyName::Tip, "round").is_err());
        assert!(
            BrushPropertyDraft::new(BrushPropertyName::Width, "1.2")
                .expect("width")
                .with_units("not-a-length")
                .is_err()
        );
        assert!(
            BrushPropertyDraft::new(BrushPropertyName::Width, " \t1.2\r\n")
                .expect("schema-whitespace decimal")
                .with_units("m")
                .is_ok()
        );
        assert!(
            BrushPropertyDraft::new(BrushPropertyName::Width, "1.2")
                .expect("width")
                .with_units("px")
                .is_err()
        );
        assert!(BrushPropertyDraft::new(BrushPropertyName::AntiAliased, " \t1\r\n").is_ok());
        let width_without_units = BrushDraft::new("brush")
            .expect("brush")
            .property(BrushPropertyDraft::new(BrushPropertyName::Width, "1.2").expect("width"))
            .expect("defer width-unit validation until brush insertion");
        assert!(
            Draft::new(AuthoringLimits::default())
                .expect("limits")
                .brush(width_without_units)
                .is_err()
        );
        assert!(BrushPropertyDraft::ink_effect(InkEffect::Pencil).is_ok());
    }

    #[test]
    fn generated_source_ids_avoid_user_ids_and_respect_limits() {
        let context = ContextDraft::new(ContextKind::WritingRegion)
            .with_xml_id("inkSrc0")
            .expect("context identifier");
        let brush = BrushDraft::new("inkSrc1").expect("brush identifier");
        let prepared = Draft::new(AuthoringLimits::default())
            .expect("limits")
            .context(context)
            .expect("context")
            .brush(brush)
            .expect("brush")
            .finish()
            .expect("collision-free generated source identifier");
        assert!(
            prepared
                .as_bytes()
                .windows(b"<inkml:inkSource xml:id=\"inkSrc2\">".len())
                .any(|window| window == b"<inkml:inkSource xml:id=\"inkSrc2\">")
        );

        let short_id_limits = AuthoringLimits {
            max_id_bytes: 6,
            ..AuthoringLimits::default()
        };
        let short_id_context = ContextDraft::new(ContextKind::WritingRegion)
            .with_xml_id("ctx0")
            .expect("context identifier");
        assert!(
            Draft::new(short_id_limits)
                .expect("limits")
                .context(short_id_context)
                .expect("context")
                .finish()
                .is_err()
        );
    }

    #[test]
    fn charges_cumulative_retained_text_before_finish() {
        let limits = AuthoringLimits {
            max_source_bytes: 16,
            max_output_bytes: 16,
            max_trace_text_bytes: 11,
            max_traces: 2,
            ..AuthoringLimits::default()
        };
        let first = TraceDraft::new("12345 67890").expect("one bounded X/Y trace");
        let second = TraceDraft::new("12345 67890").expect("another bounded X/Y trace");
        let draft = Draft::new(limits)
            .expect("limits")
            .trace(first)
            .expect("first trace retained");
        assert!(draft.trace(second).is_err());
    }

    #[test]
    fn enforces_exact_record_and_output_boundaries() {
        let limits = AuthoringLimits {
            max_contexts: 1,
            ..AuthoringLimits::default()
        };
        let draft = Draft::new(limits)
            .expect("limits")
            .context(context())
            .expect("one context");
        assert!(draft.context(context()).is_err());

        let full = Draft::new(AuthoringLimits::default())
            .expect("limits")
            .context(context())
            .expect("context")
            .finish()
            .expect("full output");
        let exact_limits = AuthoringLimits {
            max_source_bytes: full.as_bytes().len(),
            max_output_bytes: full.as_bytes().len(),
            ..AuthoringLimits::default()
        };
        let exact = Draft::new(exact_limits)
            .expect("limits")
            .context(context())
            .expect("context")
            .finish();
        assert!(exact.is_ok());
        let below_limits = AuthoringLimits {
            max_source_bytes: full.as_bytes().len() - 1,
            max_output_bytes: full.as_bytes().len() - 1,
            ..AuthoringLimits::default()
        };
        assert!(
            Draft::new(below_limits)
                .expect("limits")
                .context(context())
                .expect("context")
                .finish()
                .is_err()
        );

        let exact_text_limits = AuthoringLimits {
            max_trace_text_bytes: 5,
            ..AuthoringLimits::default()
        };
        assert!(
            Draft::new(exact_text_limits)
                .expect("limits")
                .trace(TraceDraft::new("12 34").expect("trace"))
                .is_ok()
        );
        let below_text_limits = AuthoringLimits {
            max_trace_text_bytes: 4,
            ..AuthoringLimits::default()
        };
        assert!(
            Draft::new(below_text_limits)
                .expect("limits")
                .trace(TraceDraft::new("12 34").expect("trace"))
                .is_err()
        );
    }

    #[test]
    fn rejects_invalid_limits_before_authoring() {
        assert!(Draft::default().finish().is_ok());
        let limits = AuthoringLimits {
            max_output_bytes: 0,
            ..AuthoringLimits::default()
        };
        assert!(Draft::new(limits).is_err());
    }
}
