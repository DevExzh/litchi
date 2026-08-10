//! Structural transaction values and XML mutation helpers.

use litchi_core::{Error, Metadata, Position, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{borrow::Cow, ops::Range};

use crate::{Master, style::Origin};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MAX_VALUE_BYTES: usize = 16 * 1024;
const MAX_RESOURCE_BYTES: usize = 64 * 1024 * 1024;

/// Explicit security policy for one ODM transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityPolicy {
    allow_external_targets: bool,
    allow_missing_package_targets: bool,
    active_content: ActiveContentPolicy,
    max_resource_bytes: usize,
}

/// Changed-write treatment for content which the library never executes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ActiveContentPolicy {
    /// Refuse changed output when active content is present.
    Refuse,
    /// Preserve recognized active content inertly in the package.
    PreserveInert,
}

impl SecurityPolicy {
    /// Returns a restrictive final-graph policy.
    ///
    /// External links, unresolved package targets, and changed output with
    /// active content are refused.
    #[must_use]
    pub const fn strict() -> Self {
        Self {
            allow_external_targets: false,
            allow_missing_package_targets: false,
            active_content: ActiveContentPolicy::Refuse,
            max_resource_bytes: MAX_RESOURCE_BYTES,
        }
    }

    /// Configures whether the final package may contain external link targets.
    #[must_use]
    pub const fn with_external_targets(mut self, allow: bool) -> Self {
        self.allow_external_targets = allow;
        self
    }

    /// Configures whether the final package may contain unresolved package targets.
    #[must_use]
    pub const fn with_missing_package_targets(mut self, allow: bool) -> Self {
        self.allow_missing_package_targets = allow;
        self
    }

    /// Configures changed writes for inert scripts, forms, macros, and DDE.
    #[must_use]
    pub const fn with_active_content(mut self, policy: ActiveContentPolicy) -> Self {
        self.active_content = policy;
        self
    }

    /// Configures the maximum byte size of one staged resource.
    #[must_use]
    pub const fn with_max_resource_bytes(mut self, max: usize) -> Self {
        self.max_resource_bytes = max;
        self
    }

    pub(crate) const fn allows_external_targets(self) -> bool {
        self.allow_external_targets
    }

    pub(crate) const fn allows_missing_package_targets(self) -> bool {
        self.allow_missing_package_targets
    }

    pub(crate) const fn max_resource_bytes(self) -> usize {
        self.max_resource_bytes
    }

    pub(crate) const fn active_content(self) -> ActiveContentPolicy {
        self.active_content
    }
}

impl Default for SecurityPolicy {
    fn default() -> Self {
        Self {
            allow_external_targets: true,
            allow_missing_package_targets: true,
            active_content: ActiveContentPolicy::Refuse,
            max_resource_bytes: MAX_RESOURCE_BYTES,
        }
    }
}

/// A detached inert subdocument reference for a new master section.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SubdocumentSpec {
    href: String,
    source_section: Option<String>,
    filter_name: Option<String>,
}

impl SubdocumentSpec {
    /// Creates a package or inert external subdocument reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty, excessive, or invalid XML target.
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let href = href.into();
        validate_value(&href, "ODM subdocument target", false)?;
        crate::link::validate_href(&href)?;
        Ok(Self {
            href,
            source_section: None,
            filter_name: None,
        })
    }

    /// Selects a named section within the referenced subdocument.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty, excessive, or invalid XML value.
    pub fn with_source_section(mut self, name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_value(&name, "ODM source section name", false)?;
        self.source_section = Some(name);
        Ok(self)
    }

    /// Retains the producer filter name for the referenced subdocument.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty, excessive, or invalid XML value.
    pub fn with_filter_name(mut self, name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_value(&name, "ODM source filter name", false)?;
        self.filter_name = Some(name);
        Ok(self)
    }

    /// Returns the inert target text.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Returns the source-section selector, when present.
    #[must_use]
    pub fn source_section(&self) -> Option<&str> {
        self.source_section.as_deref()
    }

    /// Returns the producer filter name, when present.
    #[must_use]
    pub fn filter_name(&self) -> Option<&str> {
        self.filter_name.as_deref()
    }
}

/// A detached root section to insert into the master body.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SectionSpec {
    name: String,
    style_name: Option<String>,
    subdocument: Option<SubdocumentSpec>,
}

impl SectionSpec {
    /// Creates a named empty root section.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty, excessive, or invalid XML name value.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_value(&name, "ODM section name", false)?;
        Ok(Self {
            name,
            style_name: None,
            subdocument: None,
        })
    }

    /// Associates an existing style name.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty, excessive, or invalid XML value.
    pub fn with_style(mut self, style_name: impl Into<String>) -> Result<Self> {
        let style_name = style_name.into();
        validate_value(&style_name, "ODM section style name", false)?;
        self.style_name = Some(style_name);
        Ok(self)
    }

    /// Makes this section a linked section with the source as its first child.
    #[must_use]
    pub fn with_subdocument(mut self, subdocument: SubdocumentSpec) -> Self {
        self.subdocument = Some(subdocument);
        self
    }

    pub(crate) fn name(&self) -> &str {
        &self.name
    }

    pub(crate) fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    pub(crate) const fn subdocument(&self) -> Option<&SubdocumentSpec> {
        self.subdocument.as_ref()
    }
}

/// A detached minimal named style definition.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StyleSpec {
    name: String,
    family: String,
    origin: Origin,
    parent: Option<String>,
    raw_fragment: Option<String>,
}

impl StyleSpec {
    /// Creates a style definition owned by `styles.xml`.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid bounded XML values.
    pub fn new(name: impl Into<String>, family: impl Into<String>) -> Result<Self> {
        let name = name.into();
        let family = family.into();
        validate_value(&name, "ODM style name", false)?;
        validate_value(&family, "ODM style family", false)?;
        Ok(Self {
            name,
            family,
            origin: Origin::Styles,
            parent: None,
            raw_fragment: None,
        })
    }

    /// Selects the package part which owns the new definition.
    #[must_use]
    pub const fn with_origin(mut self, origin: Origin) -> Self {
        self.origin = origin;
        self
    }

    pub(crate) fn name(&self) -> &str {
        &self.name
    }

    pub(crate) fn family(&self) -> &str {
        &self.family
    }

    pub(crate) const fn origin(&self) -> Origin {
        self.origin
    }

    pub(crate) fn parent(&self) -> Option<&str> {
        self.parent.as_deref()
    }

    pub(crate) fn raw_fragment(&self) -> Option<&str> {
        self.raw_fragment.as_deref()
    }

    pub(crate) fn imported(
        source_xml: &str,
        definition: &crate::style::Definition,
        name: &str,
        parent: Option<&str>,
    ) -> Result<Self> {
        let family = definition
            .family()
            .ok_or_else(|| invalid("ODM imported style has no family"))?;
        validate_value(name, "ODM imported style name", false)?;
        let raw = source_xml
            .get(definition.source_span.clone())
            .ok_or_else(|| invalid("ODM imported style source span is stale"))?;
        let raw_fragment = standalone_style_fragment(source_xml, raw)?;
        let raw_fragment = rewrite_imported_style(
            &raw_fragment,
            definition.name(),
            name,
            definition.parent(),
            parent,
        )?;
        Ok(Self {
            name: name.to_string(),
            family: family.to_string(),
            origin: definition.origin(),
            parent: parent.map(str::to_string),
            raw_fragment: Some(raw_fragment),
        })
    }
}

/// A detached inert package resource.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceSpec {
    path: String,
    media_type: String,
    bytes: Vec<u8>,
}

impl ResourceSpec {
    /// Creates a bounded resource without interpreting its bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsafe/core path or invalid media type.
    pub fn new(
        path: impl Into<String>,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<Self> {
        let path = path.into();
        let media_type = media_type.into();
        validate_resource_path(&path)?;
        validate_value(&media_type, "ODM resource media type", false)?;
        Ok(Self {
            path,
            media_type,
            bytes,
        })
    }

    /// Returns the safe package path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Returns the declared media type.
    #[must_use]
    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    /// Returns the inert bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
}

/// One staged or committed section-tree effect.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SectionChange {
    /// Inserts a detached root section.
    Add(SectionSpec),
    /// Renames an existing section and modeled local references.
    Rename {
        position: Position,
        before: String,
        after: String,
    },
    /// Removes an existing section subtree.
    Remove { position: Position, before: String },
}

/// One staged or committed generated-index effect.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum GeneratedIndexChange {
    /// Renames an existing generated index.
    Rename {
        item: Position,
        before: String,
        after: String,
    },
}

/// One staged or committed direct master-body item effect.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyItemChange {
    /// Removes an exactly addressed direct body subtree.
    Remove {
        /// The source [`crate::structure::Structure::items`] position.
        item: Position,
        /// The checked source item kind.
        kind: crate::structure::Kind,
    },
}

/// One staged or committed style-catalog effect.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum StyleChange {
    Add(StyleSpec),
    Rename {
        origin: Origin,
        before: String,
        after: String,
    },
    Remove {
        origin: Origin,
        name: String,
    },
}

/// One staged or committed package-resource effect.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ResourceChange {
    Put(ResourceSpec),
    Remove(ResourceSpec),
}

pub(crate) struct MutatedParts {
    pub(crate) content: String,
    pub(crate) styles: Option<String>,
}

pub(crate) fn mutate_xml(
    source: &Master,
    links: &[(Position, String)],
    sections: &[SectionChange],
    generated_indexes: &[GeneratedIndexChange],
    body_items: &[BodyItemChange],
    styles: &[StyleChange],
) -> Result<MutatedParts> {
    let mut content_edits = Vec::new();
    let mut styles_edits = Vec::new();
    content_edits
        .try_reserve(links.len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODM linked-section XML edits",
            source: allocation_error,
        })?;
    for (position, after) in links {
        let span = source
            .href_span(position.get())
            .ok_or_else(|| invalid("ODM linked-section source span is stale"))?;
        content_edits.push((span.clone(), escape(after)));
    }
    for intent in sections {
        stage_section(source, intent, &mut content_edits)?;
    }
    for intent in generated_indexes {
        stage_generated_index(source, intent, &mut content_edits)?;
    }
    for intent in body_items {
        stage_body_item(source, intent, &mut content_edits)?;
    }
    let removed_content_spans = removed_content_spans(source, sections, body_items)?;
    for intent in styles {
        stage_style(
            source,
            intent,
            &removed_content_spans,
            &mut content_edits,
            &mut styles_edits,
        )?;
    }
    let mut content = apply_edits(source.content_xml(), content_edits)?;
    let mut styles_xml = source.styles_xml().map(str::to_owned);
    if styles_xml.is_none()
        && styles.iter().any(
            |intent| matches!(intent, StyleChange::Add(spec) if spec.origin() == Origin::Styles),
        )
    {
        styles_xml = Some(empty_styles_part().to_owned());
    }
    if !styles_edits.is_empty() {
        let xml = styles_xml
            .as_deref()
            .ok_or_else(|| invalid("ODM styles.xml is required for this style operation"))?;
        styles_xml = Some(apply_edits(xml, styles_edits)?);
    }
    for intent in sections {
        if let SectionChange::Add(spec) = intent {
            content =
                insert_before_element_end(&content, OFFICE, b"text", &section_fragment(spec))?;
        }
    }
    if styles
        .iter()
        .any(|intent| matches!(intent, StyleChange::Add(spec) if spec.origin() == Origin::Content))
        && !contains_element(&content, OFFICE, b"automatic-styles")?
    {
        content = insert_before_element_start(
            &content,
            OFFICE,
            b"body",
            concat!(
                r#"<office:automatic-styles "#,
                r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">"#,
                r#"</office:automatic-styles>"#,
            ),
        )?;
    }
    for intent in styles {
        if let StyleChange::Add(spec) = intent {
            match spec.origin() {
                Origin::Content => {
                    content = insert_before_element_end(
                        &content,
                        OFFICE,
                        b"automatic-styles",
                        &style_fragment(spec),
                    )?;
                },
                Origin::Styles => {
                    let xml = styles_xml
                        .as_deref()
                        .ok_or_else(|| invalid("ODM styles.xml is required to add this style"))?;
                    styles_xml = Some(insert_before_element_end(
                        xml,
                        OFFICE,
                        b"styles",
                        &style_fragment(spec),
                    )?);
                },
            }
        }
    }
    Ok(MutatedParts {
        content,
        styles: styles_xml,
    })
}

fn stage_body_item(
    source: &Master,
    intent: &BodyItemChange,
    edits: &mut Vec<(Range<usize>, String)>,
) -> Result<()> {
    match intent {
        BodyItemChange::Remove { item, kind } => {
            let actual = source
                .structure()
                .items()
                .get(item.get())
                .ok_or_else(|| invalid("ODM master-body item selector is stale"))?;
            if actual != kind {
                return Err(invalid("ODM master-body item kind is stale"));
            }
            let span = source
                .structure()
                .item_spans
                .get(item.get())
                .cloned()
                .ok_or_else(|| invalid("ODM master-body item source span is stale"))?;
            edits.push((span, String::new()));
            Ok(())
        },
    }
}

fn stage_generated_index(
    source: &Master,
    intent: &GeneratedIndexChange,
    edits: &mut Vec<(Range<usize>, String)>,
) -> Result<()> {
    match intent {
        GeneratedIndexChange::Rename {
            item,
            before,
            after,
        } => {
            let index = source
                .structure()
                .generated_indexes()
                .iter()
                .find(|index| index.item() == *item)
                .ok_or_else(|| invalid("ODM generated-index selector is stale"))?;
            if index.name() != Some(before) {
                return Err(invalid("ODM generated-index identity is stale"));
            }
            let span = index
                .name_span
                .clone()
                .ok_or_else(|| invalid("ODM generated index has no addressable text:name"))?;
            edits.push((span, escape(after)));
            Ok(())
        },
    }
}

const fn empty_styles_part() -> &'static str {
    concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">"#,
        r#"<office:styles></office:styles></office:document-styles>"#,
    )
}

pub(crate) fn stage_metadata(source: &Master, metadata: &Metadata) -> Result<String> {
    use litchi_odf_common::core::{
        MetaXmlPatch, metadata::Metadata as OdfMetadata, patch_meta_xml,
    };
    let source_xml = source
        .meta_xml()?
        .ok_or_else(|| invalid("ODM metadata editing requires meta.xml"))?;
    let parsed = OdfMetadata::from_xml(&source_xml)?;
    let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&parsed, metadata);
    patch_meta_xml(&source_xml, &patch)?
        .ok_or_else(|| invalid("ODM metadata editing requires office:meta"))
}

fn stage_section(
    source: &Master,
    intent: &SectionChange,
    edits: &mut Vec<(Range<usize>, String)>,
) -> Result<()> {
    match intent {
        SectionChange::Add(_) => Ok(()),
        SectionChange::Rename {
            position,
            before,
            after,
        } => {
            let node = source
                .section_tree()
                .get(*position)
                .ok_or_else(|| invalid("ODM section selector is stale"))?;
            edits.push((node.name_span.clone(), escape(after)));
            for (target, span) in source.local_section_references() {
                if target == before {
                    edits.push((span.clone(), escape(after)));
                }
            }
            Ok(())
        },
        SectionChange::Remove { position, before } => {
            let node = source
                .section_tree()
                .get(*position)
                .ok_or_else(|| invalid("ODM section selector is stale"))?;
            if source
                .local_section_references()
                .iter()
                .any(|(target, span)| {
                    source.section_tree().sections().iter().any(|candidate| {
                        node.source_span.start <= candidate.source_span.start
                            && candidate.source_span.end <= node.source_span.end
                            && candidate.name() == target
                    }) && !(node.source_span.start <= span.start
                        && span.end <= node.source_span.end)
                })
            {
                return Err(invalid(
                    "ODM section-subtree removal is blocked by an incoming local reference",
                ));
            }
            if node.name() != before {
                return Err(invalid("ODM removed section identity is stale"));
            }
            edits.push((node.source_span.clone(), String::new()));
            Ok(())
        },
    }
}

fn stage_style(
    source: &Master,
    intent: &StyleChange,
    removed_content_spans: &[Range<usize>],
    content_edits: &mut Vec<(Range<usize>, String)>,
    styles_edits: &mut Vec<(Range<usize>, String)>,
) -> Result<()> {
    let (origin, name, replacement, remove) = match intent {
        StyleChange::Add(_) => return Ok(()),
        StyleChange::Rename {
            origin,
            before,
            after,
        } => (*origin, before.as_str(), Some(after.as_str()), false),
        StyleChange::Remove { origin, name } => (*origin, name.as_str(), None, true),
    };
    let definition = source
        .styles()
        .iter()
        .find(|definition| definition.origin() == origin && definition.name() == name)
        .ok_or_else(|| invalid("ODM style selector is stale"))?;
    let defining_edits = match origin {
        Origin::Content => &mut *content_edits,
        Origin::Styles => &mut *styles_edits,
    };
    if remove {
        if style_is_referenced(source, name, removed_content_spans)? {
            return Err(invalid(
                "ODM style removal is blocked by an incoming reference",
            ));
        }
        defining_edits.push((definition.source_span.clone(), String::new()));
    } else if let Some(after) = replacement {
        defining_edits.push((definition.name_span.clone(), escape(after)));
        content_edits.extend(
            attribute_references(source.content_xml(), name, after)?
                .into_iter()
                .filter(|(span, _)| !span_is_removed(span, removed_content_spans)),
        );
        if let Some(xml) = source.styles_xml() {
            styles_edits.extend(attribute_references(xml, name, after)?);
        }
    }
    Ok(())
}

fn style_is_referenced(
    source: &Master,
    name: &str,
    removed_content_spans: &[Range<usize>],
) -> Result<bool> {
    if attribute_references(source.content_xml(), name, name)?
        .iter()
        .any(|(span, _)| !span_is_removed(span, removed_content_spans))
    {
        return Ok(true);
    }
    source
        .styles_xml()
        .map(|xml| attribute_references(xml, name, name).map(|spans| !spans.is_empty()))
        .transpose()
        .map(Option::unwrap_or_default)
}

fn removed_content_spans(
    source: &Master,
    sections: &[SectionChange],
    body_items: &[BodyItemChange],
) -> Result<Vec<Range<usize>>> {
    let mut spans = Vec::new();
    spans
        .try_reserve(
            sections
                .iter()
                .filter(|change| matches!(change, SectionChange::Remove { .. }))
                .count()
                .saturating_add(body_items.len()),
        )
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODM removed content spans",
            source: allocation_error,
        })?;
    for change in sections {
        let SectionChange::Remove { position, .. } = change else {
            continue;
        };
        let node = source
            .section_tree()
            .get(*position)
            .ok_or_else(|| invalid("ODM removed section selector is stale"))?;
        spans.push(node.source_span.clone());
    }
    for change in body_items {
        let BodyItemChange::Remove { item, .. } = change;
        spans.push(
            source
                .structure()
                .item_spans
                .get(item.get())
                .cloned()
                .ok_or_else(|| invalid("ODM removed body-item span is stale"))?,
        );
    }
    Ok(spans)
}

fn span_is_removed(span: &Range<usize>, removed: &[Range<usize>]) -> bool {
    removed
        .iter()
        .any(|owner| owner.start <= span.start && span.end <= owner.end)
}

fn attribute_references(
    xml: &str,
    expected: &str,
    replacement: &str,
) -> Result<Vec<(Range<usize>, String)>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut edits = Vec::new();
    loop {
        let start = position(&reader)?;
        let (_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM XML reference scan: {error}")))?;
        let end = position(&reader)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let tag = xml
                    .as_bytes()
                    .get(start..end)
                    .ok_or_else(|| invalid("ODM XML reference span is outside its part"))?;
                for raw in element.attributes() {
                    let attribute = raw.map_err(|error| {
                        invalid(format!("invalid ODM reference attribute: {error}"))
                    })?;
                    let local = attribute.key.local_name();
                    if !matches!(local.as_ref(), b"style-name" | b"parent-style-name") {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map(Cow::into_owned)
                        .map_err(|error| {
                            invalid(format!("invalid ODM reference value: {error}"))
                        })?;
                    if value == expected {
                        let (value_start, value_end) =
                            attribute_value_span(tag, attribute.key.as_ref())?;
                        edits.push((start + value_start..start + value_end, escape(replacement)));
                    }
                }
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM XML")),
            Event::Eof => break,
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(edits)
}

fn apply_edits(source: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String> {
    edits.sort_by_key(|edit| std::cmp::Reverse(edit.0.start));
    let mut output = source.to_owned();
    let mut previous_start = source.len();
    for (span, replacement) in edits {
        if span.start > span.end || span.end > previous_start || span.end > output.len() {
            return Err(invalid("ODM XML edits overlap or have stale source spans"));
        }
        output.replace_range(span.clone(), &replacement);
        previous_start = span.start;
    }
    Ok(output)
}

fn insert_before_element_end(
    xml: &str,
    namespace: &[u8],
    local: &[u8],
    fragment: &str,
) -> Result<String> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM XML insertion scan: {error}")))?;
        match event {
            Event::End(element)
                if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace)
                    && element.local_name().as_ref() == local =>
            {
                let mut output = String::new();
                output
                    .try_reserve_exact(xml.len().saturating_add(fragment.len()))
                    .map_err(|source| Error::Allocation {
                        resource: "ODM XML insertion",
                        source,
                    })?;
                output.push_str(&xml[..start]);
                output.push_str(fragment);
                output.push_str(&xml[start..]);
                return Ok(output);
            },
            Event::Eof => return Err(invalid("ODM XML insertion owner was not found")),
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM XML")),
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn contains_element(xml: &str, namespace: &[u8], local: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    loop {
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM XML owner scan: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace)
                    && element.local_name().as_ref() == local
                {
                    return Ok(true);
                }
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM XML")),
            Event::Eof => return Ok(false),
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

fn insert_before_element_start(
    xml: &str,
    namespace: &[u8],
    local: &[u8],
    fragment: &str,
) -> Result<String> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM XML insertion scan: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace)
                    && element.local_name().as_ref() == local =>
            {
                let mut output = String::new();
                output
                    .try_reserve_exact(xml.len().saturating_add(fragment.len()))
                    .map_err(|source| Error::Allocation {
                        resource: "ODM XML insertion",
                        source,
                    })?;
                output.push_str(&xml[..start]);
                output.push_str(fragment);
                output.push_str(&xml[start..]);
                return Ok(output);
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM XML")),
            Event::Eof => return Err(invalid("ODM XML insertion anchor was not found")),
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn section_fragment(spec: &SectionSpec) -> String {
    let style = spec.style_name().map_or_else(String::new, |name| {
        format!(r#" text:style-name="{}""#, escape(name))
    });
    let Some(subdocument) = spec.subdocument() else {
        return format!(
            r#"<text:section xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" text:name="{}"{style}/>"#,
            escape(spec.name())
        );
    };
    let source_section = subdocument
        .source_section()
        .map_or_else(String::new, |name| {
            format!(r#" text:section-name="{}""#, escape(name))
        });
    let filter_name = subdocument.filter_name().map_or_else(String::new, |name| {
        format!(r#" text:filter-name="{}""#, escape(name))
    });
    format!(
        concat!(
            r#"<text:section xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
            r#"xmlns:xlink="http://www.w3.org/1999/xlink" text:name="{}"{style}>"#,
            r#"<text:section-source xlink:href="{}" xlink:type="simple" "#,
            r#"xlink:show="embed"{source_section}{filter_name}/></text:section>"#,
        ),
        escape(spec.name()),
        escape(subdocument.href()),
        style = style,
        source_section = source_section,
        filter_name = filter_name,
    )
}

fn style_fragment(spec: &StyleSpec) -> String {
    if let Some(fragment) = spec.raw_fragment() {
        return fragment.to_string();
    }
    format!(
        r#"<style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="{}" style:family="{}"/>"#,
        escape(spec.name()),
        escape(spec.family())
    )
}

fn standalone_style_fragment(document: &str, fragment: &str) -> Result<String> {
    use std::collections::HashSet;

    let mut reader = NsReader::from_reader(document.as_bytes());
    let mut declarations = Vec::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("invalid ODM style namespace source: {error}")))?
        {
            Event::Start(root) => {
                for raw in root.attributes() {
                    let attribute = raw.map_err(|error| {
                        invalid(format!("invalid ODM namespace declaration: {error}"))
                    })?;
                    let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| {
                        invalid(format!("ODM namespace name is not UTF-8: {error}"))
                    })?;
                    let Some(prefix) = key.strip_prefix("xmlns:") else {
                        continue;
                    };
                    if fragment.contains(&format!("{prefix}:"))
                        && !fragment.contains(&format!("xmlns:{prefix}="))
                    {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map(Cow::into_owned)
                            .map_err(|error| {
                                invalid(format!("invalid ODM namespace value: {error}"))
                            })?;
                        declarations.push((key.to_string(), value));
                    }
                }
                break;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM style XML")),
            Event::Eof => return Err(invalid("ODM style document has no root element")),
            Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    let mut seen = HashSet::new();
    declarations.retain(|(key, _)| seen.insert(key.clone()));
    if declarations.is_empty() {
        return crate::codec::compact_source_xml(fragment);
    }
    let insertion = fragment
        .find(|byte: char| byte.is_whitespace() || matches!(byte, '/' | '>'))
        .ok_or_else(|| invalid("ODM imported style start tag is malformed"))?;
    let mut output = fragment.to_string();
    let mut attributes = String::new();
    for (key, value) in declarations {
        attributes.push(' ');
        attributes.push_str(&key);
        attributes.push_str("=\"");
        attributes.push_str(&escape(&value));
        attributes.push('"');
    }
    output.insert_str(insertion, &attributes);
    crate::codec::compact_source_xml(&output)
}

fn rewrite_imported_style(
    fragment: &str,
    before: &str,
    after: &str,
    parent_before: Option<&str>,
    parent_after: Option<&str>,
) -> Result<String> {
    const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    let mut reader = NsReader::from_reader(fragment.as_bytes());
    let start = position(&reader)?;
    let (namespace, event) = reader
        .read_resolved_event()
        .map_err(|error| invalid(format!("invalid imported ODM style: {error}")))?;
    let is_style_namespace =
        matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == STYLE);
    let event = event.into_owned();
    let end = position(&reader)?;
    let element = match event {
        Event::Start(element) | Event::Empty(element)
            if is_style_namespace && element.local_name().as_ref() == b"style" =>
        {
            element
        },
        Event::Start(_)
        | Event::Empty(_)
        | Event::End(_)
        | Event::Text(_)
        | Event::CData(_)
        | Event::Comment(_)
        | Event::Decl(_)
        | Event::PI(_)
        | Event::DocType(_)
        | Event::GeneralRef(_)
        | Event::Eof => {
            return Err(invalid("ODM imported style fragment is not style:style"));
        },
    };
    let tag = fragment
        .as_bytes()
        .get(start..end)
        .ok_or_else(|| invalid("ODM imported style tag span is stale"))?;
    let mut edits = Vec::new();
    let name_key = attribute_key(&reader, &element, STYLE, b"name")?
        .ok_or_else(|| invalid("ODM imported style has no style:name"))?;
    let (name_start, name_end) = attribute_value_span(tag, &name_key)?;
    if before != after {
        edits.push((start + name_start..start + name_end, escape(after)));
    }
    if parent_before != parent_after {
        let parent_key = attribute_key(&reader, &element, STYLE, b"parent-style-name")?
            .ok_or_else(|| invalid("ODM imported parent style attribute disappeared"))?;
        let (parent_start, parent_end) = attribute_value_span(tag, &parent_key)?;
        let parent =
            parent_after.ok_or_else(|| invalid("ODM imported parent style mapping disappeared"))?;
        edits.push((start + parent_start..start + parent_end, escape(parent)));
    }
    apply_edits(fragment, edits)
}

fn validate_resource_path(path: &str) -> Result<()> {
    validate_value(path, "ODM resource path", false)?;
    if path.starts_with('/')
        || path.contains('\\')
        || path.contains(':')
        || path
            .split('/')
            .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
        || matches!(path, "mimetype" | "content.xml" | "styles.xml" | "meta.xml")
        || path.starts_with("META-INF/")
    {
        return Err(invalid("ODM resource path is unsafe or reserved"));
    }
    Ok(())
}

pub(crate) fn validate_value(value: &str, scope: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty()) || value.len() > MAX_VALUE_BYTES {
        return Err(invalid(format!(
            "{scope} is empty or exceeds the 16 KiB limit"
        )));
    }
    if value.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        return Err(invalid(format!("{scope} contains a character forbidden by XML 1.0")));
    }
    Ok(())
}

fn attribute_value_span(tag: &[u8], wanted: &[u8]) -> Result<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && !matches!(tag[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("ODM attribute is missing '='"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| invalid("ODM attribute value is not quoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return Err(invalid("ODM attribute value is unterminated"));
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start, value_end));
        }
    }
    Err(invalid("ODM attribute source span is missing"))
}

fn attribute_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<Vec<u8>>> {
    let mut key = None;
    for raw_attribute in element.attributes() {
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid ODM attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace)
            && name.as_ref() == local
            && key.replace(attribute.key.as_ref().to_vec()).is_some()
        {
            return Err(invalid("duplicate namespace-equivalent ODM attribute"));
        }
    }
    Ok(key)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_range_error| invalid("ODM XML source position exceeds the platform range"))
}

fn escape(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
