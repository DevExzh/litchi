//! Source-bound table-template snapshots and package-neutral transactions.

use super::{Template, codec, validation::MAX_AGGREGATE_BYTES};
use litchi_core::{Error, Result};
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_TEMPLATES: usize = 1_000_000;
const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_SPANS: usize = 1_048_576;
const MAX_XML_DEPTH: usize = 1_024;

/// An immutable table-template view retaining the exact source styles part.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    source: Option<String>,
    templates: Vec<Template>,
}

impl Snapshot {
    /// Parse an optional styles.xml source.
    pub fn from_source(source: Option<&str>) -> Result<Self> {
        if let Some(source) = source
            && source.len() > MAX_XML_BYTES
        {
            return Err(Error::InvalidFormat(
                "ODS styles.xml exceeds the table-template size limit".to_string(),
            ));
        }
        let templates = match source {
            Some(source) => codec::parse(source)?,
            None => Vec::new(),
        };
        validate_templates(&templates)?;
        Ok(Self {
            source: source.map(str::to_owned),
            templates,
        })
    }

    /// Return the templates in source order.
    #[must_use]
    pub fn templates(&self) -> &[Template] {
        &self.templates
    }

    /// Return the exact retained styles XML, if a styles part exists.
    #[must_use]
    pub fn source_xml(&self) -> Option<&str> {
        self.source.as_deref()
    }

    /// Start a failure-atomic edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.templates.clone(),
        }
    }
}

/// A staged table-template catalog edit.
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Vec<Template>,
}

impl Edit {
    /// Borrow the staged catalog.
    #[must_use]
    pub fn templates(&self) -> &[Template] {
        &self.draft
    }

    /// Replace the complete ordered catalog.
    pub fn replace(&mut self, templates: Vec<Template>) -> Result<()> {
        validate_templates(&templates)?;
        self.draft = templates;
        Ok(())
    }

    /// Add a template at the catalog tail.
    pub fn add(&mut self, template: Template) -> Result<()> {
        let mut candidate = self.draft.clone();
        candidate.push(template);
        self.replace(candidate)
    }

    /// Replace a template selected by zero-based source index.
    pub fn replace_at(&mut self, index: usize, template: Template) -> Result<()> {
        let mut candidate = self.draft.clone();
        let slot = candidate.get_mut(index).ok_or_else(|| {
            Error::InvalidFormat("ODS table-template index did not match".to_string())
        })?;
        *slot = template;
        self.replace(candidate)
    }

    /// Replace a template selected by exact semantic name.
    pub fn replace_named(&mut self, name: &str, template: Template) -> Result<()> {
        let index = self
            .draft
            .iter()
            .position(|candidate| candidate.name == name)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("ODS table-template '{name}' was not found"))
            })?;
        self.replace_at(index, template)
    }

    /// Remove one template by source index.
    pub fn remove_at(&mut self, index: usize) -> Result<Template> {
        let mut candidate = self.draft.clone();
        if index >= candidate.len() {
            return Err(Error::InvalidFormat(
                "ODS table-template index did not match".to_string(),
            ));
        }
        let removed = candidate.remove(index);
        self.replace(candidate)?;
        Ok(removed)
    }

    /// Remove one template by exact semantic name.
    pub fn remove_named(&mut self, name: &str) -> Result<Template> {
        let index = self
            .draft
            .iter()
            .position(|candidate| candidate.name == name)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("ODS table-template '{name}' was not found"))
            })?;
        self.remove_at(index)
    }

    /// Remove all templates while retaining an existing styles owner.
    pub fn clear(&mut self) -> Result<()> {
        self.replace(Vec::new())
    }

    /// Produce a source-checked semantic commit.
    pub fn commit(self) -> Result<Commit> {
        validate_templates(&self.draft)?;
        if self.draft == self.before.templates {
            return Ok(Commit {
                snapshot: self.before.clone(),
                patch: Patch {
                    source: self.before.source.clone(),
                    target: self.before.source.clone(),
                },
                changed: false,
            });
        }
        let target_xml = render_source(
            self.before.source.as_deref(),
            &self.before.templates,
            &self.draft,
        )?;
        let target = Snapshot::from_source(target_xml.as_deref())?;
        if target.templates != self.draft {
            return Err(Error::InvalidFormat(
                "ODS table-template commit failed typed readback".to_string(),
            ));
        }
        Ok(Commit {
            patch: Patch {
                source: self.before.source.clone(),
                target: target.source.clone(),
            },
            snapshot: target,
            changed: true,
        })
    }
}

/// An exact-source reversible styles-part patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source: Option<String>,
    target: Option<String>,
}

impl Patch {
    /// Whether this patch is physically empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Return the inverse exact-source patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
        }
    }

    /// Apply this patch only to the exact source snapshot.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if snapshot.source != self.source {
            return Err(Error::InvalidFormat(
                "ODS table-template patch source snapshot does not match".to_string(),
            ));
        }
        let target = Snapshot::from_source(self.target.as_deref())?;
        Ok(Commit {
            changed: !self.is_empty(),
            patch: self.clone(),
            snapshot: target,
        })
    }
}

/// A fully rehydrated table-template commit.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the styles part changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the resulting semantic snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}

fn validate_templates(templates: &[Template]) -> Result<()> {
    if templates.len() > MAX_TEMPLATES {
        return Err(Error::InvalidFormat(format!(
            "ODS document exceeds {MAX_TEMPLATES} table templates"
        )));
    }
    let mut names = std::collections::HashSet::with_capacity(templates.len());
    let mut aggregate = 0usize;
    for template in templates {
        template.validate()?;
        if !names.insert(template.name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate table template '{}'",
                template.name
            )));
        }
        charge_template_attribute(&mut aggregate, "name", &template.name)?;
        for (name, present) in [
            (
                "first-row-start-column",
                template.first_row_start_column.is_some(),
            ),
            (
                "first-row-end-column",
                template.first_row_end_column.is_some(),
            ),
            (
                "last-row-start-column",
                template.last_row_start_column.is_some(),
            ),
            (
                "last-row-end-column",
                template.last_row_end_column.is_some(),
            ),
        ] {
            if present {
                charge_template_attribute(&mut aggregate, name, "row")?;
            }
        }
        for (name, value) in [
            ("use-first-row-styles", template.use_first_row_styles),
            ("use-last-row-styles", template.use_last_row_styles),
            ("use-first-column-styles", template.use_first_column_styles),
            ("use-last-column-styles", template.use_last_column_styles),
            ("use-banding-rows-styles", template.use_banding_rows_styles),
            (
                "use-banding-columns-styles",
                template.use_banding_columns_styles,
            ),
        ] {
            if let Some(value) = value {
                charge_template_attribute(
                    &mut aggregate,
                    name,
                    if value { "true" } else { "false" },
                )?;
            }
        }
        for region in super::semantic::Region::ALL {
            let Some(style) = template.region(region) else {
                continue;
            };
            charge_template_attribute(&mut aggregate, "style-name", &style.style_name)?;
            if let Some(paragraph) = &style.paragraph_style_name {
                charge_template_attribute(&mut aggregate, "paragraph-style-name", paragraph)?;
            }
        }
    }
    Ok(())
}

fn charge_template_attribute(aggregate: &mut usize, name: &str, value: &str) -> Result<()> {
    *aggregate = aggregate
        .checked_add(name.len())
        .and_then(|total| total.checked_add(value.len()))
        .ok_or_else(|| {
            Error::InvalidFormat("table-template aggregate size overflow".to_string())
        })?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return Err(Error::InvalidFormat(
            "table-template metadata exceeds 16 MiB".to_string(),
        ));
    }
    Ok(())
}

fn element_prefix(qname: &str) -> &str {
    qname.split_once(':').map_or("", |(prefix, _)| prefix)
}

fn template_xml(template: &Template, prefix: &str, declare_namespace: bool) -> Result<String> {
    let mut output = String::new();
    template.write_xml_with_prefix(&mut output, prefix, declare_namespace)?;
    if output.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(
            "ODS table-template output exceeds the table-template size limit".to_string(),
        ));
    }
    Ok(output)
}

fn template_xmls(templates: &[Template], prefix: &str, declare_namespace: bool) -> Result<String> {
    let mut output = String::new();
    for template in templates {
        let fragment = template_xml(template, prefix, declare_namespace)?;
        let next = output
            .len()
            .checked_add(fragment.len())
            .ok_or_else(|| invalid("ODS table-template output size overflows usize"))?;
        if next > MAX_XML_BYTES {
            return Err(Error::InvalidFormat(
                "ODS table-template output exceeds the table-template size limit".to_string(),
            ));
        }
        output
            .try_reserve(fragment.len())
            .map_err(|_| invalid("ODS table-template output allocation failed"))?;
        output.push_str(&fragment);
    }
    Ok(output)
}

#[derive(Clone, Debug)]
struct Span {
    namespace: Option<String>,
    local: String,
    qname: String,
    start: usize,
    tag_end: usize,
    close_start: usize,
    end: usize,
    parent: Option<usize>,
    empty: bool,
    opaque_attributes: bool,
    opaque_content: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct EditRange {
    start: usize,
    end: usize,
}

fn render_source(
    source: Option<&str>,
    original: &[Template],
    draft: &[Template],
) -> Result<Option<String>> {
    if source.is_none() {
        return if draft.is_empty() {
            Ok(None)
        } else {
            Ok(Some(new_styles_xml(draft)?))
        };
    }
    let source = source.expect("source checked above");
    if source.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(
            "ODS styles.xml exceeds the table-template size limit".to_string(),
        ));
    }
    if draft.is_empty() && original.is_empty() {
        return Ok(Some(source.to_string()));
    }
    let spans = scan(source)?;
    let collections: Vec<usize> = spans
        .iter()
        .enumerate()
        .filter(|(_, span)| span.namespace.as_deref() == Some(OFFICE_NS) && span.local == "styles")
        .map(|(index, _)| index)
        .collect();
    let template_spans: Vec<usize> = spans
        .iter()
        .enumerate()
        .filter(|(_, span)| {
            span.namespace.as_deref() == Some(TABLE_NS)
                && span.local == "table-template"
                && span
                    .parent
                    .is_some_and(|parent| collections.contains(&parent))
        })
        .map(|(index, _)| index)
        .collect();
    if template_spans.len() != original.len() {
        return Err(Error::InvalidFormat(
            "ODS table-template source catalog changed before commit".to_string(),
        ));
    }

    let mut replacements = Vec::<(EditRange, String)>::new();
    let shared = original.len().min(draft.len());
    for index in 0..shared {
        if original[index] != draft[index] {
            reject_opaque_template(&spans, template_spans[index])?;
            let prefix = element_prefix(&spans[template_spans[index]].qname);
            replacements.push((
                EditRange {
                    start: spans[template_spans[index]].start,
                    end: spans[template_spans[index]].end,
                },
                template_xml(&draft[index], prefix, prefix.is_empty())?,
            ));
        }
    }
    for index in shared..original.len() {
        reject_opaque_template(&spans, template_spans[index])?;
        replacements.push((
            EditRange {
                start: spans[template_spans[index]].start,
                end: spans[template_spans[index]].end,
            },
            String::new(),
        ));
    }

    if draft.len() > original.len() {
        if let Some(collection) = collections.first().copied() {
            let additions = template_xmls(&draft[original.len()..], "table", true)?;
            if spans[collection].empty {
                let opening = source
                    .get(spans[collection].start..spans[collection].tag_end)
                    .ok_or_else(|| invalid("ODS style collection opening span is invalid"))?
                    .trim_end()
                    .strip_suffix("/>")
                    .ok_or_else(|| invalid("ODS style collection self-closing span is invalid"))?;
                let replacement = format!("{opening}>{additions}</{}>", spans[collection].qname);
                replacements.push((
                    EditRange {
                        start: spans[collection].start,
                        end: spans[collection].end,
                    },
                    replacement,
                ));
            } else {
                replacements.push((
                    EditRange {
                        start: spans[collection].close_start,
                        end: spans[collection].close_start,
                    },
                    additions,
                ));
            }
        } else {
            let root = spans
                .iter()
                .position(|span| span.parent.is_none())
                .ok_or_else(|| invalid("ODS styles.xml has no document-styles root"))?;
            let additions = template_xmls(&draft[original.len()..], "table", true)?;
            let fragment = format!(
                "<office:styles xmlns:office=\"{OFFICE_NS}\" xmlns:table=\"{TABLE_NS}\">{additions}</office:styles>"
            );
            let insertion = spans
                .iter()
                .find(|span| {
                    span.parent == Some(root)
                        && span.namespace.as_deref() == Some(OFFICE_NS)
                        && span.local == "automatic-styles"
                })
                .map_or(spans[root].close_start, |span| span.start);
            let replacement = if spans[root].empty {
                let opening = source
                    .get(spans[root].start..spans[root].tag_end)
                    .ok_or_else(|| invalid("ODS styles root opening span is invalid"))?
                    .trim_end()
                    .strip_suffix("/>")
                    .ok_or_else(|| invalid("ODS styles root self-closing span is invalid"))?;
                format!("{opening}>{fragment}</{}>", spans[root].qname)
            } else {
                fragment
            };
            replacements.push((
                EditRange {
                    start: if spans[root].empty {
                        spans[root].start
                    } else {
                        insertion
                    },
                    end: if spans[root].empty {
                        spans[root].end
                    } else {
                        insertion
                    },
                },
                replacement,
            ));
        }
    }

    if replacements.is_empty() {
        let mut output = String::new();
        output
            .try_reserve_exact(source.len())
            .map_err(|_| invalid("ODS styles XML allocation failed"))?;
        output.push_str(source);
        return Ok(Some(output));
    }
    replacements.sort_unstable_by(|left, right| {
        left.0
            .start
            .cmp(&right.0.start)
            .then(left.0.end.cmp(&right.0.end))
    });
    let mut output_len = source.len();
    let mut previous: Option<EditRange> = None;
    for (range, replacement) in &replacements {
        if range.end < range.start || range.end > source.len() {
            return Err(invalid("ODS table-template replacement span is invalid"));
        }
        if let Some(previous) = previous
            && (range.start < previous.end || range.start == previous.start)
        {
            return Err(invalid(
                "ODS table-template replacement spans overlap or share an insertion point",
            ));
        }
        previous = Some(*range);
        output_len = output_len
            .checked_sub(range.end - range.start)
            .and_then(|value| value.checked_add(replacement.len()))
            .ok_or_else(|| invalid("ODS styles XML output size overflows usize"))?;
        if output_len > MAX_XML_BYTES {
            return Err(Error::InvalidFormat(
                "patched ODS styles.xml exceeds the table-template size limit".to_string(),
            ));
        }
    }
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("ODS styles XML allocation failed"))?;
    let mut cursor = 0usize;
    for (range, replacement) in replacements {
        output.push_str(
            source
                .get(cursor..range.start)
                .ok_or_else(|| invalid("ODS table-template replacement span is invalid"))?,
        );
        output.push_str(&replacement);
        cursor = range.end;
    }
    output.push_str(
        source
            .get(cursor..)
            .ok_or_else(|| invalid("ODS table-template replacement span is invalid"))?,
    );
    if output.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(
            "patched ODS styles.xml exceeds the table-template size limit".to_string(),
        ));
    }
    Ok(Some(output))
}

fn new_styles_xml(templates: &[Template]) -> Result<String> {
    let mut output = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" \
xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"><office:styles>",
    );
    for template in templates {
        let fragment = template.to_xml()?;
        let next = output
            .len()
            .checked_add(fragment.len())
            .and_then(|length| {
                length.checked_add("</office:styles></office:document-styles>".len())
            })
            .ok_or_else(|| {
                Error::InvalidFormat("ODS styles XML size overflows usize".to_string())
            })?;
        if next > MAX_XML_BYTES {
            return Err(Error::InvalidFormat(
                "ODS styles XML exceeds the table-template size limit".to_string(),
            ));
        }
        output
            .try_reserve(fragment.len())
            .map_err(|_| Error::InvalidFormat("ODS styles XML allocation failed".to_string()))?;
        output.push_str(&fragment);
    }
    output.push_str("</office:styles></office:document-styles>");
    Ok(output)
}

pub(crate) fn styles_xml_for_templates(templates: &[Template]) -> Result<String> {
    validate_templates(templates)?;
    new_styles_xml(templates)
}

fn reject_opaque_template(spans: &[Span], template: usize) -> Result<()> {
    if spans.iter().enumerate().any(|(index, span)| {
        index != template
            && span.parent.is_some_and(|mut parent| {
                while let Some(current) = spans.get(parent) {
                    if parent == template {
                        return true;
                    }
                    parent = match current.parent {
                        Some(next) => next,
                        None => return false,
                    };
                }
                false
            })
            && (span.namespace.as_deref() != Some(TABLE_NS)
                || span.opaque_attributes
                || span.opaque_content)
    }) {
        return Err(Error::InvalidFormat(
            "changed ODS table-template contains opaque extension markup".to_string(),
        ));
    }
    if spans
        .get(template)
        .is_some_and(|span| span.opaque_attributes || span.opaque_content)
    {
        return Err(Error::InvalidFormat(
            "changed ODS table-template contains opaque extension markup".to_string(),
        ));
    }
    Ok(())
}

fn scan(xml: &str) -> Result<Vec<Span>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut spans = Vec::<Span>::new();
    let mut open = Vec::<usize>::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("ODS styles XML position overflows usize"))?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS styles XML: {error}")))?;
        let resolved = resolve_namespace(&resolved)?;
        let event = event.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("ODS styles XML position overflows usize"))?;
        match event {
            Event::Start(element) => {
                if spans.len() >= MAX_SPANS {
                    return Err(invalid("ODS styles XML exceeds the element limit"));
                }
                if open.len() >= MAX_XML_DEPTH {
                    return Err(invalid("ODS styles XML exceeds the nesting limit"));
                }
                spans
                    .try_reserve(1)
                    .map_err(|_| invalid("ODS styles span allocation failed"))?;
                open.try_reserve(1)
                    .map_err(|_| invalid("ODS styles stack allocation failed"))?;
                let index = spans.len();
                spans.push(Span {
                    namespace: resolved.clone(),
                    local: decode(element.local_name().as_ref(), "styles element local name")?,
                    qname: decode(element.name().as_ref(), "styles element qualified name")?,
                    start,
                    tag_end: end,
                    close_start: end,
                    end,
                    parent: open.last().copied(),
                    empty: false,
                    opaque_attributes: has_opaque_attributes(&reader, &element)?,
                    opaque_content: false,
                });
                open.push(index);
            },
            Event::Empty(element) => {
                if spans.len() >= MAX_SPANS {
                    return Err(invalid("ODS styles XML exceeds the element limit"));
                }
                spans
                    .try_reserve(1)
                    .map_err(|_| invalid("ODS styles span allocation failed"))?;
                spans.push(Span {
                    namespace: resolved,
                    local: decode(element.local_name().as_ref(), "styles element local name")?,
                    qname: decode(element.name().as_ref(), "styles element qualified name")?,
                    start,
                    tag_end: end,
                    close_start: end,
                    end,
                    parent: open.last().copied(),
                    empty: true,
                    opaque_attributes: has_opaque_attributes(&reader, &element)?,
                    opaque_content: false,
                });
            },
            Event::End(_) => {
                let index = open
                    .pop()
                    .ok_or_else(|| invalid("ODS styles XML span underflow"))?;
                spans[index].close_start = start;
                spans[index].end = end;
            },
            Event::Eof => break,
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in table templates"
                        .to_string(),
                ));
            },
            Event::Comment(_) => {
                if let Some(parent) = open.last().copied() {
                    spans[parent].opaque_content = true;
                }
            },
            Event::Decl(_) | Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if !open.is_empty() {
        return Err(invalid("ODS styles XML ended with open elements"));
    }
    Ok(spans)
}

fn has_opaque_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<bool> {
    let is_template = element.local_name().as_ref() == b"table-template";
    let mut opaque = false;
    let mut table_axes = [false; 4];
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid ODS styles attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        match resolved {
            ResolveResult::Bound(Namespace(uri)) => {
                let uri = std::str::from_utf8(uri)
                    .map_err(|_| invalid("ODS styles attribute namespace is not UTF-8"))?;
                if uri != TABLE_NS && uri != TEXT_NS {
                    opaque = true;
                } else if is_template && uri == TABLE_NS {
                    if let Some(axis) = template_axis(local.as_ref()) {
                        table_axes[axis] = true;
                    }
                    if is_legacy_template_flag(local.as_ref()) {
                        // These attributes were emitted by older producers on
                        // table:table-template, but the ODF grammar puts them
                        // on table:table.  The typed parser accepts them for a
                        // compatibility read; a changed rewrite must refuse
                        // rather than silently discard the source tokens.
                        opaque = true;
                    }
                } else if is_template && uri == TEXT_NS && template_axis(local.as_ref()).is_some() {
                    // Older producers used text:* aliases for the required
                    // table-template axis selectors.  They remain readable,
                    // but a canonical table:* rewrite would lose the exact
                    // source token and therefore cannot be changed in place.
                    opaque = true;
                }
            },
            ResolveResult::Unbound => opaque = true,
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unbound ODS styles attribute namespace prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        }
    }
    if is_template && table_axes.iter().any(|present| !present) {
        // ODF 1.4 requires all four selectors.  The semantic codec retains
        // the legacy source for exact no-ops, while this source scanner marks
        // an attempted rewrite as unsafe instead of synthesizing defaults.
        opaque = true;
    }
    Ok(opaque)
}

fn template_axis(local: &[u8]) -> Option<usize> {
    Some(match local {
        b"first-row-start-column" => 0,
        b"first-row-end-column" => 1,
        b"last-row-start-column" => 2,
        b"last-row-end-column" => 3,
        _ => return None,
    })
}

fn is_legacy_template_flag(local: &[u8]) -> bool {
    matches!(
        local,
        b"use-first-row-styles"
            | b"use-last-row-styles"
            | b"use-first-column-styles"
            | b"use-last-column-styles"
            | b"use-banding-rows-styles"
            | b"use-banding-columns-styles"
    )
}

fn resolve_namespace(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => Ok(Some(
            std::str::from_utf8(uri)
                .map_err(|_| invalid("ODS styles namespace is not UTF-8"))?
                .to_string(),
        )),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound ODS styles namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn decode(value: &[u8], label: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| Error::InvalidFormat(format!("{label} is not UTF-8")))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::{Snapshot, TABLE_NS};
    use crate::styles::table_template::{Region, Style, Template};

    const SOURCE: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:styles><office:font-face-decls/><table:table-template table:name="A" table:first-row-start-column="row" table:first-row-end-column="column" table:last-row-start-column="row" table:last-row-end-column="column"><table:body table:style-name="Body"/></table:table-template><x:extension xmlns:x="urn:example:ext">keep</x:extension></office:styles><office:automatic-styles/></office:document-styles>"#;

    fn template(name: &str) -> Template {
        Template::new(name).with_region(Region::Body, Style::new("Body"))
    }

    #[test]
    fn no_op_is_exact_and_changed_addition_preserves_unrelated_style_markup() {
        let snapshot = Snapshot::from_source(Some(SOURCE)).expect("valid styles source");
        let no_op = snapshot.edit().commit().expect("no-op commit");
        assert!(!no_op.changed());
        assert_eq!(no_op.snapshot().source_xml(), Some(SOURCE));

        let mut edit = snapshot.edit();
        edit.add(template("B")).expect("add template");
        let commit = edit.commit().expect("changed commit");
        assert!(commit.changed());
        let target = commit.snapshot().source_xml().expect("target source");
        assert!(target.contains("font-face-decls"));
        assert!(target.contains("x:extension"));
        assert_eq!(commit.snapshot().templates().len(), 2);
        assert_eq!(TABLE_NS, "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
    }

    #[test]
    fn inverse_and_stale_source_are_checked() {
        let snapshot = Snapshot::from_source(Some(SOURCE)).expect("valid styles source");
        let mut edit = snapshot.edit();
        edit.replace_named("A", template("A2"))
            .expect("replace template");
        let commit = edit.commit().expect("changed commit");
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("inverse applies");
        assert_eq!(restored.snapshot().source_xml(), Some(SOURCE));
        let other = Snapshot::from_source(Some(
            r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
        ))
            .expect("valid other source");
        assert!(commit.patch().apply(&other).is_err());
    }

    #[test]
    fn changed_template_preserves_source_table_namespace_alias() {
        let source = SOURCE
            .replace("xmlns:table=", "xmlns:t=")
            .replace("table:table-template", "t:table-template")
            .replace("table:name", "t:name")
            .replace("table:first-row-start-column", "t:first-row-start-column")
            .replace("table:first-row-end-column", "t:first-row-end-column")
            .replace("table:last-row-start-column", "t:last-row-start-column")
            .replace("table:last-row-end-column", "t:last-row-end-column")
            .replace("table:body", "t:body")
            .replace("table:style-name", "t:style-name");
        let snapshot = Snapshot::from_source(Some(&source)).expect("aliased source");
        let mut edit = snapshot.edit();
        edit.replace_named("A", template("A2"))
            .expect("replace template");
        let commit = edit.commit().expect("aliased commit");
        let target = commit.snapshot().source_xml().expect("target source");
        assert!(target.contains("<t:table-template"));
        assert!(target.contains("t:name=\"A2\""));
        assert!(!target.contains("<table:table-template"));
    }

    #[test]
    fn changed_template_refuses_opaque_attributes_and_comments() {
        for source in [
            r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:example:ext"><office:styles><table:table-template table:name="A" table:first-row-start-column="row" table:first-row-end-column="column" table:last-row-start-column="row" table:last-row-end-column="column" x:opaque="keep"><table:body table:style-name="Body"/></table:table-template></office:styles></office:document-styles>"#,
            r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:styles><table:table-template table:name="A" table:first-row-start-column="row" table:first-row-end-column="column" table:last-row-start-column="row" table:last-row-end-column="column"><!-- keep --><table:body table:style-name="Body"/></table:table-template></office:styles></office:document-styles>"#,
        ] {
            let snapshot = Snapshot::from_source(Some(source)).expect("valid styles source");
            let mut edit = snapshot.edit();
            edit.replace_named("A", template("A2"))
                .expect("replace template");
            assert!(edit.commit().is_err());

            let mut remove = snapshot.edit();
            remove.clear().expect("clear catalog");
            assert!(remove.commit().is_err());
        }
    }

    #[test]
    fn legacy_template_flags_are_read_but_changed_rewrites_refuse_loss() {
        let source = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:styles><table:table-template table:name="A" table:first-row-start-column="row" table:first-row-end-column="column" table:last-row-start-column="row" table:last-row-end-column="column" table:use-first-row-styles="true"><table:body table:style-name="Body"/></table:table-template></office:styles></office:document-styles>"#;
        let snapshot = Snapshot::from_source(Some(source)).expect("legacy source is readable");
        assert_eq!(snapshot.templates()[0].use_first_row_styles, Some(true));
        let no_op = snapshot
            .edit()
            .commit()
            .expect("no-op preserves legacy source");
        assert_eq!(no_op.snapshot().source_xml(), Some(source));

        let mut edit = snapshot.edit();
        edit.replace_named("A", template("A2"))
            .expect("stage replacement");
        assert!(edit.commit().is_err());
    }

    #[test]
    fn missing_or_legacy_axis_source_is_exact_noop_but_refuses_changed_rewrite() {
        for source in [
            r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:styles><table:table-template table:name="A"><table:body table:style-name="Body"/></table:table-template></office:styles></office:document-styles>"#,
            r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:styles><table:table-template table:name="A" text:first-row-start-column="row" table:first-row-end-column="column" table:last-row-start-column="row" table:last-row-end-column="column"><table:body table:style-name="Body"/></table:table-template></office:styles></office:document-styles>"#,
        ] {
            let snapshot = Snapshot::from_source(Some(source)).expect("legacy source readable");
            let no_op = snapshot.edit().commit().expect("exact no-op");
            assert_eq!(no_op.snapshot().source_xml(), Some(source));
            let mut edit = snapshot.edit();
            edit.replace_named("A", template("A2"))
                .expect("stage replacement");
            assert!(edit.commit().is_err());
        }
    }

    #[test]
    fn table_templates_are_inserted_only_under_office_styles() {
        let source = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:automatic-styles/></office:document-styles>"#;
        let snapshot = Snapshot::from_source(Some(source)).expect("styles source");
        let mut edit = snapshot.edit();
        edit.add(template("A")).expect("stage template");
        let target = edit.commit().expect("add styles owner");
        let xml = target.snapshot().source_xml().expect("target source");
        assert!(xml.contains("<office:styles"));
        assert!(
            xml.contains("<office:automatic-styles/>")
                || xml.contains("<office:automatic-styles />")
        );
        let styles_start = xml.find("<office:styles").expect("styles owner");
        let automatic_start = xml
            .find("<office:automatic-styles")
            .expect("automatic owner");
        let template_start = xml.find("<table:table-template").expect("template");
        assert!(styles_start < template_start && template_start < automatic_start);
    }
}
