//! Resolution of ODF table-cell protection styles.

use super::names::formula;
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace as XmlNamespace, NamespaceResolver, QName, ResolveResult},
    reader::NsReader,
};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_CONDITIONAL_STYLES: usize = 65_536;
const MAX_RULES_PER_STYLE: usize = 1_024;
const MAX_CONDITIONAL_RULES: usize = 262_144;
const MAX_CONDITIONAL_ATTRIBUTE_BYTES: usize = 64 * 1024;
const MAX_CONDITIONAL_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_STYLE_DEPTH: usize = 64;

/// One standard ODF conditional table-cell style.
///
/// Rules are retained in document order. Litchi does not evaluate their
/// conditions or compute an effective style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalStyle {
    pub style_name: String,
    pub parent_style_name: Option<String>,
    pub rules: Vec<Rule>,
}

/// One inert `style:map` rule belonging to a table-cell style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Rule {
    /// The decoded condition text. It is never evaluated by litchi.
    pub condition: String,
    /// Namespace bound to a condition prefix, when the condition is qualified.
    pub formula_namespace: Option<formula::Namespace>,
    /// Name of the common table-cell style referenced by the rule.
    pub apply_style_name: String,
    /// Optional lexical base cell address for relative formula references.
    pub base_cell_address: Option<String>,
}

impl ConditionalStyle {
    /// Create an inert conditional table-cell style.
    pub fn new(style_name: impl Into<String>, rules: Vec<Rule>) -> Self {
        Self {
            style_name: style_name.into(),
            parent_style_name: None,
            rules,
        }
    }

    /// Set the optional parent table-cell style name.
    pub fn with_parent_style_name(mut self, parent_style_name: impl Into<String>) -> Self {
        self.parent_style_name = Some(parent_style_name.into());
        self
    }
}

impl Rule {
    /// Create an inert conditional rule without a formula namespace or base address.
    pub fn new(condition: impl Into<String>, apply_style_name: impl Into<String>) -> Self {
        Self {
            condition: condition.into(),
            formula_namespace: None,
            apply_style_name: apply_style_name.into(),
            base_cell_address: None,
        }
    }

    /// Bind the lexical condition prefix to a namespace URI.
    pub fn with_formula_namespace(mut self, namespace: formula::Namespace) -> Self {
        self.formula_namespace = Some(namespace);
        self
    }

    /// Set the optional lexical base cell address.
    pub fn with_base_cell_address(mut self, address: impl Into<String>) -> Self {
        self.base_cell_address = Some(address.into());
        self
    }
}

/// The effective value of the ODF `style:cell-protect` property.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Protection {
    None,
    Protected,
    FormulaHidden,
    ProtectedFormulaHidden,
    HiddenAndProtected,
}

/// One automatic table-cell style with an explicit `style:cell-protect` value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableStyle {
    pub style_name: String,
    pub parent_style_name: Option<String>,
    pub protection: Protection,
}

impl TableStyle {
    pub fn new(style_name: impl Into<String>, protection: Protection) -> Self {
        Self {
            style_name: style_name.into(),
            parent_style_name: None,
            protection,
        }
    }

    pub fn with_parent_style_name(mut self, parent_style_name: impl Into<String>) -> Self {
        self.parent_style_name = Some(parent_style_name.into());
        self
    }
}

impl Protection {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "protected" => Ok(Self::Protected),
            "formula-hidden" => Ok(Self::FormulaHidden),
            "protected formula-hidden" | "formula-hidden protected" => {
                Ok(Self::ProtectedFormulaHidden)
            },
            "hidden-and-protected" => Ok(Self::HiddenAndProtected),
            _ => Err(Error::InvalidFormat(format!(
                "invalid style:cell-protect value '{value}'"
            ))),
        }
    }

    /// Whether editing is prohibited while the containing sheet is protected.
    pub fn is_protected(self) -> bool {
        matches!(
            self,
            Self::Protected | Self::ProtectedFormulaHidden | Self::HiddenAndProtected
        )
    }

    /// Whether the cell formula is hidden while protection is active.
    pub fn hides_formula(self) -> bool {
        matches!(self, Self::FormulaHidden | Self::ProtectedFormulaHidden)
    }

    /// Whether both cell content and editing are hidden/protected.
    pub fn hides_content(self) -> bool {
        self == Self::HiddenAndProtected
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Protected => "protected",
            Self::FormulaHidden => "formula-hidden",
            Self::ProtectedFormulaHidden => "protected formula-hidden",
            Self::HiddenAndProtected => "hidden-and-protected",
        }
    }
}

#[derive(Clone, Debug, Default)]
pub(crate) struct CellStyleRegistry {
    styles: HashMap<String, CellStyleDefinition>,
    default: Option<Protection>,
    common_table_cell_styles: HashSet<String>,
    style_order: Vec<String>,
    conditional_styles: Vec<ConditionalStyle>,
    conditional_style_index: HashMap<String, usize>,
    automatic_protection_styles: Vec<TableStyle>,
    parsed_conditional_styles: usize,
    parsed_conditional_rules: usize,
    conditional_text_bytes: usize,
}

#[derive(Clone, Debug)]
pub(crate) struct PreservedXmlFragment {
    pub xml: String,
    namespaces: BTreeMap<String, String>,
}

impl PreservedXmlFragment {
    pub(crate) fn namespace_prefixes(&self) -> impl Iterator<Item = &str> {
        self.namespaces.keys().map(String::as_str)
    }

    pub(crate) fn write_missing_namespaces<'a>(
        &self,
        out: &mut String,
        already_declared: impl IntoIterator<Item = &'a str>,
    ) {
        let declared = already_declared.into_iter().collect::<HashSet<_>>();
        for (prefix, uri) in &self.namespaces {
            if declared.contains(prefix.as_str()) {
                continue;
            }
            out.push_str(" xmlns:");
            out.push_str(prefix);
            out.push_str("=\"");
            out.push_str(&escape_xml(uri));
            out.push('"');
        }
    }
}

pub(crate) fn common_table_cell_style_names(styles_xml: Option<&str>) -> Result<HashSet<String>> {
    let registry = CellStyleRegistry::parse(styles_xml, "")?;
    Ok(registry.common_table_cell_styles)
}

pub(crate) fn validate_protection_style_document(
    styles_xml: Option<&str>,
    automatic_styles_xml: &str,
    authored: &[TableStyle],
) -> Result<()> {
    let registry = CellStyleRegistry::parse(styles_xml, automatic_styles_xml)?;
    for style in authored {
        let mut current = style.style_name.as_str();
        let mut visited = HashSet::new();
        loop {
            if !visited.insert(current.to_string()) {
                return Err(Error::InvalidFormat(format!(
                    "cyclic table-cell style inheritance at '{current}'"
                )));
            }
            let definition = registry.styles.get(current).ok_or_else(|| {
                Error::InvalidFormat(format!("missing table-cell style '{current}'"))
            })?;
            let Some(parent) = definition.parent.as_deref() else {
                break;
            };
            current = parent;
        }
    }
    Ok(())
}

pub(crate) fn validate_style_name(name: &str, label: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::InvalidFormat(format!("{label} must not be empty")));
    }
    check_conditional_attribute_size(label, name)
}

pub(crate) fn validate_conditional_style_collection(
    styles: &[ConditionalStyle],
    common_styles: &HashSet<String>,
) -> Result<()> {
    if styles.len() > MAX_CONDITIONAL_STYLES {
        return Err(Error::InvalidFormat(format!(
            "document exceeds the {MAX_CONDITIONAL_STYLES} conditional style limit"
        )));
    }
    let mut names = HashSet::with_capacity(styles.len());
    let mut total_rules = 0usize;
    let mut total_text = 0usize;
    for style in styles {
        validate_style_name(&style.style_name, "conditional style name")?;
        if !names.insert(style.style_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate conditional style name '{}'",
                style.style_name
            )));
        }
        if style.rules.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "conditional style '{}' must contain at least one rule",
                style.style_name
            )));
        }
        if style.rules.len() > MAX_RULES_PER_STYLE {
            return Err(Error::InvalidFormat(format!(
                "conditional style '{}' exceeds the {MAX_RULES_PER_STYLE} rule limit",
                style.style_name
            )));
        }
        if let Some(parent) = &style.parent_style_name {
            validate_style_name(parent, "parent style name")?;
        }
        total_rules = total_rules
            .checked_add(style.rules.len())
            .ok_or_else(|| Error::InvalidFormat("conditional rule count overflow".to_string()))?;
        if total_rules > MAX_CONDITIONAL_RULES {
            return Err(Error::InvalidFormat(format!(
                "document exceeds the {MAX_CONDITIONAL_RULES} conditional rule limit"
            )));
        }
        for rule in &style.rules {
            check_conditional_attribute_size("style:condition", &rule.condition)?;
            if rule.condition.is_empty() {
                return Err(Error::InvalidFormat(
                    "style:condition must not be empty".to_string(),
                ));
            }
            validate_style_name(&rule.apply_style_name, "style:apply-style-name")?;
            if !common_styles.contains(&rule.apply_style_name) {
                return Err(Error::InvalidFormat(format!(
                    "conditional style '{}' references missing, automatic, or non-table-cell common style '{}'",
                    style.style_name, rule.apply_style_name
                )));
            }
            if let Some(base) = &rule.base_cell_address {
                check_conditional_attribute_size("style:base-cell-address", base)?;
                if base.is_empty() {
                    return Err(Error::InvalidFormat(
                        "style:base-cell-address must not be empty".to_string(),
                    ));
                }
            }
            validate_formula_namespace(rule)?;
            total_text = total_text
                .checked_add(rule.condition.len())
                .and_then(|value| value.checked_add(rule.apply_style_name.len()))
                .and_then(|value| {
                    value.checked_add(rule.base_cell_address.as_deref().map_or(0, str::len))
                })
                .ok_or_else(|| {
                    Error::InvalidFormat("conditional style text size overflow".to_string())
                })?;
            if total_text > MAX_CONDITIONAL_TEXT_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "conditional style text exceeds the {MAX_CONDITIONAL_TEXT_BYTES} byte limit"
                )));
            }
        }
    }
    Ok(())
}

fn validate_formula_namespace(rule: &Rule) -> Result<()> {
    let lexical_prefix = formula_prefix(&rule.condition);
    match (lexical_prefix, &rule.formula_namespace) {
        (None, None) => Ok(()),
        (Some(prefix), Some(namespace)) if prefix == namespace.prefix => {
            validate_xml_prefix(&namespace.prefix)?;
            if namespace.uri.is_empty() {
                return Err(Error::InvalidFormat(
                    "conditional formula namespace URI must not be empty".to_string(),
                ));
            }
            check_conditional_attribute_size("formula namespace URI", &namespace.uri)
        },
        (Some(prefix), Some(namespace)) => Err(Error::InvalidFormat(format!(
            "condition prefix '{prefix}' does not match formula namespace prefix '{}'",
            namespace.prefix
        ))),
        (Some(prefix), None) => Err(Error::InvalidFormat(format!(
            "conditional style condition uses unbound namespace prefix '{prefix}'"
        ))),
        (None, Some(namespace)) => Err(Error::InvalidFormat(format!(
            "formula namespace prefix '{}' is not used by the condition",
            namespace.prefix
        ))),
    }
}

fn formula_prefix(condition: &str) -> Option<&str> {
    let (prefix, _) = condition.split_once(':')?;
    let mut characters = prefix.chars();
    if characters
        .next()
        .is_some_and(|character| character == '_' || character.is_alphabetic())
        && characters.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
    {
        Some(prefix)
    } else {
        None
    }
}

fn validate_xml_prefix(prefix: &str) -> Result<()> {
    if formula_prefix(&format!("{prefix}:x")) != Some(prefix) || matches!(prefix, "xml" | "xmlns") {
        return Err(Error::InvalidFormat(format!(
            "invalid conditional formula namespace prefix '{prefix}'"
        )));
    }
    Ok(())
}

#[cfg_attr(not(test), allow(dead_code))]
pub(crate) fn rewrite_conditional_styles(
    fragment: Option<&PreservedXmlFragment>,
    styles: &[ConditionalStyle],
) -> Result<PreservedXmlFragment> {
    let canonical = write_conditional_styles(styles);
    let Some(fragment) = fragment else {
        return Ok(PreservedXmlFragment {
            xml: format!(
                "<office:automatic-styles xmlns:office=\"{}\">{canonical}</office:automatic-styles>",
                String::from_utf8_lossy(OFFICE_NAMESPACE)
            ),
            namespaces: BTreeMap::new(),
        });
    };
    let (ranges, insertion) = conditional_style_ranges(&fragment.xml)?;
    let xml = match insertion {
        AutomaticStylesInsertion::BeforeEnd(position) => {
            let mut out = String::with_capacity(fragment.xml.len() + canonical.len());
            let mut cursor = 0usize;
            for range in ranges {
                if range.end > position {
                    return Err(Error::InvalidFormat(
                        "conditional style range exceeds automatic styles container".to_string(),
                    ));
                }
                out.push_str(&fragment.xml[cursor..range.start]);
                cursor = range.end;
            }
            out.push_str(&fragment.xml[cursor..position]);
            out.push_str(&canonical);
            out.push_str(&fragment.xml[position..]);
            out
        },
        AutomaticStylesInsertion::ExpandEmpty { slash, name } => {
            let mut out = String::with_capacity(fragment.xml.len() + canonical.len() + name.len());
            out.push_str(&fragment.xml[..slash]);
            out.push('>');
            out.push_str(&canonical);
            out.push_str("</");
            out.push_str(&name);
            out.push('>');
            out
        },
    };
    Ok(PreservedXmlFragment {
        xml,
        namespaces: fragment.namespaces.clone(),
    })
}

pub(crate) fn validate_protection_style_collection(styles: &[TableStyle]) -> Result<()> {
    if styles.len() > MAX_CONDITIONAL_STYLES {
        return Err(Error::InvalidFormat(format!(
            "document exceeds the {MAX_CONDITIONAL_STYLES} automatic protection style limit"
        )));
    }
    let mut names = HashSet::with_capacity(styles.len());
    for style in styles {
        validate_style_name(&style.style_name, "protection style name")?;
        if !names.insert(style.style_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate protection style name '{}'",
                style.style_name
            )));
        }
        if let Some(parent) = &style.parent_style_name {
            validate_style_name(parent, "parent style name")?;
        }
    }
    Ok(())
}

pub(crate) fn rewrite_managed_cell_styles(
    fragment: Option<&PreservedXmlFragment>,
    conditional_styles: &[ConditionalStyle],
    protection_styles: &[TableStyle],
) -> Result<PreservedXmlFragment> {
    validate_protection_style_collection(protection_styles)?;
    for conditional in conditional_styles {
        if let Some(protection) = protection_styles
            .iter()
            .find(|style| style.style_name == conditional.style_name)
            && conditional.parent_style_name != protection.parent_style_name
        {
            return Err(Error::InvalidFormat(format!(
                "conditional and protection definitions for '{}' have different parent styles",
                conditional.style_name
            )));
        }
    }
    let canonical = write_managed_styles(conditional_styles, protection_styles);
    let Some(fragment) = fragment else {
        return Ok(PreservedXmlFragment {
            xml: format!(
                "<office:automatic-styles xmlns:office=\"{}\">{canonical}</office:automatic-styles>",
                String::from_utf8_lossy(OFFICE_NAMESPACE)
            ),
            namespaces: BTreeMap::new(),
        });
    };
    let (ranges, insertion) = managed_style_ranges(&fragment.xml)?;
    let xml = rewrite_ranges(&fragment.xml, ranges, insertion, &canonical)?;
    Ok(PreservedXmlFragment {
        xml,
        namespaces: fragment.namespaces.clone(),
    })
}

fn rewrite_ranges(
    xml: &str,
    ranges: Vec<Range<usize>>,
    insertion: AutomaticStylesInsertion,
    canonical: &str,
) -> Result<String> {
    match insertion {
        AutomaticStylesInsertion::BeforeEnd(position) => {
            let mut out = String::with_capacity(xml.len() + canonical.len());
            let mut cursor = 0usize;
            for range in ranges {
                if range.end > position {
                    return Err(Error::InvalidFormat(
                        "managed style range exceeds automatic styles container".to_string(),
                    ));
                }
                out.push_str(&xml[cursor..range.start]);
                cursor = range.end;
            }
            out.push_str(&xml[cursor..position]);
            out.push_str(canonical);
            out.push_str(&xml[position..]);
            Ok(out)
        },
        AutomaticStylesInsertion::ExpandEmpty { slash, name } => {
            Ok(format!("{}>{canonical}</{name}>", &xml[..slash]))
        },
    }
}

fn managed_style_ranges(xml: &str) -> Result<(Vec<Range<usize>>, AutomaticStylesInsertion)> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut candidate: Option<(usize, bool, bool)> = None;
    let mut ranges = Vec::new();
    let mut insertion = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_style_namespace = is_namespace(&namespace, STYLE_NAMESPACE);
        let is_office_namespace = is_namespace(&namespace, OFFICE_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not allowed in automatic styles XML".to_string(),
                ));
            },
            Event::Start(element) => {
                if depth == 1 && is_style_namespace && element.local_name().as_ref() == b"style" {
                    let is_cell = optional_attribute(
                        reader.resolver(),
                        reader.decoder(),
                        &element,
                        b"family",
                    )?
                    .as_deref()
                        == Some("table-cell");
                    candidate = Some((event_start, is_cell, false));
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && (element.local_name().as_ref() == b"map"
                        || element.local_name().as_ref() == b"table-cell-properties"
                            && optional_attribute(
                                reader.resolver(),
                                reader.decoder(),
                                &element,
                                b"cell-protect",
                            )?
                            .is_some())
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    let slash = xml[..event_end].rfind("/>").ok_or_else(|| {
                        Error::InvalidFormat("malformed empty automatic styles".to_string())
                    })?;
                    let name =
                        String::from_utf8(element.name().as_ref().to_vec()).map_err(|_| {
                            Error::InvalidFormat("automatic styles name is not UTF-8".to_string())
                        })?;
                    insertion = Some(AutomaticStylesInsertion::ExpandEmpty { slash, name });
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && (element.local_name().as_ref() == b"map"
                        || element.local_name().as_ref() == b"table-cell-properties"
                            && optional_attribute(
                                reader.resolver(),
                                reader.decoder(),
                                &element,
                                b"cell-protect",
                            )?
                            .is_some())
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid automatic styles depth".to_string())
                })?;
                if depth == 1
                    && is_style_namespace
                    && element.local_name().as_ref() == b"style"
                    && let Some((start, is_cell, managed)) = candidate.take()
                    && is_cell
                    && managed
                {
                    ranges.push(start..event_end);
                }
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    insertion = Some(AutomaticStylesInsertion::BeforeEnd(event_start));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok((
        ranges,
        insertion.ok_or_else(|| {
            Error::InvalidFormat("missing office:automatic-styles container".to_string())
        })?,
    ))
}

fn write_managed_styles(conditionals: &[ConditionalStyle], protections: &[TableStyle]) -> String {
    let formula_prefixes = conditionals
        .iter()
        .flat_map(|style| &style.rules)
        .filter_map(|rule| rule.formula_namespace.as_ref())
        .map(|namespace| namespace.prefix.as_str())
        .collect::<HashSet<_>>();
    let mut style_prefix = "style".to_string();
    let mut suffix = 0usize;
    while formula_prefixes.contains(style_prefix.as_str()) {
        suffix += 1;
        style_prefix = format!("style{suffix}");
    }
    let mut out = String::new();
    for conditional in conditionals {
        let protection = protections
            .iter()
            .find(|style| style.style_name == conditional.style_name)
            .map(|style| style.protection);
        write_managed_style(
            &mut out,
            &style_prefix,
            &conditional.style_name,
            conditional.parent_style_name.as_deref(),
            protection,
            &conditional.rules,
        );
    }
    for protection in protections {
        if conditionals
            .iter()
            .any(|style| style.style_name == protection.style_name)
        {
            continue;
        }
        write_managed_style(
            &mut out,
            &style_prefix,
            &protection.style_name,
            protection.parent_style_name.as_deref(),
            Some(protection.protection),
            &[],
        );
    }
    out
}

fn write_managed_style(
    out: &mut String,
    prefix: &str,
    name: &str,
    parent: Option<&str>,
    protection: Option<Protection>,
    rules: &[Rule],
) {
    out.push('<');
    out.push_str(prefix);
    out.push_str(":style xmlns:");
    out.push_str(prefix);
    out.push_str("=\"");
    out.push_str(&escape_xml(&String::from_utf8_lossy(STYLE_NAMESPACE)));
    out.push_str("\" ");
    out.push_str(prefix);
    out.push_str(":name=\"");
    out.push_str(&escape_xml(name));
    out.push_str("\" ");
    out.push_str(prefix);
    out.push_str(":family=\"table-cell\"");
    if let Some(parent) = parent {
        out.push(' ');
        out.push_str(prefix);
        out.push_str(":parent-style-name=\"");
        out.push_str(&escape_xml(parent));
        out.push('"');
    }
    out.push('>');
    if let Some(protection) = protection {
        out.push('<');
        out.push_str(prefix);
        out.push_str(":table-cell-properties ");
        out.push_str(prefix);
        out.push_str(":cell-protect=\"");
        out.push_str(protection.as_str());
        out.push_str("\"/>");
    }
    for rule in rules {
        out.push('<');
        out.push_str(prefix);
        out.push_str(":map");
        if let Some(namespace) = &rule.formula_namespace {
            out.push_str(" xmlns:");
            out.push_str(&namespace.prefix);
            out.push_str("=\"");
            out.push_str(&escape_xml(&namespace.uri));
            out.push('"');
        }
        out.push(' ');
        out.push_str(prefix);
        out.push_str(":condition=\"");
        out.push_str(&escape_xml(&rule.condition));
        out.push_str("\" ");
        out.push_str(prefix);
        out.push_str(":apply-style-name=\"");
        out.push_str(&escape_xml(&rule.apply_style_name));
        out.push('"');
        if let Some(base) = &rule.base_cell_address {
            out.push(' ');
            out.push_str(prefix);
            out.push_str(":base-cell-address=\"");
            out.push_str(&escape_xml(base));
            out.push('"');
        }
        out.push_str("/>");
    }
    out.push_str("</");
    out.push_str(prefix);
    out.push_str(":style>");
}

enum AutomaticStylesInsertion {
    BeforeEnd(usize),
    ExpandEmpty { slash: usize, name: String },
}

#[cfg_attr(not(test), allow(dead_code))]
fn conditional_style_ranges(xml: &str) -> Result<(Vec<Range<usize>>, AutomaticStylesInsertion)> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut candidate: Option<(usize, bool, bool)> = None;
    let mut ranges = Vec::new();
    let mut insertion = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_style_namespace = is_namespace(&namespace, STYLE_NAMESPACE);
        let is_office_namespace = is_namespace(&namespace, OFFICE_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not allowed in automatic styles XML".to_string(),
                ));
            },
            Event::Start(element) => {
                if depth == 1 && is_style_namespace && element.local_name().as_ref() == b"style" {
                    let is_cell = optional_attribute(
                        reader.resolver(),
                        reader.decoder(),
                        &element,
                        b"family",
                    )?
                    .as_deref()
                        == Some("table-cell");
                    candidate = Some((event_start, is_cell, false));
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && element.local_name().as_ref() == b"map"
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    let slash = xml[..event_end].rfind("/>").ok_or_else(|| {
                        Error::InvalidFormat("malformed empty automatic styles".to_string())
                    })?;
                    let name =
                        String::from_utf8(element.name().as_ref().to_vec()).map_err(|_| {
                            Error::InvalidFormat("automatic styles name is not UTF-8".to_string())
                        })?;
                    insertion = Some(AutomaticStylesInsertion::ExpandEmpty { slash, name });
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && element.local_name().as_ref() == b"map"
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid automatic styles depth".to_string())
                })?;
                if depth == 1
                    && is_style_namespace
                    && element.local_name().as_ref() == b"style"
                    && let Some((start, is_cell, has_map)) = candidate.take()
                    && is_cell
                    && has_map
                {
                    ranges.push(start..event_end);
                }
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    insertion = Some(AutomaticStylesInsertion::BeforeEnd(event_start));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    let insertion = insertion.ok_or_else(|| {
        Error::InvalidFormat("missing office:automatic-styles container".to_string())
    })?;
    Ok((ranges, insertion))
}

#[cfg_attr(not(test), allow(dead_code))]
fn write_conditional_styles(styles: &[ConditionalStyle]) -> String {
    let formula_prefixes = styles
        .iter()
        .flat_map(|style| &style.rules)
        .filter_map(|rule| rule.formula_namespace.as_ref())
        .map(|namespace| namespace.prefix.as_str())
        .collect::<HashSet<_>>();
    let mut style_prefix = "style".to_string();
    let mut suffix = 0usize;
    while formula_prefixes.contains(style_prefix.as_str()) {
        suffix += 1;
        style_prefix = format!("style{suffix}");
    }
    let mut out = String::new();
    for style in styles {
        out.push('<');
        out.push_str(&style_prefix);
        out.push_str(":style xmlns:");
        out.push_str(&style_prefix);
        out.push_str("=\"");
        out.push_str(&escape_xml(&String::from_utf8_lossy(STYLE_NAMESPACE)));
        out.push_str("\" ");
        out.push_str(&style_prefix);
        out.push_str(":name=\"");
        out.push_str(&escape_xml(&style.style_name));
        out.push_str("\" ");
        out.push_str(&style_prefix);
        out.push_str(":family=\"table-cell\"");
        if let Some(parent) = &style.parent_style_name {
            out.push(' ');
            out.push_str(&style_prefix);
            out.push_str(":parent-style-name=\"");
            out.push_str(&escape_xml(parent));
            out.push('"');
        }
        out.push('>');
        for rule in &style.rules {
            out.push('<');
            out.push_str(&style_prefix);
            out.push_str(":map");
            if let Some(namespace) = &rule.formula_namespace {
                out.push_str(" xmlns:");
                out.push_str(&namespace.prefix);
                out.push_str("=\"");
                out.push_str(&escape_xml(&namespace.uri));
                out.push('"');
            }
            out.push(' ');
            out.push_str(&style_prefix);
            out.push_str(":condition=\"");
            out.push_str(&escape_xml(&rule.condition));
            out.push_str("\" ");
            out.push_str(&style_prefix);
            out.push_str(":apply-style-name=\"");
            out.push_str(&escape_xml(&rule.apply_style_name));
            out.push('"');
            if let Some(base) = &rule.base_cell_address {
                out.push(' ');
                out.push_str(&style_prefix);
                out.push_str(":base-cell-address=\"");
                out.push_str(&escape_xml(base));
                out.push('"');
            }
            out.push_str("/>");
        }
        out.push_str("</");
        out.push_str(&style_prefix);
        out.push_str(":style>");
    }
    out
}

pub(crate) fn extract_automatic_styles(xml: &str) -> Result<Option<PreservedXmlFragment>> {
    extract_office_fragment(xml, b"automatic-styles")
}

pub(crate) fn extract_font_face_decls(xml: &str) -> Result<Option<PreservedXmlFragment>> {
    extract_office_fragment(xml, b"font-face-decls")
}

fn extract_office_fragment(
    xml: &str,
    expected_local_name: &[u8],
) -> Result<Option<PreservedXmlFragment>> {
    let fragment_name = String::from_utf8_lossy(expected_local_name);
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut namespaces = BTreeMap::new();
    let mut range_start = None;
    let mut depth = 0usize;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_office_namespace = is_namespace(&namespace, OFFICE_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element)
                if is_office_namespace && element.local_name().as_ref() == b"document-content" =>
            {
                collect_namespaces(&reader, &element, &mut namespaces)?;
            },
            Event::Start(element)
                if is_office_namespace && element.local_name().as_ref() == expected_local_name =>
            {
                if range_start.is_some() {
                    return Err(Error::InvalidFormat(
                        "nested preserved office fragment".to_string(),
                    ));
                }
                range_start = Some(event_start);
                depth = 1;
            },
            Event::Empty(element)
                if is_office_namespace && element.local_name().as_ref() == expected_local_name =>
            {
                if range_start.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate preserved office fragment".to_string(),
                    ));
                }
                return Ok(Some(PreservedXmlFragment {
                    xml: xml[event_start..event_end].to_string(),
                    namespaces,
                }));
            },
            Event::Start(_) if range_start.is_some() => depth += 1,
            Event::End(element) if range_start.is_some() => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid office:{fragment_name} depth"))
                })?;
                if depth == 0 {
                    if !is_office_namespace || element.local_name().as_ref() != expected_local_name
                    {
                        return Err(Error::InvalidFormat(format!(
                            "malformed office:{fragment_name} element"
                        )));
                    }
                    let start = range_start.take().expect("checked range");
                    return Ok(Some(PreservedXmlFragment {
                        xml: xml[start..event_end].to_string(),
                        namespaces,
                    }));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if range_start.is_some() {
        return Err(Error::InvalidFormat(format!(
            "unterminated office:{fragment_name} element"
        )));
    }
    Ok(None)
}

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

fn collect_namespaces(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespaces: &mut BTreeMap<String, String>,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        let prefix = String::from_utf8(prefix.to_vec())
            .map_err(|_| Error::InvalidFormat("namespace prefix is not valid UTF-8".to_string()))?;
        let uri = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid namespace URI: {error}")))?
            .into_owned();
        namespaces.insert(prefix, uri);
    }
    Ok(())
}

#[derive(Clone, Debug)]
struct CellStyleDefinition {
    parent: Option<String>,
    protection: Option<Protection>,
    conditional_rules: Vec<Rule>,
    is_common: bool,
}

impl CellStyleRegistry {
    pub(crate) fn parse(named_styles: Option<&str>, content: &str) -> Result<Self> {
        let mut registry = Self::default();
        if let Some(named_styles) = named_styles {
            registry.parse_part(named_styles)?;
        }
        // Automatic styles are the closest scope and intentionally replace a
        // same-named definition from styles.xml.
        registry.parse_part(content)?;
        registry.finish_conditional_styles()?;
        Ok(registry)
    }

    pub(crate) fn conditional_styles(&self) -> &[ConditionalStyle] {
        &self.conditional_styles
    }

    pub(crate) fn conditional_style(&self, name: &str) -> Option<&ConditionalStyle> {
        self.conditional_style_index
            .get(name)
            .and_then(|index| self.conditional_styles.get(*index))
    }

    pub(crate) fn automatic_protection_styles(&self) -> &[TableStyle] {
        &self.automatic_protection_styles
    }

    pub(crate) fn resolve(&self, style_name: Option<&str>) -> Result<Option<Protection>> {
        let Some(mut name) = style_name else {
            return Ok(self.default);
        };
        let mut visited = HashSet::new();
        loop {
            if !visited.insert(name.to_string()) {
                return Err(Error::InvalidFormat(format!(
                    "cyclic table-cell style inheritance at '{name}'"
                )));
            }
            let style = self.styles.get(name).ok_or_else(|| {
                Error::InvalidFormat(format!("missing table-cell style '{name}'"))
            })?;
            if let Some(protection) = style.protection {
                return Ok(Some(protection));
            }
            match style.parent.as_deref() {
                Some(parent) => name = parent,
                None => return Ok(self.default),
            }
        }
    }

    fn parse_part(&mut self, xml: &str) -> Result<()> {
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        let mut current: Option<StyleBuilder> = None;
        let mut depth = 0usize;
        let mut open_map = false;
        let mut in_common_styles = false;
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::DocType(_) => {
                    return Err(Error::InvalidFormat(
                        "DOCTYPE is not allowed in ODF style XML".to_string(),
                    ));
                },
                Event::Start(element)
                    if current.is_none()
                        && is_namespace(&namespace, OFFICE_NAMESPACE)
                        && element.local_name().as_ref() == b"styles" =>
                {
                    in_common_styles = true;
                },
                Event::End(element)
                    if current.is_none()
                        && is_namespace(&namespace, OFFICE_NAMESPACE)
                        && element.local_name().as_ref() == b"styles" =>
                {
                    in_common_styles = false;
                },
                Event::Start(element)
                    if is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    if current.is_some() {
                        increment_style_depth(&mut depth)?;
                    } else if let Some(builder) =
                        StyleBuilder::new(&reader, &element, in_common_styles)?
                    {
                        current = Some(builder);
                    }
                },
                Event::Empty(element)
                    if is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    if current.is_none()
                        && let Some(builder) =
                            StyleBuilder::new(&reader, &element, in_common_styles)?
                    {
                        self.finish(builder)?;
                    }
                },
                Event::Start(element)
                    if current.is_some()
                        && depth == 0
                        && is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"map" =>
                {
                    if current.as_ref().is_some_and(|builder| builder.is_default) {
                        return Err(Error::InvalidFormat(
                            "style:map is not allowed on a default table-cell style".to_string(),
                        ));
                    }
                    let rule = parse_conditional_rule(&reader, &element)?;
                    self.record_conditional_rule(&rule)?;
                    let builder = current.as_mut().expect("checked style");
                    if builder.conditional_rules.len() >= MAX_RULES_PER_STYLE {
                        return Err(Error::InvalidFormat(format!(
                            "table-cell style exceeds the {MAX_RULES_PER_STYLE} conditional rule limit"
                        )));
                    }
                    builder.conditional_rules.push(rule);
                    increment_style_depth(&mut depth)?;
                    open_map = true;
                },
                Event::Empty(element)
                    if current.is_some()
                        && depth == 0
                        && is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"map" =>
                {
                    if current.as_ref().is_some_and(|builder| builder.is_default) {
                        return Err(Error::InvalidFormat(
                            "style:map is not allowed on a default table-cell style".to_string(),
                        ));
                    }
                    let rule = parse_conditional_rule(&reader, &element)?;
                    self.record_conditional_rule(&rule)?;
                    let builder = current.as_mut().expect("checked style");
                    if builder.conditional_rules.len() >= MAX_RULES_PER_STYLE {
                        return Err(Error::InvalidFormat(format!(
                            "table-cell style exceeds the {MAX_RULES_PER_STYLE} conditional rule limit"
                        )));
                    }
                    builder.conditional_rules.push(rule);
                },
                Event::Start(element) | Event::Empty(element)
                    if current.is_some()
                        && depth == 0
                        && is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"table-cell-properties" =>
                {
                    let protection = optional_attribute(
                        reader.resolver(),
                        reader.decoder(),
                        &element,
                        b"cell-protect",
                    )?
                    .map(|value| Protection::parse(&value))
                    .transpose()?;
                    let builder = current.as_mut().expect("checked style");
                    if builder.protection.is_some() && protection.is_some() {
                        return Err(Error::InvalidFormat(
                            "duplicate style:cell-protect property".to_string(),
                        ));
                    }
                    if protection.is_some() {
                        builder.protection = protection;
                    }
                },
                Event::Start(_) if open_map => {
                    return Err(Error::InvalidFormat(
                        "style:map must not contain child elements".to_string(),
                    ));
                },
                Event::Text(text) if open_map => {
                    let value = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid style:map text: {error}"))
                    })?;
                    if !value.trim().is_empty() {
                        return Err(Error::InvalidFormat("style:map must be empty".to_string()));
                    }
                },
                Event::CData(text) if open_map => {
                    let value = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid style:map CDATA: {error}"))
                    })?;
                    if !value.trim().is_empty() {
                        return Err(Error::InvalidFormat("style:map must be empty".to_string()));
                    }
                },
                Event::End(element)
                    if open_map
                        && is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"map" =>
                {
                    open_map = false;
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid style:map depth".to_string())
                    })?;
                },
                Event::End(_) if open_map => {
                    return Err(Error::InvalidFormat(
                        "malformed style:map element".to_string(),
                    ));
                },
                Event::Start(_) if current.is_some() => increment_style_depth(&mut depth)?,
                Event::End(element)
                    if current.is_some()
                        && is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    if depth > 0 {
                        depth -= 1;
                    } else {
                        self.finish(current.take().expect("checked style"))?;
                    }
                },
                Event::End(_) if current.is_some() && depth > 0 => depth -= 1,
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
        if current.is_some() {
            return Err(Error::InvalidFormat(
                "unterminated table-cell style".to_string(),
            ));
        }
        Ok(())
    }

    fn record_conditional_rule(&mut self, rule: &Rule) -> Result<()> {
        if self.parsed_conditional_rules >= MAX_CONDITIONAL_RULES {
            return Err(Error::InvalidFormat(format!(
                "document exceeds the {MAX_CONDITIONAL_RULES} conditional rule limit"
            )));
        }
        self.parsed_conditional_rules += 1;
        let bytes = rule
            .condition
            .len()
            .checked_add(rule.apply_style_name.len())
            .and_then(|size| {
                size.checked_add(rule.base_cell_address.as_deref().map_or(0, str::len))
            })
            .ok_or_else(|| {
                Error::InvalidFormat("conditional style text size overflow".to_string())
            })?;
        self.conditional_text_bytes =
            self.conditional_text_bytes
                .checked_add(bytes)
                .ok_or_else(|| {
                    Error::InvalidFormat("conditional style text size overflow".to_string())
                })?;
        if self.conditional_text_bytes > MAX_CONDITIONAL_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "conditional style text exceeds the {MAX_CONDITIONAL_TEXT_BYTES} byte limit"
            )));
        }
        Ok(())
    }

    fn finish_conditional_styles(&mut self) -> Result<()> {
        let mut last_position = HashMap::new();
        for (index, name) in self.style_order.iter().enumerate() {
            last_position.insert(name.as_str(), index);
        }
        for (index, name) in self.style_order.iter().enumerate() {
            if last_position.get(name.as_str()) != Some(&index) {
                continue;
            }
            let definition = self
                .styles
                .get(name)
                .expect("style order references registry");
            if !definition.is_common
                && let Some(protection) = definition.protection
            {
                self.automatic_protection_styles.push(TableStyle {
                    style_name: name.clone(),
                    parent_style_name: definition.parent.clone(),
                    protection,
                });
            }
            if definition.conditional_rules.is_empty() {
                continue;
            }
            for rule in &definition.conditional_rules {
                if !self
                    .common_table_cell_styles
                    .contains(&rule.apply_style_name)
                {
                    return Err(Error::InvalidFormat(format!(
                        "conditional style '{}' references missing, automatic, or non-table-cell common style '{}'",
                        name, rule.apply_style_name
                    )));
                }
            }
            let conditional = ConditionalStyle {
                style_name: name.clone(),
                parent_style_name: definition.parent.clone(),
                rules: definition.conditional_rules.clone(),
            };
            self.conditional_style_index
                .insert(name.clone(), self.conditional_styles.len());
            self.conditional_styles.push(conditional);
        }
        Ok(())
    }

    fn finish(&mut self, builder: StyleBuilder) -> Result<()> {
        if builder.is_default {
            if self.default.is_some() && builder.protection.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate default table-cell protection style".to_string(),
                ));
            }
            if builder.protection.is_some() {
                self.default = builder.protection;
            }
            return Ok(());
        }
        let name = builder.name.ok_or_else(|| {
            Error::InvalidFormat("table-cell style is missing style:name".to_string())
        })?;
        if !builder.conditional_rules.is_empty() {
            if self.parsed_conditional_styles >= MAX_CONDITIONAL_STYLES {
                return Err(Error::InvalidFormat(format!(
                    "document exceeds the {MAX_CONDITIONAL_STYLES} conditional style limit"
                )));
            }
            self.parsed_conditional_styles += 1;
        }
        if builder.is_common {
            self.common_table_cell_styles.insert(name.clone());
        }
        self.style_order.push(name.clone());
        self.styles.insert(
            name,
            CellStyleDefinition {
                parent: builder.parent,
                protection: builder.protection,
                conditional_rules: builder.conditional_rules,
                is_common: builder.is_common,
            },
        );
        Ok(())
    }
}

struct StyleBuilder {
    name: Option<String>,
    parent: Option<String>,
    protection: Option<Protection>,
    conditional_rules: Vec<Rule>,
    is_default: bool,
    is_common: bool,
}

impl StyleBuilder {
    fn new(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        is_common: bool,
    ) -> Result<Option<Self>> {
        if optional_attribute(reader.resolver(), reader.decoder(), element, b"family")?.as_deref()
            != Some("table-cell")
        {
            return Ok(None);
        }
        Ok(Some(Self {
            name: optional_attribute(reader.resolver(), reader.decoder(), element, b"name")?,
            parent: optional_attribute(
                reader.resolver(),
                reader.decoder(),
                element,
                b"parent-style-name",
            )?,
            protection: None,
            conditional_rules: Vec::new(),
            is_default: element.local_name().as_ref() == b"default-style",
            is_common,
        }))
    }
}

fn increment_style_depth(depth: &mut usize) -> Result<()> {
    *depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("table-cell style depth overflow".to_string()))?;
    if *depth > MAX_STYLE_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "table-cell style exceeds the {MAX_STYLE_DEPTH} level nesting limit"
        )));
    }
    Ok(())
}

fn parse_conditional_rule(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Rule> {
    let condition = required_conditional_attribute(reader, element, b"condition")?;
    let apply_style_name = required_conditional_attribute(reader, element, b"apply-style-name")?;
    let base_cell_address = optional_attribute(
        reader.resolver(),
        reader.decoder(),
        element,
        b"base-cell-address",
    )?;
    if let Some(value) = &base_cell_address {
        check_conditional_attribute_size("style:base-cell-address", value)?;
        if value.is_empty() {
            return Err(Error::InvalidFormat(
                "style:base-cell-address must not be empty".to_string(),
            ));
        }
    }
    let formula_namespace = condition_formula_namespace(reader.resolver(), &condition)?;
    Ok(Rule {
        condition,
        formula_namespace,
        apply_style_name,
        base_cell_address,
    })
}

fn required_conditional_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<String> {
    let qualified_name = format!("style:{}", String::from_utf8_lossy(local_name));
    let value = optional_attribute(reader.resolver(), reader.decoder(), element, local_name)?
        .ok_or_else(|| Error::InvalidFormat(format!("style:map is missing {qualified_name}")))?;
    check_conditional_attribute_size(&qualified_name, &value)?;
    if value.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{qualified_name} must not be empty"
        )));
    }
    Ok(value)
}

fn check_conditional_attribute_size(name: &str, value: &str) -> Result<()> {
    if value.len() > MAX_CONDITIONAL_ATTRIBUTE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds the {MAX_CONDITIONAL_ATTRIBUTE_BYTES} byte limit"
        )));
    }
    Ok(())
}

fn condition_formula_namespace(
    resolver: &NamespaceResolver,
    condition: &str,
) -> Result<Option<formula::Namespace>> {
    let Some((prefix, _)) = condition.split_once(':') else {
        return Ok(None);
    };
    let mut characters = prefix.chars();
    if !characters
        .next()
        .is_some_and(|character| character == '_' || character.is_alphabetic())
        || !characters.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
    {
        return Ok(None);
    }
    let (namespace, _) = resolver.resolve_attribute(QName(condition.as_bytes()));
    let ResolveResult::Bound(XmlNamespace(uri)) = namespace else {
        return Err(Error::InvalidFormat(format!(
            "conditional style condition uses unbound namespace prefix '{prefix}'"
        )));
    };
    let uri = String::from_utf8(uri.to_vec()).map_err(|_| {
        Error::InvalidFormat("conditional style namespace URI is not valid UTF-8".to_string())
    })?;
    Ok(Some(formula::Namespace {
        prefix: prefix.to_string(),
        uri,
    }))
}

fn optional_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    let mut found = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace, STYLE_NAMESPACE) && local.as_ref() == local_name {
            if found.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate style:{} attribute",
                    String::from_utf8_lossy(local_name)
                )));
            }
            found = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map(|value| value.into_owned())
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML attribute: {error}"))
                    })?,
            );
        }
    }
    Ok(found)
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(XmlNamespace(value)) if *value == expected)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn conditional_editor_preserves_unrelated_automatic_styles_with_arbitrary_prefixes() {
        let xml = r#"<o:automatic-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:x="urn:example:formula"><draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep&amp;exact"/><s:style s:name="old" s:family="table-cell"><s:map s:condition="x:old()" s:apply-style-name="Red"/></s:style><s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style></o:automatic-styles>"#;
        let fragment = PreservedXmlFragment {
            xml: xml.to_string(),
            namespaces: BTreeMap::new(),
        };
        let style = ConditionalStyle::new(
            "new&style",
            vec![
                Rule::new("x:test()<2", "Red").with_formula_namespace(formula::Namespace {
                    prefix: "x".to_string(),
                    uri: "urn:example:formula".to_string(),
                }),
            ],
        );
        let common = HashSet::from(["Red".to_string()]);
        validate_conditional_style_collection(std::slice::from_ref(&style), &common).unwrap();
        let rewritten = rewrite_conditional_styles(Some(&fragment), &[style]).unwrap();
        assert!(rewritten.xml.contains(
            r#"<draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep&amp;exact"/>"#
        ));
        assert!(rewritten.xml.contains(
            r#"<s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style>"#
        ));
        assert!(!rewritten.xml.contains("s:name=\"old\""));
        assert!(rewritten.xml.contains("style:name=\"new&amp;style\""));
    }

    #[test]
    fn managed_protection_editor_preserves_unrelated_xml_and_merges_maps() {
        let xml = r#"<o:automatic-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep"/><s:style s:name="combo" s:family="table-cell"><s:table-cell-properties s:cell-protect="protected"/><s:map s:condition="cell-content()>0" s:apply-style-name="Red"/></s:style><s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style></o:automatic-styles>"#;
        let fragment = PreservedXmlFragment {
            xml: xml.to_string(),
            namespaces: BTreeMap::new(),
        };
        let conditional =
            ConditionalStyle::new("combo", vec![Rule::new("cell-content()>0", "Red")]);
        let protection = TableStyle::new("combo", Protection::HiddenAndProtected);
        let rewritten =
            rewrite_managed_cell_styles(Some(&fragment), &[conditional], &[protection]).unwrap();
        assert!(rewritten.xml.contains(
            r#"<draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="keep"/>"#
        ));
        assert!(rewritten.xml.contains(
            r#"<s:style s:name="plain" s:family="table-cell"><s:table-cell-properties/></s:style>"#
        ));
        assert_eq!(rewritten.xml.matches("name=\"combo\"").count(), 1);
        assert!(
            rewritten
                .xml
                .contains("cell-protect=\"hidden-and-protected\"")
        );
        assert!(rewritten.xml.contains("condition=\"cell-content()&gt;0\""));
    }

    #[test]
    fn resolves_defaults_named_parents_and_automatic_overrides() {
        let named = r#"<o:styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:default-style s:family="table-cell"><s:table-cell-properties s:cell-protect="none"/></s:default-style><s:style s:name="Locked" s:family="table-cell"><s:table-cell-properties s:cell-protect="protected"/></s:style><s:style s:name="Child" s:family="table-cell" s:parent-style-name="Locked"/></o:styles>"#;
        let content = r#"<o:automatic-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:style s:name="Auto" s:family="table-cell" s:parent-style-name="Child"><s:table-cell-properties s:cell-protect="formula-hidden protected"/></s:style></o:automatic-styles>"#;
        let registry = CellStyleRegistry::parse(Some(named), content).unwrap();
        assert_eq!(registry.resolve(None).unwrap(), Some(Protection::None));
        assert_eq!(
            registry.resolve(Some("Child")).unwrap(),
            Some(Protection::Protected)
        );
        assert_eq!(
            registry.resolve(Some("Auto")).unwrap(),
            Some(Protection::ProtectedFormulaHidden)
        );
    }

    #[test]
    fn rejects_invalid_values_missing_parents_and_cycles() {
        let invalid = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="Bad" s:family="table-cell"><s:table-cell-properties s:cell-protect="locked"/></s:style>"#;
        assert!(CellStyleRegistry::parse(None, invalid).is_err());

        let styles = r#"<o:styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:style s:name="Missing" s:family="table-cell" s:parent-style-name="Nope"/><s:style s:name="A" s:family="table-cell" s:parent-style-name="B"/><s:style s:name="B" s:family="table-cell" s:parent-style-name="A"/></o:styles>"#;
        let registry = CellStyleRegistry::parse(None, styles).unwrap();
        assert!(registry.resolve(Some("Missing")).is_err());
        assert!(registry.resolve(Some("A")).is_err());
    }

    #[test]
    fn extracts_automatic_styles_and_required_namespace_bindings() {
        let xml = r##"<?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><o:font-face-decls/><o:automatic-styles><s:style s:name="ce1" s:family="table-cell"><s:table-cell-properties f:background-color="#fff"/></s:style></o:automatic-styles><o:body/></o:document-content>"##;
        let fragment = extract_automatic_styles(xml).unwrap().unwrap();
        assert!(fragment.xml.starts_with("<o:automatic-styles>"));
        assert!(fragment.xml.ends_with("</o:automatic-styles>"));
        let mut declarations = String::new();
        fragment.write_missing_namespaces(&mut declarations, ["o"]);
        assert!(!declarations.contains("xmlns:o="));
        assert!(declarations.contains("xmlns:s="));
        assert!(declarations.contains("xmlns:f="));
        let font_faces = extract_font_face_decls(xml).unwrap().unwrap();
        assert!(font_faces.xml.starts_with("<o:font-face-decls"));
    }

    #[test]
    fn parses_ordered_inert_conditional_cell_styles_and_overrides() {
        let named = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:example:formula"><o:styles><s:style s:name="Red" s:family="table-cell"/><s:style s:name="Blue" s:family="table-cell"/></o:styles></o:document-styles>"#;
        let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:example:formula"><o:automatic-styles><s:style s:name="ce1" s:family="table-cell" s:parent-style-name="Default"><s:map s:condition="cell-content()&lt;1" s:apply-style-name="Red" s:base-cell-address="Sheet1.A1"/><s:map s:condition="f:is-true-formula([.A1]&gt;0)" s:apply-style-name="Blue"></s:map></s:style></o:automatic-styles></o:document-content>"#;
        let registry = CellStyleRegistry::parse(Some(named), content).unwrap();
        let style = registry.conditional_style("ce1").unwrap();
        assert_eq!(style.parent_style_name.as_deref(), Some("Default"));
        assert_eq!(style.rules.len(), 2);
        assert_eq!(style.rules[0].condition, "cell-content()<1");
        assert_eq!(style.rules[0].apply_style_name, "Red");
        assert_eq!(
            style.rules[0].base_cell_address.as_deref(),
            Some("Sheet1.A1")
        );
        assert_eq!(
            style.rules[1].formula_namespace,
            Some(formula::Namespace {
                prefix: "f".to_string(),
                uri: "urn:example:formula".to_string(),
            })
        );
        assert_eq!(registry.conditional_styles(), std::slice::from_ref(style));

        let override_content = format!(
            "{content}<o:automatic-styles xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\"><s:style s:name=\"ce1\" s:family=\"table-cell\"><s:map s:condition=\"cell-content()=2\" s:apply-style-name=\"Blue\"/></s:style></o:automatic-styles>"
        );
        let overridden = CellStyleRegistry::parse(Some(named), &override_content).unwrap();
        assert_eq!(
            overridden.conditional_style("ce1").unwrap().rules[0].condition,
            "cell-content()=2"
        );
    }

    #[test]
    fn rejects_malformed_or_active_conditional_style_inputs() {
        let common = r#"<o:styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:style s:name="Target" s:family="table-cell"/></o:styles>"#;
        let missing_target = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="ce1" s:family="table-cell"><s:map s:condition="cell-content()=1" s:apply-style-name="Missing"/></s:style>"#;
        assert!(CellStyleRegistry::parse(Some(common), missing_target).is_err());

        let non_empty_map = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="ce1" s:family="table-cell"><s:map s:condition="cell-content()=1" s:apply-style-name="Target">run</s:map></s:style>"#;
        assert!(CellStyleRegistry::parse(Some(common), non_empty_map).is_err());

        let dtd = r#"<!DOCTYPE x [<!ENTITY run SYSTEM "file:///etc/passwd">]><s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" s:name="ce1" s:family="table-cell"/>"#;
        assert!(CellStyleRegistry::parse(Some(common), dtd).is_err());

        let oversized = "x".repeat(MAX_CONDITIONAL_ATTRIBUTE_BYTES + 1);
        let oversized = format!(
            "<s:style xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" s:name=\"ce1\" s:family=\"table-cell\"><s:map s:condition=\"{oversized}\" s:apply-style-name=\"Target\"/></s:style>"
        );
        assert!(CellStyleRegistry::parse(Some(common), &oversized).is_err());

        let extension_only = r#"<s:style xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:c="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" s:name="ce1" s:family="table-cell"><c:conditional-formats><c:condition c:value="1"/></c:conditional-formats></s:style>"#;
        let registry = CellStyleRegistry::parse(Some(common), extension_only).unwrap();
        assert!(registry.conditional_styles().is_empty());
    }

    #[test]
    fn parses_libreoffice_flat_conditional_style_fixture_when_available() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../test-data/libreoffice-core/sc/qa/unit/data/functions/financial/fods/couppcd.fods",
        );
        if !path.exists() {
            return;
        }
        let xml = std::fs::read_to_string(path).unwrap();
        let registry = CellStyleRegistry::parse(None, &xml).unwrap();
        let ce6 = registry.conditional_style("ce6").unwrap();
        assert_eq!(ce6.rules.len(), 3);
        assert_eq!(ce6.rules[0].condition, "cell-content()=\"\"");
        assert_eq!(ce6.rules[1].apply_style_name, "Untitled1");
        assert_eq!(ce6.rules[2].base_cell_address.as_deref(), Some("Sheet1.B3"));
    }
}
