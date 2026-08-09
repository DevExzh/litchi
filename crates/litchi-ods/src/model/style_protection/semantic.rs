//! Semantic resolution and XML parsing for table-cell protection styles.

use super::{
    OFFICE_NAMESPACE, STYLE_NAMESPACE,
    model::{
        CellStyleDefinition, CellStyleRegistry, ConditionalStyle, MAX_CONDITIONAL_RULES,
        MAX_CONDITIONAL_STYLES, MAX_CONDITIONAL_TEXT_BYTES, MAX_RULES_PER_STYLE, MAX_STYLE_DEPTH,
        Protection, Rule, StyleBuilder, TableStyle,
    },
    validation::check_conditional_attribute_size,
};
use crate::model::names::formula;
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace as XmlNamespace, NamespaceResolver, QName, ResolveResult},
    reader::NsReader,
};
use std::collections::{HashMap, HashSet};

/// # Errors
///
/// Returns an error when the styles XML is malformed.
pub fn common_table_cell_style_names(styles_xml: Option<&str>) -> Result<HashSet<String>> {
    let registry = CellStyleRegistry::parse(styles_xml, "")?;
    Ok(registry.common_table_cell_styles)
}

/// # Errors
///
/// Returns an error when a value violates the format or resource constraints.
pub fn validate_protection_style_document(
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

    #[must_use]
    pub fn conditional_style(&self, name: &str) -> Option<&ConditionalStyle> {
        self.conditional_style_index
            .get(name)
            .and_then(|index| self.conditional_styles.get(*index))
    }

    pub(crate) fn automatic_protection_styles(&self) -> &[TableStyle] {
        &self.automatic_protection_styles
    }

    /// # Errors
    ///
    /// Returns an error when the style inheritance chain is cyclic.
    pub fn resolve(&self, style_name: Option<&str>) -> Result<Option<Protection>> {
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
                Event::Start(_)
                | Event::End(_)
                | Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => {},
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
    let uri = String::from_utf8(uri.to_vec()).map_err(|_error| {
        Error::InvalidFormat("conditional style namespace URI is not valid UTF-8".to_string())
    })?;
    Ok(Some(formula::Namespace {
        prefix: prefix.to_string(),
        uri,
    }))
}

pub(super) fn optional_attribute(
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
                    .map(std::borrow::Cow::into_owned)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML attribute: {error}"))
                    })?,
            );
        }
    }
    Ok(found)
}

pub(super) fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(XmlNamespace(value)) if *value == expected)
}
