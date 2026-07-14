//! Resolution of ODF table-cell protection styles.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, NamespaceResolver, ResolveResult},
    reader::NsReader,
};
use std::collections::{BTreeMap, HashMap, HashSet};

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/// The effective value of the ODF `style:cell-protect` property.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellStyleProtection {
    None,
    Protected,
    FormulaHidden,
    ProtectedFormulaHidden,
    HiddenAndProtected,
}

impl CellStyleProtection {
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
}

#[derive(Clone, Debug, Default)]
pub(crate) struct CellStyleRegistry {
    styles: HashMap<String, CellStyleDefinition>,
    default: Option<CellStyleProtection>,
}

#[derive(Clone, Debug)]
pub(crate) struct AutomaticStylesFragment {
    pub xml: String,
    namespaces: BTreeMap<String, String>,
}

impl AutomaticStylesFragment {
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

pub(crate) fn extract_automatic_styles(xml: &str) -> Result<Option<AutomaticStylesFragment>> {
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
                if is_office_namespace && element.local_name().as_ref() == b"automatic-styles" =>
            {
                if range_start.is_some() {
                    return Err(Error::InvalidFormat(
                        "nested office:automatic-styles element".to_string(),
                    ));
                }
                range_start = Some(event_start);
                depth = 1;
            },
            Event::Empty(element)
                if is_office_namespace && element.local_name().as_ref() == b"automatic-styles" =>
            {
                if range_start.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate office:automatic-styles element".to_string(),
                    ));
                }
                return Ok(Some(AutomaticStylesFragment {
                    xml: xml[event_start..event_end].to_string(),
                    namespaces,
                }));
            },
            Event::Start(_) if range_start.is_some() => depth += 1,
            Event::End(element) if range_start.is_some() => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid automatic-styles depth".to_string())
                })?;
                if depth == 0 {
                    if !is_office_namespace || element.local_name().as_ref() != b"automatic-styles"
                    {
                        return Err(Error::InvalidFormat(
                            "malformed office:automatic-styles element".to_string(),
                        ));
                    }
                    let start = range_start.take().expect("checked range");
                    return Ok(Some(AutomaticStylesFragment {
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
        return Err(Error::InvalidFormat(
            "unterminated office:automatic-styles element".to_string(),
        ));
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
    protection: Option<CellStyleProtection>,
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
        Ok(registry)
    }

    pub(crate) fn resolve(&self, style_name: Option<&str>) -> Result<Option<CellStyleProtection>> {
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
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(element)
                    if is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    if current.is_some() {
                        depth += 1;
                    } else if let Some(builder) = StyleBuilder::new(&reader, &element)? {
                        current = Some(builder);
                    }
                },
                Event::Empty(element)
                    if is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    if current.is_none()
                        && let Some(builder) = StyleBuilder::new(&reader, &element)?
                    {
                        self.finish(builder)?;
                    }
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
                    .map(|value| CellStyleProtection::parse(&value))
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
                Event::Start(_) if current.is_some() => depth += 1,
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
        self.styles.insert(
            name,
            CellStyleDefinition {
                parent: builder.parent,
                protection: builder.protection,
            },
        );
        Ok(())
    }
}

struct StyleBuilder {
    name: Option<String>,
    parent: Option<String>,
    protection: Option<CellStyleProtection>,
    is_default: bool,
}

impl StyleBuilder {
    fn new(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<Self>> {
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
            is_default: element.local_name().as_ref() == b"default-style",
        }))
    }
}

fn optional_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace, STYLE_NAMESPACE) && local.as_ref() == local_name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")));
        }
    }
    Ok(None)
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_defaults_named_parents_and_automatic_overrides() {
        let named = r#"<o:styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:default-style s:family="table-cell"><s:table-cell-properties s:cell-protect="none"/></s:default-style><s:style s:name="Locked" s:family="table-cell"><s:table-cell-properties s:cell-protect="protected"/></s:style><s:style s:name="Child" s:family="table-cell" s:parent-style-name="Locked"/></o:styles>"#;
        let content = r#"<o:automatic-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><s:style s:name="Auto" s:family="table-cell" s:parent-style-name="Child"><s:table-cell-properties s:cell-protect="formula-hidden protected"/></s:style></o:automatic-styles>"#;
        let registry = CellStyleRegistry::parse(Some(named), content).unwrap();
        assert_eq!(
            registry.resolve(None).unwrap(),
            Some(CellStyleProtection::None)
        );
        assert_eq!(
            registry.resolve(Some("Child")).unwrap(),
            Some(CellStyleProtection::Protected)
        );
        assert_eq!(
            registry.resolve(Some("Auto")).unwrap(),
            Some(CellStyleProtection::ProtectedFormulaHidden)
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
    }
}
