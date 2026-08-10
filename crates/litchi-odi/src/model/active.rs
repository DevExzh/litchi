//! Inert active-content inventory for ODF script and external-data constructs.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SCRIPT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_DEPTH: usize = 256;
pub(crate) const MAX_ITEMS: usize = 100_000;

/// Semantic kind of inert active content discovered in an ODI artifact.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ActiveContentKind {
    EmbeddedScript,
    EventListener,
    MacroExecution,
    DdeConnection,
    PackageScriptMember,
}

/// Artifact location containing one active-content item.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ActiveContentLocation {
    FlatXml,
    ContentXml,
    StylesXml,
    PackageMember,
}

/// One inert active-content reference. No target is resolved or executed.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ActiveContent {
    kind: ActiveContentKind,
    location: ActiveContentLocation,
    source: String,
    language: Option<String>,
    event_name: Option<String>,
    target: Option<String>,
    dde_application: Option<String>,
    dde_topic: Option<String>,
    dde_item: Option<String>,
}

impl ActiveContent {
    /// Returns the semantic active-content kind.
    #[must_use]
    pub const fn kind(&self) -> ActiveContentKind {
        self.kind
    }

    /// Returns the artifact location where the item was found.
    #[must_use]
    pub const fn location(&self) -> ActiveContentLocation {
        self.location
    }

    /// Returns the source element `QName` or package-member path.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Returns the declared script language, if present.
    #[must_use]
    pub fn language(&self) -> Option<&str> {
        self.language.as_deref()
    }

    /// Returns the declared event name, if present.
    #[must_use]
    pub fn event_name(&self) -> Option<&str> {
        self.event_name.as_deref()
    }

    /// Returns an inert macro, link, action, or DDE target, if present.
    #[must_use]
    pub fn target(&self) -> Option<&str> {
        self.target.as_deref()
    }

    /// Returns the inert DDE application coordinate, if present.
    #[must_use]
    pub fn dde_application(&self) -> Option<&str> {
        self.dde_application.as_deref()
    }

    /// Returns the inert DDE topic coordinate, if present.
    #[must_use]
    pub fn dde_topic(&self) -> Option<&str> {
        self.dde_topic.as_deref()
    }

    /// Returns the inert DDE item coordinate, if present.
    #[must_use]
    pub fn dde_item(&self) -> Option<&str> {
        self.dde_item.as_deref()
    }

    pub(crate) fn package_member(path: String) -> Self {
        Self {
            kind: ActiveContentKind::PackageScriptMember,
            location: ActiveContentLocation::PackageMember,
            source: path,
            language: None,
            event_name: None,
            target: None,
            dde_application: None,
            dde_topic: None,
            dde_item: None,
        }
    }
}

pub(crate) fn scan_xml(xml: &str, location: ActiveContentLocation) -> Result<Vec<ActiveContent>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut items = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid(format!("invalid ODI active-content XML: {error}")))?;
            (element_namespace(&namespace), event)
        };
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .filter(|value| *value <= MAX_DEPTH)
                    .ok_or_else(|| invalid("ODI active-content XML exceeds the depth limit"))?;
                if let Some(item) = classify_element(&reader, namespace, &element, location)? {
                    push_item(&mut items, item)?;
                }
            },
            Event::Empty(element) => {
                if let Some(item) = classify_element(&reader, namespace, &element, location)? {
                    push_item(&mut items, item)?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODI active-content XML depth underflow"))?;
            },
            Event::DocType(_) => {
                return Err(invalid("DOCTYPE is not allowed in ODI active-content XML"));
            },
            Event::Eof => {
                if depth != 0 {
                    return Err(invalid("unterminated ODI active-content XML"));
                }
                return Ok(items);
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

pub(crate) fn is_package_script_member(path: &str) -> bool {
    matches!(
        path.split('/').next(),
        Some("Basic" | "Dialogs" | "Scripts")
    )
}

fn classify_element(
    reader: &NsReader<&[u8]>,
    namespace: ElementNamespace,
    element: &BytesStart<'_>,
    location: ActiveContentLocation,
) -> Result<Option<ActiveContent>> {
    let local = element.local_name();
    let kind = if matches!(namespace, ElementNamespace::Office | ElementNamespace::Text)
        && local.as_ref() == b"script"
    {
        ActiveContentKind::EmbeddedScript
    } else if matches!(
        namespace,
        ElementNamespace::Script | ElementNamespace::Presentation
    ) && local.as_ref() == b"event-listener"
    {
        ActiveContentKind::EventListener
    } else if namespace == ElementNamespace::Text && local.as_ref() == b"execute-macro" {
        ActiveContentKind::MacroExecution
    } else if (namespace == ElementNamespace::Text
        && matches!(local.as_ref(), b"dde-connection" | b"dde-connection-decl"))
        || (namespace == ElementNamespace::Table && local.as_ref() == b"dde-link")
        || (namespace == ElementNamespace::Office && local.as_ref() == b"dde-source")
    {
        ActiveContentKind::DdeConnection
    } else {
        return Ok(None);
    };
    let source = format!(
        "{}:{}",
        namespace.prefix(),
        String::from_utf8_lossy(local.as_ref())
    );
    let language = attribute(reader, element, SCRIPT, b"language")?;
    let event_name = attribute(reader, element, SCRIPT, b"event-name")?;
    let target = first_present([
        attribute(reader, element, SCRIPT, b"macro-name")?,
        attribute(reader, element, XLINK, b"href")?,
        attribute(reader, element, PRESENTATION, b"action")?,
        attribute(reader, element, TEXT, b"name")?,
    ]);
    let dde_application = attribute(reader, element, OFFICE, b"dde-application")?;
    let dde_topic = attribute(reader, element, OFFICE, b"dde-topic")?;
    let dde_item = attribute(reader, element, OFFICE, b"dde-item")?;
    Ok(Some(ActiveContent {
        kind,
        location,
        source,
        language,
        event_name,
        target,
        dde_application,
        dde_topic,
        dde_item,
    }))
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut result = None;
    for raw in element.attributes() {
        let value =
            raw.map_err(|error| invalid(format!("invalid ODI active-content attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(value.key);
        if bound_to(&resolved, namespace) && name.as_ref() == local {
            if result.is_some() {
                return Err(invalid("duplicate expanded ODI active-content attribute"));
            }
            result = Some(
                value
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| invalid(format!("invalid ODI active-content value: {error}")))?
                    .into_owned(),
            );
        }
    }
    Ok(result)
}

fn first_present<const N: usize>(values: [Option<String>; N]) -> Option<String> {
    values.into_iter().flatten().next()
}

fn push_item(items: &mut Vec<ActiveContent>, item: ActiveContent) -> Result<()> {
    if items.len() >= MAX_ITEMS {
        return Err(invalid(
            "ODI active-content inventory exceeds the item limit",
        ));
    }
    items.push(item);
    Ok(())
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ElementNamespace {
    Office,
    Presentation,
    Script,
    Table,
    Text,
    Other,
}

impl ElementNamespace {
    const fn prefix(self) -> &'static str {
        match self {
            Self::Office => "office",
            Self::Presentation => "presentation",
            Self::Script => "script",
            Self::Table => "table",
            Self::Text => "text",
            Self::Other => "unknown",
        }
    }
}

fn element_namespace(namespace: &ResolveResult<'_>) -> ElementNamespace {
    if bound_to(namespace, OFFICE) {
        ElementNamespace::Office
    } else if bound_to(namespace, PRESENTATION) {
        ElementNamespace::Presentation
    } else if bound_to(namespace, SCRIPT) {
        ElementNamespace::Script
    } else if bound_to(namespace, TABLE) {
        ElementNamespace::Table
    } else if bound_to(namespace, TEXT) {
        ElementNamespace::Text
    } else {
        ElementNamespace::Other
    }
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

    use super::*;

    #[test]
    fn inventory_is_namespace_aware_and_retains_inert_coordinates() {
        let xml = concat!(
            r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">"#,
            r#"<o:script s:language="python"/>"#,
            r#"<s:event-listener s:event-name="dom:click" s:macro-name="Macro.Main"/>"#,
            r#"<t:execute-macro t:name="Run"/>"#,
            r#"<o:dde-source o:dde-application="Calc" o:dde-topic="Sheet1" o:dde-item="A1"/>"#,
            r#"</o:document>"#,
        );
        let items = scan_xml(xml, ActiveContentLocation::FlatXml).unwrap();
        assert_eq!(items.len(), 4);
        assert_eq!(items[0].kind(), ActiveContentKind::EmbeddedScript);
        assert_eq!(items[0].source(), "office:script");
        assert_eq!(items[0].language(), Some("python"));
        assert_eq!(items[1].event_name(), Some("dom:click"));
        assert_eq!(items[1].target(), Some("Macro.Main"));
        assert_eq!(items[2].kind(), ActiveContentKind::MacroExecution);
        assert_eq!(items[2].target(), Some("Run"));
        assert_eq!(items[3].dde_application(), Some("Calc"));
        assert_eq!(items[3].dde_topic(), Some("Sheet1"));
        assert_eq!(items[3].dde_item(), Some("A1"));
    }

    #[test]
    fn package_script_container_detection_is_segment_exact() {
        assert!(is_package_script_member("Scripts/python/main.py"));
        assert!(is_package_script_member("Basic/Standard/Module1.xml"));
        assert!(!is_package_script_member("Pictures/Scripts.png"));
        assert!(!is_package_script_member("Scripts-backup/main.py"));
    }
}
