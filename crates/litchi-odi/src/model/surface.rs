//! Bounded inert inventory of form and producer-extension XML surfaces.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const FORM: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const XFORMS: &[u8] = b"http://www.w3.org/2002/xforms";
const MAX_DEPTH: usize = 256;
pub(crate) const MAX_ITEMS: usize = 100_000;

const KNOWN_NAMESPACES: &[&[u8]] = &[
    b"http://purl.org/dc/elements/1.1/",
    b"http://www.w3.org/1998/Math/MathML",
    b"http://www.w3.org/1999/02/22-rdf-syntax-ns#",
    b"http://www.w3.org/1999/xhtml",
    b"http://www.w3.org/1999/xlink",
    b"http://www.w3.org/2000/xmlns/",
    b"http://www.w3.org/2001/XMLSchema",
    b"http://www.w3.org/2001/XMLSchema-instance",
    b"http://www.w3.org/2001/xml-events",
    b"http://www.w3.org/2002/xforms",
    b"http://www.w3.org/XML/1998/namespace",
    b"urn:oasis:names:tc:opendocument:xmlns:animation:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:database:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:of:1.2",
    b"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
];

/// Kind of unexecuted form or producer-extension surface.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SurfaceKind {
    FormContainer,
    FormElement,
    XFormElement,
    ExtensionElement,
    ExtensionAttribute,
}

/// XML part containing an inert surface.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SurfaceLocation {
    FlatXml,
    ContentXml,
    StylesXml,
}

/// Supported lifecycle for an inventoried unmodeled surface.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SurfaceDisposition {
    /// Preserve the exact retained XML inertly during non-overlapping edits.
    PreserveExact,
    /// Direct semantic mutation is not supported and must be refused.
    RefuseMutation,
}

/// One inert form or producer-extension XML occurrence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Surface {
    kind: SurfaceKind,
    location: SurfaceLocation,
    source: String,
    namespace: String,
    name: Option<String>,
    control_id: Option<String>,
    target: Option<String>,
}

impl Surface {
    /// Returns the semantic surface kind.
    #[must_use]
    pub const fn kind(&self) -> SurfaceKind {
        self.kind
    }

    /// Returns the XML part containing this occurrence.
    #[must_use]
    pub const fn location(&self) -> SurfaceLocation {
        self.location
    }

    /// Returns the exact lexical element or attribute `QName`.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Returns the resolved namespace URI.
    #[must_use]
    pub fn namespace(&self) -> &str {
        &self.namespace
    }

    /// Returns the inert form name, if declared.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the inert form control identifier, if declared.
    #[must_use]
    pub fn control_id(&self) -> Option<&str> {
        self.control_id.as_deref()
    }

    /// Returns the inert form target location or link, if declared.
    #[must_use]
    pub fn target(&self) -> Option<&str> {
        self.target.as_deref()
    }

    /// Returns the exact-preservation disposition for read-only inventory.
    #[must_use]
    pub const fn preservation_disposition(&self) -> SurfaceDisposition {
        SurfaceDisposition::PreserveExact
    }

    /// Returns the unsupported disposition for direct semantic mutation.
    #[must_use]
    pub const fn mutation_disposition(&self) -> SurfaceDisposition {
        SurfaceDisposition::RefuseMutation
    }
}

#[derive(Default)]
pub(crate) struct Inventory {
    pub(crate) forms: Vec<Surface>,
    pub(crate) extensions: Vec<Surface>,
}

pub(crate) fn scan_xml(xml: &str, location: SurfaceLocation) -> Result<Inventory> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut inventory = Inventory::default();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid(format!("invalid ODI surface XML: {error}")))?;
            (namespace_uri(&namespace).map(<[u8]>::to_vec), event)
        };
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .filter(|value| *value <= MAX_DEPTH)
                    .ok_or_else(|| invalid("ODI surface XML exceeds the depth limit"))?;
                inspect_element(
                    &reader,
                    namespace.as_deref(),
                    &element,
                    location,
                    &mut inventory,
                )?;
            },
            Event::Empty(element) => {
                inspect_element(
                    &reader,
                    namespace.as_deref(),
                    &element,
                    location,
                    &mut inventory,
                )?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODI surface XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI surface XML")),
            Event::Eof => {
                if depth != 0 {
                    return Err(invalid("unterminated ODI surface XML"));
                }
                return Ok(inventory);
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

fn inspect_element(
    reader: &NsReader<&[u8]>,
    namespace: Option<&[u8]>,
    element: &BytesStart<'_>,
    location: SurfaceLocation,
    inventory: &mut Inventory,
) -> Result<()> {
    if namespace == Some(OFFICE) && element.local_name().as_ref() == b"forms" {
        push(
            &mut inventory.forms,
            form_surface(
                reader,
                SurfaceKind::FormContainer,
                location,
                namespace,
                element,
            )?,
        )?;
    } else if namespace == Some(FORM) {
        push(
            &mut inventory.forms,
            form_surface(
                reader,
                SurfaceKind::FormElement,
                location,
                namespace,
                element,
            )?,
        )?;
    } else if namespace == Some(XFORMS) {
        push(
            &mut inventory.forms,
            form_surface(
                reader,
                SurfaceKind::XFormElement,
                location,
                namespace,
                element,
            )?,
        )?;
    } else if namespace.is_some_and(is_extension_namespace) {
        push(
            &mut inventory.extensions,
            surface(
                SurfaceKind::ExtensionElement,
                location,
                namespace,
                element.name().as_ref(),
            )?,
        )?;
    }
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI surface attribute: {error}")))?;
        let (attribute_namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_uri(&attribute_namespace).is_some_and(is_extension_namespace) {
            push(
                &mut inventory.extensions,
                surface(
                    SurfaceKind::ExtensionAttribute,
                    location,
                    namespace_uri(&attribute_namespace),
                    attribute.key.as_ref(),
                )?,
            )?;
        }
    }
    Ok(())
}

fn surface(
    kind: SurfaceKind,
    location: SurfaceLocation,
    namespace: Option<&[u8]>,
    qname: &[u8],
) -> Result<Surface> {
    let Some(uri) = namespace else {
        return Err(invalid("ODI inventoried surface has no bound namespace"));
    };
    Ok(Surface {
        kind,
        location,
        source: std::str::from_utf8(qname)
            .map_err(|error| invalid(format!("ODI surface QName is not UTF-8: {error}")))?
            .to_owned(),
        namespace: std::str::from_utf8(uri)
            .map_err(|error| invalid(format!("ODI surface namespace is not UTF-8: {error}")))?
            .to_owned(),
        name: None,
        control_id: None,
        target: None,
    })
}

fn form_surface(
    reader: &NsReader<&[u8]>,
    kind: SurfaceKind,
    location: SurfaceLocation,
    namespace: Option<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Surface> {
    let mut result = surface(kind, location, namespace, element.name().as_ref())?;
    result.name = attribute(reader, element, FORM, b"name")?;
    result.control_id = attribute(reader, element, FORM, b"id")?;
    result.target = first_present([
        attribute(reader, element, FORM, b"target-location")?,
        attribute(reader, element, XLINK, b"href")?,
    ]);
    Ok(result)
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut result = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI surface attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_uri(&resolved) == Some(namespace) && name.as_ref() == local {
            if result.is_some() {
                return Err(invalid("duplicate expanded ODI surface attribute"));
            }
            result = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| invalid(format!("invalid ODI surface value: {error}")))?
                    .into_owned(),
            );
        }
    }
    Ok(result)
}

fn first_present<const N: usize>(values: [Option<String>; N]) -> Option<String> {
    values.into_iter().flatten().next()
}

fn push(items: &mut Vec<Surface>, item: Surface) -> Result<()> {
    if items.len() >= MAX_ITEMS {
        return Err(invalid("ODI surface inventory exceeds the item limit"));
    }
    items.push(item);
    Ok(())
}

fn is_extension_namespace(namespace: &[u8]) -> bool {
    !KNOWN_NAMESPACES.contains(&namespace)
}

fn namespace_uri<'namespace>(namespace: &ResolveResult<'namespace>) -> Option<&'namespace [u8]> {
    if let ResolveResult::Bound(Namespace(uri)) = namespace {
        Some(*uri)
    } else {
        None
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn form_and_unknown_namespace_occurrences_remain_separate_and_inert() {
        let xml = concat!(
            r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:v="urn:vendor:test">"#,
            r##"<o:forms><f:form f:name="Search"><f:text f:id="query" x:href="#target"/></f:form></o:forms>"##,
            r#"<v:item v:flag="1"/>"#,
            r#"</o:document>"#,
        );
        let inventory = match scan_xml(xml, SurfaceLocation::FlatXml) {
            Ok(inventory) => inventory,
            Err(error) => panic!("surface inventory should parse: {error}"),
        };
        assert_eq!(inventory.forms.len(), 3);
        assert_eq!(inventory.extensions.len(), 2);
        assert_eq!(inventory.forms[0].kind(), SurfaceKind::FormContainer);
        assert_eq!(inventory.forms[1].name(), Some("Search"));
        assert_eq!(inventory.forms[2].control_id(), Some("query"));
        assert_eq!(inventory.forms[2].target(), Some("#target"));
        assert_eq!(
            inventory.extensions[0].kind(),
            SurfaceKind::ExtensionElement
        );
        assert_eq!(
            inventory.extensions[1].mutation_disposition(),
            SurfaceDisposition::RefuseMutation
        );
        assert_eq!(inventory.extensions[1].namespace(), "urn:vendor:test");
    }
}
