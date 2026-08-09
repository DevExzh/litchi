//! Strict non-executing chart styles-part validation.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::BTreeSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Style,
    Other,
}

struct State {
    depth: usize,
    root_seen: bool,
    root_closed: bool,
    style_scope: Option<u8>,
    office_containers: BTreeSet<u8>,
    style_names: BTreeSet<(u8, String, String)>,
}

impl State {
    fn new() -> Self {
        Self {
            depth: 0,
            root_seen: false,
            root_closed: false,
            style_scope: None,
            office_containers: BTreeSet::new(),
            style_names: BTreeSet::new(),
        }
    }

    fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        limits: crate::Limits,
    ) -> Result<()> {
        self.depth = checked_depth(self.depth, limits.max_depth())?;
        if self.depth == 1 {
            self.validate_root(namespace, element)?;
            self.root_seen = true;
        }
        reject_scripts(namespace, element)?;
        if self.depth == 2 && namespace == NamespaceKind::Office {
            self.style_scope = style_scope_id(element.local_name().as_ref());
            if let Some(scope) = self.style_scope {
                self.insert_container(scope)?;
            }
        }
        validate_style_element(
            reader,
            namespace,
            element,
            self.style_scope,
            &mut self.style_names,
            limits,
        )
    }

    fn empty(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        limits: crate::Limits,
    ) -> Result<()> {
        if self.depth == 0 {
            self.validate_root(namespace, element)?;
            self.root_seen = true;
            self.root_closed = true;
        }
        reject_scripts(namespace, element)?;
        let empty_scope = if self.depth == 1 && namespace == NamespaceKind::Office {
            style_scope_id(element.local_name().as_ref())
        } else {
            None
        };
        if let Some(scope) = empty_scope {
            self.insert_container(scope)?;
        }
        validate_style_element(
            reader,
            namespace,
            element,
            self.style_scope.or(empty_scope),
            &mut self.style_names,
            limits,
        )
    }

    fn end(&mut self) -> Result<()> {
        if self.depth == 0 {
            return Err(invalid("ODC styles XML depth underflow"));
        }
        if self.depth == 1 {
            self.root_closed = true;
        }
        if self.depth == 2 {
            self.style_scope = None;
        }
        self.depth -= 1;
        Ok(())
    }

    fn validate_root(&self, namespace: NamespaceKind, element: &BytesStart<'_>) -> Result<()> {
        if self.root_seen
            || namespace != NamespaceKind::Office
            || element.local_name().as_ref() != b"document-styles"
        {
            return Err(invalid(
                "ODC styles require one office:document-styles root",
            ));
        }
        Ok(())
    }

    fn insert_container(&mut self, scope: u8) -> Result<()> {
        if !self.office_containers.insert(scope) {
            return Err(invalid("ODC styles contain a duplicate office container"));
        }
        Ok(())
    }

    fn finish(self) -> Result<()> {
        if self.depth != 0 || !self.root_seen || !self.root_closed {
            return Err(invalid("ODC styles structure is incomplete"));
        }
        Ok(())
    }
}

pub(crate) fn validate(xml: &str, limits: crate::Limits) -> Result<()> {
    if xml.len() > limits.max_content_bytes() {
        return Err(invalid(
            "ODC styles exceed the caller-selected content limit",
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut state = State::new();
    loop {
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODC styles XML: {error}")))?;
        let namespace = namespace_kind(&resolved_namespace);
        let event = borrowed_event.into_owned();
        match event {
            Event::Start(element) => state.start(&reader, namespace, &element, limits)?,
            Event::Empty(element) => state.empty(&reader, namespace, &element, limits)?,
            Event::End(_) => state.end()?,
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODC styles")),
            Event::Eof => break,
            Event::Text(text) => {
                if state.depth == 0 && !text.iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("ODC styles contain text outside the root"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if state.depth == 0 => {
                return Err(invalid("ODC styles contain data outside the root"));
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    state.finish()
}

fn reject_scripts(namespace: NamespaceKind, element: &BytesStart<'_>) -> Result<()> {
    if namespace == NamespaceKind::Office && element.local_name().as_ref() == b"scripts" {
        return Err(Error::Unsupported(
            "ODC styles refuse executable office:scripts content".into(),
        ));
    }
    Ok(())
}

fn validate_style_element(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    scope: Option<u8>,
    names: &mut BTreeSet<(u8, String, String)>,
    limits: crate::Limits,
) -> Result<()> {
    let mut name = None;
    let mut family = None;
    for result in element.attributes().with_checks(true) {
        let attribute =
            result.map_err(|error| invalid(format!("invalid ODC styles attribute: {error}")))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid ODC styles value: {error}")))?
            .into_owned();
        if value.len() > limits.max_scalar_bytes() {
            return Err(invalid(
                "ODC styles attribute exceeds the caller-selected scalar limit",
            ));
        }
        let (attribute_namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved(&attribute_namespace, STYLE) {
            match local.as_ref() {
                b"name" => name = Some(value),
                b"family" => family = Some(value),
                _ => {},
            }
        }
    }
    if namespace == NamespaceKind::Style && element.local_name().as_ref() == b"style" {
        let required_scope =
            scope.ok_or_else(|| invalid("style:style is outside a style collection"))?;
        let style_name = name
            .filter(|value| !value.is_empty())
            .ok_or_else(|| invalid("style:style requires style:name"))?;
        let style_family = family
            .filter(|value| !value.is_empty())
            .ok_or_else(|| invalid("style:style requires style:family"))?;
        if !names.insert((required_scope, style_family, style_name)) {
            return Err(invalid(
                "ODC styles contain a duplicate style name and family",
            ));
        }
    }
    Ok(())
}

fn style_scope_id(local: &[u8]) -> Option<u8> {
    match local {
        b"font-face-decls" => Some(1),
        b"styles" => Some(2),
        b"automatic-styles" => Some(3),
        b"master-styles" => Some(4),
        _ => None,
    }
}

fn checked_depth(depth: usize, limit: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODC styles depth overflow"))?;
    if next > limit {
        return Err(invalid("ODC styles exceed the caller-selected depth limit"));
    }
    Ok(next)
}

fn resolved(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    if resolved(namespace, OFFICE) {
        NamespaceKind::Office
    } else if resolved(namespace, STYLE) {
        NamespaceKind::Style
    } else {
        NamespaceKind::Other
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
