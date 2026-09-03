#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
#![expect(
    clippy::struct_excessive_bools,
    reason = "the public model preserves independent OOXML flags"
)]
//! Bounded and lossless `w:sectPr` `WordprocessingML` codec.

use super::model::{
    Column, Columns, Emu, MAX_TWIPS, MAX_XML_BYTES, MAX_XML_DEPTH, MAX_XML_NODES, Margins,
    Orientation, PageSize, Reference, Start, State,
};
use crate::error::{Error, Result};
use crate::header_footer::Kind;
use crate::namespace::is_wordprocessing_namespace;
use litchi_ooxml_common::xml_name::{is_ncname, is_qualified_name};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};
use std::fmt::Write;

const WORD_ROOT: &[u8] = b"sectPr";

#[derive(Debug, Clone)]
pub(crate) struct Snapshot {
    pub(crate) raw: Raw,
    pub(crate) state: State,
    dirty: Dirty,
}

impl Snapshot {
    pub(crate) fn is_dirty(&self) -> bool {
        !self.dirty.is_empty()
    }

    pub(crate) fn mark_dirty_since(&mut self, previous: &State) {
        self.dirty.page_size = self.state.page_size != previous.page_size;
        self.dirty.margins = self.state.margins != previous.margins;
        self.dirty.start = self.state.start != previous.start;
        self.dirty.columns = self.state.columns != previous.columns;
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct Dirty {
    page_size: bool,
    margins: bool,
    start: bool,
    columns: bool,
}

impl Dirty {
    fn is_empty(self) -> bool {
        !self.page_size && !self.margins && !self.start && !self.columns
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Raw {
    prefix: Vec<u8>,
    root_open: Vec<u8>,
    root_close: Vec<u8>,
    root_name: Vec<u8>,
    root_prefix: Option<Vec<u8>>,
    attribute_prefix: Option<Vec<u8>>,
    generated_attribute_prefix: Option<Vec<u8>>,
    word_namespace: Option<Vec<u8>>,
    namespace_bindings: Vec<NamespaceBinding>,
    suffix: Vec<u8>,
    children: Vec<Node>,
}

#[derive(Debug, Clone)]
struct NamespaceBinding {
    prefix: Option<Vec<u8>>,
    uri: String,
}

#[derive(Debug, Clone)]
enum Node {
    Element {
        local_name: String,
        raw: Vec<u8>,
        word: bool,
    },
    Raw(Vec<u8>),
}

/// Validate the section fragment without decoding semantic properties.
pub(crate) fn validate(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "section XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    parse_raw(xml).map(|_| ())
}

/// Decode a validated section fragment into its owned raw topology and lazy
/// semantic state.
pub(crate) fn decode(xml: &[u8]) -> Result<Snapshot> {
    let raw = parse_raw(xml)?;
    let state = decode_state(&raw)?;
    Ok(Snapshot {
        raw,
        state,
        dirty: Dirty::default(),
    })
}

pub(crate) fn validate_page_size(page_size: &PageSize) -> Result<()> {
    validate_optional_measurement(page_size.width, "page width", false, true)?;
    validate_optional_measurement(page_size.height, "page height", false, true)
}

pub(crate) fn validate_margins(margins: &Margins) -> Result<()> {
    validate_optional_measurement(margins.top, "top margin", true, false)?;
    validate_optional_measurement(margins.bottom, "bottom margin", true, false)?;
    validate_optional_measurement(margins.right, "right margin", false, false)?;
    validate_optional_measurement(margins.left, "left margin", false, false)?;
    validate_optional_measurement(margins.header, "header distance", false, false)?;
    validate_optional_measurement(margins.footer, "footer distance", false, false)?;
    validate_optional_measurement(margins.gutter, "gutter", false, false)
}

pub(crate) fn validate_columns(columns: &Columns) -> Result<()> {
    if !(1..=45).contains(&columns.count) {
        return Err(Error::InvalidFormat(
            "section column count must be in 1..=45".into(),
        ));
    }
    validate_optional_measurement(columns.space, "column space", false, false)?;
    if !columns.equal_width && usize::from(columns.count) != columns.columns.len() {
        return Err(Error::InvalidFormat(
            "unequal section columns require one width per column".into(),
        ));
    }
    for column in &columns.columns {
        validate_optional_measurement(Some(column.width), "column width", false, false)?;
        validate_optional_measurement(column.space, "column space", false, false)?;
    }
    Ok(())
}

fn validate_optional_measurement(
    value: Option<Emu>,
    description: &str,
    signed: bool,
    page_size: bool,
) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    let twips = value.try_to_twips()?;
    let valid = if signed {
        (-MAX_TWIPS..=MAX_TWIPS).contains(&twips)
    } else if page_size {
        (1..=MAX_TWIPS).contains(&twips)
    } else {
        (0..=MAX_TWIPS).contains(&twips)
    };
    if !valid {
        return Err(Error::InvalidFormat(format!(
            "{description} {twips} twips is outside the Word domain"
        )));
    }
    Ok(())
}

/// Encode a changed snapshot. Unchanged children, comments, whitespace, and
/// foreign extension elements retain their original bytes and relative order.
pub(crate) fn encode(snapshot: &Snapshot) -> Result<Vec<u8>> {
    if snapshot.dirty.is_empty() {
        return Ok(snapshot.raw.original_bytes());
    }
    ensure_no_unknown_word_insertion_barrier(&snapshot.raw, &snapshot.state, snapshot.dirty)?;

    let mut output = Vec::new();
    output.extend_from_slice(&snapshot.raw.prefix);
    output.extend_from_slice(&snapshot.raw.open_tag());

    let mut inserted = [false; 4];
    for node in &snapshot.raw.children {
        let rank = node_rank(node);
        insert_missing_before(
            &mut output,
            &snapshot.raw,
            &snapshot.state,
            snapshot.dirty,
            rank,
            &mut inserted,
        )?;

        match node {
            Node::Element {
                local_name,
                raw,
                word,
            } if *word => {
                if local_name == "type" && snapshot.dirty.start {
                    if !inserted[0] {
                        if let Some(start) = snapshot.state.start {
                            output.extend_from_slice(&write_start(&snapshot.raw, start)?);
                        }
                        inserted[0] = true;
                    }
                } else if local_name == "pgSz" && snapshot.dirty.page_size {
                    if !inserted[1] {
                        if let Some(page_size) = snapshot.state.page_size {
                            output.extend_from_slice(&write_page_size(&snapshot.raw, &page_size)?);
                        }
                        inserted[1] = true;
                    }
                } else if local_name == "pgMar" && snapshot.dirty.margins {
                    if !inserted[2] {
                        if let Some(margins) = snapshot.state.margins {
                            output.extend_from_slice(&write_margins(&snapshot.raw, &margins)?);
                        }
                        inserted[2] = true;
                    }
                } else if local_name == "cols" && snapshot.dirty.columns {
                    if !inserted[3] {
                        if let Some(columns) = &snapshot.state.columns {
                            output.extend_from_slice(&write_columns(&snapshot.raw, columns)?);
                        }
                        inserted[3] = true;
                    }
                } else {
                    output.extend_from_slice(raw);
                }
            },
            Node::Element { raw, .. } | Node::Raw(raw) => output.extend_from_slice(raw),
        }
    }

    insert_missing_before(
        &mut output,
        &snapshot.raw,
        &snapshot.state,
        snapshot.dirty,
        Some(u8::MAX),
        &mut inserted,
    )?;
    output.extend_from_slice(&snapshot.raw.close_tag());
    output.extend_from_slice(&snapshot.raw.suffix);
    Ok(output)
}

impl Raw {
    fn original_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&self.prefix);
        bytes.extend_from_slice(&self.root_open);
        for child in &self.children {
            match child {
                Node::Element { raw, .. } | Node::Raw(raw) => bytes.extend_from_slice(raw),
            }
        }
        bytes.extend_from_slice(&self.root_close);
        bytes.extend_from_slice(&self.suffix);
        bytes
    }

    fn open_tag(&self) -> Vec<u8> {
        if !self.root_close.is_empty() {
            return self.root_open.clone();
        }
        let mut open = self.root_open.clone();
        if let Some(index) = open.windows(2).rposition(|window| window == b"/>") {
            open.remove(index);
        }
        open
    }

    fn close_tag(&self) -> Vec<u8> {
        if !self.root_close.is_empty() {
            return self.root_close.clone();
        }
        let mut close = Vec::with_capacity(self.root_name.len() + 3);
        close.extend_from_slice(b"</");
        close.extend_from_slice(&self.root_name);
        close.extend_from_slice(b">");
        close
    }

    fn element_name(&self, local_name: &str) -> String {
        match &self.root_prefix {
            Some(prefix) if !prefix.is_empty() => {
                format!("{}:{local_name}", String::from_utf8_lossy(prefix))
            },
            _ => local_name.to_owned(),
        }
    }

    fn attribute_name(&self, local_name: &str) -> Result<String> {
        match &self.attribute_prefix {
            Some(prefix) if !prefix.is_empty() => {
                Ok(format!("{}:{local_name}", String::from_utf8_lossy(prefix)))
            },
            _ => Err(Error::InvalidFormat(format!(
                "cannot generate WordprocessingML attribute '{local_name}' without a named WordprocessingML prefix"
            ))),
        }
    }

    fn write_generated_attribute_namespace(&self, xml: &mut String) -> Result<()> {
        let Some(prefix) = &self.generated_attribute_prefix else {
            return Ok(());
        };
        let Some(namespace) = &self.word_namespace else {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "edit_section_layout",
                reason: "the generated section property has no resolved owning Word namespace",
            });
        };
        write!(
            xml,
            " xmlns:{}=\"{}\"",
            String::from_utf8_lossy(prefix),
            escape_namespace_uri(&String::from_utf8_lossy(namespace)),
        )?;
        Ok(())
    }

    fn has_namespace_binding(&self, prefix: &[u8]) -> bool {
        self.namespace_bindings
            .iter()
            .any(|binding| binding.prefix.as_deref() == Some(prefix))
    }

    fn detached_fragment(&self, xml: &[u8]) -> Vec<u8> {
        let mut detached = Vec::with_capacity(xml.len() + 128);
        detached.extend_from_slice(b"<litchiRoot");
        for binding in &self.namespace_bindings {
            detached.extend_from_slice(b" xmlns");
            if let Some(prefix) = &binding.prefix {
                detached.push(b':');
                detached.extend_from_slice(prefix);
            }
            detached.extend_from_slice(b"=\"");
            detached.extend_from_slice(escape_namespace_uri(&binding.uri).as_bytes());
            detached.push(b'"');
        }
        detached.extend_from_slice(b">");
        detached.extend_from_slice(xml);
        detached.extend_from_slice(b"</litchiRoot>");
        detached
    }
}

fn insert_missing_before(
    output: &mut Vec<u8>,
    raw: &Raw,
    state: &State,
    dirty: Dirty,
    before_rank: Option<u8>,
    inserted: &mut [bool; 4],
) -> Result<()> {
    let start = if dirty.start {
        state
            .start
            .map(|value| write_start(raw, value))
            .transpose()?
    } else {
        None
    };
    let page_size = if dirty.page_size {
        state
            .page_size
            .map(|value| write_page_size(raw, &value))
            .transpose()?
    } else {
        None
    };
    let margins = if dirty.margins {
        state
            .margins
            .map(|value| write_margins(raw, &value))
            .transpose()?
    } else {
        None
    };
    let columns = if dirty.columns {
        state
            .columns
            .as_ref()
            .map(|value| write_columns(raw, value))
            .transpose()?
    } else {
        None
    };
    let fields = [
        (
            0,
            section_child_rank(b"type").expect("section type rank"),
            dirty.start,
            start,
        ),
        (
            1,
            section_child_rank(b"pgSz").expect("section page size rank"),
            dirty.page_size,
            page_size,
        ),
        (
            2,
            section_child_rank(b"pgMar").expect("section margin rank"),
            dirty.margins,
            margins,
        ),
        (
            3,
            section_child_rank(b"cols").expect("section columns rank"),
            dirty.columns,
            columns,
        ),
    ];
    for (index, rank, dirty, bytes) in fields {
        if dirty && !inserted[index] && before_rank.is_some_and(|before_rank| rank < before_rank) {
            if let Some(bytes) = bytes {
                output.extend_from_slice(&bytes);
            }
            inserted[index] = true;
        }
    }
    Ok(())
}

fn ensure_no_unknown_word_insertion_barrier(raw: &Raw, state: &State, dirty: Dirty) -> Result<()> {
    let has_unknown_word_child = raw.children.iter().any(|node| {
        matches!(
            node,
            Node::Element {
                local_name,
                word: true,
                ..
            } if section_child_rank(local_name.as_bytes()).is_none()
        )
    });
    if !has_unknown_word_child {
        return Ok(());
    }
    let insertions = [
        (b"type".as_slice(), dirty.start, state.start.is_some()),
        (
            b"pgSz".as_slice(),
            dirty.page_size,
            state.page_size.is_some(),
        ),
        (b"pgMar".as_slice(), dirty.margins, state.margins.is_some()),
        (b"cols".as_slice(), dirty.columns, state.columns.is_some()),
    ];
    let needs_insertion = insertions.iter().any(|(local_name, dirty, present)| {
        *dirty
            && *present
            && !raw.children.iter().any(|node| {
                matches!(
                    node,
                    Node::Element {
                        local_name: child_name,
                        word: true,
                        ..
                    } if child_name.as_bytes() == *local_name
                )
            })
    });
    if needs_insertion {
        return Err(Error::UnsafeEdit {
            format: "DOCX",
            operation: "edit_section_layout",
            reason: "an unknown direct Word section child blocks safe modeled insertion",
        });
    }
    Ok(())
}

fn node_rank(node: &Node) -> Option<u8> {
    match node {
        Node::Element {
            local_name,
            word: true,
            ..
        } => section_child_rank(local_name.as_bytes()),
        Node::Element { .. } | Node::Raw(_) => None,
    }
}

fn write_start(raw: &Raw, start: Start) -> Result<Vec<u8>> {
    let mut xml = String::new();
    write!(
        xml,
        "<{} {}=\"{}\"",
        raw.element_name("type"),
        raw.attribute_name("val")?,
        start.to_xml()
    )?;
    raw.write_generated_attribute_namespace(&mut xml)?;
    xml.push_str("/>");
    Ok(xml.into_bytes())
}

fn write_page_size(raw: &Raw, page_size: &PageSize) -> Result<Vec<u8>> {
    validate_page_size(page_size)?;
    let mut xml = String::new();
    write!(xml, "<{}", raw.element_name("pgSz"))?;
    raw.write_generated_attribute_namespace(&mut xml)?;
    if let Some(width) = page_size.width {
        write!(
            xml,
            " {}=\"{}\"",
            raw.attribute_name("w")?,
            width.try_to_twips()?
        )?;
    }
    if let Some(height) = page_size.height {
        write!(
            xml,
            " {}=\"{}\"",
            raw.attribute_name("h")?,
            height.try_to_twips()?
        )?;
    }
    write!(
        xml,
        " {}=\"{}\"/>",
        raw.attribute_name("orient")?,
        page_size.orientation.to_xml()
    )?;
    Ok(xml.into_bytes())
}

fn write_margins(raw: &Raw, margins: &Margins) -> Result<Vec<u8>> {
    validate_margins(margins)?;
    let mut xml = String::new();
    write!(xml, "<{}", raw.element_name("pgMar"))?;
    raw.write_generated_attribute_namespace(&mut xml)?;
    write_measurement(raw, &mut xml, "top", margins.top)?;
    write_measurement(raw, &mut xml, "right", margins.right)?;
    write_measurement(raw, &mut xml, "bottom", margins.bottom)?;
    write_measurement(raw, &mut xml, "left", margins.left)?;
    write_measurement(raw, &mut xml, "header", margins.header)?;
    write_measurement(raw, &mut xml, "footer", margins.footer)?;
    write_measurement(raw, &mut xml, "gutter", margins.gutter)?;
    xml.push_str("/>");
    Ok(xml.into_bytes())
}

fn write_measurement(raw: &Raw, xml: &mut String, name: &str, value: Option<Emu>) -> Result<()> {
    if let Some(value) = value {
        write!(
            xml,
            " {}=\"{}\"",
            raw.attribute_name(name)?,
            value.try_to_twips()?
        )?;
    }
    Ok(())
}

fn write_columns(raw: &Raw, columns: &Columns) -> Result<Vec<u8>> {
    validate_columns(columns)?;
    let mut xml = String::new();
    write!(
        xml,
        "<{} {}=\"{}\" {}=\"{}\"",
        raw.element_name("cols"),
        raw.attribute_name("equalWidth")?,
        i32::from(columns.equal_width),
        raw.attribute_name("num")?,
        columns.count
    )?;
    raw.write_generated_attribute_namespace(&mut xml)?;
    if let Some(space) = columns.space {
        write!(
            xml,
            " {}=\"{}\"",
            raw.attribute_name("space")?,
            space.try_to_twips()?
        )?;
    }
    if columns.separator {
        write!(xml, " {}=\"1\"", raw.attribute_name("sep")?)?;
    }
    if columns.columns.is_empty() {
        xml.push_str("/>");
        return Ok(xml.into_bytes());
    }
    xml.push('>');
    for column in &columns.columns {
        write!(
            xml,
            "<{} {}=\"{}\"",
            raw.element_name("col"),
            raw.attribute_name("w")?,
            column.width.try_to_twips()?
        )?;
        if let Some(space) = column.space {
            write!(
                xml,
                " {}=\"{}\"",
                raw.attribute_name("space")?,
                space.try_to_twips()?
            )?;
        }
        xml.push_str("/>");
    }
    write!(xml, "</{}>", raw.element_name("cols"))?;
    Ok(xml.into_bytes())
}

pub(crate) fn validate_element_qname(value: &[u8], owner: &str) -> Result<()> {
    validate_qname(value, owner, "element")?;
    if value
        .split(|byte| *byte == b':')
        .next()
        .is_some_and(|prefix| prefix == b"xmlns")
        && value.contains(&b':')
    {
        return Err(Error::InvalidFormat(format!(
            "{owner} element cannot use the reserved xmlns prefix"
        )));
    }
    Ok(())
}

pub(crate) fn validate_qname(value: &[u8], owner: &str, kind: &str) -> Result<()> {
    let value = std::str::from_utf8(value)
        .map_err(|error| Error::Xml(format!("invalid UTF-8 in {owner} {kind} QName: {error}")))?;
    if !is_qualified_name(value) {
        return Err(Error::InvalidFormat(format!(
            "{owner} {kind} has an invalid QName '{value}'"
        )));
    }
    Ok(())
}

fn validate_attribute_qnames(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    owner: &str,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        validate_qname(attribute.key.as_ref(), owner, "attribute")?;
        let prefix = attribute.key.prefix();
        let is_namespace_declaration = (prefix.is_none()
            && attribute.key.local_name().as_ref() == b"xmlns")
            || prefix
                .as_ref()
                .is_some_and(|value| value.as_ref() == b"xmlns");
        if !is_namespace_declaration {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_namespace_declaration(attribute.key.as_ref(), value.as_ref(), owner)?;
    }
    Ok(())
}

pub(crate) fn validate_namespace_declaration(name: &[u8], value: &str, owner: &str) -> Result<()> {
    const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
    const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
    let prefix = name.strip_prefix(b"xmlns:");
    if let Some(prefix) = prefix {
        let prefix = std::str::from_utf8(prefix).map_err(|error| {
            Error::Xml(format!(
                "invalid UTF-8 in {owner} namespace prefix: {error}"
            ))
        })?;
        if !is_ncname(prefix) || prefix == "xmlns" {
            return Err(Error::InvalidFormat(format!(
                "{owner} namespace declaration has an invalid prefix '{prefix}'"
            )));
        }
        if value.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "{owner} namespace prefix '{prefix}' cannot be undeclared"
            )));
        }
        if (prefix == "xml") != (value == XML_NAMESPACE) {
            return Err(Error::InvalidFormat(
                "the XML namespace URI may be bound only to the xml prefix".into(),
            ));
        }
    } else if value == XML_NAMESPACE {
        return Err(Error::InvalidFormat(
            "the XML namespace URI may be bound only to the xml prefix".into(),
        ));
    }
    if value == XMLNS_NAMESPACE {
        return Err(Error::InvalidFormat(
            "the xmlns namespace URI cannot be rebound".into(),
        ));
    }
    Ok(())
}

fn parse_raw(xml: &[u8]) -> Result<Raw> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "section XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut children = Vec::new();
    let mut root_open = None;
    let mut root_close = Vec::new();
    let mut root_name = Vec::new();
    let mut root_prefix = None;
    let mut attribute_prefix: Option<Vec<u8>> = None;
    let mut generated_attribute_prefix: Option<Vec<u8>> = None;
    let mut word_namespace: Option<Vec<u8>> = None;
    let mut namespace_bindings = Vec::new();
    let mut root_start = 0usize;
    let mut root_end = 0usize;
    let mut root_seen = false;
    let mut direct_child_start = None;
    let mut nodes = 0usize;

    loop {
        let event_start = offset(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                validate_element_qname(element.name().as_ref(), "section")?;
                validate_attribute_qnames(element, decoder, "section")?;
            },
            Event::End(element) => {
                validate_element_qname(element.name().as_ref(), "section end")?;
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Eof => {},
        }
        if stack.is_empty() && matches!(event, Event::Start(_) | Event::Empty(_)) {
            validate_root_namespace(&namespace)?;
        }
        let word_element = is_word_element(&namespace, &fragment_prefix);
        let event_word_namespace = resolved_word_namespace(&namespace);
        let event_end = offset(&reader)?;
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("section XML element counter overflow".into())
            })?;
            if nodes > MAX_XML_NODES {
                return Err(Error::InvalidFormat(format!(
                    "section XML exceeds {MAX_XML_NODES} elements"
                )));
            }
        }

        match event {
            Event::Start(element) => {
                if stack.is_empty() {
                    if root_seen || element.local_name().as_ref() != WORD_ROOT {
                        return Err(Error::InvalidFormat(
                            "section properties have an invalid root".into(),
                        ));
                    }
                    fragment_prefix = Some(element_prefix(&element));
                    root_seen = true;
                    root_start = event_start;
                    root_name = element.name().as_ref().to_vec();
                    root_prefix = element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec());
                    namespace_bindings = root_namespace_bindings(&element, decoder)?;
                    word_namespace = event_word_namespace.clone();
                    attribute_prefix = root_attribute_prefix(
                        root_prefix.as_deref(),
                        &namespace_bindings,
                        word_namespace.as_deref(),
                    );
                    if attribute_prefix.is_none() && word_namespace.is_some() {
                        let prefix = fresh_namespace_prefix(&namespace_bindings);
                        attribute_prefix = Some(prefix.clone());
                        generated_attribute_prefix = Some(prefix);
                    }
                    root_open = Some(xml[event_start..event_end].to_vec());
                } else if stack.len() == 1 {
                    direct_child_start = Some((
                        event_start,
                        String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        word_element,
                    ));
                }
                stack.push(element.name().as_ref().to_vec());
                if stack.len() > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "section XML nesting exceeds {MAX_XML_DEPTH}"
                    )));
                }
            },
            Event::Empty(element) => {
                if stack.is_empty() {
                    if root_seen || element.local_name().as_ref() != WORD_ROOT {
                        return Err(Error::InvalidFormat(
                            "section properties have an invalid root".into(),
                        ));
                    }
                    fragment_prefix = Some(element_prefix(&element));
                    root_seen = true;
                    root_start = event_start;
                    root_end = event_end;
                    root_name = element.name().as_ref().to_vec();
                    root_prefix = element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec());
                    root_open = Some(xml[event_start..event_end].to_vec());
                    namespace_bindings = root_namespace_bindings(&element, decoder)?;
                    word_namespace = event_word_namespace.clone();
                    attribute_prefix = root_attribute_prefix(
                        root_prefix.as_deref(),
                        &namespace_bindings,
                        word_namespace.as_deref(),
                    );
                    if attribute_prefix.is_none() && word_namespace.is_some() {
                        let prefix = fresh_namespace_prefix(&namespace_bindings);
                        attribute_prefix = Some(prefix.clone());
                        generated_attribute_prefix = Some(prefix);
                    }
                } else if stack.len() == 1 {
                    children.push(Node::Element {
                        local_name: String::from_utf8_lossy(element.local_name().as_ref())
                            .into_owned(),
                        raw: xml[event_start..event_end].to_vec(),
                        word: word_element,
                    });
                }
            },
            Event::End(element) => {
                let expected = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("section XML has an unmatched end element".into())
                })?;
                if expected != element.name().as_ref() {
                    return Err(Error::InvalidFormat(
                        "section XML has mismatched element nesting".into(),
                    ));
                }
                if stack.is_empty() {
                    if element.local_name().as_ref() != WORD_ROOT || !root_seen {
                        return Err(Error::InvalidFormat(
                            "section properties have an invalid root".into(),
                        ));
                    }
                    root_close = xml[event_start..event_end].to_vec();
                    root_end = event_end;
                } else if stack.len() == 1 {
                    let (child_start, local_name, word) =
                        direct_child_start.take().ok_or_else(|| {
                            Error::InvalidFormat("section XML child tracking failed".into())
                        })?;
                    children.push(Node::Element {
                        local_name,
                        raw: xml[child_start..event_end].to_vec(),
                        word,
                    });
                }
            },
            Event::Eof => break,
            _ if stack.len() == 1 => {
                children.push(Node::Raw(xml[event_start..event_end].to_vec()));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !root_seen || !stack.is_empty() || root_open.is_none() {
        return Err(Error::InvalidFormat(
            "unterminated section properties".into(),
        ));
    }
    if root_end == 0 {
        return Err(Error::InvalidFormat(
            "section properties have no end".into(),
        ));
    }
    if xml[..root_start]
        .iter()
        .any(|byte| !byte.is_ascii_whitespace())
        && !is_xml_declaration_or_misc(&xml[..root_start])
    {
        return Err(Error::InvalidFormat(
            "section XML has invalid content before sectPr".into(),
        ));
    }
    if xml[root_end..]
        .iter()
        .any(|byte| !byte.is_ascii_whitespace())
        && !is_xml_declaration_or_misc(&xml[root_end..])
    {
        return Err(Error::InvalidFormat(
            "section XML has invalid trailing content".into(),
        ));
    }

    Ok(Raw {
        prefix: xml[..root_start].to_vec(),
        root_open: root_open.expect("root present"),
        root_close,
        root_name,
        root_prefix,
        attribute_prefix,
        generated_attribute_prefix,
        word_namespace,
        namespace_bindings,
        suffix: xml[root_end..].to_vec(),
        children,
    })
}

fn decode_state(raw: &Raw) -> Result<State> {
    let mut state = State::default();
    let mut seen = [false; 21];
    let mut last_rank = None;
    for node in &raw.children {
        let Node::Element {
            local_name,
            raw: child,
            word: true,
        } = node
        else {
            continue;
        };
        if let Some(rank) = section_child_rank(local_name.as_bytes()) {
            if last_rank.is_some_and(|last_rank| rank < last_rank) {
                return Err(Error::InvalidFormat(format!(
                    "section property '{local_name}' is out of schema order"
                )));
            }
            last_rank = Some(rank);
            let rank = usize::from(rank);
            if seen[rank] && !matches!(local_name.as_str(), "headerReference" | "footerReference") {
                return Err(Error::InvalidFormat(format!(
                    "section properties contain duplicate '{local_name}'"
                )));
            }
            seen[rank] = true;
        }
        match local_name.as_str() {
            "type" => {
                if state.start.is_some() {
                    return Err(Error::InvalidFormat("duplicate section type".into()));
                }
                let value = required_attribute(raw, child, b"val")?;
                state.start = Some(Start::from_xml(&value).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid section type '{value}'"))
                })?);
            },
            "pgSz" => {
                if state.page_size.is_some() {
                    return Err(Error::InvalidFormat("duplicate section page size".into()));
                }
                state.page_size = Some(parse_page_size(raw, child)?);
            },
            "pgMar" => {
                if state.margins.is_some() {
                    return Err(Error::InvalidFormat("duplicate section margins".into()));
                }
                state.margins = Some(parse_margins(raw, child)?);
            },
            "cols" => {
                if state.columns.is_some() {
                    return Err(Error::InvalidFormat("duplicate section columns".into()));
                }
                state.columns = Some(parse_columns(raw, child)?);
            },
            "headerReference" => state.headers.push(parse_reference(raw, child)?),
            "footerReference" => state.footers.push(parse_reference(raw, child)?),
            _ => {},
        }
    }
    validate_page_size(&state.page_size.unwrap_or_default())?;
    if let Some(margins) = state.margins {
        validate_margins(&margins)?;
    }
    if let Some(columns) = &state.columns {
        validate_columns(columns)?;
    }
    validate_references(&state.headers)?;
    validate_references(&state.footers)?;
    Ok(state)
}

pub(crate) const fn section_child_rank(name: &[u8]) -> Option<u8> {
    match name {
        b"headerReference" | b"footerReference" => Some(0),
        b"footnotePr" => Some(1),
        b"endnotePr" => Some(2),
        b"type" => Some(3),
        b"pgSz" => Some(4),
        b"pgMar" => Some(5),
        b"paperSrc" => Some(6),
        b"pgBorders" => Some(7),
        b"lnNumType" => Some(8),
        b"pgNumType" => Some(9),
        b"cols" => Some(10),
        b"formProt" => Some(11),
        b"vAlign" => Some(12),
        b"noEndnote" => Some(13),
        b"titlePg" => Some(14),
        b"textDirection" => Some(15),
        b"bidi" => Some(16),
        b"rtlGutter" => Some(17),
        b"docGrid" => Some(18),
        b"printerSettings" => Some(19),
        b"sectPrChange" => Some(20),
        _ => None,
    }
}

fn parse_page_size(raw: &Raw, xml: &[u8]) -> Result<PageSize> {
    let attrs = attributes(raw, xml, AttributeFamily::Word)?;
    let width = attr(&attrs, "w")
        .map(|value| parse_measurement(value, "page width", false, true))
        .transpose()?;
    let height = attr(&attrs, "h")
        .map(|value| parse_measurement(value, "page height", false, true))
        .transpose()?;
    let orientation = attr(&attrs, "orient")
        .map(|value| {
            Orientation::from_xml(value).ok_or_else(|| {
                Error::InvalidFormat(format!("invalid section orientation '{value}'"))
            })
        })
        .transpose()?
        .unwrap_or_default();
    Ok(PageSize {
        width,
        height,
        orientation,
    })
}

fn parse_margins(raw: &Raw, xml: &[u8]) -> Result<Margins> {
    let attrs = attributes(raw, xml, AttributeFamily::Word)?;
    Ok(Margins {
        top: parse_attr_measurement(&attrs, "top", true, false)?,
        right: parse_attr_measurement(&attrs, "right", false, false)?,
        bottom: parse_attr_measurement(&attrs, "bottom", true, false)?,
        left: parse_attr_measurement(&attrs, "left", false, false)?,
        header: parse_attr_measurement(&attrs, "header", false, false)?,
        footer: parse_attr_measurement(&attrs, "footer", false, false)?,
        gutter: parse_attr_measurement(&attrs, "gutter", false, false)?,
    })
}

fn parse_columns(raw: &Raw, xml: &[u8]) -> Result<Columns> {
    let attrs = attributes(raw, xml, AttributeFamily::Word)?;
    let mut columns = Columns {
        equal_width: attr(&attrs, "equalWidth")
            .is_none_or(|value| value != "0" && value != "false"),
        count: attr(&attrs, "num")
            .map(|value| {
                value.parse::<u16>().map_err(|_source_error| {
                    Error::InvalidFormat("invalid section column count".into())
                })
            })
            .transpose()?
            .unwrap_or(1),
        space: parse_attr_measurement(&attrs, "space", false, false)?,
        separator: attr(&attrs, "sep").is_some_and(|value| value == "1" || value == "true"),
        columns: Vec::new(),
    };
    for (name, child) in direct_children(xml)? {
        if name != "col" {
            continue;
        }
        let attrs = attributes(raw, &child, AttributeFamily::Word)?;
        let width = parse_attr_measurement(&attrs, "w", false, false)?
            .ok_or_else(|| Error::InvalidFormat("section column omits required width".into()))?;
        columns.columns.push(Column {
            width,
            space: parse_attr_measurement(&attrs, "space", false, false)?,
        });
    }
    validate_columns(&columns)?;
    Ok(columns)
}

fn parse_reference(raw: &Raw, xml: &[u8]) -> Result<Reference> {
    let attrs = attributes(raw, xml, AttributeFamily::Reference)?;
    let kind = Kind::from_xml(required_attr(&attrs, "type")?.as_str())
        .ok_or_else(|| Error::InvalidFormat("invalid section header/footer type".into()))?;
    let relationship_id = required_attr(&attrs, "id")?;
    if relationship_id.is_empty() {
        return Err(Error::InvalidFormat(
            "section header/footer relationship ID is empty".into(),
        ));
    }
    if !is_xml_id(&relationship_id) {
        return Err(Error::InvalidFormat(
            "section header/footer relationship ID is not an XML ID".into(),
        ));
    }
    Ok(Reference {
        kind,
        relationship_id,
    })
}

fn is_xml_id(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_char)
}

fn is_ncname_start(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00c0}'..='\u{00d6}'
            | '\u{00d8}'..='\u{00f6}'
            | '\u{00f8}'..='\u{02ff}'
            | '\u{0370}'..='\u{037d}'
            | '\u{037f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

fn validate_references(references: &[Reference]) -> Result<()> {
    let mut kinds = Vec::new();
    for reference in references {
        if kinds.contains(&reference.kind) {
            return Err(Error::InvalidFormat(
                "section has duplicate header/footer reference type".into(),
            ));
        }
        kinds.push(reference.kind);
    }
    Ok(())
}

fn parse_attr_measurement(
    attrs: &[(String, String)],
    name: &str,
    signed: bool,
    page_size: bool,
) -> Result<Option<Emu>> {
    attr(attrs, name)
        .map(|value| parse_measurement(value, name, signed, page_size))
        .transpose()
}

fn parse_measurement(value: &str, description: &str, signed: bool, page_size: bool) -> Result<Emu> {
    let twips = value.parse::<i64>().map_err(|_source_error| {
        Error::InvalidFormat(format!("invalid {description} twip value '{value}'"))
    })?;
    let valid = if signed {
        (-MAX_TWIPS..=MAX_TWIPS).contains(&twips)
    } else if page_size {
        (1..=MAX_TWIPS).contains(&twips)
    } else {
        (0..=MAX_TWIPS).contains(&twips)
    };
    if !valid {
        return Err(Error::InvalidFormat(format!(
            "{description} {twips} twips is outside the Word domain"
        )));
    }
    Emu::try_from_twips(twips)
}

fn direct_children(xml: &[u8]) -> Result<Vec<(String, Vec<u8>)>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root = false;
    let mut start = None;
    let mut output = Vec::new();
    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    root = true;
                } else if depth == 1 {
                    start = Some((
                        String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        event_start,
                    ));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("section XML nesting is too deep".into())
                })?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    let _ = xml.get(event_start..event_end).ok_or_else(|| {
                        Error::InvalidFormat(
                            "section child XML range is outside the fragment".into(),
                        )
                    })?;
                    root = true;
                } else if depth == 1 {
                    let raw = xml.get(event_start..event_end).ok_or_else(|| {
                        Error::InvalidFormat(
                            "section child XML range is outside the fragment".into(),
                        )
                    })?;
                    output.push((
                        String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        raw.to_vec(),
                    ));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid section XML nesting".into()))?;
                if depth == 1 {
                    let (name, start) = start.take().ok_or_else(|| {
                        Error::InvalidFormat("section child tracking failed".into())
                    })?;
                    let raw = xml.get(start..event_end).ok_or_else(|| {
                        Error::InvalidFormat(
                            "section child XML range is outside the fragment".into(),
                        )
                    })?;
                    output.push((name, raw.to_vec()));
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root || depth != 0 {
        return Err(Error::InvalidFormat("invalid section child XML".into()));
    }
    Ok(output)
}

#[derive(Clone, Copy)]
enum AttributeFamily {
    Word,
    Reference,
}

fn attributes(raw: &Raw, xml: &[u8], family: AttributeFamily) -> Result<Vec<(String, String)>> {
    let detached = raw.detached_fragment(xml);
    let mut reader = NsReader::from_reader(detached.as_slice());
    reader.config_mut().trim_text(false);
    let (element, resolver) = loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!(
                "section child namespace resolution failed: {error}"
            ))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() == b"litchiRoot" {
                    continue;
                }
                break (element, reader.resolver().clone());
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "section property has no element".into(),
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    };
    let fragment_prefix = element
        .name()
        .prefix()
        .map(|prefix| prefix.into_inner().to_vec());
    let element_namespace = resolver.resolve_element(element.name()).0;
    let decoder = reader.decoder();
    validate_element_qname(element.name().as_ref(), "section")?;
    validate_attribute_qnames(&element, decoder, "section")?;
    let mut result = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let same_fragment_prefix = matches!(
            &namespace,
            ResolveResult::Unknown(prefix)
                if fragment_prefix.as_deref() == Some(prefix.as_slice())
                    && !raw.has_namespace_binding(prefix.as_slice())
        );
        let relevant = match family {
            AttributeFamily::Word => {
                same_word_namespace(&element_namespace, &namespace)
                    || matches!(namespace, ResolveResult::Unbound)
                    || same_fragment_prefix
            },
            AttributeFamily::Reference if name == "id" => {
                is_relationship_namespace(&namespace)
                    || matches!(&namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r" && !raw.has_namespace_binding(prefix.as_slice()))
            },
            AttributeFamily::Reference => {
                same_word_namespace(&element_namespace, &namespace)
                    || matches!(namespace, ResolveResult::Unbound)
                    || same_fragment_prefix
            },
        };
        if !relevant {
            continue;
        }
        if result.iter().any(|(candidate, _)| candidate == &name) {
            return Err(Error::InvalidFormat(format!(
                "duplicate section property attribute '{name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        result.push((name, value));
    }
    Ok(result)
}

fn is_relationship_namespace(namespace: &ResolveResult<'_>) -> bool {
    const TRANSITIONAL: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
    matches!(namespace, ResolveResult::Bound(quick_xml::name::Namespace(value)) if *value == TRANSITIONAL || *value == STRICT)
}

fn same_word_namespace(
    element_namespace: &ResolveResult<'_>,
    attribute_namespace: &ResolveResult<'_>,
) -> bool {
    matches!(
        (element_namespace, attribute_namespace),
        (
            ResolveResult::Bound(quick_xml::name::Namespace(element)),
            ResolveResult::Bound(quick_xml::name::Namespace(attribute)),
        ) if is_word_namespace_bytes(element) && element == attribute
    )
}

fn attr<'a>(attrs: &'a [(String, String)], name: &str) -> Option<&'a str> {
    attrs
        .iter()
        .find_map(|(candidate, value)| (candidate == name).then_some(value.as_str()))
}

fn required_attr(attrs: &[(String, String)], name: &str) -> Result<String> {
    attr(attrs, name)
        .map(ToOwned::to_owned)
        .ok_or_else(|| Error::InvalidFormat(format!("missing section attribute '{name}'")))
}

fn required_attribute(raw: &Raw, xml: &[u8], name: &[u8]) -> Result<String> {
    let detached = raw.detached_fragment(xml);
    let mut reader = NsReader::from_reader(detached.as_slice());
    reader.config_mut().trim_text(false);
    let (element, decoder, resolver) = loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!(
                "section child namespace resolution failed: {error}"
            ))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() == b"litchiRoot" {
                    continue;
                }
                break (element, reader.decoder(), reader.resolver().clone());
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "section property has no element".into(),
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    };
    let element_prefix = element
        .name()
        .prefix()
        .map(|prefix| prefix.into_inner().to_vec());
    let element_namespace = resolver.resolve_element(element.name()).0;
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let same_fragment_prefix = matches!(
            &namespace,
            ResolveResult::Unknown(prefix)
                if element_prefix.as_deref() == Some(prefix.as_slice())
                    && !raw.has_namespace_binding(prefix.as_slice())
        );
        if !same_word_namespace(&element_namespace, &namespace)
            && !matches!(namespace, ResolveResult::Unbound)
            && !same_fragment_prefix
        {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate section attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    value.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "missing section attribute '{}'",
            String::from_utf8_lossy(name)
        ))
    })
}

fn offset(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))
}

fn element_prefix(element: &BytesStart<'_>) -> Option<Vec<u8>> {
    element
        .name()
        .prefix()
        .map(|prefix| prefix.into_inner().to_vec())
}

fn root_namespace_bindings(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<NamespaceBinding>> {
    let mut bindings = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            Some(None)
        } else {
            name.strip_prefix(b"xmlns:")
                .map(|prefix| Some(prefix.to_vec()))
        };
        let Some(prefix) = prefix else {
            continue;
        };
        let uri = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        bindings.push(NamespaceBinding { prefix, uri });
    }
    Ok(bindings)
}

fn root_attribute_prefix(
    root_prefix: Option<&[u8]>,
    bindings: &[NamespaceBinding],
    word_namespace: Option<&[u8]>,
) -> Option<Vec<u8>> {
    if let Some(prefix) = root_prefix.filter(|prefix| !prefix.is_empty()) {
        return Some(prefix.to_vec());
    }
    let word_namespace = word_namespace?;
    if let Some(prefix) = bindings.iter().find_map(|binding| {
        let prefix = binding.prefix.as_deref()?;
        (binding.uri.as_bytes() == word_namespace).then(|| prefix.to_vec())
    }) {
        return Some(prefix);
    }
    None
}

fn fresh_namespace_prefix(bindings: &[NamespaceBinding]) -> Vec<u8> {
    let mut index = 0usize;
    loop {
        let prefix = if index == 0 {
            b"w".to_vec()
        } else {
            format!("w{index}").into_bytes()
        };
        if !bindings
            .iter()
            .any(|binding| binding.prefix.as_deref() == Some(prefix.as_slice()))
        {
            return prefix;
        }
        index = index.saturating_add(1);
    }
}

fn resolved_word_namespace(namespace: &ResolveResult<'_>) -> Option<Vec<u8>> {
    match namespace {
        ResolveResult::Bound(quick_xml::name::Namespace(namespace))
            if is_word_namespace_bytes(namespace) =>
        {
            Some(namespace.to_vec())
        },
        _ => None,
    }
}

fn is_word_namespace_bytes(value: &[u8]) -> bool {
    matches!(
        value,
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            | b"http://purl.oclc.org/ooxml/wordprocessingml/main"
    )
}

fn escape_namespace_uri(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

fn validate_root_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if matches!(namespace, ResolveResult::Bound(_)) && !is_wordprocessing_namespace(namespace) {
        return Err(Error::InvalidFormat(
            "section properties use a non-WordprocessingML namespace".into(),
        ));
    }
    Ok(())
}

fn is_word_element(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix.as_ref().and_then(|value| value.as_deref()) == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
        ResolveResult::Bound(_) => false,
    }
}

fn is_xml_declaration_or_misc(bytes: &[u8]) -> bool {
    let trimmed = bytes
        .iter()
        .copied()
        .filter(|byte| !byte.is_ascii_whitespace());
    let collected: Vec<u8> = trimmed.collect();
    collected.is_empty()
        || collected.starts_with(b"<?xml")
        || collected.starts_with(b"<!--")
        || collected.starts_with(b"<?")
}

#[cfg(test)]
mod tests {
    use super::{Dirty, Error, Node, decode, encode};

    fn section_xml(children: &str) -> Vec<u8> {
        format!(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{children}</w:sectPr>"#
        )
        .into_bytes()
    }

    #[test]
    fn interleaved_header_footer_references_decode_before_page_size() {
        let xml = section_xml(
            r#"<w:headerReference w:type="default" r:id="rHdDefault"/><w:footerReference w:type="default" r:id="rFtDefault"/><w:headerReference w:type="first" r:id="rHdFirst"/><w:footerReference w:type="first" r:id="rFtFirst"/><w:headerReference w:type="even" r:id="rHdEven"/><w:footerReference w:type="even" r:id="rFtEven"/><w:pgSz w:w="12240" w:h="15840"/>"#,
        );
        let snapshot = decode(&xml).expect("interleaved references are schema-valid");
        assert_eq!(snapshot.state.headers.len(), 3);
        assert_eq!(snapshot.state.footers.len(), 3);
        assert!(snapshot.state.page_size.is_some());
    }

    #[test]
    fn reference_after_page_size_is_rejected() {
        let xml = section_xml(
            r#"<w:pgSz w:w="12240" w:h="15840"/><w:headerReference w:type="default" r:id="rHdDefault"/>"#,
        );
        assert!(matches!(
            decode(&xml),
            Err(Error::InvalidFormat(message)) if message.contains("out of schema order")
        ));
    }

    #[test]
    fn duplicate_same_header_or_footer_type_is_rejected() {
        for children in [
            r#"<w:headerReference w:type="default" r:id="rHdOne"/><w:headerReference w:type="default" r:id="rHdTwo"/>"#,
            r#"<w:footerReference w:type="default" r:id="rFtOne"/><w:footerReference w:type="default" r:id="rFtTwo"/>"#,
        ] {
            assert!(decode(&section_xml(children)).is_err());
        }
    }

    fn snapshot_without(children: &str, removed: &[&str]) -> super::Snapshot {
        let mut snapshot = decode(&section_xml(children)).expect("valid section fixture");
        snapshot.raw.children.retain(|node| {
            !matches!(
                node,
                Node::Element { local_name, .. }
                    if removed.contains(&local_name.as_str())
            )
        });
        snapshot.dirty = Dirty {
            page_size: true,
            margins: true,
            start: true,
            columns: true,
        };
        snapshot
    }

    #[test]
    fn opaque_barriers_stay_before_late_modeled_insertions() {
        let snapshot = snapshot_without(
            r#"<x:foreign xmlns:x="urn:foreign"/><!--opaque--><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/><w:cols w:num="2" w:equalWidth="1"/><w:docGrid/><w:sectPrChange/>"#,
            &["pgSz", "pgMar", "cols"],
        );
        let output = String::from_utf8(encode(&snapshot).expect("encode section")).unwrap();
        assert!(output.contains(r#"<x:foreign xmlns:x="urn:foreign"/>"#));
        assert!(output.contains("<!--opaque-->"));
        let foreign = output.find("<x:foreign").expect("foreign child");
        let comment = output.find("<!--opaque-->").expect("opaque comment");
        let page_size = output.find("<w:pgSz").expect("inserted page size");
        let margins = output.find("<w:pgMar").expect("inserted margins");
        let columns = output.find("<w:cols").expect("inserted columns");
        let doc_grid = output.find("<w:docGrid").expect("doc grid");
        let section_change = output.find("<w:sectPrChange").expect("section change");
        assert!(foreign < comment);
        assert!(comment < page_size);
        assert!(page_size < margins);
        assert!(margins < columns);
        assert!(columns < doc_grid);
        assert!(doc_grid < section_change);
    }

    #[test]
    fn late_schema_barriers_preserve_pg_borders_no_endnote_and_change_order() {
        let snapshot = snapshot_without(
            r#"<w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/><w:pgBorders/><w:cols w:num="2" w:equalWidth="1"/><w:noEndnote/><w:sectPrChange/>"#,
            &["type", "pgSz", "pgMar", "cols"],
        );
        let output = String::from_utf8(encode(&snapshot).expect("encode section")).unwrap();
        let borders = output.find("<w:pgBorders").expect("page borders");
        let page_size = output.find("<w:pgSz").expect("inserted page size");
        let margins = output.find("<w:pgMar").expect("inserted margins");
        let columns = output.find("<w:cols").expect("inserted columns");
        let no_endnote = output.find("<w:noEndnote").expect("noEndnote");
        let section_change = output.find("<w:sectPrChange").expect("section change");
        assert!(page_size < margins);
        assert!(margins < borders);
        assert!(borders < columns);
        assert!(columns < no_endnote);
        assert!(no_endnote < section_change);
    }

    #[test]
    fn only_opaque_children_receive_all_modeled_fields_at_root_close() {
        let snapshot = snapshot_without(
            r#"<x:foreign xmlns:x="urn:foreign"/><!--opaque--><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/><w:cols w:num="2" w:equalWidth="1"/>"#,
            &["type", "pgSz", "pgMar", "cols"],
        );
        let output = String::from_utf8(encode(&snapshot).expect("encode section")).unwrap();
        assert!(output.contains(r#"<x:foreign xmlns:x="urn:foreign"/>"#));
        assert!(output.contains("<!--opaque-->"));
        let foreign_end = output.find("/><!--opaque-->").expect("opaque order");
        let type_start = output.find("<w:type").expect("inserted type");
        let page_size = output.find("<w:pgSz").expect("inserted page size");
        let margins = output.find("<w:pgMar").expect("inserted margins");
        let columns = output.find("<w:cols").expect("inserted columns");
        assert!(foreign_end < type_start);
        assert!(type_start < page_size);
        assert!(page_size < margins);
        assert!(margins < columns);
    }

    #[test]
    fn reversed_no_endnote_and_section_change_are_rejected() {
        let xml = section_xml(r#"<w:sectPrChange/><w:noEndnote/>"#);
        assert!(matches!(
            decode(&xml),
            Err(Error::InvalidFormat(message)) if message.contains("out of schema order")
        ));
    }

    fn dirty_page_size(xml: &str) -> Result<Vec<u8>, Error> {
        let mut snapshot = decode(xml.as_bytes())?;
        snapshot.dirty.page_size = true;
        encode(&snapshot)
    }

    #[test]
    fn generated_qnames_use_named_root_word_prefix() {
        let output = String::from_utf8(dirty_page_size(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
        )
        .expect("named root prefix encodes"))
        .unwrap();
        assert!(output.contains("<w:pgSz w:w=\""));
    }

    #[test]
    fn generated_qnames_use_named_alias_with_default_element_namespace() {
        let output = String::from_utf8(dirty_page_size(
            r#"<alias:sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:alias="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><pgSz w="12240" h="15840"/></alias:sectPr>"#,
        )
        .expect("named alias encodes"))
        .unwrap();
        assert!(output.contains("alias:w=\""));
        assert!(!output.contains("r:w=\""));
        assert!(!output.contains("x:w=\""));
    }

    #[test]
    fn child_relationship_or_foreign_attributes_cannot_hijack_generated_prefix() {
        let output = String::from_utf8(dirty_page_size(
            r#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:foreign"><x:foreign r:id="r1" x:value="v"/><pgSz w="12240" h="15840"/></sectPr>"#,
        )
        .expect("root binding controls generated prefix"))
        .unwrap();
        assert!(output.contains("w:w=\""));
        assert!(!output.contains("r:w=\""));
        assert!(!output.contains("x:w=\""));
    }

    #[test]
    fn explicit_root_bindings_decode_child_word_and_relationship_attributes_exactly() {
        let xml = br#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><headerReference w:type="default" r:id="rHeader"/><pgSz w:w="12240" w:h="15840"/></sectPr>"#;
        let snapshot = decode(xml).expect("explicit root bindings decode");
        assert!(snapshot.state.page_size.is_some());
        assert_eq!(snapshot.state.headers[0].relationship_id, "rHeader");
        assert_eq!(encode(&snapshot).expect("unchanged encode"), xml);
    }

    #[test]
    fn detached_lexical_prefix_fallback_remains_byte_exact() {
        let xml = br#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#;
        let snapshot = decode(xml).expect("detached lexical prefix decodes");
        assert_eq!(encode(&snapshot).expect("unchanged encode"), xml);
    }

    #[test]
    fn explicit_foreign_relationship_binding_cannot_satisfy_reference_id() {
        let xml = br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="urn:foreign"><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>"#;
        assert!(matches!(
            decode(xml),
            Err(Error::InvalidFormat(message))
                if message.contains("missing section attribute 'id'")
        ));
    }

    #[test]
    fn default_only_page_size_deletion_preserves_margin_bytes() {
        let xml = br#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><pgSz w="12240" h="15840"/><pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></sectPr>"#;
        let mut snapshot = decode(xml).expect("default namespace section decodes");
        snapshot.state.page_size = None;
        snapshot.dirty.page_size = true;
        let output = encode(&snapshot).expect("page size deletion does not need attributes");
        assert!(
            !output
                .windows(b"<pgSz".len())
                .any(|window| window == b"<pgSz")
        );
        assert!(
            output
                .windows(
                    b"<pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\"/>".len()
                )
                .any(|window| {
                    window
                        == b"<pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\"/>"
                })
        );
    }

    #[test]
    fn named_prefix_page_size_addition_precedes_existing_margins() {
        let xml = section_xml(
            r#"<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>"#,
        );
        let mut snapshot = decode(&xml).expect("named section decodes");
        snapshot.raw.children.retain(|node| {
            !matches!(
                node,
                Node::Element { local_name, .. } if local_name == "pgSz"
            )
        });
        snapshot.dirty.page_size = true;
        let output = String::from_utf8(encode(&snapshot).expect("page size insertion")).unwrap();
        let page_size = output.find("<w:pgSz").expect("inserted page size");
        let margins = output.find("<w:pgMar").expect("existing margins");
        assert!(page_size < margins);
    }
}
