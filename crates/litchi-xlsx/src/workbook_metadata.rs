//! Typed, bounded SpreadsheetML workbook future-metadata codec.
//!
//! The owner implements the `metadata`, `metadataTypes`, `futureMetadata`,
//! `cellMetadata`, and `valueMetadata` grammar from [MS-XLSX] §2.2.4.4
//! and the referenced ISO metadata structures. Package relationship and
//! content-type discovery remains in the OOXML host adapter.

use crate::error::{Error, Result, invalid};
use litchi_ooxml_common::XmlError;
use litchi_ooxml_common::mce::{MceCapabilities, MceLimits, process_markup_compatibility};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::{NsReader, XmlVersion};
use std::collections::HashSet;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SML_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const XDA: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
const XLRD: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
const MAX_XML: usize = 16 * 1024 * 1024;
const MAX_OUTPUT: usize = 32 * 1024 * 1024;
const MAX_EXTENSION_XML: usize = 8 * 1024 * 1024;
const MAX_TYPES: usize = 65_536;
const MAX_FUTURE: usize = 65_536;
const MAX_BLOCKS: usize = 1_000_000;
const MAX_RECORDS: usize = 1_000_000;
const MAX_STRING: usize = 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 2_000_000;
const OFFICE_MAX_COUNT: u32 = 2_147_483_647;

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MetadataBehavior {
    pub ghost_row: bool,
    pub ghost_column: bool,
    pub edit: bool,
    pub delete: bool,
    pub copy: bool,
    pub paste_all: bool,
    pub paste_formulas: bool,
    pub paste_values: bool,
    pub paste_formats: bool,
    pub paste_comments: bool,
    pub paste_data_validation: bool,
    pub paste_borders: bool,
    pub paste_column_widths: bool,
    pub paste_number_formats: bool,
    pub merge: bool,
    pub split_first: bool,
    pub split_all: bool,
    pub row_column_shift: bool,
    pub clear_all: bool,
    pub clear_formats: bool,
    pub clear_contents: bool,
    pub clear_comments: bool,
    pub assign: bool,
    pub coerce: bool,
    pub cell_metadata: bool,
    pub adjust: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MetadataType {
    pub name: String,
    pub minimum_supported_version: u32,
    pub behavior: MetadataBehavior,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MetadataRecord {
    /// One-based index into `WorkbookMetadata::types`.
    pub type_index: u32,
    /// Zero-based index into the matching future-metadata store.
    pub value_index: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueMetadataExtension {
    pub uri: String,
    /// Deterministically normalized, inert child XML from the extension.
    pub payload_xml: Vec<u8>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MetadataBlock {
    pub records: Vec<MetadataRecord>,
    pub extensions: Vec<OpaqueMetadataExtension>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FutureMetadata {
    pub name: String,
    pub blocks: Vec<MetadataBlock>,
    pub extensions: Vec<OpaqueMetadataExtension>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorkbookMetadata {
    pub types: Vec<MetadataType>,
    pub future: Vec<FutureMetadata>,
    pub cell_blocks: Vec<MetadataBlock>,
    pub value_blocks: Vec<MetadataBlock>,
    pub extensions: Vec<OpaqueMetadataExtension>,
}

impl WorkbookMetadata {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML {
            return Err(limit("metadata XML bytes"));
        }
        let root = parse_mce_dom(xml)?;
        if root.local != "metadata" || !is_sml(&root.ns) {
            return Err(invalid("expected SpreadsheetML metadata root"));
        }
        attrs(&root, &[])?;
        let mut position = 0usize;
        let types_node = root
            .children
            .get(position)
            .ok_or_else(|| invalid("metadataTypes is required"))?;
        if !core(types_node, "metadataTypes") {
            return Err(invalid("metadataTypes must be first"));
        }
        position += 1;
        let types = parse_types(types_node)?;
        if root
            .children
            .get(position)
            .is_some_and(|n| core(n, "metadataStrings"))
        {
            return Err(invalid(
                "metadataStrings is outside the future-metadata subset",
            ));
        }
        if root
            .children
            .get(position)
            .is_some_and(|n| core(n, "mdxMetadata"))
        {
            return Err(invalid(
                "MDX metadata is outside the future-metadata subset",
            ));
        }
        let mut future = Vec::new();
        while root
            .children
            .get(position)
            .is_some_and(|n| core(n, "futureMetadata"))
        {
            if future.len() >= MAX_FUTURE {
                return Err(limit("future metadata stores"));
            }
            future.push(parse_future(&root.children[position])?);
            position += 1;
        }
        let mut cell_blocks = Vec::new();
        if root
            .children
            .get(position)
            .is_some_and(|n| core(n, "cellMetadata"))
        {
            cell_blocks = parse_block_store(&root.children[position], "cellMetadata")?;
            position += 1;
        }
        let mut value_blocks = Vec::new();
        if root
            .children
            .get(position)
            .is_some_and(|n| core(n, "valueMetadata"))
        {
            value_blocks = parse_block_store(&root.children[position], "valueMetadata")?;
            position += 1;
        }
        let mut extensions = Vec::new();
        if root
            .children
            .get(position)
            .is_some_and(|n| core(n, "extLst"))
        {
            extensions = parse_extensions(&root.children[position])?;
            position += 1;
        }
        if position != root.children.len() {
            return Err(invalid("unexpected or out-of-order metadata child"));
        }
        let value = Self {
            types,
            future,
            cell_blocks,
            value_blocks,
            extensions,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        self.validate()?;
        let ns = if strict { SML_STRICT } else { SML };
        let mut out = String::with_capacity(1024);
        out.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><metadata xmlns=\"",
        );
        out.push_str(ns);
        out.push_str("\"><metadataTypes");
        count_attr(&mut out, self.types.len());
        out.push('>');
        for ty in &self.types {
            write_type(&mut out, ty);
        }
        out.push_str("</metadataTypes>");
        for item in &self.future {
            out.push_str("<futureMetadata");
            qattr(&mut out, "name", &item.name);
            count_attr(&mut out, item.blocks.len());
            out.push('>');
            for block in &item.blocks {
                write_block(&mut out, block);
            }
            write_extensions(&mut out, &item.extensions);
            out.push_str("</futureMetadata>");
        }
        if !self.cell_blocks.is_empty() {
            out.push_str("<cellMetadata");
            count_attr(&mut out, self.cell_blocks.len());
            out.push('>');
            for block in &self.cell_blocks {
                write_block(&mut out, block);
            }
            out.push_str("</cellMetadata>");
        }
        if !self.value_blocks.is_empty() {
            out.push_str("<valueMetadata");
            count_attr(&mut out, self.value_blocks.len());
            out.push('>');
            for block in &self.value_blocks {
                write_block(&mut out, block);
            }
            out.push_str("</valueMetadata>");
        }
        write_extensions(&mut out, &self.extensions);
        out.push_str("</metadata>");
        if out.len() > MAX_OUTPUT {
            return Err(limit("canonical metadata XML bytes"));
        }
        Ok(out.into_bytes())
    }

    pub fn cell_block(&self, one_based_index: u32) -> Option<&MetadataBlock> {
        one_based_index
            .checked_sub(1)
            .and_then(|v| self.cell_blocks.get(v as usize))
    }
    pub fn value_block(&self, one_based_index: u32) -> Option<&MetadataBlock> {
        one_based_index
            .checked_sub(1)
            .and_then(|v| self.value_blocks.get(v as usize))
    }

    fn validate(&self) -> Result<()> {
        if self.types.is_empty() || self.types.len() > MAX_TYPES {
            return Err(invalid("metadata requires 1..65536 metadata types"));
        }
        if self.future.len() > MAX_FUTURE
            || self.cell_blocks.len() > MAX_BLOCKS
            || self.value_blocks.len() > MAX_BLOCKS
        {
            return Err(limit("metadata collections"));
        }
        let mut type_names = HashSet::new();
        for ty in &self.types {
            bounded_nonempty(&ty.name, "metadata type name")?;
            if !type_names.insert(ty.name.as_str()) {
                return Err(invalid("duplicate metadata type name"));
            }
            let paste = &ty.behavior;
            if !paste.copy
                && (paste.paste_all
                    || paste.paste_formulas
                    || paste.paste_values
                    || paste.paste_formats
                    || paste.paste_comments
                    || paste.paste_data_validation
                    || paste.paste_borders
                    || paste.paste_column_widths
                    || paste.paste_number_formats)
            {
                return Err(invalid("metadata paste behavior requires copy=true"));
            }
        }
        let mut future_names = HashSet::new();
        for item in &self.future {
            bounded_nonempty(&item.name, "future metadata name")?;
            if item.name == "XLMDX" {
                return Err(invalid("XLMDX requires the excluded MDX grammar"));
            }
            if !type_names.contains(item.name.as_str()) {
                return Err(invalid(
                    "futureMetadata name does not identify a metadataType",
                ));
            }
            if !future_names.insert(item.name.as_str()) {
                return Err(invalid("duplicate futureMetadata name"));
            }
            if item.blocks.len() > OFFICE_MAX_COUNT as usize {
                return Err(limit("future metadata blocks"));
            }
            validate_extensions(&item.extensions)?;
            for block in &item.blocks {
                validate_block(block, &self.types, &self.future, None)?;
            }
        }
        for block in &self.cell_blocks {
            validate_block(block, &self.types, &self.future, Some(true))?;
        }
        for block in &self.value_blocks {
            validate_block(block, &self.types, &self.future, Some(false))?;
        }
        validate_extensions(&self.extensions)
    }
}

fn parse_types(node: &Node) -> Result<Vec<MetadataType>> {
    attrs(node, &[("", "count")])?;
    check_count(node, node.children.len())?;
    if node.children.is_empty() || node.children.len() > MAX_TYPES {
        return Err(invalid(
            "metadataTypes requires 1..65536 metadataType children",
        ));
    }
    node.children.iter().map(parse_type).collect()
}

const BOOL_ATTRS: &[&str] = &[
    "ghostRow",
    "ghostCol",
    "edit",
    "delete",
    "copy",
    "pasteAll",
    "pasteFormulas",
    "pasteValues",
    "pasteFormats",
    "pasteComments",
    "pasteDataValidation",
    "pasteBorders",
    "pasteColWidths",
    "pasteNumberFormats",
    "merge",
    "splitFirst",
    "splitAll",
    "rowColShift",
    "clearAll",
    "clearFormats",
    "clearContents",
    "clearComments",
    "assign",
    "coerce",
    "cellMeta",
    "adjust",
];
fn parse_type(node: &Node) -> Result<MetadataType> {
    if !core(node, "metadataType") || !node.children.is_empty() {
        return Err(invalid("metadataTypes contains invalid child"));
    }
    let mut allowed = vec![("", "name"), ("", "minSupportedVersion")];
    for name in BOOL_ATTRS {
        allowed.push(("", *name));
    }
    attrs(node, &allowed)?;
    let name = required(node, "name")?;
    let minimum_supported_version = required_u32(node, "minSupportedVersion")?;
    let b = |name| boolean(node, name).map(|v| v.unwrap_or(false));
    Ok(MetadataType {
        name,
        minimum_supported_version,
        behavior: MetadataBehavior {
            ghost_row: b("ghostRow")?,
            ghost_column: b("ghostCol")?,
            edit: b("edit")?,
            delete: b("delete")?,
            copy: b("copy")?,
            paste_all: b("pasteAll")?,
            paste_formulas: b("pasteFormulas")?,
            paste_values: b("pasteValues")?,
            paste_formats: b("pasteFormats")?,
            paste_comments: b("pasteComments")?,
            paste_data_validation: b("pasteDataValidation")?,
            paste_borders: b("pasteBorders")?,
            paste_column_widths: b("pasteColWidths")?,
            paste_number_formats: b("pasteNumberFormats")?,
            merge: b("merge")?,
            split_first: b("splitFirst")?,
            split_all: b("splitAll")?,
            row_column_shift: b("rowColShift")?,
            clear_all: b("clearAll")?,
            clear_formats: b("clearFormats")?,
            clear_contents: b("clearContents")?,
            clear_comments: b("clearComments")?,
            assign: b("assign")?,
            coerce: b("coerce")?,
            cell_metadata: b("cellMeta")?,
            adjust: b("adjust")?,
        },
    })
}

fn parse_future(node: &Node) -> Result<FutureMetadata> {
    attrs(node, &[("", "name"), ("", "count")])?;
    let name = required(node, "name")?;
    let mut blocks = Vec::new();
    let mut extensions = Vec::new();
    let mut ext_seen = false;
    for child in &node.children {
        if core(child, "bk") && !ext_seen {
            if blocks.len() >= MAX_BLOCKS {
                return Err(limit("future metadata blocks"));
            }
            blocks.push(parse_block(child)?);
        } else if core(child, "extLst") && !ext_seen {
            ext_seen = true;
            extensions = parse_extensions(child)?;
        } else {
            return Err(invalid("unexpected or out-of-order futureMetadata child"));
        }
    }
    check_count(node, blocks.len())?;
    Ok(FutureMetadata {
        name,
        blocks,
        extensions,
    })
}
fn parse_block_store(node: &Node, expected: &str) -> Result<Vec<MetadataBlock>> {
    if !core(node, expected) {
        return Err(invalid("invalid metadata block store"));
    }
    attrs(node, &[("", "count")])?;
    check_count(node, node.children.len())?;
    if node.children.len() > MAX_BLOCKS {
        return Err(limit("metadata blocks"));
    }
    node.children.iter().map(parse_block).collect()
}
fn parse_block(node: &Node) -> Result<MetadataBlock> {
    if !core(node, "bk") {
        return Err(invalid("metadata block must be bk"));
    }
    attrs(node, &[])?;
    let mut records = Vec::new();
    let mut extensions = Vec::new();
    let mut ext_seen = false;
    for child in &node.children {
        if core(child, "rc") && !ext_seen {
            if !child.children.is_empty() {
                return Err(invalid("metadata record must be empty"));
            }
            attrs(child, &[("", "t"), ("", "v")])?;
            if records.len() >= MAX_RECORDS {
                return Err(limit("metadata records"));
            }
            records.push(MetadataRecord {
                type_index: required_u32(child, "t")?,
                value_index: required_u32(child, "v")?,
            });
        } else if core(child, "extLst") && !ext_seen {
            ext_seen = true;
            extensions = parse_extensions(child)?;
        } else {
            return Err(invalid("unexpected or out-of-order metadata block child"));
        }
    }
    Ok(MetadataBlock {
        records,
        extensions,
    })
}
fn parse_extensions(node: &Node) -> Result<Vec<OpaqueMetadataExtension>> {
    if !core(node, "extLst") {
        return Err(invalid("expected extLst"));
    }
    attrs(node, &[])?;
    let mut total = 0usize;
    let mut result = Vec::with_capacity(node.children.len());
    for ext in &node.children {
        if !core(ext, "ext") {
            return Err(invalid("extLst may contain only ext"));
        }
        attrs(ext, &[("", "uri")])?;
        let uri = required(ext, "uri")?;
        bounded_nonempty(&uri, "extension URI")?;
        let mut payload = String::new();
        for child in &ext.children {
            write_opaque(&mut payload, child);
        }
        total = total
            .checked_add(payload.len())
            .ok_or_else(|| limit("extension XML"))?;
        if total > MAX_EXTENSION_XML {
            return Err(limit("extension XML"));
        }
        result.push(OpaqueMetadataExtension {
            uri,
            payload_xml: payload.into_bytes(),
        });
    }
    Ok(result)
}

fn validate_block(
    block: &MetadataBlock,
    types: &[MetadataType],
    future: &[FutureMetadata],
    cell: Option<bool>,
) -> Result<()> {
    if block.records.len() > MAX_RECORDS {
        return Err(limit("metadata records"));
    }
    validate_extensions(&block.extensions)?;
    for record in &block.records {
        if record.type_index == 0 || record.type_index as usize > types.len() {
            return Err(invalid("metadata record type index is out of range"));
        }
        let ty = &types[record.type_index as usize - 1];
        if let Some(cell) = cell
            && ty.behavior.cell_metadata != cell
        {
            return Err(invalid(
                "metadata record type does not match cell/value store",
            ));
        }
        let store = future
            .iter()
            .find(|v| v.name == ty.name)
            .ok_or_else(|| invalid("metadata record type has no futureMetadata store"))?;
        if record.value_index as usize >= store.blocks.len() {
            return Err(invalid("metadata record value index is out of range"));
        }
    }
    Ok(())
}
fn validate_extensions(values: &[OpaqueMetadataExtension]) -> Result<()> {
    let mut total = 0usize;
    for value in values {
        bounded_nonempty(&value.uri, "extension URI")?;
        total = total
            .checked_add(value.payload_xml.len())
            .ok_or_else(|| limit("extension XML"))?;
        if total > MAX_EXTENSION_XML {
            return Err(limit("extension XML"));
        }
        std::str::from_utf8(&value.payload_xml).map_err(xml_error)?;
    }
    Ok(())
}

#[derive(Debug, Clone)]
struct Attribute {
    ns: String,
    local: String,
    value: String,
}
#[derive(Debug, Clone)]
struct Node {
    ns: String,
    local: String,
    attrs: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}
fn parse_mce_dom(xml: &[u8]) -> Result<Node> {
    let mut capabilities = MceCapabilities::ooxml_baseline();
    capabilities
        .understand_namespace(XDA)
        .understand_namespace(XLRD);
    let limits = MceLimits {
        max_input_bytes: MAX_XML,
        max_output_bytes: MAX_OUTPUT,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &limits)?;
    parse_dom(processed.xml.as_ref())
}
fn parse_dom(xml: &[u8]) -> Result<Node> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text = 0usize;
    loop {
        let decoder = reader.decoder();
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(e) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                let ns = namespace(&resolved)?;
                drop(resolved);
                let resolver = reader.resolver().clone();
                stack.push(node(&resolver, ns, &e, decoder)?);
            },
            Event::Empty(e) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                let ns = namespace(&resolved)?;
                drop(resolved);
                let resolver = reader.resolver().clone();
                append(&mut stack, &mut root, node(&resolver, ns, &e, decoder)?)?;
            },
            Event::End(_) => {
                let value = stack.pop().ok_or_else(|| invalid("unexpected XML end"))?;
                append(&mut stack, &mut root, value)?;
            },
            Event::Text(e) => {
                let decoded = e.xml_content(XmlVersion::Explicit1_0).map_err(xml_error)?;
                let value = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                push_text(&mut stack, &mut text, &value)?;
            },
            Event::CData(e) => {
                let value = e.xml_content(XmlVersion::Explicit1_0).map_err(xml_error)?;
                push_text(&mut stack, &mut text, &value)?;
            },
            Event::GeneralRef(e) => {
                let name = e.decode().map_err(xml_error)?;
                let value = if let Some(c) = e.resolve_char_ref().map_err(xml_error)? {
                    c.to_string()
                } else {
                    match name.as_ref() {
                        "amp" => "&",
                        "lt" => "<",
                        "gt" => ">",
                        "apos" => "'",
                        "quot" => "\"",
                        _ => return Err(invalid("custom entity is rejected")),
                    }
                    .into()
                };
                push_text(&mut stack, &mut text, &value)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}
fn node(
    resolver: &NamespaceResolver,
    ns: String,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Node> {
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .into();
    let mut values = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        if item.key.as_ref() == b"xmlns" || item.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, _) = resolver.resolve_attribute(item.key);
        let ans = namespace(&resolved)?;
        let alocal = std::str::from_utf8(item.key.local_name().as_ref())
            .map_err(xml_error)?
            .into();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if value.len() > MAX_STRING {
            return Err(limit("attribute string"));
        }
        if values
            .iter()
            .any(|a: &Attribute| a.ns == ans && a.local == alocal)
        {
            return Err(invalid("duplicate expanded attribute"));
        }
        values.push(Attribute {
            ns: ans,
            local: alocal,
            value,
        });
    }
    Ok(Node {
        ns,
        local,
        attrs: values,
        children: Vec::new(),
        text: String::new(),
    })
}
fn namespace(value: &ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(v)) => std::str::from_utf8(v)
            .map(str::to_string)
            .map_err(xml_error),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(p) => Err(invalid(format!(
            "unbound namespace prefix {}",
            String::from_utf8_lossy(p)
        ))),
    }
}
fn append(stack: &mut [Node], root: &mut Option<Node>, value: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(value);
    } else if root.replace(value).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
fn push_text(stack: &mut [Node], total: &mut usize, value: &str) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| limit("XML text"))?;
    if *total > MAX_OUTPUT {
        return Err(limit("XML text"));
    }
    if let Some(parent) = stack.last_mut() {
        parent.text.push_str(value);
    } else if !value.trim().is_empty() {
        return Err(invalid("text outside XML root"));
    }
    Ok(())
}

fn write_type(out: &mut String, value: &MetadataType) {
    out.push_str("<metadataType");
    qattr(out, "name", &value.name);
    qattr(
        out,
        "minSupportedVersion",
        &value.minimum_supported_version.to_string(),
    );
    let b = &value.behavior;
    for (name, value) in [
        ("ghostRow", b.ghost_row),
        ("ghostCol", b.ghost_column),
        ("edit", b.edit),
        ("delete", b.delete),
        ("copy", b.copy),
        ("pasteAll", b.paste_all),
        ("pasteFormulas", b.paste_formulas),
        ("pasteValues", b.paste_values),
        ("pasteFormats", b.paste_formats),
        ("pasteComments", b.paste_comments),
        ("pasteDataValidation", b.paste_data_validation),
        ("pasteBorders", b.paste_borders),
        ("pasteColWidths", b.paste_column_widths),
        ("pasteNumberFormats", b.paste_number_formats),
        ("merge", b.merge),
        ("splitFirst", b.split_first),
        ("splitAll", b.split_all),
        ("rowColShift", b.row_column_shift),
        ("clearAll", b.clear_all),
        ("clearFormats", b.clear_formats),
        ("clearContents", b.clear_contents),
        ("clearComments", b.clear_comments),
        ("assign", b.assign),
        ("coerce", b.coerce),
        ("cellMeta", b.cell_metadata),
        ("adjust", b.adjust),
    ] {
        if value {
            qattr(out, name, "1");
        }
    }
    out.push_str("/>");
}
fn write_block(out: &mut String, value: &MetadataBlock) {
    out.push_str("<bk>");
    for record in &value.records {
        out.push_str("<rc");
        qattr(out, "t", &record.type_index.to_string());
        qattr(out, "v", &record.value_index.to_string());
        out.push_str("/>");
    }
    write_extensions(out, &value.extensions);
    out.push_str("</bk>");
}
fn write_extensions(out: &mut String, values: &[OpaqueMetadataExtension]) {
    if values.is_empty() {
        return;
    }
    out.push_str("<extLst>");
    for value in values {
        out.push_str("<ext");
        qattr(out, "uri", &value.uri);
        out.push('>');
        out.push_str(std::str::from_utf8(&value.payload_xml).expect("validated UTF-8"));
        out.push_str("</ext>");
    }
    out.push_str("</extLst>");
}
fn write_opaque(out: &mut String, node: &Node) {
    let prefixed = !node.ns.is_empty();
    out.push('<');
    if prefixed {
        out.push_str("p:");
    }
    out.push_str(&node.local);
    if prefixed {
        qattr(out, "xmlns:p", &node.ns);
    }
    let mut index = 0usize;
    for a in &node.attrs {
        if a.ns.is_empty() {
            qattr(out, &a.local, &a.value);
        } else {
            let prefix = format!("a{index}");
            index += 1;
            qattr(out, &format!("xmlns:{prefix}"), &a.ns);
            qattr(out, &format!("{prefix}:{}", a.local), &a.value);
        }
    }
    if node.children.is_empty() && node.text.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    escape_text(out, &node.text);
    for child in &node.children {
        write_opaque(out, child);
    }
    out.push_str("</");
    if prefixed {
        out.push_str("p:");
    }
    out.push_str(&node.local);
    out.push('>');
}
fn count_attr(out: &mut String, count: usize) {
    qattr(out, "count", &count.to_string())
}
fn qattr(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    escape_attr(out, value);
    out.push('"')
}
fn escape_attr(out: &mut String, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            '\r' => out.push_str("&#xD;"),
            '\n' => out.push_str("&#xA;"),
            '\t' => out.push_str("&#x9;"),
            _ => out.push(c),
        }
    }
}
fn escape_text(out: &mut String, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(c),
        }
    }
}

fn attrs(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for a in &node.attrs {
        if !allowed.iter().any(|(ns, n)| *ns == a.ns && *n == a.local) {
            return Err(invalid(format!(
                "unexpected attribute {{{}}}{} on {}",
                a.ns, a.local, node.local
            )));
        }
    }
    Ok(())
}
fn attribute<'a>(node: &'a Node, name: &str) -> Option<&'a str> {
    node.attrs
        .iter()
        .find(|a| a.ns.is_empty() && a.local == name)
        .map(|a| a.value.as_str())
}
fn required(node: &Node, name: &str) -> Result<String> {
    attribute(node, name)
        .filter(|v| !v.is_empty())
        .map(str::to_string)
        .ok_or_else(|| invalid(format!("{} requires {name}", node.local)))
}
fn required_u32(node: &Node, name: &str) -> Result<u32> {
    required(node, name)?
        .parse()
        .map_err(|_| invalid(format!("invalid unsigned integer {name}")))
}
fn boolean(node: &Node, name: &str) -> Result<Option<bool>> {
    attribute(node, name)
        .map(|v| match v {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid boolean {name}"))),
        })
        .transpose()
}
fn check_count(node: &Node, actual: usize) -> Result<()> {
    let declared = required_u32(node, "count")?;
    if declared > OFFICE_MAX_COUNT || declared as usize != actual {
        Err(invalid(format!("{} count mismatch", node.local)))
    } else {
        Ok(())
    }
}
fn core(node: &Node, name: &str) -> bool {
    node.local == name && is_sml(&node.ns)
}
fn is_sml(ns: &str) -> bool {
    matches!(ns, SML | SML_STRICT)
}
fn bounded_nonempty(value: &str, what: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_STRING {
        Err(limit(what))
    } else {
        Ok(())
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(XmlError::Malformed(error.to_string()))
}

fn limit(value: impl Into<String>) -> Error {
    Error::Invalid(format!(
        "workbook metadata resource limit exceeded: {}",
        value.into()
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> WorkbookMetadata {
        WorkbookMetadata {
            types: vec![MetadataType {
                name: "XLDAPR".into(),
                minimum_supported_version: 120000,
                behavior: MetadataBehavior {
                    copy: true,
                    paste_all: true,
                    paste_values: true,
                    cell_metadata: true,
                    ..Default::default()
                },
            }],
            future: vec![FutureMetadata {
                name: "XLDAPR".into(),
                blocks: vec![MetadataBlock {
                    records: Vec::new(),
                    extensions: vec![OpaqueMetadataExtension {
                        uri: "u".into(),
                        payload_xml: format!(r#"<p:x xmlns:p="{XDA}" a="1"/>"#).into_bytes(),
                    }],
                }],
                extensions: Vec::new(),
            }],
            cell_blocks: vec![MetadataBlock {
                records: vec![MetadataRecord {
                    type_index: 1,
                    value_index: 0,
                }],
                extensions: Vec::new(),
            }],
            value_blocks: Vec::new(),
            extensions: Vec::new(),
        }
    }

    #[test]
    fn strict_round_trip_preserves_indices_and_extensions() {
        let value = sample();
        let xml = value.to_xml(true).unwrap();
        let parsed = WorkbookMetadata::parse(&xml).unwrap();
        assert_eq!(parsed.cell_block(1).unwrap().records[0].type_index, 1);
        assert!(parsed.cell_block(0).is_none());
        assert_eq!(parsed.to_xml(true).unwrap(), xml);
    }

    #[test]
    fn mce_choice_selects_understood_metadata_branch() {
        let body = String::from_utf8(sample().to_xml(false).unwrap()).unwrap();
        let body = body
            .replace(
                "<metadataTypes",
                r#"<mc:AlternateContent><mc:Choice Requires="xda"><metadataTypes"#,
            )
            .replace(
                "</metadataTypes>",
                "</metadataTypes></mc:Choice><mc:Fallback/></mc:AlternateContent>",
            );
        let xml = body.replace(
            r#"<metadata xmlns=""#,
            &format!(
                r#"<metadata xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xda="{XDA}" xmlns=""#
            ),
        );
        assert_eq!(
            WorkbookMetadata::parse(xml.as_bytes()).unwrap().types.len(),
            1
        );
    }

    #[test]
    fn rejects_malformed_and_out_of_bounds_values() {
        assert!(WorkbookMetadata::parse(br#"<!DOCTYPE x><metadata/>"#).is_err());
        assert!(WorkbookMetadata::parse(
            format!(
                r#"<metadata xmlns="{SML}"><metadataTypes count="2"><metadataType name="x" minSupportedVersion="1"/></metadataTypes></metadata>"#
            )
            .as_bytes(),
        )
        .is_err());

        let mut value = sample();
        value.cell_blocks[0].records[0].type_index = 2;
        assert!(value.to_xml(false).is_err());

        let mut value = sample();
        value.types[0].name = "x".repeat(MAX_STRING + 1);
        assert!(value.to_xml(false).is_err());
    }
}
