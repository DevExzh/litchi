use super::parser::{Parser, Section, SectionDisplay};
use base64::Engine as _;
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::ops::Range;

const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_VALUE: usize = 65_536;

/// Stable whole-block location used to wrap existing content in a section.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Block {
    BodyParagraph(usize),
    BodyTable(usize),
    TableCellParagraph {
        table: usize,
        row: usize,
        cell: usize,
        paragraph: usize,
    },
}

impl Section {
    /// Validate authorable section metadata and inert source declarations.
    pub fn validate(&self) -> Result<()> {
        validate_text(&self.name, "section name", false)?;
        if let Some(value) = &self.style {
            validate_text(value, "section style", false)?;
        }
        if let Some(value) = &self.xml_id {
            validate_ncname(value, "section xml:id")?;
        }
        if self.content.len() > 16 * 1024 * 1024 {
            return invalid("section content exceeds 16 MiB");
        }
        match self.display {
            SectionDisplay::Condition if self.condition.is_none() => {
                return invalid("conditional section requires a condition");
            },
            SectionDisplay::Visible | SectionDisplay::Hidden if self.condition.is_some() => {
                return invalid("section condition requires conditional display");
            },
            _ => {},
        }
        if let Some(value) = &self.condition {
            validate_text(value, "section condition", false)?;
        }
        match (&self.protection_key, &self.protection_key_digest_algorithm) {
            (None, Some(_)) => return invalid("section protection digest requires a key"),
            (Some(key), digest) => {
                if !self.protected {
                    return invalid("section protection key requires protected=true");
                }
                validate_text(key, "section protection key", false)?;
                base64::engine::general_purpose::STANDARD
                    .decode(key)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid section protection key: {error}"))
                    })?;
                if let Some(value) = digest {
                    validate_uri(value, "section digest algorithm")?;
                }
            },
            _ => {},
        }
        if self.source.is_some() && self.dde_source.is_some() {
            return invalid("section may have only one source declaration");
        }
        if let Some(source) = &self.source {
            if let Some(value) = &source.href {
                validate_uri(value, "section source href")?;
            }
            if let Some(value) = &source.section_name {
                validate_text(value, "source section name", false)?;
            }
            if let Some(value) = &source.filter_name {
                validate_text(value, "source filter name", false)?;
            }
        }
        if let Some(source) = &self.dde_source {
            if let Some(value) = &source.name {
                validate_text(value, "DDE source name", false)?;
            }
            if let Some(value) = &source.conversion_mode
                && !matches!(
                    value.as_str(),
                    "into-default-style-data-style" | "into-english-number" | "keep-text"
                )
            {
                return invalid("unsupported section DDE conversion mode");
            }
        }
        Ok(())
    }

    /// Serialize a complete canonical `text:section` with inert plain-text content.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = section_opening(self)?;
        output.push('>');
        write_source(&mut output, self);
        for paragraph in self.content.split('\n') {
            output.push_str("<text:p>");
            escape(&mut output, paragraph, false);
            output.push_str("</text:p>");
        }
        output.push_str("</text:section>");
        Ok(output)
    }
}

pub fn add_section_xml(xml: &str, section: &Section) -> Result<String> {
    let scan = scan(xml)?;
    ensure_unique(&scan, None, &section.name, section.xml_id.as_deref())?;
    let fragment = section.to_xml_fragment()?;
    validate_candidate(apply(
        xml,
        vec![(scan.office_text_close..scan.office_text_close, fragment)],
    )?)
}

pub fn update_section_xml(xml: &str, name: &str, replacement: &Section) -> Result<String> {
    let scan = scan(xml)?;
    let index = scan
        .sections
        .iter()
        .position(|site| site.name == name)
        .ok_or_else(|| Error::InvalidFormat(format!("section '{name}' was not found")))?;
    ensure_unique(
        &scan,
        Some(index),
        &replacement.name,
        replacement.xml_id.as_deref(),
    )?;
    replacement.validate()?;
    let site = &scan.sections[index];
    let replacement_xml = if let Some(close) = &site.close {
        let mut value = section_opening(replacement)?;
        value.push('>');
        write_source(&mut value, replacement);
        value.push_str(&xml[site.content_start..close.start]);
        value.push_str("</text:section>");
        value
    } else {
        replacement.to_xml_fragment()?
    };
    validate_candidate(apply(xml, vec![(site.whole.clone(), replacement_xml)])?)
}

pub fn remove_section_xml(xml: &str, name: &str) -> Result<String> {
    let scan = scan(xml)?;
    let site = find_section(&scan, name)?;
    validate_candidate(apply(xml, vec![(site.whole.clone(), String::new())])?)
}

pub fn unwrap_section_xml(xml: &str, name: &str) -> Result<String> {
    let scan = scan(xml)?;
    let site = find_section(&scan, name)?;
    let replacement = site.close.as_ref().map_or_else(String::new, |close| {
        xml[site.content_start..close.start].to_string()
    });
    validate_candidate(apply(xml, vec![(site.whole.clone(), replacement)])?)
}

pub fn clear_sections_xml(xml: &str) -> Result<String> {
    let scan = scan(xml)?;
    let mut edits = Vec::new();
    for site in &scan.sections {
        if let Some(close) = &site.close {
            edits.push((close.clone(), String::new()));
            if site.content_start > site.open.end {
                edits.push((site.open.end..site.content_start, String::new()));
            }
            edits.push((site.open.clone(), String::new()));
        } else {
            edits.push((site.whole.clone(), String::new()));
        }
    }
    validate_candidate(apply(xml, edits)?)
}

pub fn wrap_section_xml(
    xml: &str,
    section: &Section,
    start: &Block,
    end: &Block,
) -> Result<String> {
    section.validate()?;
    let scan = scan(xml)?;
    ensure_unique(&scan, None, &section.name, section.xml_id.as_deref())?;
    let start_site = scan
        .blocks
        .iter()
        .find(|site| &site.block == start)
        .ok_or_else(|| Error::InvalidFormat("section start block was not found".to_string()))?;
    let end_site = scan
        .blocks
        .iter()
        .find(|site| &site.block == end)
        .ok_or_else(|| Error::InvalidFormat("section end block was not found".to_string()))?;
    compatible_blocks(start, end)?;
    if start_site.span.start > end_site.span.start {
        return invalid("section wrap start must not follow its end");
    }
    let (start_offset, end_offset) = (start_site.span.start, end_site.span.end);
    for existing in &scan.sections {
        let (existing_start, existing_end) = (existing.whole.start, existing.whole.end);
        if (existing_start < start_offset
            && start_offset < existing_end
            && existing_end < end_offset)
            || (start_offset < existing_start
                && existing_start < end_offset
                && end_offset < existing_end)
        {
            return invalid("section ranges must be nested or disjoint, not crossing");
        }
    }
    let mut opening = section_opening(section)?;
    opening.push('>');
    write_source(&mut opening, section);
    let output = apply(
        xml,
        vec![
            (
                end_site.span.end..end_site.span.end,
                "</text:section>".to_string(),
            ),
            (start_site.span.start..start_site.span.start, opening),
        ],
    )?;
    validate_candidate(output)
}

fn compatible_blocks(start: &Block, end: &Block) -> Result<()> {
    match (start, end) {
        (
            Block::BodyParagraph(_) | Block::BodyTable(_),
            Block::BodyParagraph(_) | Block::BodyTable(_),
        ) => Ok(()),
        (
            Block::TableCellParagraph {
                table: a,
                row: b,
                cell: c,
                ..
            },
            Block::TableCellParagraph {
                table: x,
                row: y,
                cell: z,
                ..
            },
        ) if (a, b, c) == (x, y, z) => Ok(()),
        _ => invalid("section wrap endpoints must share the same body or table-cell story"),
    }
}

fn section_opening(section: &Section) -> Result<String> {
    let mut output = String::from(
        "<text:section xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"",
    );
    attr(&mut output, "text:name", &section.name);
    if let Some(value) = &section.style {
        attr(&mut output, "text:style-name", value);
    }
    if section.protected {
        attr(&mut output, "text:protected", "true");
    }
    if let Some(value) = &section.xml_id {
        attr(&mut output, "xml:id", value);
    }
    if let Some(value) = &section.protection_key {
        attr(&mut output, "text:protection-key", value);
    }
    if let Some(value) = &section.protection_key_digest_algorithm {
        attr(&mut output, "text:protection-key-digest-algorithm", value);
    }
    match section.display {
        SectionDisplay::Visible => {},
        SectionDisplay::Hidden => attr(&mut output, "text:display", "none"),
        SectionDisplay::Condition => {
            attr(&mut output, "text:display", "condition");
            attr(
                &mut output,
                "text:condition",
                section.condition.as_deref().ok_or_else(|| {
                    Error::InvalidFormat("conditional section has no condition".to_string())
                })?,
            );
        },
    }
    Ok(output)
}

fn write_source(output: &mut String, section: &Section) {
    if let Some(source) = &section.source {
        output.push_str("<text:section-source xmlns:xlink=\"http://www.w3.org/1999/xlink\"");
        if let Some(value) = &source.href {
            attr(output, "xlink:href", value);
            attr(output, "xlink:type", "simple");
            attr(output, "xlink:show", "embed");
        }
        if let Some(value) = &source.section_name {
            attr(output, "text:section-name", value);
        }
        if let Some(value) = &source.filter_name {
            attr(output, "text:filter-name", value);
        }
        output.push_str("/>");
    } else if let Some(source) = &section.dde_source {
        output.push_str("<office:dde-source");
        if let Some(value) = &source.name {
            attr(output, "office:name", value);
        }
        if let Some(value) = &source.conversion_mode {
            attr(output, "office:conversion-mode", value);
        }
        if let Some(value) = source.automatic_update {
            attr(
                output,
                "office:automatic-update",
                if value { "true" } else { "false" },
            );
        }
        output.push_str("/>");
    }
}

#[derive(Clone)]
struct SectionSite {
    name: String,
    xml_id: Option<String>,
    whole: Range<usize>,
    open: Range<usize>,
    close: Option<Range<usize>>,
    content_start: usize,
}
struct BlockSite {
    block: Block,
    span: Range<usize>,
}
struct Scan {
    office_text_close: usize,
    sections: Vec<SectionSite>,
    blocks: Vec<BlockSite>,
    xml_ids: HashMap<String, usize>,
}
struct OpenSection {
    name: String,
    xml_id: Option<String>,
    depth: usize,
    start: usize,
    open: Range<usize>,
    content_start: usize,
}
struct OpenBlock {
    block: Block,
    depth: usize,
    start: usize,
}
struct TableState {
    table: usize,
    next_row: usize,
    row: Option<usize>,
    next_cell: usize,
    cell: Option<usize>,
    next_paragraph: usize,
    depth: usize,
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > 256 * 1024 * 1024 {
        return invalid("section XML exceeds 256 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut depth = 0usize;
    let mut office_text_depth = None;
    let mut office_text_close = None;
    let mut sections = Vec::new();
    let mut open_sections = Vec::new();
    let mut blocks = Vec::new();
    let mut open_blocks = Vec::new();
    let mut body_p = 0usize;
    let mut next_table = 0usize;
    let mut table: Option<TableState> = None;
    let mut tracked_depth = None;
    let mut annotation_depth = None;
    let mut xml_ids = HashMap::new();
    loop {
        let (ns, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| Error::InvalidFormat(format!("invalid mutable section XML: {e}")))?;
        let text = crate::elements::xml::is_bound(&ns, TEXT_NS);
        let office = crate::elements::xml::is_bound(&ns, OFFICE_NS);
        let table_ns = crate::elements::xml::is_bound(&ns, TABLE_NS);
        drop(ns);
        let end = reader.buffer_position() as usize;
        let span = previous..end;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                let local = element.local_name();
                record_xml_id(&reader, element, &mut xml_ids)?;
                if office && local.as_ref() == b"text" && office_text_depth.is_none() {
                    office_text_depth = Some(depth);
                }
                if text && local.as_ref() == b"tracked-changes" {
                    tracked_depth = Some(depth);
                }
                if office && local.as_ref() == b"annotation" {
                    annotation_depth = Some(depth);
                }
                if text && local.as_ref() == b"section" {
                    let section = section_identity(&reader, element)?;
                    open_sections.push(OpenSection {
                        name: section.0,
                        xml_id: section.1,
                        depth,
                        start: span.start,
                        open: span.clone(),
                        content_start: span.end,
                    });
                } else if open_sections.last().is_some_and(|s| depth == s.depth + 1)
                    && ((text && local.as_ref() == b"section-source")
                        || (office && local.as_ref() == b"dde-source"))
                {
                    return invalid("section source declarations must be empty");
                }
                if tracked_depth.is_none() && annotation_depth.is_none() {
                    update_table_start(
                        table_ns,
                        local.as_ref(),
                        depth,
                        &mut table,
                        &mut next_table,
                    );
                    if let Some(block) =
                        block_start(text, table_ns, local.as_ref(), &mut table, &mut body_p)
                    {
                        open_blocks.push(OpenBlock {
                            block,
                            depth,
                            start: span.start,
                        });
                    }
                }
            },
            Event::Empty(ref element) => {
                let local = element.local_name();
                record_xml_id(&reader, element, &mut xml_ids)?;
                if text && local.as_ref() == b"section" {
                    let section = section_identity(&reader, element)?;
                    sections.push(SectionSite {
                        name: section.0,
                        xml_id: section.1,
                        whole: span.clone(),
                        open: span.clone(),
                        close: None,
                        content_start: span.end,
                    });
                } else if open_sections.last().is_some_and(|s| depth == s.depth)
                    && ((text && local.as_ref() == b"section-source")
                        || (office && local.as_ref() == b"dde-source"))
                {
                    open_sections
                        .last_mut()
                        .ok_or_else(|| {
                            Error::InvalidFormat("section source has no parent section".to_string())
                        })?
                        .content_start = span.end;
                }
                if tracked_depth.is_none()
                    && annotation_depth.is_none()
                    && let Some(block) =
                        block_start(text, table_ns, local.as_ref(), &mut table, &mut body_p)
                {
                    blocks.push(BlockSite {
                        block,
                        span: span.clone(),
                    });
                }
            },
            Event::End(ref element) => {
                let local = element.local_name();
                if let Some(index) = open_blocks.iter().rposition(|b| b.depth == depth) {
                    let block = open_blocks.remove(index);
                    blocks.push(BlockSite {
                        block: block.block,
                        span: block.start..span.end,
                    });
                }
                if let Some(open) = open_sections.pop_if(|section| section.depth == depth) {
                    sections.push(SectionSite {
                        name: open.name,
                        xml_id: open.xml_id,
                        whole: open.start..span.end,
                        open: open.open,
                        close: Some(span.clone()),
                        content_start: open.content_start,
                    });
                }
                if office_text_depth == Some(depth) {
                    office_text_close = Some(span.start);
                }
                if tracked_depth == Some(depth) {
                    tracked_depth = None;
                }
                if annotation_depth == Some(depth) {
                    annotation_depth = None;
                }
                update_table_end(table_ns, local.as_ref(), depth, &mut table);
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML stack underflow".into()))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are prohibited");
            },
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !open_sections.is_empty() || !open_blocks.is_empty() {
        return invalid("incomplete mutable section XML");
    }
    sections.sort_by_key(|s| s.whole.start);
    blocks.sort_by_key(|b| b.span.start);
    Ok(Scan {
        office_text_close: office_text_close
            .ok_or_else(|| Error::InvalidFormat("document has no office:text".into()))?,
        sections,
        blocks,
        xml_ids,
    })
}

fn record_xml_id(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    ids: &mut HashMap<String, usize>,
) -> Result<()> {
    if let Some(id) = crate::elements::xml::namespaced_attribute(
        reader,
        element,
        b"http://www.w3.org/XML/1998/namespace",
        b"id",
        "xml:id",
    )? {
        validate_ncname(&id, "xml:id")?;
        let count = ids.entry(id.clone()).or_insert(0);
        *count += 1;
        if *count > 1 {
            return invalid(format!("duplicate xml:id '{id}'"));
        }
    }
    Ok(())
}

fn section_identity(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<(String, Option<String>)> {
    let name =
        crate::elements::xml::namespaced_attribute(reader, element, TEXT_NS, b"name", "section")?
            .ok_or_else(|| Error::InvalidFormat("section requires text:name".into()))?;
    let xml_id = crate::elements::xml::namespaced_attribute(
        reader,
        element,
        b"http://www.w3.org/XML/1998/namespace",
        b"id",
        "section",
    )?;
    Ok((name, xml_id))
}

fn block_start(
    text: bool,
    table_ns: bool,
    local: &[u8],
    table: &mut Option<TableState>,
    body: &mut usize,
) -> Option<Block> {
    if text && matches!(local, b"p" | b"h") {
        if let Some(t) = table {
            let block = Block::TableCellParagraph {
                table: t.table,
                row: t.row?,
                cell: t.cell?,
                paragraph: t.next_paragraph,
            };
            t.next_paragraph += 1;
            Some(block)
        } else {
            let i = *body;
            *body += 1;
            Some(Block::BodyParagraph(i))
        }
    } else if table_ns && local == b"table" && table.as_ref().is_some_and(|t| t.depth > 0) {
        Some(Block::BodyTable(table.as_ref()?.table))
    } else {
        None
    }
}
fn update_table_start(
    ns: bool,
    local: &[u8],
    depth: usize,
    table: &mut Option<TableState>,
    next: &mut usize,
) {
    if !ns {
        return;
    }
    match local {
        b"table" if table.is_none() => {
            *table = Some(TableState {
                table: *next,
                next_row: 0,
                row: None,
                next_cell: 0,
                cell: None,
                next_paragraph: 0,
                depth,
            });
            *next += 1;
        },
        b"table-row" => {
            if let Some(t) = table {
                t.row = Some(t.next_row);
                t.next_row += 1;
                t.next_cell = 0;
            }
        },
        b"table-cell" | b"covered-table-cell" => {
            if let Some(t) = table {
                t.cell = Some(t.next_cell);
                t.next_cell += 1;
                t.next_paragraph = 0;
            }
        },
        _ => {},
    }
}
fn update_table_end(ns: bool, local: &[u8], depth: usize, table: &mut Option<TableState>) {
    if !ns {
        return;
    }
    match local {
        b"table" if table.as_ref().is_some_and(|t| t.depth == depth) => *table = None,
        b"table-row" => {
            if let Some(t) = table {
                t.row = None;
            }
        },
        b"table-cell" | b"covered-table-cell" => {
            if let Some(t) = table {
                t.cell = None;
            }
        },
        _ => {},
    }
}

fn find_section<'a>(scan: &'a Scan, name: &str) -> Result<&'a SectionSite> {
    scan.sections
        .iter()
        .find(|s| s.name == name)
        .ok_or_else(|| Error::InvalidFormat(format!("section '{name}' was not found")))
}
fn ensure_unique(
    scan: &Scan,
    except: Option<usize>,
    name: &str,
    xml_id: Option<&str>,
) -> Result<()> {
    validate_text(name, "section name", false)?;
    if scan
        .sections
        .iter()
        .enumerate()
        .any(|(i, s)| Some(i) != except && s.name == name)
    {
        return invalid(format!("duplicate section name '{name}'"));
    }
    if let Some(id) = xml_id {
        validate_ncname(id, "section xml:id")?;
        let same_existing = except
            .and_then(|i| scan.sections.get(i))
            .and_then(|s| s.xml_id.as_deref())
            == Some(id);
        if scan.xml_ids.get(id).copied().unwrap_or(0) > usize::from(same_existing) {
            return invalid(format!("duplicate section xml:id '{id}'"));
        }
    }
    Ok(())
}
fn validate_candidate(xml: String) -> Result<String> {
    let _ = scan(&xml)?;
    let sections = Parser::parse_sections(&xml)?;
    let mut names = HashSet::new();
    for section in sections {
        section.validate()?;
        if !names.insert(section.name) {
            return invalid("duplicate section name");
        }
    }
    Ok(xml)
}
fn apply(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String> {
    edits.sort_by_key(|edit| std::cmp::Reverse(edit.0.start));
    let mut output = xml.to_string();
    let mut previous = xml.len();
    for (span, value) in edits {
        if span.start > span.end || span.end > previous {
            return invalid("overlapping section mutation spans");
        }
        output.replace_range(span.clone(), &value);
        previous = span.start;
    }
    Ok(output)
}
fn attr(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape(output, value, true);
    output.push('"');
}
fn escape(output: &mut String, value: &str, attribute: bool) {
    for c in value.chars() {
        match c {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' if attribute => output.push_str("&quot;"),
            '\'' if attribute => output.push_str("&apos;"),
            _ => output.push(c),
        }
    }
}
fn validate_text(value: &str, label: &str, empty: bool) -> Result<()> {
    if !empty && value.is_empty() {
        return invalid(format!("{label} cannot be empty"));
    }
    if value.len() > MAX_VALUE {
        return invalid(format!("{label} exceeds 64 KiB"));
    }
    if value
        .chars()
        .any(|c| c == '\0' || c.is_control() && !matches!(c, '\t' | '\n' | '\r'))
    {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}
fn validate_uri(value: &str, label: &str) -> Result<()> {
    validate_text(value, label, false)?;
    if value.chars().any(char::is_whitespace) {
        return invalid(format!("{label} contains whitespace"));
    }
    Ok(())
}
fn validate_ncname(value: &str, label: &str) -> Result<()> {
    validate_text(value, label, false)?;
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return invalid(format!("{label} must not be empty"));
    };
    if !(first == '_' || first.is_alphabetic())
        || chars.any(|c| !(c == '_' || c == '-' || c == '.' || c.is_alphanumeric()))
    {
        return invalid(format!("{label} must be an XML NCName"));
    }
    Ok(())
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
