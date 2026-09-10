//! ODS host navigation and bounded database-range XML replacement.

use crate::model::database_range::{Range, write_database_ranges};
use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TABLE_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:table:1.0";
const CALC_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;
const MAX_DEPTH: usize = 512;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;

/// A checked byte span for one XML element.
#[derive(Debug, Clone)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) end: usize,
    pub(crate) empty: bool,
    pub(crate) qname: String,
}

/// The legal spreadsheet host and its optional `database-range` child.
#[derive(Debug, Clone)]
pub(crate) struct Location {
    pub(crate) spreadsheet: Span,
    pub(crate) container: Option<Span>,
    pub(crate) insert_at: usize,
    pub(crate) opaque: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Document,
    Body,
    Spreadsheet,
    Container,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    TableExt,
    CalcExt,
    Other,
}

/// The table vocabulary names whose placement is owned by this facade.
///
/// Keeping the expanded table name separate from the serialized QName is
/// important here.  A known local name in the wrong parent is still authored
/// XML that the typed database-range graph cannot represent; treating it as
/// an ordinary known child would let a later rewrite silently drop it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum KnownElement {
    DatabaseRanges,
    DatabaseRange,
    SourceSql,
    SourceTable,
    SourceQuery,
    Filter,
    FilterCondition,
    FilterAnd,
    FilterOr,
    FilterSetItem,
    Sort,
    SortBy,
    SubtotalRules,
    SortGroups,
    SubtotalRule,
    SubtotalField,
}

#[derive(Debug)]
struct OpenElement {
    kind: Kind,
    known: Option<KnownElement>,
    start: usize,
    tag_end: usize,
    qname: String,
}

/// Locate the direct spreadsheet `database-range` owner and its safe insertion point.
pub(crate) fn locate(xml: &str) -> Result<Location> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("ODS database-range source exceeds the size limit"));
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<OpenElement>::new();
    let mut spreadsheet = None;
    let mut container = None;
    let mut insertion = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut events = 0usize;
    let mut opaque = false;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("ODS database-range XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("ODS database-range source exceeds the event limit"));
        }

        let event_start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid ODS database-range XML: {error}")))?;
        let namespace = namespace_kind(&resolved);
        let event_end = position(&reader)?;

        // Validate the borrowed raw attributes before converting the event to
        // an owned value.  `quick-xml` can decode an attribute value while
        // materializing the event, so doing this in the owned match would let
        // an oversized value allocate before the admission check runs.
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                let parent = stack.last().map(|open| open.kind);
                let kind = classify(parent, namespace, element.local_name().as_ref());
                let inside_container = stack.iter().any(|open| open.kind == Kind::Container);
                if kind == Kind::Container || inside_container {
                    let known = known_element(namespace, element.local_name().as_ref());
                    if known.is_none() {
                        opaque = true;
                    }
                    if let Some(known) = known {
                        if !allowed_child(stack.last(), known) {
                            opaque = true;
                        }
                    }
                    if !validate_attributes(
                        element,
                        reader.resolver(),
                        element.local_name().as_ref(),
                    )? {
                        opaque = true;
                    }
                }
            },
            _ => {},
        }

        let event = event.into_owned();

        match event {
            Event::Start(element) => {
                if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("ODS content.xml has more than one root element"));
                    }
                    root_seen = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid(
                        "ODS database-range source exceeds the nesting limit",
                    ));
                }
                let parent = stack.last().map(|open| open.kind);
                let kind = classify(parent, namespace, element.local_name().as_ref());
                let known = known_element(namespace, element.local_name().as_ref());
                if parent == Some(Kind::Spreadsheet)
                    && is_insertion_anchor(namespace, element.local_name().as_ref())
                    && insertion.is_none()
                {
                    insertion = Some(event_start);
                }
                let qname = element_name(&element)?;
                stack.push(OpenElement {
                    kind,
                    known,
                    start: event_start,
                    tag_end: event_end,
                    qname,
                });
            },
            Event::Empty(element) => {
                if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("ODS content.xml has more than one root element"));
                    }
                    root_seen = true;
                    root_closed = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid(
                        "ODS database-range source exceeds the nesting limit",
                    ));
                }
                let parent = stack.last().map(|open| open.kind);
                let kind = classify(parent, namespace, element.local_name().as_ref());
                if parent == Some(Kind::Spreadsheet)
                    && is_insertion_anchor(namespace, element.local_name().as_ref())
                    && insertion.is_none()
                {
                    insertion = Some(event_start);
                }
                let qname = element_name(&element)?;
                record(
                    kind,
                    Span {
                        start: event_start,
                        tag_end: event_end,
                        close_start: event_start,
                        end: event_end,
                        empty: true,
                        qname,
                    },
                    &mut spreadsheet,
                    &mut container,
                )?;
            },
            Event::End(_) => {
                let open = stack.pop().ok_or_else(|| invalid("unbalanced ODS XML"))?;
                let span = Span {
                    start: open.start,
                    tag_end: open.tag_end,
                    close_start: event_start,
                    end: event_end,
                    empty: false,
                    qname: open.qname,
                };
                record(open.kind, span, &mut spreadsheet, &mut container)?;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) if stack.iter().any(|open| open.kind == Kind::Container) => {
                let value = text
                    .xml_content(quick_xml::XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        invalid(format!("invalid ODS database-range text: {error}"))
                    })?;
                if !value.trim().is_empty() {
                    opaque = true;
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_)
                if stack.iter().any(|open| open.kind == Kind::Container) =>
            {
                opaque = true;
            },
            Event::GeneralRef(_) if stack.iter().any(|open| open.kind == Kind::Container) => {
                opaque = true;
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
        buffer.clear();
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("incomplete ODS content.xml document"));
    }
    let spreadsheet =
        spreadsheet.ok_or_else(|| invalid("ODS content.xml has no office:spreadsheet host"))?;
    let insert_at = insertion.unwrap_or(spreadsheet.close_start);
    Ok(Location {
        spreadsheet,
        container,
        insert_at,
        opaque,
    })
}

fn classify(parent: Option<Kind>, namespace: NamespaceKind, local: &[u8]) -> Kind {
    if parent.is_none() && is_element(namespace, NamespaceKind::Office, local, b"document-content")
    {
        Kind::Document
    } else if parent == Some(Kind::Document)
        && is_element(namespace, NamespaceKind::Office, local, b"body")
    {
        Kind::Body
    } else if parent == Some(Kind::Body)
        && is_element(namespace, NamespaceKind::Office, local, b"spreadsheet")
    {
        Kind::Spreadsheet
    } else if parent == Some(Kind::Spreadsheet)
        && is_element(namespace, NamespaceKind::Table, local, b"database-ranges")
    {
        Kind::Container
    } else {
        Kind::Other
    }
}

fn record(
    kind: Kind,
    span: Span,
    spreadsheet: &mut Option<Span>,
    container: &mut Option<Span>,
) -> Result<()> {
    match kind {
        Kind::Spreadsheet => {
            if spreadsheet.replace(span).is_some() {
                return Err(invalid(
                    "ODS content.xml has more than one office:spreadsheet host",
                ));
            }
        },
        Kind::Container if container.replace(span).is_some() => {
            return Err(invalid(
                "ODS content.xml has duplicate table:database-ranges",
            ));
        },
        Kind::Document | Kind::Body | Kind::Container | Kind::Other => {},
    }
    Ok(())
}

/// Replace only the owned `database-range` container, or insert/remove it as a
/// direct child of the spreadsheet host.
pub(crate) fn replace(
    source: &str,
    location: &Location,
    ranges: Option<&[Range]>,
) -> Result<String> {
    let fragment = ranges.map(render).transpose()?;
    if let Some(container) = &location.container {
        return splice(
            source,
            container.start,
            container.end,
            fragment.as_deref().unwrap_or_default(),
        );
    }
    let Some(fragment) = fragment else {
        return Ok(source.to_owned());
    };
    let spreadsheet = &location.spreadsheet;
    if spreadsheet.empty {
        let opening = source
            .get(spreadsheet.start..spreadsheet.tag_end)
            .ok_or_else(|| invalid("invalid ODS spreadsheet XML span"))?;
        let opening = opening
            .strip_suffix("/>")
            .ok_or_else(|| invalid("empty ODS spreadsheet has no close token"))?;
        let mut expanded = String::with_capacity(opening.len() + fragment.len() + 32);
        expanded.push_str(opening);
        expanded.push('>');
        expanded.push_str(&fragment);
        expanded.push_str("</");
        expanded.push_str(&spreadsheet.qname);
        expanded.push('>');
        splice(source, spreadsheet.start, spreadsheet.end, &expanded)
    } else {
        splice(source, location.insert_at, location.insert_at, &fragment)
    }
}

/// Add a local table namespace declaration to an owner slice before parsing.
///
/// The owner normally inherits its prefix declaration from
/// `office:document-content`. Parsing the slice independently would otherwise
/// lose that namespace binding. The declaration is added only to the temporary
/// parsing buffer; source bytes are never changed by this helper.
pub(crate) fn owner_fragment(source: &str, span: &Span) -> Result<String> {
    let fragment = source
        .get(span.start..span.end)
        .ok_or_else(|| invalid("invalid ODS database-range owner span"))?;
    let open_end = opening_tag_end(fragment)
        .ok_or_else(|| invalid("database-range owner has no complete opening tag"))?;
    let open = &fragment[..open_end];
    let (open, close_token) = open
        .strip_suffix('/')
        .map_or((open, ">"), |open| (open, "/>"));
    let name = span.qname.as_str();
    let (prefix, declaration) = name.split_once(':').map_or(
        (
            None,
            " xmlns=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"".to_string(),
        ),
        |(prefix, _)| {
            (
                Some(prefix),
                format!(" xmlns:{prefix}=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\""),
            )
        },
    );
    let already_declared = prefix.map_or_else(
        || has_namespace_declaration(open, "xmlns"),
        |prefix| has_namespace_declaration(open, &format!("xmlns:{prefix}")),
    );
    if already_declared {
        return Ok(fragment.to_owned());
    }
    let mut output = String::new();
    output
        .try_reserve(fragment.len() + declaration.len())
        .map_err(|source| Error::Allocation {
            resource: "ODS database-range owner parse buffer",
            source,
        })?;
    output.push_str(open);
    output.push_str(&declaration);
    output.push_str(close_token);
    output.push_str(&fragment[open_end + 1..]);
    Ok(output)
}

fn has_namespace_declaration(open: &str, name: &str) -> bool {
    let bytes = open.as_bytes();
    let expected = name.as_bytes();
    let mut index = 0;

    // Skip the opening '<' and the element QName.
    if bytes.first() == Some(&b'<') {
        index += 1;
    }
    while index < bytes.len() && !is_xml_space(bytes[index]) {
        index += 1;
    }

    // Scan attribute names while skipping their quoted values.  A simple
    // whitespace split can mistake text such as `xmlns:t=...` inside an
    // attribute value for a declaration on the owner itself.
    loop {
        while index < bytes.len() && is_xml_space(bytes[index]) {
            index += 1;
        }
        if index >= bytes.len() || bytes[index] == b'/' {
            return false;
        }
        let name_start = index;
        while index < bytes.len()
            && !is_xml_space(bytes[index])
            && bytes[index] != b'='
            && bytes[index] != b'/'
        {
            index += 1;
        }
        if &bytes[name_start..index] == expected {
            return true;
        }
        while index < bytes.len() && is_xml_space(bytes[index]) {
            index += 1;
        }
        if bytes.get(index) == Some(&b'=') {
            index += 1;
            while index < bytes.len() && is_xml_space(bytes[index]) {
                index += 1;
            }
            if let Some(&quote @ (b'"' | b'\'')) = bytes.get(index) {
                index += 1;
                while index < bytes.len() && bytes[index] != quote {
                    index += 1;
                }
                if index < bytes.len() {
                    index += 1;
                }
            } else {
                while index < bytes.len() && !is_xml_space(bytes[index]) {
                    index += 1;
                }
            }
        }
    }
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

fn opening_tag_end(fragment: &str) -> Option<usize> {
    let mut quote = None;
    for (index, byte) in fragment.bytes().enumerate() {
        match quote {
            Some(delimiter) if byte == delimiter => quote = None,
            Some(_) => {},
            None if byte == b'"' || byte == b'\'' => quote = Some(byte),
            None if byte == b'>' => return Some(index),
            None => {},
        }
    }
    None
}

fn render(ranges: &[Range]) -> Result<String> {
    if ranges.is_empty() {
        return Ok(format!(
            "<table:database-ranges xmlns:table=\"{}\"/>",
            String::from_utf8_lossy(TABLE_NAMESPACE)
        ));
    }
    let mut output = String::new();
    write_database_ranges(&mut output, ranges)?;
    const ROOT: &str = "<table:database-ranges>";
    if !output.starts_with(ROOT) {
        return Err(invalid("database-range writer emitted an unexpected root"));
    }
    output.insert_str(
        ROOT.len() - 1,
        &format!(
            " xmlns:table=\"{}\"",
            String::from_utf8_lossy(TABLE_NAMESPACE)
        ),
    );
    Ok(output)
}

fn validate_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    local: &[u8],
) -> Result<bool> {
    let mut known = true;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid ODS database-range attribute: {error}")))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(invalid(
                "ODS database-range attribute exceeds the size limit",
            ));
        }
        let key = attribute.key.as_ref();
        if key == b"xmlns"
            || attribute
                .key
                .prefix()
                .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
        {
            continue;
        }
        let (namespace, name) = resolver.resolve_attribute(attribute.key);
        if !allowed_attribute(namespace, local, name.as_ref()) {
            known = false;
        }
    }
    Ok(known)
}

fn allowed_attribute(namespace: ResolveResult<'_>, element: &[u8], attribute: &[u8]) -> bool {
    let namespace = match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == *TABLE_NAMESPACE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(uri)) if *uri == *TABLE_EXT_NAMESPACE => {
            NamespaceKind::TableExt
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == *CALC_EXT_NAMESPACE => {
            NamespaceKind::CalcExt
        },
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    };
    match namespace {
        NamespaceKind::Table => match element {
            b"database-ranges" => false,
            b"database-range" => matches!(
                attribute,
                b"contains-header"
                    | b"display-filter-buttons"
                    | b"has-persistent-data"
                    | b"is-selection"
                    | b"name"
                    | b"on-update-keep-size"
                    | b"on-update-keep-styles"
                    | b"orientation"
                    | b"refresh-delay"
                    | b"target-range-address"
            ),
            b"database-source-sql" => {
                matches!(
                    attribute,
                    b"database-name" | b"parse-sql-statement" | b"sql-statement"
                )
            },
            b"database-source-table" => {
                matches!(attribute, b"database-name" | b"database-table-name")
            },
            b"database-source-query" => matches!(attribute, b"database-name" | b"query-name"),
            b"filter" => matches!(
                attribute,
                b"condition-source"
                    | b"condition-source-range-address"
                    | b"display-duplicates"
                    | b"target-range-address"
            ),
            b"filter-condition" => matches!(
                attribute,
                b"case-sensitive" | b"data-type" | b"field-number" | b"operator" | b"value"
            ),
            b"filter-set-item" => attribute == b"value",
            b"filter-and" | b"filter-or" => false,
            b"sort" => matches!(
                attribute,
                b"algorithm"
                    | b"bind-styles-to-content"
                    | b"case-sensitive"
                    | b"country"
                    | b"embedded-number-behavior"
                    | b"language"
                    | b"rfc-language-tag"
                    | b"script"
                    | b"target-range-address"
            ),
            b"sort-by" => matches!(attribute, b"data-type" | b"field-number" | b"order"),
            b"subtotal-rules" => matches!(
                attribute,
                b"bind-styles-to-content" | b"case-sensitive" | b"page-breaks-on-group-change"
            ),
            b"sort-groups" => matches!(attribute, b"data-type" | b"order"),
            b"subtotal-rule" => attribute == b"group-by-field-number",
            b"subtotal-field" => matches!(attribute, b"field-number" | b"function"),
            _ => false,
        },
        NamespaceKind::TableExt
        | NamespaceKind::CalcExt
        | NamespaceKind::Office
        | NamespaceKind::Other => false,
    }
}

fn known_element(namespace: NamespaceKind, local: &[u8]) -> Option<KnownElement> {
    if namespace != NamespaceKind::Table {
        return None;
    }
    Some(match local {
        b"database-ranges" => KnownElement::DatabaseRanges,
        b"database-range" => KnownElement::DatabaseRange,
        b"database-source-sql" => KnownElement::SourceSql,
        b"database-source-table" => KnownElement::SourceTable,
        b"database-source-query" => KnownElement::SourceQuery,
        b"filter" => KnownElement::Filter,
        b"filter-condition" => KnownElement::FilterCondition,
        b"filter-and" => KnownElement::FilterAnd,
        b"filter-or" => KnownElement::FilterOr,
        b"filter-set-item" => KnownElement::FilterSetItem,
        b"sort" => KnownElement::Sort,
        b"sort-by" => KnownElement::SortBy,
        b"subtotal-rules" => KnownElement::SubtotalRules,
        b"sort-groups" => KnownElement::SortGroups,
        b"subtotal-rule" => KnownElement::SubtotalRule,
        b"subtotal-field" => KnownElement::SubtotalField,
        _ => return None,
    })
}

/// Return whether a known table element is legal directly under its expanded
/// name parent.  This is deliberately a parent grammar rather than a prefix
/// check: producers may use any table prefix (or the default namespace), and
/// the same local name in another namespace is opaque by construction.
fn allowed_child(parent: Option<&OpenElement>, child: KnownElement) -> bool {
    let parent_kind = parent.map(|open| open.kind);
    let parent_known = parent.and_then(|open| open.known);
    match child {
        KnownElement::DatabaseRanges => parent_kind == Some(Kind::Spreadsheet),
        KnownElement::DatabaseRange => parent_known == Some(KnownElement::DatabaseRanges),
        KnownElement::SourceSql
        | KnownElement::SourceTable
        | KnownElement::SourceQuery
        | KnownElement::Filter
        | KnownElement::Sort
        | KnownElement::SubtotalRules => parent_known == Some(KnownElement::DatabaseRange),
        KnownElement::FilterCondition => matches!(
            parent_known,
            Some(KnownElement::Filter)
                | Some(KnownElement::FilterAnd)
                | Some(KnownElement::FilterOr)
        ),
        KnownElement::FilterAnd => matches!(
            parent_known,
            Some(KnownElement::Filter) | Some(KnownElement::FilterOr)
        ),
        KnownElement::FilterOr => matches!(
            parent_known,
            Some(KnownElement::Filter) | Some(KnownElement::FilterAnd)
        ),
        KnownElement::FilterSetItem => parent_known == Some(KnownElement::FilterCondition),
        KnownElement::SortBy => parent_known == Some(KnownElement::Sort),
        KnownElement::SortGroups | KnownElement::SubtotalRule => {
            parent_known == Some(KnownElement::SubtotalRules)
        },
        KnownElement::SubtotalField => parent_known == Some(KnownElement::SubtotalRule),
    }
}

fn is_insertion_anchor(namespace: NamespaceKind, local: &[u8]) -> bool {
    namespace == NamespaceKind::Table
        && matches!(
            local,
            b"data-pilot-tables" | b"consolidation" | b"dde-links" | b"shapes"
        )
}

fn element_name(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_error| invalid("ODS database-range element name is not UTF-8"))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid("ODS database-range XML position overflows usize"))
}

fn is_element(
    namespace: NamespaceKind,
    expected: NamespaceKind,
    local: &[u8],
    expected_local: &[u8],
) -> bool {
    namespace == expected && local == expected_local
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_EXT_NAMESPACE => {
            NamespaceKind::TableExt
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == CALC_EXT_NAMESPACE => {
            NamespaceKind::CalcExt
        },
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn splice(source: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end
        || end > source.len()
        || !source.is_char_boundary(start)
        || !source.is_char_boundary(end)
    {
        return Err(invalid("invalid ODS database-range XML span"));
    }
    let capacity = source
        .len()
        .checked_sub(end - start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| invalid("ODS database-range replacement size overflow"))?;
    if capacity > 128 * 1_048_576 {
        return Err(invalid(
            "ODS database-range replacement exceeds the size limit",
        ));
    }
    let mut output = String::new();
    output
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODS database-range replacement XML",
            source,
        })?;
    output.push_str(&source[..start]);
    output.push_str(replacement);
    output.push_str(&source[end..]);
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
