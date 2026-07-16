//! Worksheet data-consolidation settings (`CT_DataConsolidate`).

use std::fmt::Write;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::XmlVersion;

use crate::common::mce::process_str;
use crate::error::{OoxmlError, Result};

const TRANSITIONAL_MAIN: &[u8] =
    b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_REL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_DATA_REFERENCES: usize = 65_536;
const MAX_XSTRING_CHARS: usize = 32_767;
const MAX_RELATIONSHIP_ID_CHARS: usize = 1_024;

/// Namespace form used when serializing a consolidation fragment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetDataConsolidationConformance {
    Transitional,
    Strict,
}

impl WorksheetDataConsolidationConformance {
    fn main_namespace(self) -> &'static str {
        match self {
            Self::Transitional => std::str::from_utf8(TRANSITIONAL_MAIN).unwrap(),
            Self::Strict => std::str::from_utf8(STRICT_MAIN).unwrap(),
        }
    }

    fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => std::str::from_utf8(TRANSITIONAL_REL).unwrap(),
            Self::Strict => std::str::from_utf8(STRICT_REL).unwrap(),
        }
    }
}

/// Mathematical aggregator selected by `ST_DataConsolidateFunction`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetDataConsolidationFunction {
    Average,
    Count,
    CountNumbers,
    Maximum,
    Minimum,
    Product,
    StandardDeviation,
    PopulationStandardDeviation,
    Sum,
    Variance,
    PopulationVariance,
}

impl WorksheetDataConsolidationFunction {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "average" => Ok(Self::Average),
            "count" => Ok(Self::Count),
            "countNums" => Ok(Self::CountNumbers),
            "max" => Ok(Self::Maximum),
            "min" => Ok(Self::Minimum),
            "product" => Ok(Self::Product),
            "stdDev" => Ok(Self::StandardDeviation),
            "stdDevp" => Ok(Self::PopulationStandardDeviation),
            "sum" => Ok(Self::Sum),
            "var" => Ok(Self::Variance),
            "varp" => Ok(Self::PopulationVariance),
            _ => Err(invalid(format!("invalid dataConsolidate function {value:?}"))),
        }
    }

    pub fn as_str(self) -> &'static str {
        match self {
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNumbers => "countNums",
            Self::Maximum => "max",
            Self::Minimum => "min",
            Self::Product => "product",
            Self::StandardDeviation => "stdDev",
            Self::PopulationStandardDeviation => "stdDevp",
            Self::Sum => "sum",
            Self::Variance => "var",
            Self::PopulationVariance => "varp",
        }
    }
}

impl Default for WorksheetDataConsolidationFunction {
    fn default() -> Self { Self::Sum }
}

/// A validated A1 cell or rectangular range reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetDataConsolidationRangeReference(String);

impl WorksheetDataConsolidationRangeReference {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if !valid_range_reference(&value) {
            return Err(invalid(format!("invalid dataRef ref {value:?}")));
        }
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str { &self.0 }
}

/// The mutually exclusive source forms of `CT_DataRef`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum WorksheetDataReferenceSource {
    DefinedName(String),
    Range {
        sheet: String,
        reference: WorksheetDataConsolidationRangeReference,
    },
}

/// A single consolidation source, optionally in an external workbook.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetDataReference {
    source: WorksheetDataReferenceSource,
    relationship_id: Option<String>,
}

impl WorksheetDataReference {
    pub fn named(name: impl Into<String>) -> Result<Self> {
        let name = checked_xstring(name.into(), "dataRef name")?;
        Ok(Self { source: WorksheetDataReferenceSource::DefinedName(name), relationship_id: None })
    }

    pub fn range(
        sheet: impl Into<String>,
        reference: WorksheetDataConsolidationRangeReference,
    ) -> Result<Self> {
        let sheet = checked_xstring(sheet.into(), "dataRef sheet")?;
        Ok(Self {
            source: WorksheetDataReferenceSource::Range { sheet, reference },
            relationship_id: None,
        })
    }

    pub fn with_relationship_id(mut self, relationship_id: impl Into<String>) -> Result<Self> {
        self.relationship_id = Some(checked_relationship_id(relationship_id.into())?);
        Ok(self)
    }

    pub fn source(&self) -> &WorksheetDataReferenceSource { &self.source }
    pub fn relationship_id(&self) -> Option<&str> { self.relationship_id.as_deref() }
}

/// Bounded `dataRefs` collection and its optional source count.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetDataReferences {
    references: Vec<WorksheetDataReference>,
    declared_count: Option<u32>,
}

impl WorksheetDataReferences {
    pub fn new(references: Vec<WorksheetDataReference>) -> Result<Self> {
        validate_reference_count(references.len())?;
        Ok(Self { declared_count: Some(references.len() as u32), references })
    }

    pub fn references(&self) -> &[WorksheetDataReference] { &self.references }
    pub fn declared_count(&self) -> Option<u32> { self.declared_count }
}

/// Complete immutable worksheet `dataConsolidate` settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetDataConsolidation {
    function: WorksheetDataConsolidationFunction,
    left_labels: bool,
    start_labels: bool,
    top_labels: bool,
    link: bool,
    data_references: Option<WorksheetDataReferences>,
}

impl WorksheetDataConsolidation {
    pub fn new(
        function: WorksheetDataConsolidationFunction,
        data_references: Option<WorksheetDataReferences>,
    ) -> Self {
        Self {
            function,
            left_labels: false,
            start_labels: false,
            top_labels: false,
            link: false,
            data_references,
        }
    }

    pub fn with_left_labels(mut self, value: bool) -> Self { self.left_labels = value; self }
    pub fn with_start_labels(mut self, value: bool) -> Self { self.start_labels = value; self }
    pub fn with_top_labels(mut self, value: bool) -> Self { self.top_labels = value; self }
    pub fn with_link(mut self, value: bool) -> Self { self.link = value; self }
    pub fn function(&self) -> WorksheetDataConsolidationFunction { self.function }
    pub fn left_labels(&self) -> bool { self.left_labels }
    pub fn start_labels(&self) -> bool { self.start_labels }
    pub fn top_labels(&self) -> bool { self.top_labels }
    pub fn link(&self) -> bool { self.link }
    pub fn data_references(&self) -> Option<&WorksheetDataReferences> {
        self.data_references.as_ref()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope { Worksheet, Consolidate, DataRefs, DataRef, Other }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind { Unbound, Main, Relationship, Other }

#[derive(Default)]
struct ConsolidationBuilder {
    function: Option<WorksheetDataConsolidationFunction>,
    left_labels: Option<bool>,
    start_labels: Option<bool>,
    top_labels: Option<bool>,
    link: Option<bool>,
    data_references: Option<WorksheetDataReferences>,
}

/// Parses the direct worksheet `dataConsolidate` child after applying shared MCE processing.
pub fn parse_worksheet_data_consolidation(
    xml: &[u8],
) -> Result<Option<WorksheetDataConsolidation>> {
    let source = std::str::from_utf8(xml)
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    let processed = process_str(source)?;
    let mut reader = NsReader::from_reader(processed.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut scopes = Vec::new();
    let mut builder: Option<ConsolidationBuilder> = None;
    let mut declared_count: Option<u32> = None;
    let mut references = Vec::new();
    let mut seen_consolidation = false;
    let mut passed_consolidation_position = false;

    loop {
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid worksheet XML: {error}")))?;
        let namespace = namespace_kind(resolved)?;
        match event {
            Event::Start(element) => {
                let scope = begin_element(
                    &reader, &element, namespace, scopes.last().copied(), &mut builder,
                    &mut declared_count, &mut references, &mut seen_consolidation,
                    &mut passed_consolidation_position,
                )?;
                scopes.push(scope);
            }
            Event::Empty(element) => {
                let scope = begin_element(
                    &reader, &element, namespace, scopes.last().copied(), &mut builder,
                    &mut declared_count, &mut references, &mut seen_consolidation,
                    &mut passed_consolidation_position,
                )?;
                end_scope(scope, &mut builder, &mut declared_count, &mut references)?;
            }
            Event::End(_) => {
                let scope = scopes.pop().ok_or_else(|| invalid("unexpected worksheet end element"))?;
                end_scope(scope, &mut builder, &mut declared_count, &mut references)?;
            }
            Event::Text(text) => {
                if matches!(scopes.last(), Some(Scope::Consolidate | Scope::DataRefs | Scope::DataRef))
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("dataConsolidate family cannot contain text"));
                }
            }
            Event::CData(text) => {
                if matches!(scopes.last(), Some(Scope::Consolidate | Scope::DataRefs | Scope::DataRef))
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("dataConsolidate family cannot contain CDATA"));
                }
            }
            Event::DocType(_) => return Err(invalid("worksheet XML cannot contain a document type")),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !scopes.is_empty() { return Err(invalid("unterminated worksheet XML")); }
    Ok(builder.map(finish_builder).transpose()?)
}

#[allow(clippy::too_many_arguments)]
fn begin_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    parent: Option<Scope>,
    builder: &mut Option<ConsolidationBuilder>,
    declared_count: &mut Option<u32>,
    references: &mut Vec<WorksheetDataReference>,
    seen_consolidation: &mut bool,
    passed_consolidation_position: &mut bool,
) -> Result<Scope> {
    let local = element.local_name();
    let local = local.as_ref();
    let main = namespace == NamespaceKind::Main;
    match parent {
        None => {
            if !main || local != b"worksheet" { return Err(invalid("expected SpreadsheetML worksheet root")); }
            Ok(Scope::Worksheet)
        }
        Some(Scope::Worksheet) => {
            if local == b"dataConsolidate" {
                if !main { return Err(invalid("spoofed dataConsolidate element namespace")); }
                if *seen_consolidation { return Err(invalid("duplicate worksheet dataConsolidate element")); }
                if *passed_consolidation_position {
                    return Err(invalid("dataConsolidate is out of worksheet schema order"));
                }
                *seen_consolidation = true;
                *builder = Some(parse_consolidation_attributes(reader, element)?);
                Ok(Scope::Consolidate)
            } else {
                if main {
                    let position = worksheet_child_position(local);
                    if *seen_consolidation && position == Some(false) {
                        return Err(invalid("worksheet child precedes dataConsolidate in schema order"));
                    }
                    if !*seen_consolidation && position == Some(true) {
                        *passed_consolidation_position = true;
                    }
                }
                Ok(Scope::Other)
            }
        }
        Some(Scope::Consolidate) => {
            if local != b"dataRefs" || !main {
                return Err(invalid(if local == b"dataRefs" {
                    "spoofed dataRefs element namespace"
                } else { "unknown dataConsolidate child element" }));
            }
            let current = builder.as_ref().ok_or_else(|| invalid("missing dataConsolidate state"))?;
            if current.data_references.is_some() || declared_count.is_some() || !references.is_empty() {
                return Err(invalid("duplicate dataRefs element"));
            }
            *declared_count = parse_data_refs_attributes(reader, element)?;
            Ok(Scope::DataRefs)
        }
        Some(Scope::DataRefs) => {
            if local != b"dataRef" || !main {
                return Err(invalid(if local == b"dataRef" {
                    "spoofed dataRef element namespace"
                } else { "unknown dataRefs child element" }));
            }
            if references.len() >= MAX_DATA_REFERENCES {
                return Err(invalid(format!("dataRefs exceeds safety limit {MAX_DATA_REFERENCES}")));
            }
            references.push(parse_data_ref_attributes(reader, element)?);
            Ok(Scope::DataRef)
        }
        Some(Scope::DataRef) => Err(invalid("dataRef must be a leaf element")),
        Some(Scope::Other) => Ok(Scope::Other),
    }
}

fn end_scope(
    scope: Scope,
    builder: &mut Option<ConsolidationBuilder>,
    declared_count: &mut Option<u32>,
    references: &mut Vec<WorksheetDataReference>,
) -> Result<()> {
    if scope == Scope::DataRefs {
        validate_reference_count(references.len())?;
        if let Some(count) = *declared_count {
            if count as usize != references.len() {
                return Err(invalid(format!(
                    "dataRefs count {count} does not match {} dataRef children", references.len()
                )));
            }
        }
        let collection = WorksheetDataReferences {
            references: std::mem::take(references),
            declared_count: declared_count.take(),
        };
        let target = builder.as_mut().ok_or_else(|| invalid("missing dataConsolidate state"))?;
        if target.data_references.replace(collection).is_some() {
            return Err(invalid("duplicate dataRefs element"));
        }
    }
    Ok(())
}

fn finish_builder(builder: ConsolidationBuilder) -> Result<WorksheetDataConsolidation> {
    Ok(WorksheetDataConsolidation {
        function: builder.function.unwrap_or_default(),
        left_labels: builder.left_labels.unwrap_or(false),
        start_labels: builder.start_labels.unwrap_or(false),
        top_labels: builder.top_labels.unwrap_or(false),
        link: builder.link.unwrap_or(false),
        data_references: builder.data_references,
    })
}

fn parse_consolidation_attributes(
    reader: &NsReader<&[u8]>, element: &BytesStart<'_>,
) -> Result<ConsolidationBuilder> {
    let mut value = ConsolidationBuilder::default();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(format!("invalid dataConsolidate attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) { continue; }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(namespace)?;
        if namespace != NamespaceKind::Unbound { return Err(invalid("unknown namespaced dataConsolidate attribute")); }
        let text = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid dataConsolidate attribute value: {error}")))?;
        match local.as_ref() {
            b"function" => set_once(&mut value.function, WorksheetDataConsolidationFunction::parse(&text)?, "function")?,
            b"leftLabels" => set_once(&mut value.left_labels, parse_bool(&text, "leftLabels")?, "leftLabels")?,
            b"startLabels" => set_once(&mut value.start_labels, parse_bool(&text, "startLabels")?, "startLabels")?,
            b"topLabels" => set_once(&mut value.top_labels, parse_bool(&text, "topLabels")?, "topLabels")?,
            b"link" => set_once(&mut value.link, parse_bool(&text, "link")?, "link")?,
            _ => return Err(invalid("unknown dataConsolidate attribute")),
        }
    }
    Ok(value)
}

fn parse_data_refs_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<u32>> {
    let mut count = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(format!("invalid dataRefs attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) { continue; }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(namespace)? != NamespaceKind::Unbound || local.as_ref() != b"count" {
            return Err(invalid("unknown dataRefs attribute"));
        }
        let text = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid dataRefs count: {error}")))?;
        let parsed = text.parse::<u32>().map_err(|_| invalid("dataRefs count must be unsignedInt"))?;
        if parsed as usize > MAX_DATA_REFERENCES {
            return Err(invalid(format!("dataRefs count exceeds safety limit {MAX_DATA_REFERENCES}")));
        }
        set_once(&mut count, parsed, "count")?;
    }
    Ok(count)
}

fn parse_data_ref_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<WorksheetDataReference> {
    let mut name = None;
    let mut sheet = None;
    let mut reference = None;
    let mut relationship_id = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(format!("invalid dataRef attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) { continue; }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(namespace)?;
        let text = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid dataRef attribute value: {error}")))?.into_owned();
        match (namespace, local.as_ref()) {
            (NamespaceKind::Unbound, b"name") => set_once(&mut name, checked_xstring(text, "dataRef name")?, "name")?,
            (NamespaceKind::Unbound, b"sheet") => set_once(&mut sheet, checked_xstring(text, "dataRef sheet")?, "sheet")?,
            (NamespaceKind::Unbound, b"ref") => set_once(&mut reference, WorksheetDataConsolidationRangeReference::new(text)?, "ref")?,
            (NamespaceKind::Relationship, b"id") => {
                set_once(&mut relationship_id, checked_relationship_id(text)?, "r:id")?
            }
            _ => return Err(invalid("unknown or spoofed dataRef attribute")),
        }
    }
    let source = match (name, sheet, reference) {
        (Some(name), None, None) => WorksheetDataReferenceSource::DefinedName(name),
        (None, Some(sheet), Some(reference)) => WorksheetDataReferenceSource::Range { sheet, reference },
        _ => return Err(invalid("dataRef requires exactly name or the sheet and ref pair")),
    };
    Ok(WorksheetDataReference { source, relationship_id })
}

/// Serializes one canonical, namespace-complete `dataConsolidate` fragment.
pub fn write_worksheet_data_consolidation(
    value: &WorksheetDataConsolidation,
    conformance: WorksheetDataConsolidationConformance,
) -> Result<String> {
    if let Some(data_refs) = &value.data_references {
        validate_reference_count(data_refs.references.len())?;
    }
    let has_relationships = value.data_references.as_ref().is_some_and(|refs|
        refs.references.iter().any(|reference| reference.relationship_id.is_some()));
    let mut xml = String::new();
    write!(xml, "<dataConsolidate xmlns=\"{}\"", conformance.main_namespace()).unwrap();
    if has_relationships {
        write!(xml, " xmlns:r=\"{}\"", conformance.relationship_namespace()).unwrap();
    }
    if value.function != WorksheetDataConsolidationFunction::Sum {
        write!(xml, " function=\"{}\"", value.function.as_str()).unwrap();
    }
    write_true_attribute(&mut xml, "leftLabels", value.left_labels);
    write_true_attribute(&mut xml, "startLabels", value.start_labels);
    write_true_attribute(&mut xml, "topLabels", value.top_labels);
    write_true_attribute(&mut xml, "link", value.link);
    let Some(data_refs) = &value.data_references else {
        xml.push_str("/>");
        return Ok(xml);
    };
    xml.push('>');
    write!(xml, "<dataRefs count=\"{}\">", data_refs.references.len()).unwrap();
    for reference in &data_refs.references {
        xml.push_str("<dataRef");
        match &reference.source {
            WorksheetDataReferenceSource::DefinedName(name) => write_attribute(&mut xml, "name", name),
            WorksheetDataReferenceSource::Range { sheet, reference } => {
                write_attribute(&mut xml, "ref", reference.as_str());
                write_attribute(&mut xml, "sheet", sheet);
            }
        }
        if let Some(id) = &reference.relationship_id { write_attribute(&mut xml, "r:id", id); }
        xml.push_str("/>");
    }
    xml.push_str("</dataRefs></dataConsolidate>");
    Ok(xml)
}

fn write_true_attribute(xml: &mut String, name: &str, value: bool) {
    if value { write!(xml, " {name}=\"1\"").unwrap(); }
}

fn write_attribute(xml: &mut String, name: &str, value: &str) {
    write!(xml, " {name}=\"").unwrap();
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"), '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"), '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"), _ => xml.push(character),
        }
    }
    xml.push('"');
}

fn namespace_kind(result: ResolveResult<'_>) -> Result<NamespaceKind> {
    match result {
        ResolveResult::Unbound => Ok(NamespaceKind::Unbound),
        ResolveResult::Bound(namespace) if is_main_namespace(namespace.as_ref()) => Ok(NamespaceKind::Main),
        ResolveResult::Bound(namespace) if is_relationship_namespace(namespace.as_ref()) => Ok(NamespaceKind::Relationship),
        ResolveResult::Bound(_) => Ok(NamespaceKind::Other),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix {}", String::from_utf8_lossy(&prefix)
        ))),
    }
}

fn is_main_namespace(namespace: &[u8]) -> bool {
    namespace == TRANSITIONAL_MAIN || namespace == STRICT_MAIN
}

fn is_relationship_namespace(namespace: &[u8]) -> bool {
    namespace == TRANSITIONAL_REL || namespace == STRICT_REL
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn worksheet_child_position(local: &[u8]) -> Option<bool> {
    const BEFORE: &[&[u8]] = &[
        b"sheetPr", b"dimension", b"sheetViews", b"sheetFormatPr", b"cols", b"sheetData",
        b"sheetCalcPr", b"sheetProtection", b"protectedRanges", b"scenarios", b"autoFilter",
        b"sortState",
    ];
    const AFTER: &[&[u8]] = &[
        b"customSheetViews", b"mergeCells", b"phoneticPr", b"conditionalFormatting",
        b"dataValidations", b"hyperlinks", b"printOptions", b"pageMargins", b"pageSetup",
        b"headerFooter", b"rowBreaks", b"colBreaks", b"customProperties", b"cellWatches",
        b"ignoredErrors", b"smartTags", b"drawing", b"legacyDrawing", b"legacyDrawingHF",
        b"picture", b"oleObjects", b"controls", b"webPublishItems", b"tableParts", b"extLst",
    ];
    if BEFORE.contains(&local) { Some(false) } else if AFTER.contains(&local) { Some(true) } else { None }
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value { "true" | "1" => Ok(true), "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{name} must be an XML boolean"))) }
}

fn checked_xstring(value: String, name: &str) -> Result<String> {
    if value.is_empty() { return Err(invalid(format!("{name} must not be empty"))); }
    if value.chars().count() > MAX_XSTRING_CHARS {
        return Err(invalid(format!("{name} exceeds {MAX_XSTRING_CHARS} characters")));
    }
    Ok(value)
}

fn checked_relationship_id(value: String) -> Result<String> {
    if value.is_empty() || value.chars().count() > MAX_RELATIONSHIP_ID_CHARS {
        return Err(invalid("dataRef r:id is empty or exceeds the safety limit"));
    }
    Ok(value)
}

fn validate_reference_count(count: usize) -> Result<()> {
    if count > MAX_DATA_REFERENCES {
        Err(invalid(format!("dataRefs exceeds safety limit {MAX_DATA_REFERENCES}")))
    } else { Ok(()) }
}

fn valid_range_reference(value: &str) -> bool {
    let mut parts = value.split(':');
    let Some(first) = parts.next() else { return false; };
    let second = parts.next();
    parts.next().is_none() && valid_cell_reference(first) && second.is_none_or(valid_cell_reference)
}

fn valid_cell_reference(value: &str) -> bool {
    let value = value.strip_prefix('$').unwrap_or(value);
    let letter_count = value.bytes().take_while(u8::is_ascii_alphabetic).count();
    if !(1..=3).contains(&letter_count) { return false; }
    let (letters, row) = value.split_at(letter_count);
    let row = row.strip_prefix('$').unwrap_or(row);
    if row.is_empty() || !row.bytes().all(|byte| byte.is_ascii_digit()) || row.starts_with('0') {
        return false;
    }
    let column = letters.bytes().try_fold(0u32, |value, byte| {
        value.checked_mul(26)?.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
    });
    let row = row.parse::<u32>().ok();
    column.is_some_and(|column| column <= 16_384)
        && row.is_some_and(|row| (1..=1_048_576).contains(&row))
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() { Err(invalid(format!("duplicate {name} attribute"))) } else { Ok(()) }
}

fn invalid(message: impl Into<String>) -> OoxmlError { OoxmlError::InvalidFormat(message.into()) }

#[cfg(test)]
mod tests {
    use super::*;

    const T: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn worksheet(namespace: &str, body: &str) -> Vec<u8> {
        format!(r#"<worksheet xmlns="{namespace}" xmlns:r="{R}"><sheetData/>{body}</worksheet>"#).into_bytes()
    }

    #[test]
    fn parses_every_function_and_effective_defaults() {
        let values = ["average", "count", "countNums", "max", "min", "product", "stdDev", "stdDevp", "sum", "var", "varp"];
        for value in values {
            let xml = worksheet(T, &format!(r#"<dataConsolidate function="{value}"/>"#));
            assert_eq!(parse_worksheet_data_consolidation(&xml).unwrap().unwrap().function().as_str(), value);
        }
        let value = parse_worksheet_data_consolidation(&worksheet(T, "<dataConsolidate/>")).unwrap().unwrap();
        assert_eq!(value.function(), WorksheetDataConsolidationFunction::Sum);
        assert!(!value.left_labels() && !value.start_labels() && !value.top_labels() && !value.link());
        assert!(value.data_references().is_none());
    }

    #[test]
    fn canonical_writer_round_trips_range_name_relationships_and_flags() {
        let references = WorksheetDataReferences::new(vec![
            WorksheetDataReference::range("Sales & West", WorksheetDataConsolidationRangeReference::new("$A$1:XFD1048576").unwrap()).unwrap(),
            WorksheetDataReference::named("Workbook_Name").unwrap().with_relationship_id("rId7").unwrap(),
        ]).unwrap();
        let value = WorksheetDataConsolidation::new(WorksheetDataConsolidationFunction::CountNumbers, Some(references))
            .with_left_labels(true).with_start_labels(true).with_top_labels(true).with_link(true);
        let fragment = write_worksheet_data_consolidation(&value, WorksheetDataConsolidationConformance::Transitional).unwrap();
        assert_eq!(fragment, format!(r#"<dataConsolidate xmlns="{T}" xmlns:r="{R}" function="countNums" leftLabels="1" startLabels="1" topLabels="1" link="1"><dataRefs count="2"><dataRef ref="$A$1:XFD1048576" sheet="Sales &amp; West"/><dataRef name="Workbook_Name" r:id="rId7"/></dataRefs></dataConsolidate>"#));
        let parsed = parse_worksheet_data_consolidation(&worksheet(T, &fragment)).unwrap().unwrap();
        assert_eq!(parsed, value);
        assert_eq!(write_worksheet_data_consolidation(&parsed, WorksheetDataConsolidationConformance::Transitional).unwrap(), fragment);
    }

    #[test]
    fn supports_strict_mce_preservation_and_exact_schema_position() {
        let strict = format!(r#"<worksheet xmlns="{S}" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheetData/><dataConsolidate link="true"><dataRefs><dataRef name="N" r:id="rId1"/></dataRefs></dataConsolidate><phoneticPr fontId="0"/></worksheet>"#);
        let value = parse_worksheet_data_consolidation(strict.as_bytes()).unwrap().unwrap();
        assert!(value.link());
        assert_eq!(value.data_references().unwrap().declared_count(), None);

        let mce = format!(r#"<worksheet xmlns="{T}" xmlns:mc="{MC}" xmlns:x="urn:test" mc:Ignorable="x" mc:PreserveAttributes="x:*" mc:PreserveElements="x:keep"><sheetData/><x:wrapper mc:Ignorable="x" mc:ProcessContent="x:wrapper"><dataConsolidate><dataRefs count="1"><dataRef sheet="S" ref="A1:B2"/></dataRefs></dataConsolidate></x:wrapper></worksheet>"#);
        assert_eq!(parse_worksheet_data_consolidation(mce.as_bytes()).unwrap().unwrap().data_references().unwrap().references().len(), 1);

        assert!(parse_worksheet_data_consolidation(&worksheet(T, "<phoneticPr fontId=\"0\"/><dataConsolidate/>" )).is_err());
        assert!(parse_worksheet_data_consolidation(&worksheet(T, "<dataConsolidate/><sortState ref=\"A1\"/>" )).is_err());
        assert!(parse_worksheet_data_consolidation(&worksheet(T, "<extLst><ext uri=\"u\"><dataConsolidate/></ext></extLst>" )).unwrap().is_none());
    }

    #[test]
    fn rejects_malformed_counts_choices_spoofing_duplicates_unknowns_and_bounds() {
        let invalid_bodies = [
            r#"<dataConsolidate function="median"/>"#,
            r#"<dataConsolidate link="maybe"/>"#,
            r#"<dataConsolidate function="sum" function="count"/>"#,
            r#"<dataConsolidate bogus="1"/>"#,
            r#"<dataConsolidate><dataRefs count="2"><dataRef name="N"/></dataRefs></dataConsolidate>"#,
            r#"<dataConsolidate><dataRefs count="65537"/></dataConsolidate>"#,
            r#"<dataConsolidate><dataRefs><dataRef/></dataRefs></dataConsolidate>"#,
            r#"<dataConsolidate><dataRefs><dataRef name="N" sheet="S" ref="A1"/></dataRefs></dataConsolidate>"#,
            r#"<dataConsolidate><dataRefs><dataRef sheet="S"/></dataRefs></dataConsolidate>"#,
            r#"<dataConsolidate><dataRefs><dataRef sheet="S" ref="XFE1"/></dataRefs></dataConsolidate>"#,
            r#"<dataConsolidate><dataRefs><dataRef name="N"><child/></dataRef></dataRefs></dataConsolidate>"#,
            r#"<dataConsolidate><unknown/></dataConsolidate>"#,
            r#"<dataConsolidate><dataRefs/><dataRefs/></dataConsolidate>"#,
            r#"<x:dataConsolidate xmlns:x="urn:fake"/>"#,
            r#"<dataConsolidate xmlns:f="urn:fake" f:link="1"/>"#,
            r#"<dataConsolidate><dataRefs><dataRef xmlns:f="urn:fake" name="N" f:id="rId1"/></dataRefs></dataConsolidate>"#,
        ];
        for body in invalid_bodies {
            assert!(parse_worksheet_data_consolidation(&worksheet(T, body)).is_err(), "accepted {body}");
        }
    }

    // Provenance: LibreOffice's sc/source/filter/excel/ooxml-export-TODO.txt explicitly lists
    // dataConsolidate/dataRefs/dataRef, while POI exposes the same function enum only for pivot
    // tables. This deterministic package is therefore synthetic rather than mislabeled corpus data.
    #[test]
    fn immutable_worksheet_accessor_reads_synthetic_package() {
        use crate::xlsx::{Workbook, Worksheet, WorksheetInfo};
        use std::fs;
        use std::time::{SystemTime, UNIX_EPOCH};

        let sheet = format!(r#"<?xml version="1.0"?><worksheet xmlns="{T}"><sheetData/><dataConsolidate function="average" topLabels="1"><dataRefs count="1"><dataRef sheet="Input" ref="A1:C9"/></dataRefs></dataConsolidate></worksheet>"#);
        let package = make_package(&[
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#),
            ("_rels/.rels", r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
            ("xl/workbook.xml", &format!(r#"<?xml version="1.0"?><workbook xmlns="{T}" xmlns:r="{R}"><sheets><sheet name="Result" sheetId="1" r:id="rId1"/></sheets></workbook>"#)),
            ("xl/_rels/workbook.xml.rels", r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#),
            ("xl/worksheets/sheet1.xml", &sheet),
        ]);
        let path = std::env::temp_dir().join(format!("litchi-data-consolidate-{}-{}.xlsx", std::process::id(), SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_nanos()));
        fs::write(&path, package).unwrap();
        let workbook = Workbook::open(&path).unwrap();
        let mut worksheet = Worksheet::new(&workbook, WorksheetInfo {
            name: "Result".into(), relationship_id: "rId1".into(), sheet_id: 1,
            is_active: true, print_area: None, repeating_rows: None, repeating_columns: None,
        });
        worksheet.load_data().unwrap();
        let consolidation = worksheet.data_consolidation().unwrap();
        assert_eq!(consolidation.function(), WorksheetDataConsolidationFunction::Average);
        assert!(consolidation.top_labels());
        assert_eq!(consolidation.data_references().unwrap().references().len(), 1);
        fs::remove_file(path).unwrap();
    }

    fn make_package(entries: &[(&str, &str)]) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut central = Vec::new();
        for (name, value) in entries {
            let offset = bytes.len() as u32;
            let data = value.as_bytes();
            let crc = crc32(data);
            push_u32(&mut bytes, 0x04034b50); push_u16(&mut bytes, 20); push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0); push_u16(&mut bytes, 0); push_u16(&mut bytes, 0);
            push_u32(&mut bytes, crc); push_u32(&mut bytes, data.len() as u32); push_u32(&mut bytes, data.len() as u32);
            push_u16(&mut bytes, name.len() as u16); push_u16(&mut bytes, 0); bytes.extend_from_slice(name.as_bytes()); bytes.extend_from_slice(data);
            push_u32(&mut central, 0x02014b50); push_u16(&mut central, 20); push_u16(&mut central, 20);
            push_u16(&mut central, 0); push_u16(&mut central, 0); push_u16(&mut central, 0); push_u16(&mut central, 0);
            push_u32(&mut central, crc); push_u32(&mut central, data.len() as u32); push_u32(&mut central, data.len() as u32);
            push_u16(&mut central, name.len() as u16); push_u16(&mut central, 0); push_u16(&mut central, 0);
            push_u16(&mut central, 0); push_u16(&mut central, 0); push_u32(&mut central, 0); push_u32(&mut central, offset);
            central.extend_from_slice(name.as_bytes());
        }
        let central_offset = bytes.len() as u32;
        let central_size = central.len() as u32;
        bytes.extend_from_slice(&central);
        push_u32(&mut bytes, 0x06054b50); push_u16(&mut bytes, 0); push_u16(&mut bytes, 0);
        push_u16(&mut bytes, entries.len() as u16); push_u16(&mut bytes, entries.len() as u16);
        push_u32(&mut bytes, central_size); push_u32(&mut bytes, central_offset); push_u16(&mut bytes, 0);
        bytes
    }

    fn crc32(data: &[u8]) -> u32 {
        let mut crc = !0u32;
        for byte in data {
            crc ^= u32::from(*byte);
            for _ in 0..8 { crc = if crc & 1 != 0 { (crc >> 1) ^ 0xedb88320 } else { crc >> 1 }; }
        }
        !crc
    }
    fn push_u16(out: &mut Vec<u8>, value: u16) { out.extend_from_slice(&value.to_le_bytes()); }
    fn push_u32(out: &mut Vec<u8>, value: u32) { out.extend_from_slice(&value.to_le_bytes()); }
}
