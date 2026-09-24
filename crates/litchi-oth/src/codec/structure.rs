//! Bounded, read-only projection of ODF body structures that are not ordinary
//! paragraphs.  The source XML remains authoritative: this module only owns
//! semantic read values and never serializes or rewrites a structure.

use litchi_core::{Error, Result};
use litchi_odf_common::datatype::{Boolean, Duration};
use quick_xml::{XmlVersion, events::Event, name::QName, reader::Reader};
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

mod date;
mod namespace;
use namespace::{NamespaceId, Ns as ResolvedNs, SemanticNamespaceContext, namespace_declaration};

const MAX_STRUCTURE_TEXT: usize = 16 * 1024 * 1024;
const MAX_STRUCTURE_NODES: usize = 1_000_000;
const MAX_REPEAT: usize = 1_000_000;
const MAX_STRUCTURE_SITE_DEPTH: usize = 256;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Text,
    Table,
    Draw,
    Dr3d,
    Xlink,
    Svg,
    Xml,
    Dc,
    Meta,
    Xhtml,
    Fo,
    Style,
    Other,
    None,
}

#[derive(Clone, Debug)]
struct Attribute {
    local: String,
    namespace: Ns,
    namespace_uri: Option<Arc<str>>,
    value: String,
}

#[derive(Clone, Debug)]
enum Child {
    Element(Box<Node>),
    Text(String),
}

#[derive(Clone, Debug)]
struct Node {
    children: Vec<Child>,
    local: String,
    namespace: Ns,
    namespace_uri: Option<Arc<str>>,
    attributes: Vec<Attribute>,
}

impl Node {
    fn attribute(&self, namespace: Ns, local: &[u8]) -> Option<&str> {
        let local = std::str::from_utf8(local).ok()?;
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace == namespace
                    && (namespace != Ns::Other || attribute.namespace_uri.is_some())
                    && attribute.local == local
            })
            .map(|attribute| attribute.value.as_str())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RootKind {
    Table,
    Section,
    Note,
    Annotation,
    TrackedChanges,
    ChangedRegion,
    ChangeStart,
    ChangeEnd,
    Change,
    Index,
    Frame,
    Ruby,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ContextKind {
    OfficeBody,
    OfficeText,
    TextSection,
    TextParagraph,
    TextHeading,
    TextSpan,
    TextNote,
    TextNoteBody,
    TextTrackedChanges,
    TextChangedRegion,
    TextInsertion,
    TextDeletion,
    TextFormatChange,
    TextIndex,
    TextIndexBody,
    TextIndexSource,
    Table,
    TableColumns,
    TableColumn,
    TableRows,
    TableRow,
    TableCell,
    CoveredTableCell,
    OfficeAnnotation,
    DrawFrame,
    DrawTextBox,
    DrawImage,
    DrawObject,
    DrawObjectOle,
    DrawPlugin,
    DrawFloatingFrame,
    TextRuby,
    TextRubyBase,
    TextRubyText,
    BodyContainer,
    Other,
}

/// The admitted direct-element prefix of `office:text`.
///
/// ODF's `office-text-content-prelude` is an ordered sequence of optional
/// declarations.  Tracking changes may follow `office:forms` and must precede
/// the text and table declaration groups, while any declaration that
/// appears after ordinary body content remains malformed.  Keeping this state
/// per open element also makes nested look-alikes fail closed without treating
/// an arbitrary foreign subtree as part of the prelude.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OfficeTextPreludeState {
    Start,
    Forms,
    TrackedChanges,
    VariableDeclarations,
    SequenceDeclarations,
    UserFieldDeclarations,
    DdeConnectionDeclarations,
    AlphabeticalIndexAutoMarkFile,
    CalculationSettings,
    ContentValidations,
    LabelRanges,
    Body,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OfficeTextPreludeElement {
    Forms,
    TrackedChanges,
    VariableDeclarations,
    SequenceDeclarations,
    UserFieldDeclarations,
    DdeConnectionDeclarations,
    AlphabeticalIndexAutoMarkFile,
    CalculationSettings,
    ContentValidations,
    LabelRanges,
}

impl OfficeTextPreludeElement {
    const fn rank(self) -> u8 {
        match self {
            Self::Forms => 0,
            Self::TrackedChanges => 1,
            Self::VariableDeclarations => 2,
            Self::SequenceDeclarations => 3,
            Self::UserFieldDeclarations => 4,
            Self::DdeConnectionDeclarations => 5,
            Self::AlphabeticalIndexAutoMarkFile => 6,
            Self::CalculationSettings => 7,
            Self::ContentValidations => 8,
            Self::LabelRanges => 9,
        }
    }

    const fn state(self) -> OfficeTextPreludeState {
        match self {
            Self::Forms => OfficeTextPreludeState::Forms,
            Self::TrackedChanges => OfficeTextPreludeState::TrackedChanges,
            Self::VariableDeclarations => OfficeTextPreludeState::VariableDeclarations,
            Self::SequenceDeclarations => OfficeTextPreludeState::SequenceDeclarations,
            Self::UserFieldDeclarations => OfficeTextPreludeState::UserFieldDeclarations,
            Self::DdeConnectionDeclarations => OfficeTextPreludeState::DdeConnectionDeclarations,
            Self::AlphabeticalIndexAutoMarkFile => {
                OfficeTextPreludeState::AlphabeticalIndexAutoMarkFile
            },
            Self::CalculationSettings => OfficeTextPreludeState::CalculationSettings,
            Self::ContentValidations => OfficeTextPreludeState::ContentValidations,
            Self::LabelRanges => OfficeTextPreludeState::LabelRanges,
        }
    }
}

impl OfficeTextPreludeState {
    const fn rank(self) -> Option<u8> {
        match self {
            Self::Start => None,
            Self::Forms => Some(0),
            Self::TrackedChanges => Some(1),
            Self::VariableDeclarations => Some(2),
            Self::SequenceDeclarations => Some(3),
            Self::UserFieldDeclarations => Some(4),
            Self::DdeConnectionDeclarations => Some(5),
            Self::AlphabeticalIndexAutoMarkFile => Some(6),
            Self::CalculationSettings => Some(7),
            Self::ContentValidations => Some(8),
            Self::LabelRanges => Some(9),
            Self::Body => None,
        }
    }
}

impl ContextKind {
    fn permits_body_structure(self) -> bool {
        !matches!(self, Self::Other | Self::OfficeBody | Self::OfficeText)
    }
}

/// Additional semantic values projected from the OTH body.
pub(crate) struct BodyStructures {
    pub(crate) annotations: Vec<crate::annotation::Annotation>,
    pub(crate) changes: Vec<crate::change::Change>,
    pub(crate) change_tracking: Option<crate::change::ChangeTracking>,
    pub(crate) frames: Vec<crate::frame::Frame>,
    pub(crate) indexes: Vec<crate::index::Index>,
    pub(crate) notes: Vec<crate::note::Note>,
    pub(crate) rubies: Vec<crate::ruby::Ruby>,
    pub(crate) sections: Vec<crate::section::Section>,
    pub(crate) tables: Vec<crate::table::Table>,
    retained_bytes: usize,
}

/// A source-backed text span inside one admitted body-structure family.
///
/// The scanner intentionally keeps this separate from [`BodyStructures`].
/// The latter is a semantic projection and is lazy; this value is created only
/// when an edit explicitly resolves a structural selector.
#[derive(Clone, Debug)]
pub(crate) struct StructureSite {
    pub(crate) kind: EditableStructureKind,
    pub(crate) index: usize,
    pub(crate) identity: Option<String>,
    pub(crate) full: Range<usize>,
    pub(crate) text: Option<crate::codec::ReplacementSite>,
    pub(crate) text_value: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum EditableStructureKind {
    Section,
    Note,
    Annotation,
    Table,
}

impl BodyStructures {
    fn new() -> Self {
        Self {
            annotations: Vec::new(),
            changes: Vec::new(),
            change_tracking: None,
            frames: Vec::new(),
            indexes: Vec::new(),
            notes: Vec::new(),
            rubies: Vec::new(),
            sections: Vec::new(),
            tables: Vec::new(),
            retained_bytes: 0,
        }
    }

    fn account(&mut self, bytes: usize) -> Result<()> {
        self.retained_bytes = self.checked_total(bytes)?;
        Ok(())
    }

    fn ensure(&self, bytes: usize) -> Result<()> {
        self.checked_total(bytes).map(|_| ())
    }

    fn checked_total(&self, bytes: usize) -> Result<usize> {
        let total = self
            .retained_bytes
            .checked_add(bytes)
            .ok_or_else(|| Error::InvalidFormat("OTH projected body size overflow".to_string()))?;
        if total > MAX_STRUCTURE_TEXT {
            return invalid("OTH projected body structures exceed the aggregate text limit");
        }
        Ok(total)
    }
}

fn add_layout<T>(total: &mut usize, count: usize) -> Result<()> {
    let bytes = size_of::<T>()
        .checked_mul(count)
        .ok_or_else(|| Error::InvalidFormat("OTH projected body storage overflow".to_string()))?;
    *total = total
        .checked_add(bytes)
        .ok_or_else(|| Error::InvalidFormat("OTH projected body storage overflow".to_string()))?;
    Ok(())
}

fn measure_table_node(node: &Node) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::table::Table>(&mut total, 1)?;
    add_layout::<crate::table::TableProperties>(&mut total, 1)?;
    validate_table_properties(node)?;
    for local in [
        b"name".as_slice(),
        b"style-name".as_slice(),
        b"template-name".as_slice(),
        b"protection-key".as_slice(),
        b"protection-key-digest-algorithm".as_slice(),
        b"print-ranges".as_slice(),
    ] {
        add_string(&mut total, attr(node, Ns::Table, local))?;
    }
    add_string(&mut total, attr(node, Ns::Xml, b"id"))?;
    if let Some(source) = unique_direct_child(node, Ns::Table, "table-source", "table-source")? {
        add_layout::<crate::table::TableSource>(&mut total, 1)?;
        measure_table_source(source, &mut total)?;
    }
    for child in elements(node) {
        if child.namespace != Ns::Table {
            continue;
        }
        match child.local.as_str() {
            "table-column" | "table-columns" | "table-column-group" | "table-header-columns" => {
                measure_table_columns(child, &mut total)?;
            },
            "table-row" | "table-rows" | "table-row-group" | "table-header-rows" => {
                measure_table_rows(child, &mut total)?;
            },
            _ => {},
        }
    }
    Ok(total)
}

fn validate_table_properties(node: &Node) -> Result<()> {
    for local in [
        b"use-first-row-styles".as_slice(),
        b"use-last-row-styles".as_slice(),
        b"use-first-column-styles".as_slice(),
        b"use-last-column-styles".as_slice(),
        b"use-banding-rows-styles".as_slice(),
        b"use-banding-columns-styles".as_slice(),
        b"protected".as_slice(),
        b"print".as_slice(),
        b"is-sub-table".as_slice(),
    ] {
        optional_schema_bool_attr(node, Ns::Table, local)?;
    }
    if let Some(value) = attr(node, Ns::Table, b"protection-key-digest-algorithm") {
        validate_any_iri(value, "table:protection-key-digest-algorithm")?;
    }
    Ok(())
}

fn measure_table_source(source: &Node, total: &mut usize) -> Result<()> {
    if elements(source).next().is_some() {
        return invalid("OTH table-source must be empty");
    }
    if required_attr(source, Ns::Xlink, b"type", "xlink:type")? != "simple" {
        return invalid("OTH table-source xlink:type must be 'simple'");
    }
    let href = required_attr(source, Ns::Xlink, b"href", "xlink:href")?;
    validate_any_iri(href, "xlink:href")?;
    if let Some(value) = attr(source, Ns::Table, b"mode")
        && !matches!(value, "copy-all" | "copy-results-only")
    {
        return invalid("OTH table-source mode is invalid");
    }
    if let Some(value) = attr(source, Ns::Xlink, b"actuate")
        && value != "onRequest"
    {
        return invalid("OTH table-source xlink:actuate is invalid");
    }
    add_string(total, Some(href))?;
    add_string(total, attr(source, Ns::Table, b"table-name"))?;
    add_string(total, attr(source, Ns::Table, b"filter-name"))?;
    add_string(total, attr(source, Ns::Table, b"filter-options"))?;
    if let Some(value) = attr(source, Ns::Table, b"refresh-delay") {
        add_duration_lexical_bytes(total, value)?;
    }
    Ok(())
}

fn measure_table_columns(node: &Node, total: &mut usize) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table-column" {
        add_layout::<crate::table::Column>(total, 1)?;
        optional_count_attr(node, b"number-columns-repeated")?;
        optional_visibility_attr(node)?;
        add_string(total, attr(node, Ns::Table, b"style-name"))?;
        add_string(total, attr(node, Ns::Table, b"default-cell-style-name"))?;
        add_string(total, attr(node, Ns::Xml, b"id"))?;
        return Ok(());
    }
    for child in elements(node) {
        if child.namespace == Ns::Table
            && matches!(
                child.local.as_str(),
                "table-column" | "table-columns" | "table-column-group" | "table-header-columns"
            )
        {
            measure_table_columns(child, total)?;
        }
    }
    Ok(())
}

fn measure_table_rows(node: &Node, total: &mut usize) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table-row" {
        add_layout::<crate::table::Row>(total, 1)?;
        optional_count_attr(node, b"number-rows-repeated")?;
        optional_visibility_attr(node)?;
        add_string(total, attr(node, Ns::Table, b"style-name"))?;
        add_string(total, attr(node, Ns::Table, b"default-cell-style-name"))?;
        add_string(total, attr(node, Ns::Xml, b"id"))?;
        for child in elements(node) {
            if child.namespace == Ns::Table
                && matches!(child.local.as_str(), "table-cell" | "covered-table-cell")
            {
                measure_table_cell(child, total)?;
            }
        }
        return Ok(());
    }
    for child in elements(node) {
        if child.namespace == Ns::Table
            && matches!(
                child.local.as_str(),
                "table-row" | "table-rows" | "table-row-group" | "table-header-rows"
            )
        {
            measure_table_rows(child, total)?;
        }
    }
    Ok(())
}

fn measure_table_cell(node: &Node, total: &mut usize) -> Result<()> {
    add_layout::<crate::table::Cell>(total, 1)?;
    optional_count_attr(node, b"number-columns-repeated")?;
    optional_count_attr(node, b"number-columns-spanned")?;
    optional_count_attr(node, b"number-rows-spanned")?;
    optional_count_attr(node, b"number-matrix-columns-spanned")?;
    optional_count_attr(node, b"number-matrix-rows-spanned")?;
    optional_schema_bool_attr(node, Ns::Table, b"protect")?;
    optional_schema_bool_attr(node, Ns::Table, b"protected")?;
    for local in [
        b"style-name".as_slice(),
        b"formula".as_slice(),
        b"content-validation-name".as_slice(),
    ] {
        add_string(total, attr(node, Ns::Table, local))?;
    }
    add_string(total, attr(node, Ns::Xml, b"id"))?;
    add_string(total, attr(node, Ns::Office, b"value-type"))?;
    add_string(total, attr(node, Ns::Office, b"value"))?;
    add_len(total, plain_text_len(node)?)?;
    if attr(node, Ns::Xhtml, b"about").is_some() || attr(node, Ns::Xhtml, b"property").is_some() {
        add_layout::<crate::table::InContentMeta>(total, 1)?;
    }
    measure_in_content_meta(node, total)?;
    measure_cell_value(node, total)?;
    for child in elements(node) {
        if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
            add_layout::<crate::paragraph::Paragraph>(total, 1)?;
            add_len(total, plain_text_len(child)?)?;
        }
    }
    Ok(())
}

fn measure_in_content_meta(node: &Node, total: &mut usize) -> Result<()> {
    let about = attr(node, Ns::Xhtml, b"about");
    let property = attr(node, Ns::Xhtml, b"property");
    let has_other =
        attr(node, Ns::Xhtml, b"datatype").is_some() || attr(node, Ns::Xhtml, b"content").is_some();
    if about.is_none() && property.is_none() {
        if has_other {
            return invalid("OTH RDFa metadata requires xhtml:about and xhtml:property");
        }
        return Ok(());
    }
    let about = about.ok_or_else(|| {
        Error::InvalidFormat("OTH RDFa metadata requires xhtml:about".to_string())
    })?;
    let property = property.ok_or_else(|| {
        Error::InvalidFormat("OTH RDFa metadata requires xhtml:property".to_string())
    })?;
    validate_uri_or_safe_curie(about, "xhtml:about")?;
    validate_curies(property, "xhtml:property")?;
    if let Some(datatype) = attr(node, Ns::Xhtml, b"datatype") {
        validate_curie(datatype, "xhtml:datatype")?;
    }
    add_string(total, Some(about))?;
    add_string(total, Some(property))?;
    add_string(total, attr(node, Ns::Xhtml, b"datatype"))?;
    add_string(total, attr(node, Ns::Xhtml, b"content"))?;
    Ok(())
}

fn measure_cell_value(node: &Node, total: &mut usize) -> Result<()> {
    let Some(value_type) = attr(node, Ns::Office, b"value-type") else {
        reject_cell_value_companions(node, None)?;
        return Ok(());
    };
    reject_cell_value_companions(node, Some(value_type))?;
    add_layout::<crate::table::CellValue>(total, 1)?;
    match value_type {
        "float" | "percentage" | "currency" => {
            let value = required_attr(node, Ns::Office, b"value", "office:value")?;
            validate_double(value, "office:value")?;
            add_string(total, Some(value))?;
            if value_type == "currency" {
                add_string(total, attr(node, Ns::Office, b"currency"))?;
            }
        },
        "date" => {
            let value = required_attr(node, Ns::Office, b"date-value", "office:date-value")?;
            add_string(total, Some(value))?;
        },
        "time" => {
            let value = required_attr(node, Ns::Office, b"time-value", "office:time-value")?;
            add_duration_lexical_bytes(total, value)?;
        },
        "boolean" => {
            let value = required_attr(node, Ns::Office, b"boolean-value", "office:boolean-value")?;
            parse_odf_bool(value, "office:boolean-value")?;
            add_string(total, Some(value))?;
        },
        "string" | "error" => add_string(total, attr(node, Ns::Office, b"string-value"))?,
        _ => return invalid("OTH office:value-type value is invalid"),
    }
    Ok(())
}

fn measure_index_node(node: &Node, source: &Node) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::index::Index>(&mut total, 1)?;
    let name = required_attr(node, Ns::Text, b"name", "text:name")?;
    add_string(&mut total, Some(name))?;
    optional_schema_bool_attr(node, Ns::Text, b"protected")?;
    if let Some(value) = attr(node, Ns::Text, b"protection-key-digest-algorithm") {
        validate_any_iri(value, "text:protection-key-digest-algorithm")?;
    }
    add_string(&mut total, attr(node, Ns::Text, b"style-name"))?;
    add_string(&mut total, attr(node, Ns::Text, b"protection-key"))?;
    add_string(
        &mut total,
        attr(node, Ns::Text, b"protection-key-digest-algorithm"),
    )?;
    add_string(&mut total, attr(node, Ns::Xml, b"id"))?;
    add_len(&mut total, plain_text_len(source)?)?;
    if let Some(body) = unique_direct_child(node, Ns::Text, "index-body", "index-body")? {
        add_len(&mut total, plain_text_len(body)?)?;
    } else {
        add_len(&mut total, 0)?;
    }
    add_layout::<crate::index::IndexSource>(&mut total, 1)?;
    measure_index_source(node, source, &mut total)?;
    Ok(total)
}

fn measure_index_source(node: &Node, source: &Node, total: &mut usize) -> Result<()> {
    let scope = attr(source, Ns::Text, b"index-scope");
    if let Some(value) = scope
        && !matches!(value, "document" | "chapter")
    {
        return invalid("OTH index-scope value is invalid");
    }
    optional_schema_bool_attr(source, Ns::Text, b"relative-tab-stop-position")?;
    match node.local.as_str() {
        "table-of-content" => {
            optional_positive_attr(source, b"outline-level")?;
            for local in [
                b"use-outline-level".as_slice(),
                b"use-index-marks".as_slice(),
                b"use-index-source-styles".as_slice(),
            ] {
                optional_schema_bool_attr(source, Ns::Text, local)?;
            }
        },
        "illustration-index" | "table-index" => {
            optional_schema_bool_attr(source, Ns::Text, b"use-caption")?;
            if let Some(value) = attr(source, Ns::Text, b"caption-sequence-format")
                && !matches!(value, "text" | "category-and-value" | "caption")
            {
                return invalid("OTH caption-sequence-format value is invalid");
            }
            add_string(total, attr(source, Ns::Text, b"caption-sequence-name"))?;
        },
        "object-index" => {
            for local in [
                b"use-spreadsheet-objects".as_slice(),
                b"use-math-objects".as_slice(),
                b"use-draw-objects".as_slice(),
                b"use-chart-objects".as_slice(),
                b"use-other-objects".as_slice(),
            ] {
                optional_schema_bool_attr(source, Ns::Text, local)?;
            }
        },
        "user-index" => {
            let index_name = required_attr(source, Ns::Text, b"index-name", "text:index-name")?;
            if index_name.is_empty() {
                return invalid("OTH text:index-name must not be empty");
            }
            add_string(total, Some(index_name))?;
            for local in [
                b"use-index-marks".as_slice(),
                b"use-index-source-styles".as_slice(),
                b"use-graphics".as_slice(),
                b"use-tables".as_slice(),
                b"use-floating-frames".as_slice(),
                b"use-objects".as_slice(),
                b"copy-outline-levels".as_slice(),
            ] {
                optional_schema_bool_attr(source, Ns::Text, local)?;
            }
        },
        "alphabetical-index" => {
            for local in [
                b"ignore-case".as_slice(),
                b"alphabetical-separators".as_slice(),
                b"combine-entries".as_slice(),
                b"combine-entries-with-dash".as_slice(),
                b"combine-entries-with-pp".as_slice(),
                b"use-keys-as-entries".as_slice(),
                b"capitalize-entries".as_slice(),
                b"comma-separated".as_slice(),
            ] {
                optional_schema_bool_attr(source, Ns::Text, local)?;
            }
            for (namespace, local) in [
                (Ns::Text, b"main-entry-style-name".as_slice()),
                (Ns::Fo, b"language".as_slice()),
                (Ns::Fo, b"country".as_slice()),
                (Ns::Fo, b"script".as_slice()),
                (Ns::Style, b"rfc-language-tag".as_slice()),
                (Ns::Text, b"sort-algorithm".as_slice()),
            ] {
                add_string(total, attr(source, namespace, local))?;
            }
        },
        "bibliography" | "bibliography-index" => {},
        _ => return invalid("OTH index source family is invalid"),
    }
    Ok(())
}

fn measure_change_node(node: &Node, kind_name: &str) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::change::Change>(&mut total, 1)?;
    let is_region = node.namespace == Ns::Text && node.local == "changed-region";
    if is_region {
        let xml_id = attr(node, Ns::Xml, b"id");
        let region_id = attr(node, Ns::Text, b"id");
        if let Some(value) = xml_id {
            validate_ncname(value, "xml:id")?;
            add_string(&mut total, Some(value))?;
        }
        if let Some(value) = region_id {
            validate_ncname(value, "text:id")?;
            if let Some(xml_id) = xml_id
                && xml_id != value
            {
                return invalid("OTH changed-region text:id must equal xml:id");
            }
            add_string(&mut total, Some(value))?;
        } else if xml_id.is_none() {
            return invalid("OTH changed-region requires xml:id");
        }
        add_string(&mut total, xml_id.or(region_id))?;
        let content = direct_change_child(node)?;
        let info = validate_change_content(content)?;
        measure_change_info(info, &mut total)?;
        let creator_len = plain_text_len(elements(info).next().ok_or_else(|| {
            Error::InvalidFormat("OTH office:change-info children have invalid order".to_string())
        })?)?;
        let date_len = plain_text_len(elements(info).nth(1).ok_or_else(|| {
            Error::InvalidFormat("OTH office:change-info children have invalid order".to_string())
        })?)?;
        add_len(&mut total, creator_len)?;
        add_len(&mut total, date_len)?;
        if content.local == "deletion" {
            add_len(&mut total, change_text_len(node)?)?;
        } else {
            add_len(&mut total, 0)?;
        }
    } else {
        let marker = required_attr(node, Ns::Text, b"change-id", "text:change-id")?;
        validate_ncname(marker, "text:change-id")?;
        add_string(&mut total, Some(marker))?;
        add_string(&mut total, Some(marker))?;
        add_len(&mut total, plain_text_len(node)?)?;
    }
    if kind_name == "change" {
        add_string(&mut total, Some(kind_name))?;
    }
    Ok(total)
}

fn measure_change_info(node: &Node, total: &mut usize) -> Result<()> {
    add_layout::<crate::change::ChangeInfo>(total, 1)?;
    if node
        .children
        .iter()
        .any(|child| matches!(child, Child::Text(text) if !text.trim().is_empty()))
    {
        return invalid("OTH office:change-info has unexpected direct text");
    }
    let mut children = elements(node);
    let Some(creator) = children.next() else {
        return invalid("OTH office:change-info children have invalid order");
    };
    let Some(date) = children.next() else {
        return invalid("OTH office:change-info children have invalid order");
    };
    if creator.namespace != Ns::Dc
        || creator.local != "creator"
        || date.namespace != Ns::Dc
        || date.local != "date"
    {
        return invalid("OTH office:change-info children have invalid order");
    }
    let creator_len = plain_text_len(creator)?;
    let date_len = plain_text_len(date)?;
    add_len(total, creator_len)?;
    add_len(total, date_len)?;
    for paragraph in children {
        if paragraph.namespace != Ns::Text || paragraph.local != "p" {
            return invalid("OTH office:change-info children have invalid order");
        }
        add_layout::<crate::paragraph::Paragraph>(total, 1)?;
        add_len(total, plain_text_len(paragraph)?)?;
    }
    Ok(())
}

fn add_len(total: &mut usize, bytes: usize) -> Result<()> {
    let bytes = bytes
        .checked_add(size_of::<String>())
        .ok_or_else(|| Error::InvalidFormat("OTH projected body size overflow".to_string()))?;
    add_len_without_string_slot(total, bytes)
}

fn add_len_without_string_slot(total: &mut usize, bytes: usize) -> Result<()> {
    *total = total
        .checked_add(bytes)
        .ok_or_else(|| Error::InvalidFormat("OTH projected body size overflow".to_string()))?;
    Ok(())
}

fn plain_text_len(node: &Node) -> Result<usize> {
    if matches!(node.namespace, Ns::Other | Ns::None) {
        return Ok(0);
    }
    if node.namespace == Ns::Text {
        match node.local.as_str() {
            "s" => {
                let count = node.attribute(Ns::Text, b"c").map_or(Ok(1), |value| {
                    value.parse::<usize>().map_err(|_| {
                        Error::InvalidFormat("OTH text:s count is invalid".to_string())
                    })
                })?;
                if count == 0 || count > MAX_STRUCTURE_TEXT {
                    return invalid("OTH text:s count is outside the supported range");
                }
                return Ok(count);
            },
            "tab" => return Ok(1),
            "line-break" => return Ok(1),
            _ => {},
        }
    }
    let mut total = 0usize;
    for child in &node.children {
        match child {
            Child::Element(element) => add_len(&mut total, plain_text_len(element)?)?,
            Child::Text(text) => add_len(&mut total, text.len())?,
        }
    }
    if total > MAX_STRUCTURE_TEXT {
        return invalid("OTH structure text exceeds the limit");
    }
    Ok(total)
}

fn plain_text_content_len(node: &Node) -> Result<usize> {
    if matches!(node.namespace, Ns::Other | Ns::None) {
        return Ok(0);
    }
    if node.namespace == Ns::Text {
        match node.local.as_str() {
            "s" => {
                let count = node.attribute(Ns::Text, b"c").map_or(Ok(1), |value| {
                    value.parse::<usize>().map_err(|_| {
                        Error::InvalidFormat("OTH text:s count is invalid".to_string())
                    })
                })?;
                if count == 0 || count > MAX_STRUCTURE_TEXT {
                    return invalid("OTH text:s count is outside the supported range");
                }
                return Ok(count);
            },
            "tab" | "line-break" => return Ok(1),
            _ => {},
        }
    }
    let mut total = 0usize;
    for child in &node.children {
        let length = match child {
            Child::Element(element) => plain_text_content_len(element)?,
            Child::Text(text) => text.len(),
        };
        total = total
            .checked_add(length)
            .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    }
    if total > MAX_STRUCTURE_TEXT {
        return invalid("OTH structure text exceeds the limit");
    }
    Ok(total)
}

fn add_duration_lexical_bytes(total: &mut usize, value: &str) -> Result<()> {
    let (bytes, components) = duration_lexical_bytes(value)?;
    let slots = components
        .checked_add(1)
        .and_then(|count| size_of::<String>().checked_mul(count))
        .ok_or_else(|| Error::InvalidFormat("OTH duration storage size overflow".to_string()))?;
    add_len_without_string_slot(total, bytes)?;
    add_len_without_string_slot(total, slots)?;
    Ok(())
}

fn duration_lexical_bytes(value: &str) -> Result<(usize, usize)> {
    if value.len() > 1_048_576 {
        return invalid("OTH duration exceeds 1 MiB");
    }
    let unsigned = value.strip_prefix('-').unwrap_or(value);
    let body = unsigned
        .strip_prefix('P')
        .ok_or_else(|| Error::InvalidFormat("OTH duration is invalid".to_string()))?;
    let bytes = body.as_bytes();
    let mut position = 0usize;
    let mut in_time = false;
    let mut last_rank = 0u8;
    let mut component_count = 0usize;
    let mut time_component_count = 0usize;
    let mut component_bytes = 0usize;
    while position < bytes.len() {
        if bytes[position] == b'T' {
            if in_time {
                return invalid("OTH duration has duplicate T designators");
            }
            in_time = true;
            last_rank = 0;
            position += 1;
            continue;
        }
        let start = position;
        while position < bytes.len() && bytes[position].is_ascii_digit() {
            position += 1;
        }
        if position == start {
            return invalid("OTH duration has an invalid numeric component");
        }
        if position < bytes.len() && bytes[position] == b'.' {
            position += 1;
            let fraction_start = position;
            while position < bytes.len() && bytes[position].is_ascii_digit() {
                position += 1;
            }
            if position == fraction_start {
                return invalid("OTH duration has an empty fractional component");
            }
        }
        if position == bytes.len() {
            return invalid("OTH duration component has no designator");
        }
        let component = &bytes[start..position];
        let designator = bytes[position];
        position += 1;
        let (rank, allows_fraction) = match (in_time, designator) {
            (false, b'Y') => (1, false),
            (false, b'M') => (2, false),
            (false, b'D') => (3, false),
            (true, b'H') => (1, false),
            (true, b'M') => (2, false),
            (true, b'S') => (3, true),
            _ => return invalid("OTH duration designator is invalid"),
        };
        if !allows_fraction && component.contains(&b'.') {
            return invalid("OTH duration fraction is only valid for seconds");
        }
        if rank <= last_rank {
            return invalid("OTH duration components are unordered or duplicated");
        }
        last_rank = rank;
        component_count += 1;
        component_bytes = component_bytes
            .checked_add(component.len())
            .ok_or_else(|| Error::InvalidFormat("OTH duration size overflow".to_string()))?;
        if in_time {
            time_component_count += 1;
        }
    }
    if component_count == 0 || (in_time && time_component_count == 0) {
        return invalid("OTH duration has no components");
    }
    let bytes = value
        .len()
        .checked_add(component_bytes)
        .ok_or_else(|| Error::InvalidFormat("OTH duration size overflow".to_string()))?;
    Ok((bytes, component_count))
}

fn change_text_len(node: &Node) -> Result<usize> {
    if node.namespace != Ns::Text || node.local != "changed-region" {
        return plain_text_len(node);
    }
    let content = direct_change_child(node)?;
    validate_change_content(content)?;
    if content.local != "deletion" {
        return Ok(0);
    }
    let mut info_seen = false;
    let mut total = 0usize;
    for child in &content.children {
        match child {
            Child::Element(child)
                if child.namespace == Ns::Office && child.local == "change-info" =>
            {
                info_seen = true;
            },
            Child::Element(child) if info_seen => add_len(&mut total, plain_text_len(child)?)?,
            Child::Text(value) if !value.trim().is_empty() => {
                return invalid("OTH deletion contains unexpected direct text");
            },
            Child::Text(_) | Child::Element(_) => {},
        }
    }
    Ok(total)
}

fn measure_section_node(node: &Node) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::section::Section>(&mut total, 1)?;
    bool_attr(node, Ns::Text, b"protected")?;
    optional_bool_attr(node, Ns::Text, b"display")?;
    add_string(&mut total, attr(node, Ns::Text, b"name"))?;
    add_string(&mut total, attr(node, Ns::Text, b"style-name"))?;
    add_string(&mut total, attr(node, Ns::Text, b"condition"))?;
    add_len(&mut total, plain_text_len(node)?)?;
    for child in elements(node) {
        if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
            add_layout::<crate::paragraph::Paragraph>(&mut total, 1)?;
            add_len(&mut total, plain_text_len(child)?)?;
        }
    }
    Ok(total)
}

fn measure_note_node(node: &Node) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::note::Note>(&mut total, 1)?;
    let class = attr(node, Ns::Text, b"note-class").unwrap_or_default();
    add_string(&mut total, Some(class))?;
    add_string(&mut total, attr(node, Ns::Text, b"id"))?;
    let citation_node = direct_child(node, Ns::Text, "note-citation");
    let body_node = direct_child(node, Ns::Text, "note-body");
    if let Some(citation) = citation_node {
        add_string(&mut total, attr(citation, Ns::Text, b"label"))?;
        add_len(&mut total, plain_text_len(citation)?)?;
    } else {
        add_len(&mut total, 0)?;
    }
    if let Some(body) = body_node {
        add_len(&mut total, plain_text_len(body)?)?;
        for child in elements(body) {
            if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
                add_layout::<crate::paragraph::Paragraph>(&mut total, 1)?;
                add_len(&mut total, plain_text_len(child)?)?;
            }
        }
    } else {
        add_len(&mut total, 0)?;
    }
    Ok(total)
}

fn measure_annotation_node(node: &Node) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::annotation::Annotation>(&mut total, 1)?;
    optional_bool_attr(node, Ns::Office, b"display")?;
    add_string(&mut total, attr(node, Ns::Office, b"name"))?;
    if let Some(creator) = direct_child(node, Ns::Dc, "creator") {
        add_len(&mut total, plain_text_len(creator)?)?;
    }
    if let Some(date) = direct_child(node, Ns::Dc, "date") {
        add_len(&mut total, plain_text_len(date)?)?;
    }
    if let Some(date_string) = direct_child(node, Ns::Meta, "date-string") {
        add_len(&mut total, plain_text_len(date_string)?)?;
    }
    if let Some(initials) = direct_child(node, Ns::Meta, "creator-initials")
        .or_else(|| direct_child(node, Ns::Text, "sender-initials"))
    {
        add_len(&mut total, plain_text_len(initials)?)?;
    }
    add_len(&mut total, annotation_text_len(node)?)?;
    Ok(total)
}

fn annotation_text_len(node: &Node) -> Result<usize> {
    let mut total = 0usize;
    for child in elements(node) {
        if child.namespace == Ns::Dc
            || child.namespace == Ns::Meta
            || (child.namespace == Ns::Text && child.local == "sender-initials")
        {
            continue;
        }
        total = total
            .checked_add(plain_text_len(child)?)
            .ok_or_else(|| Error::InvalidFormat("OTH annotation text size overflow".to_string()))?;
    }
    Ok(total)
}

fn measure_frame_node(node: &Node) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::frame::Frame>(&mut total, 1)?;
    add_string(&mut total, attr(node, Ns::Draw, b"name"))?;
    add_string(&mut total, attr(node, Ns::Draw, b"style-name"))?;
    add_string(&mut total, attr(node, Ns::Text, b"anchor-type"))?;
    add_string(&mut total, attr(node, Ns::Svg, b"x"))?;
    add_string(&mut total, attr(node, Ns::Svg, b"y"))?;
    add_string(&mut total, attr(node, Ns::Svg, b"width"))?;
    add_string(&mut total, attr(node, Ns::Svg, b"height"))?;
    let href = find_descendant(node, Ns::Draw, "image")
        .or_else(|| find_descendant(node, Ns::Draw, "object"))
        .or_else(|| find_descendant(node, Ns::Draw, "object-ole"))
        .or_else(|| find_descendant(node, Ns::Draw, "plugin"))
        .or_else(|| find_descendant(node, Ns::Draw, "floating-frame"))
        .and_then(|value| attr(value, Ns::Xlink, b"href"));
    add_string(&mut total, href)?;
    add_len(&mut total, plain_text_len(node)?)?;
    Ok(total)
}

fn measure_ruby_node(node: &Node) -> Result<usize> {
    let mut total = 0usize;
    add_layout::<crate::ruby::Ruby>(&mut total, 1)?;
    add_string(&mut total, attr(node, Ns::Text, b"style-name"))?;
    let base = direct_child(node, Ns::Text, "ruby-base");
    let text = direct_child(node, Ns::Text, "ruby-text");
    add_string(
        &mut total,
        text.and_then(|value| attr(value, Ns::Text, b"style-name")),
    )?;
    add_len(
        &mut total,
        base.map(plain_text_len).transpose()?.unwrap_or_default(),
    )?;
    add_len(
        &mut total,
        text.map(plain_text_len).transpose()?.unwrap_or_default(),
    )?;
    Ok(total)
}

struct Budget {
    nodes: usize,
    text_bytes: usize,
    storage_bytes: usize,
}

impl Budget {
    fn count_node(&mut self) -> Result<()> {
        let nodes = self.nodes.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure node count overflow".to_string())
        })?;
        if nodes > MAX_STRUCTURE_NODES {
            return invalid("OTH body structure node count exceeds the limit");
        }
        self.nodes = nodes;
        Ok(())
    }

    fn node(&mut self) -> Result<()> {
        self.count_node()?;
        self.storage(size_of::<Node>())?;
        Ok(())
    }

    fn text(&mut self, bytes: usize) -> Result<()> {
        let text_bytes = self.text_bytes.checked_add(bytes).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure text size overflow".to_string())
        })?;
        self.check_total(
            text_bytes,
            self.storage_bytes,
            "OTH body structure text exceeds the limit",
        )?;
        self.text_bytes = text_bytes;
        Ok(())
    }

    fn storage(&mut self, bytes: usize) -> Result<()> {
        let storage_bytes = self.storage_bytes.checked_add(bytes).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure storage size overflow".to_string())
        })?;
        self.check_total(
            self.text_bytes,
            storage_bytes,
            "OTH body structure storage exceeds the limit",
        )?;
        self.storage_bytes = storage_bytes;
        Ok(())
    }

    fn text_after_raw(&mut self, raw: usize, decoded: usize) -> Result<()> {
        self.text(decoded.saturating_sub(raw))
    }

    fn check_total(&self, text_bytes: usize, storage_bytes: usize, message: &str) -> Result<()> {
        let total = text_bytes.checked_add(storage_bytes).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure budget overflow".to_string())
        })?;
        if total > MAX_STRUCTURE_TEXT {
            return invalid(message);
        }
        Ok(())
    }

    fn reserve_vec<T>(
        &mut self,
        values: &mut Vec<T>,
        required_len: usize,
        resource: &'static str,
    ) -> Result<()> {
        reserve_vec_charged(values, required_len, resource, |bytes| self.storage(bytes))
    }

    fn reserve_vec_exact<T>(
        &mut self,
        values: &mut Vec<T>,
        required_len: usize,
        resource: &'static str,
    ) -> Result<()> {
        reserve_vec_exact_charged(values, required_len, resource, |bytes| self.storage(bytes))
    }

    fn reserve_set<T>(
        &mut self,
        values: &mut std::collections::HashSet<T>,
        required_len: usize,
        resource: &'static str,
    ) -> Result<()>
    where
        T: Eq + std::hash::Hash,
    {
        let current_capacity = values.capacity();
        let target_capacity = geometric_capacity(current_capacity, required_len)?;
        if target_capacity <= current_capacity {
            return Ok(());
        }
        let growth = target_capacity
            .checked_sub(current_capacity)
            .ok_or_else(|| Error::InvalidFormat("OTH set capacity overflow".to_string()))?;
        let bucket_bytes = size_of::<T>()
            .checked_add(size_of::<usize>())
            .and_then(|bytes| bytes.checked_mul(growth))
            .ok_or_else(|| Error::InvalidFormat("OTH set storage size overflow".to_string()))?;
        self.storage(bucket_bytes)?;
        let additional = target_capacity
            .checked_sub(values.len())
            .ok_or_else(|| Error::InvalidFormat("OTH set reservation underflow".to_string()))?;
        values
            .try_reserve(additional)
            .map_err(|source| Error::Allocation { resource, source })
    }
}

/// Projects tables, sections, notes, annotations, review metadata, indexes,
/// frames, and ruby pairs from a validated `content.xml` source.
pub(crate) fn project_structures(xml: &str) -> Result<BodyStructures> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut namespaces = SemanticNamespaceContext::new();
    let mut budget = Budget {
        nodes: 0,
        text_bytes: 0,
        storage_bytes: 0,
    };
    let tracked_changes = document_tracked_changes(xml, &mut budget)?;
    validate_xml_id_and_change_references(xml, tracked_changes.is_some(), &mut budget)?;
    let mut stack = Vec::<Node>::new();
    let mut contexts = Vec::<ContextKind>::new();
    let mut body_text_depth = 0usize;
    let mut output = BodyStructures::new();
    output.change_tracking = tracked_changes.clone();
    loop {
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                debug_assert_eq!(namespaces.depth(), contexts.len() + 1);
                let (_namespace, context, root) = classify_element(&namespaces, start.name())?;
                let owns_text = context == ContextKind::OfficeText
                    && contexts.last() == Some(&ContextKind::OfficeBody);
                let capture_root = stack.is_empty()
                    && body_text_depth > 0
                    && owns_structure_root(&contexts)
                    && root.is_some_and(|root| {
                        !root_requires_tracked_changes(root) || tracked_changes.is_some()
                    });
                if !stack.is_empty() || capture_root {
                    reserve_stack_node(&mut stack, &mut budget)?;
                    let node = parse_node(&namespaces, reader.decoder(), &start, &mut budget)?;
                    stack.push(node);
                }
                if owns_text {
                    body_text_depth = body_text_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("OTH office:text depth overflow".to_string())
                    })?;
                }
                if contexts.len() >= MAX_STRUCTURE_NODES {
                    return invalid("OTH body structure nesting exceeds the limit");
                }
                let context_len = contexts.len().checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("OTH body context length overflow".to_string())
                })?;
                budget.reserve_vec(
                    &mut contexts,
                    context_len,
                    "OTH body structure context stack",
                )?;
                contexts.push(context);
            },
            Event::Empty(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                let (_namespace, _context, root) = classify_element(&namespaces, start.name())?;
                let capture_root = stack.is_empty()
                    && body_text_depth > 0
                    && owns_structure_root(&contexts)
                    && root.is_some_and(|root| {
                        !root_requires_tracked_changes(root) || tracked_changes.is_some()
                    });
                if !stack.is_empty() || capture_root {
                    let node = parse_node(&namespaces, reader.decoder(), &start, &mut budget)?;
                    if stack.is_empty() {
                        if let Some(kind) = root_kind(node.namespace, &node.local) {
                            finish_root(node, kind, &mut output)?;
                        }
                    } else {
                        append_node(&mut stack, node, &mut budget)?;
                    }
                }
                namespaces.end()?;
            },
            Event::End(end) => {
                let _ = classify_element(&namespaces, end.name())?;
                if let Some(node) = stack.pop() {
                    if !stack.is_empty() {
                        append_node(&mut stack, node, &mut budget)?;
                    } else if let Some(kind) = root_kind(node.namespace, &node.local) {
                        finish_root(node, kind, &mut output)?;
                    }
                }
                let owns_text_end = contexts.len() >= 2
                    && contexts[contexts.len() - 1] == ContextKind::OfficeText
                    && contexts[contexts.len() - 2] == ContextKind::OfficeBody;
                if contexts.pop() == Some(ContextKind::OfficeText)
                    && owns_text_end
                    && body_text_depth > 0
                {
                    body_text_depth -= 1;
                }
                namespaces.end()?;
            },
            Event::Text(text) => {
                if let Some(node) = stack.last_mut() {
                    reserve_child_slot(node, &mut budget)?;
                    let raw = text.as_ref().len();
                    budget.text(raw)?;
                    budget.storage(raw)?;
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid OTH structure text: {error}"))
                    })?;
                    budget.text_after_raw(raw, value.len())?;
                    if value.len() > raw {
                        budget.storage(value.len() - raw)?;
                    }
                    append_reserved_child(node, Child::Text(value.into_owned()));
                }
            },
            Event::CData(text) => {
                if let Some(node) = stack.last_mut() {
                    reserve_child_slot(node, &mut budget)?;
                    let raw = text.as_ref().len();
                    budget.text(raw)?;
                    budget.storage(raw)?;
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid OTH structure CDATA: {error}"))
                    })?;
                    budget.text_after_raw(raw, value.len())?;
                    if value.len() > raw {
                        budget.storage(value.len() - raw)?;
                    }
                    append_reserved_child(node, Child::Text(value.into_owned()));
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(node) = stack.last_mut() {
                    reserve_child_slot(node, &mut budget)?;
                    let raw = reference.as_ref().len();
                    budget.text(raw)?;
                    budget.storage(raw)?;
                    let value = reference_value(&reference)?;
                    budget.text_after_raw(raw, value.len())?;
                    if value.len() > raw {
                        budget.storage(value.len() - raw)?;
                    }
                    append_reserved_child(node, Child::Text(value));
                }
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                if !stack.is_empty() {
                    return invalid("OTH body structure stack is not empty");
                }
                if !namespaces.is_empty() {
                    return invalid("OTH body structure namespace stack is not empty");
                }
                return Ok(output);
            },
        }
    }
}

fn document_tracked_changes(
    xml: &str,
    budget: &mut Budget,
) -> Result<Option<crate::change::ChangeTracking>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut namespaces = SemanticNamespaceContext::new();
    let mut contexts = Vec::<ContextKind>::new();
    let mut prelude_states = Vec::<OfficeTextPreludeState>::new();
    let mut tracked_changes = None;
    loop {
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                let (namespace, context, root) = classify_element(&namespaces, start.name())?;
                let prelude_element =
                    office_text_prelude_element(namespace, start.name(), &namespaces)?;
                let admitted_prelude =
                    admit_office_text_child(&contexts, &mut prelude_states, prelude_element)?;
                if root == Some(RootKind::TrackedChanges) {
                    if tracked_changes.is_some() {
                        return invalid("duplicate OTH text:tracked-changes declaration");
                    }
                    let track_changes = optional_attr_from_start(
                        &namespaces,
                        reader.decoder(),
                        &start,
                        Ns::Text,
                        b"track-changes",
                        budget,
                    )?
                    .map(|value| parse_odf_bool(&value, "text:track-changes"))
                    .transpose()?;
                    tracked_changes = Some(crate::change::ChangeTracking::projected(track_changes));
                }
                if contexts.len() >= MAX_STRUCTURE_SITE_DEPTH {
                    return invalid("OTH tracked-change context scan exceeds the depth limit");
                }
                let context_len = contexts.len().checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("OTH tracked-change context length overflow".to_string())
                })?;
                budget.reserve_vec(
                    &mut contexts,
                    context_len,
                    "OTH tracked-change context scan",
                )?;
                let prelude_len = prelude_states.len().checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("OTH tracked-change prelude length overflow".to_string())
                })?;
                budget.reserve_vec(
                    &mut prelude_states,
                    prelude_len,
                    "OTH tracked-change prelude scan",
                )?;
                contexts.push(context);
                prelude_states.push(if context == ContextKind::OfficeText {
                    OfficeTextPreludeState::Start
                } else {
                    OfficeTextPreludeState::Body
                });
                debug_assert!(admitted_prelude || prelude_element.is_none());
            },
            Event::Empty(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                let (namespace, _context, root) = classify_element(&namespaces, start.name())?;
                let prelude_element =
                    office_text_prelude_element(namespace, start.name(), &namespaces)?;
                let admitted_prelude =
                    admit_office_text_child(&contexts, &mut prelude_states, prelude_element)?;
                if root == Some(RootKind::TrackedChanges) {
                    if tracked_changes.is_some() {
                        return invalid("duplicate OTH text:tracked-changes declaration");
                    }
                    let track_changes = optional_attr_from_start(
                        &namespaces,
                        reader.decoder(),
                        &start,
                        Ns::Text,
                        b"track-changes",
                        budget,
                    )?
                    .map(|value| parse_odf_bool(&value, "text:track-changes"))
                    .transpose()?;
                    tracked_changes = Some(crate::change::ChangeTracking::projected(track_changes));
                }
                namespaces.end()?;
                debug_assert!(admitted_prelude || prelude_element.is_none());
            },
            Event::End(end) => {
                let (_namespace, context, _root) = classify_element(&namespaces, end.name())?;
                let Some(open) = contexts.pop() else {
                    return invalid("OTH tracked-change context scan has an unmatched end tag");
                };
                prelude_states.pop();
                if open != context {
                    return invalid("OTH tracked-change context scan has mismatched tags");
                }
                namespaces.end()?;
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::Eof => {
                if !contexts.is_empty() || !namespaces.is_empty() {
                    return invalid("OTH tracked-change context scan is not closed");
                }
                return Ok(tracked_changes);
            },
            _ => {},
        }
    }
}

fn office_text_prelude_element(
    namespace: Ns,
    name: QName<'_>,
    namespaces: &SemanticNamespaceContext<'_>,
) -> Result<Option<OfficeTextPreludeElement>> {
    let local = namespaces.resolve_element(name)?.local_bytes();
    Ok(match (namespace, local) {
        (Ns::Office, b"forms") => Some(OfficeTextPreludeElement::Forms),
        (Ns::Text, b"tracked-changes") => Some(OfficeTextPreludeElement::TrackedChanges),
        // The text-decls and table-decls groups are represented by their
        // direct XML children.  They remain inert here; the structure
        // projection only needs to enforce their direct owner, uniqueness, and
        // schema order before ordinary body content starts.
        (Ns::Text, b"variable-decls") => Some(OfficeTextPreludeElement::VariableDeclarations),
        (Ns::Text, b"sequence-decls") => Some(OfficeTextPreludeElement::SequenceDeclarations),
        (Ns::Text, b"user-field-decls") => Some(OfficeTextPreludeElement::UserFieldDeclarations),
        (Ns::Text, b"dde-connection-decls") => {
            Some(OfficeTextPreludeElement::DdeConnectionDeclarations)
        },
        (Ns::Text, b"alphabetical-index-auto-mark-file") => {
            Some(OfficeTextPreludeElement::AlphabeticalIndexAutoMarkFile)
        },
        (Ns::Table, b"calculation-settings") => Some(OfficeTextPreludeElement::CalculationSettings),
        (Ns::Table, b"content-validations") => Some(OfficeTextPreludeElement::ContentValidations),
        (Ns::Table, b"label-ranges") => Some(OfficeTextPreludeElement::LabelRanges),
        _ => None,
    })
}

fn admit_office_text_child(
    contexts: &[ContextKind],
    prelude_states: &mut [OfficeTextPreludeState],
    element: Option<OfficeTextPreludeElement>,
) -> Result<bool> {
    let direct = direct_office_text_parent(contexts);
    let Some(element) = element else {
        if direct {
            let state = prelude_states.last_mut().ok_or_else(|| {
                Error::InvalidFormat("OTH office:text prelude state is missing".to_string())
            })?;
            *state = OfficeTextPreludeState::Body;
        }
        return Ok(false);
    };
    if !direct {
        return invalid("OTH office:text prelude declaration must be a direct office:text child");
    }
    let state = prelude_states.last_mut().ok_or_else(|| {
        Error::InvalidFormat("OTH office:text prelude state is missing".to_string())
    })?;
    let current_rank = state.rank();
    if *state == OfficeTextPreludeState::Body
        || current_rank.is_some_and(|rank| element.rank() <= rank)
    {
        return invalid("OTH office:text prelude declaration is duplicated or out of order");
    }
    *state = element.state();
    Ok(true)
}

fn direct_office_text_parent(contexts: &[ContextKind]) -> bool {
    contexts.len() >= 2
        && contexts[contexts.len() - 2] == ContextKind::OfficeBody
        && contexts[contexts.len() - 1] == ContextKind::OfficeText
}

fn validate_xml_id_and_change_references(
    xml: &str,
    tracking_present: bool,
    budget: &mut Budget,
) -> Result<()> {
    use std::collections::HashSet;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut namespaces = SemanticNamespaceContext::new();
    let mut contexts = Vec::<ContextKind>::new();
    let mut ids = HashSet::<String>::new();
    let mut regions = HashSet::<String>::new();
    let mut marker_references = Vec::<String>::new();

    loop {
        let event = reader.read_event().map_err(|error| xml_error(&error))?;
        match event {
            Event::Start(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                scan_identity_element(
                    &namespaces,
                    reader.decoder(),
                    &start,
                    true,
                    tracking_present,
                    &mut contexts,
                    &mut ids,
                    &mut regions,
                    &mut marker_references,
                    budget,
                )?;
            },
            Event::Empty(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                scan_identity_element(
                    &namespaces,
                    reader.decoder(),
                    &start,
                    false,
                    tracking_present,
                    &mut contexts,
                    &mut ids,
                    &mut regions,
                    &mut marker_references,
                    budget,
                )?;
                namespaces.end()?;
            },
            Event::End(end) => {
                let (_namespace, context, _root) = classify_element(&namespaces, end.name())?;
                let Some(open) = contexts.pop() else {
                    return invalid("OTH change identity scan has an unmatched end tag");
                };
                if open != context {
                    return invalid("OTH change identity scan has mismatched tags");
                }
                namespaces.end()?;
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::Eof => {
                if marker_references
                    .iter()
                    .any(|reference| !regions.contains(reference))
                {
                    return invalid("OTH change marker references an unresolved changed-region");
                }
                if !contexts.is_empty() || !namespaces.is_empty() {
                    return invalid("OTH change identity scan is not closed");
                }
                return Ok(());
            },
            _ => {},
        }
    }
}

fn scan_identity_element(
    namespaces: &SemanticNamespaceContext<'_>,
    decoder: quick_xml::encoding::Decoder,
    start: &quick_xml::events::BytesStart<'_>,
    push_context: bool,
    tracking_present: bool,
    contexts: &mut Vec<ContextKind>,
    ids: &mut std::collections::HashSet<String>,
    regions: &mut std::collections::HashSet<String>,
    marker_references: &mut Vec<String>,
    budget: &mut Budget,
) -> Result<()> {
    let (namespace, context, root) = classify_element(namespaces, start.name())?;
    let owns_body = owns_structure_root(contexts);
    // A known OTH-looking element inside a foreign wrapper remains opaque.
    // Only an element in the admitted body ancestry participates in the
    // changed-region identity grammar; this preserves inert same-named
    // descendants while still refusing direct or nested regions in body
    // structures.
    let is_region = owns_body && namespace == Ns::Text && root == Some(RootKind::ChangedRegion);
    if is_region && contexts.last() != Some(&ContextKind::TextTrackedChanges) {
        return invalid("OTH text:changed-region must be a direct text:tracked-changes child");
    }
    if is_region && !tracking_present {
        return invalid("OTH text:changed-region requires text:tracked-changes");
    }
    let is_marker = tracking_present
        && owns_body
        && namespace == Ns::Text
        && matches!(
            root,
            Some(RootKind::ChangeStart | RootKind::ChangeEnd | RootKind::Change)
        );
    let xml_id = optional_attr_from_start(namespaces, decoder, start, Ns::Xml, b"id", budget)?;
    if let Some(xml_id) = xml_id.as_deref() {
        validate_ncname(xml_id, "xml:id")?;
        budget.storage(size_of::<String>())?;
        budget.storage(size_of::<usize>())?;
        budget.storage(xml_id.len())?;
        let id_len = ids
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("OTH XML identity length overflow".to_string()))?;
        budget.reserve_set(&mut *ids, id_len, "OTH XML identity registry")?;
        if !ids.insert(xml_id.to_owned()) {
            return invalid("duplicate OTH xml:id value");
        }
        if is_region {
            budget.storage(size_of::<String>())?;
            budget.storage(size_of::<usize>())?;
            budget.storage(xml_id.len())?;
            let region_len = regions.len().checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("OTH changed-region identity length overflow".to_string())
            })?;
            budget.reserve_set(
                &mut *regions,
                region_len,
                "OTH changed-region identity registry",
            )?;
            regions.insert(xml_id.to_owned());
        }
    }
    if is_region {
        let Some(xml_id) = xml_id.as_deref() else {
            return invalid("OTH changed-region requires xml:id");
        };
        let region_id =
            optional_attr_from_start(namespaces, decoder, start, Ns::Text, b"id", budget)?;
        if let Some(region_id) = region_id.as_deref() {
            validate_ncname(region_id, "text:id")?;
            if region_id != xml_id {
                return invalid("OTH changed-region text:id must equal xml:id");
            }
        }
        if !regions.contains(xml_id) {
            budget.storage(size_of::<String>())?;
            budget.storage(size_of::<usize>())?;
            budget.storage(xml_id.len())?;
            let region_len = regions.len().checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("OTH changed-region identity length overflow".to_string())
            })?;
            budget.reserve_set(regions, region_len, "OTH changed-region identity registry")?;
            regions.insert(xml_id.to_owned());
        }
    }
    if is_marker {
        let Some(change_id) =
            optional_attr_from_start(namespaces, decoder, start, Ns::Text, b"change-id", budget)?
        else {
            return invalid("OTH change marker requires text:change-id");
        };
        validate_ncname(&change_id, "text:change-id")?;
        budget.storage(size_of::<String>())?;
        let marker_len = marker_references
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("OTH change marker length overflow".to_string()))?;
        budget.reserve_vec(
            marker_references,
            marker_len,
            "OTH change marker references",
        )?;
        marker_references.push(change_id);
    }
    if push_context {
        if contexts.len() >= MAX_STRUCTURE_SITE_DEPTH {
            return invalid("OTH change identity scan exceeds the depth limit");
        }
        let context_len = contexts.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH change identity context length overflow".to_string())
        })?;
        budget.reserve_vec(contexts, context_len, "OTH change identity context scan")?;
        contexts.push(context);
    }
    Ok(())
}

fn validate_ncname(value: &str, field: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(format!(
            "invalid OTH {field} lexical value"
        )));
    };
    if !is_ncname_start(first) || chars.any(|character| !is_ncname_character(character)) {
        return Err(Error::InvalidFormat(format!(
            "invalid OTH {field} lexical value"
        )));
    }
    Ok(())
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

fn is_ncname_character(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

fn classify_element(
    namespaces: &SemanticNamespaceContext<'_>,
    name: QName<'_>,
) -> Result<(Ns, ContextKind, Option<RootKind>)> {
    let resolved = namespaces.resolve_element(name)?;
    let namespace = local_namespace(resolved.namespace.as_ref());
    let local = resolved.local_bytes();
    Ok((
        namespace,
        context_kind(namespace, local),
        root_kind_bytes(namespace, local),
    ))
}

fn local_namespace(namespace: Option<&NamespaceId>) -> Ns {
    let Some(namespace) = namespace else {
        return Ns::None;
    };
    if namespace.unknown_uri().is_some() {
        debug_assert!(namespace.is_unknown());
        return Ns::Other;
    }
    let Some(namespace) = namespace.known() else {
        return Ns::Other;
    };
    match namespace {
        ResolvedNs::Office => Ns::Office,
        ResolvedNs::Text => Ns::Text,
        ResolvedNs::Table => Ns::Table,
        ResolvedNs::Draw => Ns::Draw,
        ResolvedNs::Dr3d => Ns::Dr3d,
        ResolvedNs::Xlink => Ns::Xlink,
        ResolvedNs::Svg => Ns::Svg,
        ResolvedNs::Xml => Ns::Xml,
        ResolvedNs::Dc => Ns::Dc,
        ResolvedNs::Meta => Ns::Meta,
        ResolvedNs::Xhtml => Ns::Xhtml,
        ResolvedNs::Fo => Ns::Fo,
        ResolvedNs::Style => Ns::Style,
    }
}

fn root_kind_bytes(namespace: Ns, local: &[u8]) -> Option<RootKind> {
    match (namespace, local) {
        (Ns::Table, b"table") => Some(RootKind::Table),
        (Ns::Text, b"section") => Some(RootKind::Section),
        (Ns::Text, b"note") => Some(RootKind::Note),
        (Ns::Office, b"annotation") => Some(RootKind::Annotation),
        (Ns::Text, b"tracked-changes") => Some(RootKind::TrackedChanges),
        (Ns::Text, b"changed-region") => Some(RootKind::ChangedRegion),
        (Ns::Text, b"change-start") => Some(RootKind::ChangeStart),
        (Ns::Text, b"change-end") => Some(RootKind::ChangeEnd),
        (Ns::Text, b"change") => Some(RootKind::Change),
        (
            Ns::Text,
            b"table-of-content"
            | b"illustration-index"
            | b"table-index"
            | b"object-index"
            | b"user-index"
            | b"alphabetical-index"
            | b"bibliography"
            | b"bibliography-index",
        ) => Some(RootKind::Index),
        (Ns::Draw, b"frame" | b"text-box") => Some(RootKind::Frame),
        (Ns::Text, b"ruby") => Some(RootKind::Ruby),
        _ => None,
    }
}

const fn root_requires_tracked_changes(kind: RootKind) -> bool {
    matches!(
        kind,
        RootKind::ChangedRegion | RootKind::ChangeStart | RootKind::ChangeEnd | RootKind::Change
    )
}

fn context_kind(namespace: Ns, local: &[u8]) -> ContextKind {
    match (namespace, local) {
        (Ns::Office, b"body") => ContextKind::OfficeBody,
        (Ns::Office, b"text") => ContextKind::OfficeText,
        (Ns::Text, b"section") => ContextKind::TextSection,
        (Ns::Text, b"p") => ContextKind::TextParagraph,
        (Ns::Text, b"h") => ContextKind::TextHeading,
        (Ns::Text, b"span") => ContextKind::TextSpan,
        (Ns::Text, b"note") => ContextKind::TextNote,
        (Ns::Text, b"note-body") => ContextKind::TextNoteBody,
        (Ns::Text, b"tracked-changes") => ContextKind::TextTrackedChanges,
        (Ns::Text, b"changed-region") => ContextKind::TextChangedRegion,
        (Ns::Text, b"insertion") => ContextKind::TextInsertion,
        (Ns::Text, b"deletion") => ContextKind::TextDeletion,
        (Ns::Text, b"format-change") => ContextKind::TextFormatChange,
        (
            Ns::Text,
            b"list"
            | b"list-header"
            | b"list-item"
            | b"numbered-paragraph"
            | b"bookmark-start"
            | b"bookmark-end"
            | b"bookmark"
            | b"reference-mark-start"
            | b"reference-mark-end"
            | b"reference-mark"
            | b"change"
            | b"note-citation"
            | b"index-title"
            | b"index-title-template"
            | b"index-entry"
            | b"table-of-content-entry-template"
            | b"illustration-index-entry-template"
            | b"table-index-entry-template"
            | b"object-index-entry-template"
            | b"user-index-entry-template"
            | b"alphabetical-index-entry-template"
            | b"bibliography-entry-template",
        ) => ContextKind::BodyContainer,
        (Ns::Text, b"index-body") => ContextKind::TextIndexBody,
        (
            Ns::Text,
            b"section-source"
            | b"table-of-content-source"
            | b"illustration-index-source"
            | b"table-index-source"
            | b"object-index-source"
            | b"user-index-source"
            | b"alphabetical-index-source"
            | b"bibliography-source",
        ) => ContextKind::TextIndexSource,
        (Ns::Text, local) if root_kind_bytes(namespace, local) == Some(RootKind::Index) => {
            ContextKind::TextIndex
        },
        (Ns::Table, b"table") => ContextKind::Table,
        (Ns::Table, b"table-columns" | b"table-column-group" | b"table-header-columns") => {
            ContextKind::TableColumns
        },
        (Ns::Table, b"table-column") => ContextKind::TableColumn,
        (Ns::Table, b"table-rows" | b"table-row-group" | b"table-header-rows") => {
            ContextKind::TableRows
        },
        (Ns::Table, b"table-row") => ContextKind::TableRow,
        (Ns::Table, b"table-cell") => ContextKind::TableCell,
        (Ns::Table, b"covered-table-cell") => ContextKind::CoveredTableCell,
        (Ns::Office, b"annotation") => ContextKind::OfficeAnnotation,
        (Ns::Draw, b"frame") => ContextKind::DrawFrame,
        (Ns::Draw, b"text-box") => ContextKind::DrawTextBox,
        (Ns::Draw, b"image") => ContextKind::DrawImage,
        (Ns::Draw, b"object") => ContextKind::DrawObject,
        (Ns::Draw, b"object-ole") => ContextKind::DrawObjectOle,
        (Ns::Draw, b"plugin") => ContextKind::DrawPlugin,
        (Ns::Draw, b"floating-frame") => ContextKind::DrawFloatingFrame,
        (Ns::Text, b"ruby") => ContextKind::TextRuby,
        (Ns::Text, b"ruby-base") => ContextKind::TextRubyBase,
        (Ns::Text, b"ruby-text") => ContextKind::TextRubyText,
        _ => ContextKind::Other,
    }
}

fn owns_structure_root(contexts: &[ContextKind]) -> bool {
    let Some(text_index) = contexts.iter().enumerate().rposition(|(index, context)| {
        *context == ContextKind::OfficeText
            && index > 0
            && contexts[index - 1] == ContextKind::OfficeBody
    }) else {
        return false;
    };
    contexts[text_index + 1..]
        .iter()
        .copied()
        .all(ContextKind::permits_body_structure)
}

fn parse_node(
    namespaces: &SemanticNamespaceContext<'_>,
    decoder: quick_xml::encoding::Decoder,
    start: &quick_xml::events::BytesStart<'_>,
    budget: &mut Budget,
) -> Result<Node> {
    budget.node()?;
    let resolved = namespaces.resolve_element(start.name())?;
    let local = std::str::from_utf8(resolved.local_bytes()).map_err(|error| {
        Error::InvalidFormat(format!("invalid OTH structure element name: {error}"))
    })?;
    let namespace = local_namespace(resolved.namespace.as_ref());
    budget.text(start.name().as_ref().len())?;
    budget.text(local.len())?;
    let mut attribute_count = 0usize;
    for raw in start.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute: {error}"))
        })?;
        let raw_attribute_bytes = attribute
            .key
            .as_ref()
            .len()
            .checked_add(attribute.value.len())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH structure attribute size overflow".to_string())
            })?;
        budget.text(raw_attribute_bytes)?;
        if namespace_declaration(attribute.key).is_some() {
            budget.storage(size_of::<usize>() * 2)?;
            continue;
        }
        let resolved = namespaces.resolve_attribute(attribute.key)?;
        let _local = std::str::from_utf8(resolved.local_bytes()).map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute name: {error}"))
        })?;
        attribute_count = attribute_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure attribute count overflow".to_string())
        })?;
    }
    let mut attributes = Vec::new();
    budget.reserve_vec_exact(
        &mut attributes,
        attribute_count,
        "OTH body structure attributes",
    )?;

    // The attribute vector is admitted before any element name or namespace
    // URI is cloned.  Unknown namespace URIs stay interned Arc identities
    // supplied by the resolver; aliases never allocate one URI per node.
    budget.storage(size_of::<String>())?;
    budget.storage(local.len())?;
    let mut node = Node {
        children: Vec::new(),
        local: local.to_owned(),
        namespace,
        namespace_uri: unknown_namespace_arc(resolved.namespace.as_ref()),
        attributes,
    };
    for raw in start.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute: {error}"))
        })?;
        if namespace_declaration(attribute.key).is_some() {
            continue;
        }
        let resolved = namespaces.resolve_attribute(attribute.key)?;
        let local = std::str::from_utf8(resolved.local_bytes()).map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute name: {error}"))
        })?;
        let namespace = local_namespace(resolved.namespace.as_ref());
        let namespace_uri = unknown_namespace_arc(resolved.namespace.as_ref());
        let raw_value_len = attribute.value.len();
        // The raw value bounds quick_xml's temporary Cow allocation.  Charge
        // it before invoking the decoder, then account any decoded expansion
        // before converting a borrowed value into an owned String.
        budget.storage(raw_value_len)?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid OTH structure attribute value: {error}"))
            })?;
        budget.text_after_raw(
            attribute
                .key
                .as_ref()
                .len()
                .checked_add(raw_value_len)
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH structure attribute size overflow".to_string())
                })?,
            value.len(),
        )?;
        if value.len() > raw_value_len {
            budget.storage(value.len() - raw_value_len)?;
        }
        budget.storage(local.len())?;
        node.attributes.push(Attribute {
            local: local.to_owned(),
            namespace,
            namespace_uri,
            value: value.into_owned(),
        });
    }
    Ok(node)
}

fn unknown_namespace_arc(namespace: Option<&NamespaceId>) -> Option<Arc<str>> {
    match namespace {
        Some(NamespaceId::Unknown(uri)) => Some(Arc::clone(uri)),
        _ => None,
    }
}

fn reserve_stack_node(stack: &mut Vec<Node>, budget: &mut Budget) -> Result<()> {
    let required_len = stack.len().checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("OTH body structure stack length overflow".to_string())
    })?;
    budget.reserve_vec(stack, required_len, "OTH body structure stack")
}

fn append_node(stack: &mut [Node], node: Node, budget: &mut Budget) -> Result<()> {
    let parent = stack
        .last_mut()
        .ok_or_else(|| Error::InvalidFormat("OTH body structure parent is missing".to_string()))?;
    reserve_child_slot(parent, budget)?;
    budget.storage(size_of::<Node>())?;
    append_reserved_child(parent, Child::Element(Box::new(node)));
    Ok(())
}

fn reserve_child_slot(parent: &mut Node, budget: &mut Budget) -> Result<()> {
    let required_len = parent.children.len().checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("OTH body structure child length overflow".to_string())
    })?;
    budget.reserve_vec(
        &mut parent.children,
        required_len,
        "OTH body structure children",
    )
}

fn append_reserved_child(parent: &mut Node, child: Child) {
    parent.children.push(child);
}

fn root_kind(namespace: Ns, local: &str) -> Option<RootKind> {
    match (namespace, local) {
        (Ns::Table, "table") => Some(RootKind::Table),
        (Ns::Text, "section") => Some(RootKind::Section),
        (Ns::Text, "note") => Some(RootKind::Note),
        (Ns::Office, "annotation") => Some(RootKind::Annotation),
        (Ns::Text, "tracked-changes") => Some(RootKind::TrackedChanges),
        (Ns::Text, "changed-region") => Some(RootKind::ChangedRegion),
        (Ns::Text, "change-start") => Some(RootKind::ChangeStart),
        (Ns::Text, "change-end") => Some(RootKind::ChangeEnd),
        (Ns::Text, "change") => Some(RootKind::Change),
        (
            Ns::Text,
            "table-of-content" | "illustration-index" | "table-index" | "object-index"
            | "user-index" | "alphabetical-index" | "bibliography" | "bibliography-index",
        ) => Some(RootKind::Index),
        (Ns::Draw, "frame" | "text-box") => Some(RootKind::Frame),
        (Ns::Text, "ruby") => Some(RootKind::Ruby),
        _ => None,
    }
}

fn finish_root(node: Node, kind: RootKind, output: &mut BodyStructures) -> Result<()> {
    match kind {
        RootKind::Table => collect_tables(&node, output, false),
        RootKind::Section => collect_sections(&node, 1, output, false),
        RootKind::Note => collect_notes(&node, output, false),
        RootKind::Annotation => collect_annotations(&node, output, false),
        RootKind::TrackedChanges
        | RootKind::ChangedRegion
        | RootKind::ChangeStart
        | RootKind::ChangeEnd
        | RootKind::Change => collect_changes(&node, output, true),
        RootKind::Index => collect_indexes(&node, output, false),
        RootKind::Frame => collect_frames(&node, output, false),
        RootKind::Ruby => collect_rubies(&node, output, false),
    }
}

fn collect_tables(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table" {
        let planned = measure_table_node(node)?;
        output.ensure(planned)?;
        let table_len = output.tables.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH table projection length overflow".to_string())
        })?;
        reserve_retained_vec(
            &mut output.tables,
            table_len,
            &mut output.retained_bytes,
            "OTH table projection",
        )?;
        output.account(planned)?;
        let table = project_table(node)?;
        output.tables.push(table);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn collect_descendant_roots(
    node: &Node,
    section_depth: usize,
    output: &mut BodyStructures,
    tracked_scope: bool,
) -> Result<()> {
    let tracked_scope =
        tracked_scope || (node.namespace == Ns::Text && node.local == "tracked-changes");
    for child in elements(node) {
        if !valid_structure_child(child) {
            continue;
        }
        match root_kind(child.namespace, &child.local) {
            Some(RootKind::Table) => collect_tables(child, output, tracked_scope)?,
            Some(RootKind::Section) => {
                collect_sections(child, section_depth, output, tracked_scope)?
            },
            Some(RootKind::Note) => collect_notes(child, output, tracked_scope)?,
            Some(RootKind::Annotation) => collect_annotations(child, output, tracked_scope)?,
            Some(
                RootKind::TrackedChanges
                | RootKind::ChangedRegion
                | RootKind::ChangeStart
                | RootKind::ChangeEnd
                | RootKind::Change,
            ) => collect_changes(child, output, tracked_scope)?,
            Some(RootKind::Index) => collect_indexes(child, output, tracked_scope)?,
            Some(RootKind::Frame) => collect_frames(child, output, tracked_scope)?,
            Some(RootKind::Ruby) => collect_rubies(child, output, tracked_scope)?,
            None => collect_descendant_roots(child, section_depth, output, tracked_scope)?,
        }
    }
    Ok(())
}

fn collect_table_columns(
    node: &Node,
    columns: &mut Vec<crate::table::Column>,
    declared_columns: &mut usize,
) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table-column" {
        let repeated = count_attr(node, b"number-columns-repeated")?;
        *declared_columns = checked_dimension_add(*declared_columns, repeated)?;
        columns.push(crate::table::Column::projected(
            attr(node, Ns::Table, b"style-name").map(str::to_owned),
            attr(node, Ns::Table, b"default-cell-style-name").map(str::to_owned),
            optional_count_attr(node, b"number-columns-repeated")?,
            repeated,
            optional_visibility_attr(node)?,
            attr(node, Ns::Xml, b"id").map(str::to_owned),
        ));
        return Ok(());
    }
    for child in elements(node) {
        match (child.namespace, child.local.as_str()) {
            (Ns::Table, "table-column") => {
                collect_table_columns(child, columns, declared_columns)?;
            },
            (Ns::Table, "table-columns" | "table-column-group" | "table-header-columns") => {
                collect_table_columns(child, columns, declared_columns)?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn collect_table_rows(
    node: &Node,
    rows: &mut Vec<crate::table::Row>,
    logical_columns: &mut usize,
) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table-row" {
        let row = project_row(node)?;
        *logical_columns = (*logical_columns).max(row_width(&row)?);
        rows.push(row);
        return Ok(());
    }
    for child in elements(node) {
        match (child.namespace, child.local.as_str()) {
            (Ns::Table, "table-row") => {
                collect_table_rows(child, rows, logical_columns)?;
            },
            (Ns::Table, "table-rows" | "table-row-group" | "table-header-rows") => {
                collect_table_rows(child, rows, logical_columns)?
            },
            _ => {},
        }
    }
    Ok(())
}

fn count_table_columns(node: &Node) -> Result<usize> {
    if node.namespace == Ns::Table && node.local == "table-column" {
        return Ok(1);
    }
    let mut count = 0usize;
    for child in elements(node) {
        if child.namespace == Ns::Table
            && matches!(
                child.local.as_str(),
                "table-columns" | "table-column-group" | "table-header-columns"
            )
        {
            count = count
                .checked_add(count_table_columns(child)?)
                .ok_or_else(|| {
                    Error::InvalidFormat("OTH table column count overflow".to_string())
                })?;
        } else if child.namespace == Ns::Table && child.local == "table-column" {
            count = count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("OTH table column count overflow".to_string())
            })?;
        }
    }
    Ok(count)
}

fn count_table_rows(node: &Node) -> Result<usize> {
    if node.namespace == Ns::Table && node.local == "table-row" {
        return Ok(1);
    }
    let mut count = 0usize;
    for child in elements(node) {
        if child.namespace == Ns::Table
            && matches!(
                child.local.as_str(),
                "table-rows" | "table-row-group" | "table-header-rows"
            )
        {
            count = count
                .checked_add(count_table_rows(child)?)
                .ok_or_else(|| Error::InvalidFormat("OTH table row count overflow".to_string()))?;
        } else if child.namespace == Ns::Table && child.local == "table-row" {
            count = count
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("OTH table row count overflow".to_string()))?;
        }
    }
    Ok(count)
}

fn project_table(node: &Node) -> Result<crate::table::Table> {
    let mut columns = Vec::new();
    let mut rows = Vec::new();
    let column_count = count_table_columns(node)?;
    reserve_exact(&mut columns, column_count, "OTH table columns")?;
    let row_count = count_table_rows(node)?;
    reserve_exact(&mut rows, row_count, "OTH table rows")?;
    let mut declared_columns = 0usize;
    let mut logical_columns = 0usize;
    for child in elements(node) {
        if child.namespace != Ns::Table {
            continue;
        }
        match child.local.as_str() {
            "table-column" => {
                collect_table_columns(child, &mut columns, &mut declared_columns)?;
            },
            "table-columns" | "table-column-group" | "table-header-columns" => {
                collect_table_columns(child, &mut columns, &mut declared_columns)?;
            },
            "table-row" => collect_table_rows(child, &mut rows, &mut logical_columns)?,
            "table-rows" | "table-row-group" | "table-header-rows" => {
                collect_table_rows(child, &mut rows, &mut logical_columns)?;
            },
            _ => {},
        }
    }
    Ok(crate::table::Table::projected(
        attr(node, Ns::Table, b"name").map(str::to_owned),
        attr(node, Ns::Table, b"style-name").map(str::to_owned),
        project_table_properties(node)?,
        project_table_source(node)?,
        columns,
        rows,
        declared_columns,
        checked_dimension_max(declared_columns, logical_columns)?,
    ))
}

fn project_table_properties(node: &Node) -> Result<crate::table::TableProperties> {
    Ok(crate::table::TableProperties::projected(
        attr(node, Ns::Table, b"template-name").map(str::to_owned),
        optional_schema_bool_attr(node, Ns::Table, b"use-first-row-styles")?,
        optional_schema_bool_attr(node, Ns::Table, b"use-last-row-styles")?,
        optional_schema_bool_attr(node, Ns::Table, b"use-first-column-styles")?,
        optional_schema_bool_attr(node, Ns::Table, b"use-last-column-styles")?,
        optional_schema_bool_attr(node, Ns::Table, b"use-banding-rows-styles")?,
        optional_schema_bool_attr(node, Ns::Table, b"use-banding-columns-styles")?,
        optional_schema_bool_attr(node, Ns::Table, b"protected")?,
        attr(node, Ns::Table, b"protection-key").map(str::to_owned),
        optional_any_iri_attr(node, Ns::Table, b"protection-key-digest-algorithm")?,
        optional_schema_bool_attr(node, Ns::Table, b"print")?,
        attr(node, Ns::Table, b"print-ranges").map(str::to_owned),
        attr(node, Ns::Xml, b"id").map(str::to_owned),
        optional_schema_bool_attr(node, Ns::Table, b"is-sub-table")?,
    ))
}

fn project_table_source(node: &Node) -> Result<Option<crate::table::TableSource>> {
    let Some(source) = unique_direct_child(node, Ns::Table, "table-source", "table-source")? else {
        return Ok(None);
    };
    if elements(source).next().is_some() {
        return invalid("OTH table-source must be empty");
    }
    let source_type = attr(source, Ns::Xlink, b"type")
        .ok_or_else(|| Error::InvalidFormat("OTH table-source requires xlink:type".to_string()))?;
    if source_type != "simple" {
        return invalid("OTH table-source xlink:type must be 'simple'");
    }
    let href = attr(source, Ns::Xlink, b"href")
        .ok_or_else(|| Error::InvalidFormat("OTH table-source requires xlink:href".to_string()))?;
    validate_any_iri(href, "xlink:href")?;
    let mode = attr(source, Ns::Table, b"mode")
        .map(|value| match value {
            "copy-all" => Ok(crate::table::TableSourceMode::CopyAll),
            "copy-results-only" => Ok(crate::table::TableSourceMode::CopyResultsOnly),
            _ => invalid("OTH table-source mode is invalid"),
        })
        .transpose()?;
    let actuate = attr(source, Ns::Xlink, b"actuate")
        .map(|value| {
            if value == "onRequest" {
                Ok(crate::table::TableSourceActuate::OnRequest)
            } else {
                invalid("OTH table-source xlink:actuate is invalid")
            }
        })
        .transpose()?;
    let refresh_delay = attr(source, Ns::Table, b"refresh-delay")
        .map(|value| {
            Duration::decode_exact(value).map_err(|_| {
                Error::InvalidFormat("OTH table-source refresh-delay is invalid".to_string())
            })
        })
        .transpose()?;
    Ok(Some(crate::table::TableSource::projected(
        mode,
        attr(source, Ns::Table, b"table-name").map(str::to_owned),
        href.to_owned(),
        actuate,
        attr(source, Ns::Table, b"filter-name").map(str::to_owned),
        attr(source, Ns::Table, b"filter-options").map(str::to_owned),
        refresh_delay,
    )))
}

fn project_in_content_meta(node: &Node) -> Result<Option<crate::table::InContentMeta>> {
    let about = attr(node, Ns::Xhtml, b"about");
    let property = attr(node, Ns::Xhtml, b"property");
    if about.is_none() && property.is_none() {
        if attr(node, Ns::Xhtml, b"datatype").is_some()
            || attr(node, Ns::Xhtml, b"content").is_some()
        {
            return invalid("OTH RDFa metadata requires xhtml:about and xhtml:property");
        }
        return Ok(None);
    }
    let about = about.ok_or_else(|| {
        Error::InvalidFormat("OTH RDFa metadata requires xhtml:about".to_string())
    })?;
    let property = property.ok_or_else(|| {
        Error::InvalidFormat("OTH RDFa metadata requires xhtml:property".to_string())
    })?;
    validate_uri_or_safe_curie(about, "xhtml:about")?;
    validate_curies(property, "xhtml:property")?;
    if let Some(datatype) = attr(node, Ns::Xhtml, b"datatype") {
        validate_curie(datatype, "xhtml:datatype")?;
    }
    Ok(Some(crate::table::InContentMeta::projected(
        about.to_owned(),
        property.to_owned(),
        attr(node, Ns::Xhtml, b"datatype").map(str::to_owned),
        attr(node, Ns::Xhtml, b"content").map(str::to_owned),
    )))
}

fn project_cell_value(node: &Node) -> Result<Option<crate::table::CellValue>> {
    let Some(value_type) = attr(node, Ns::Office, b"value-type") else {
        reject_cell_value_companions(node, None)?;
        return Ok(None);
    };
    reject_cell_value_companions(node, Some(value_type))?;
    match value_type {
        "float" => {
            let lexical = required_attr(node, Ns::Office, b"value", "office:value")?;
            validate_double(lexical, "office:value")?;
            Ok(Some(crate::table::CellValue::Float {
                lexical: lexical.to_owned(),
            }))
        },
        "percentage" => {
            let lexical = required_attr(node, Ns::Office, b"value", "office:value")?;
            validate_double(lexical, "office:value")?;
            Ok(Some(crate::table::CellValue::Percentage {
                lexical: lexical.to_owned(),
            }))
        },
        "currency" => {
            let lexical = required_attr(node, Ns::Office, b"value", "office:value")?;
            validate_double(lexical, "office:value")?;
            Ok(Some(crate::table::CellValue::Currency {
                lexical: lexical.to_owned(),
                currency: attr(node, Ns::Office, b"currency").map(str::to_owned),
            }))
        },
        "date" => {
            let lexical = required_attr(node, Ns::Office, b"date-value", "office:date-value")?;
            validate_date_or_datetime(lexical)?;
            Ok(Some(crate::table::CellValue::Date {
                lexical: lexical.to_owned(),
            }))
        },
        "time" => {
            let lexical = required_attr(node, Ns::Office, b"time-value", "office:time-value")?;
            let value = Duration::decode_exact(lexical).map_err(|_| {
                Error::InvalidFormat("OTH office:time-value is invalid".to_string())
            })?;
            Ok(Some(crate::table::CellValue::Time { value }))
        },
        "boolean" => {
            let lexical =
                required_attr(node, Ns::Office, b"boolean-value", "office:boolean-value")?;
            let value = parse_odf_bool(lexical, "office:boolean-value")?;
            Ok(Some(crate::table::CellValue::Boolean {
                lexical: lexical.to_owned(),
                value,
            }))
        },
        "string" => Ok(Some(crate::table::CellValue::String {
            value: attr(node, Ns::Office, b"string-value").map(str::to_owned),
        })),
        "error" => Ok(Some(crate::table::CellValue::Error {
            value: attr(node, Ns::Office, b"string-value").map(str::to_owned),
        })),
        _ => invalid("OTH office:value-type value is invalid"),
    }
}

fn reject_cell_value_companions(node: &Node, value_type: Option<&str>) -> Result<()> {
    let companion = |local: &[u8]| attr(node, Ns::Office, local).is_some();
    let allowed = match value_type {
        Some("float") | Some("percentage") | Some("currency") => {
            Some((true, value_type == Some("currency"), false, false, false))
        },
        Some("date") => Some((false, false, true, false, false)),
        Some("time") => Some((false, false, false, true, false)),
        Some("boolean") => Some((false, false, false, false, true)),
        Some("string") | Some("error") => Some((false, false, false, false, false)),
        Some(_) => None,
        None => Some((false, false, false, false, false)),
    };
    let Some((value, currency, date, time, boolean)) = allowed else {
        return Ok(());
    };
    let present = [
        (b"value".as_slice(), value),
        (b"currency".as_slice(), currency),
        (b"date-value".as_slice(), date),
        (b"time-value".as_slice(), time),
        (b"boolean-value".as_slice(), boolean),
        (
            b"string-value".as_slice(),
            matches!(value_type, Some("string" | "error")),
        ),
    ];
    for (local, is_allowed) in present {
        if companion(local) && !is_allowed {
            return invalid(
                "OTH cell value companion attribute is not valid for office:value-type",
            );
        }
    }
    Ok(())
}

fn row_width(row: &crate::table::Row) -> Result<usize> {
    let mut width = 0usize;
    for cell in row.cells() {
        let span = cell
            .columns_spanned()
            .checked_mul(cell.repeat_count())
            .ok_or_else(|| Error::InvalidFormat("OTH table logical width overflow".to_string()))?;
        width = checked_dimension_add(width, span)?;
    }
    Ok(width)
}

fn checked_dimension_add(left: usize, right: usize) -> Result<usize> {
    let value = left
        .checked_add(right)
        .ok_or_else(|| Error::InvalidFormat("OTH table logical width overflow".to_string()))?;
    if value > MAX_REPEAT {
        return invalid("OTH table logical width exceeds the supported limit");
    }
    Ok(value)
}

fn checked_dimension_max(left: usize, right: usize) -> Result<usize> {
    let value = left.max(right);
    if value > MAX_REPEAT {
        return invalid("OTH table logical width exceeds the supported limit");
    }
    Ok(value)
}

fn project_row(node: &Node) -> Result<crate::table::Row> {
    let mut cells = Vec::new();
    let cell_count = node
        .children
        .iter()
        .filter(|child| {
            matches!(
                child,
                Child::Element(element)
                    if element.namespace == Ns::Table
                        && matches!(element.local.as_str(), "table-cell" | "covered-table-cell")
            )
        })
        .count();
    reserve_exact(&mut cells, cell_count, "OTH table cells")?;
    for child in elements(node) {
        if child.namespace == Ns::Table
            && matches!(child.local.as_str(), "table-cell" | "covered-table-cell")
        {
            cells.push(project_cell(child)?);
        }
    }
    Ok(crate::table::Row::projected(
        attr(node, Ns::Table, b"style-name").map(str::to_owned),
        attr(node, Ns::Table, b"default-cell-style-name").map(str::to_owned),
        optional_count_attr(node, b"number-rows-repeated")?,
        count_attr(node, b"number-rows-repeated")?,
        optional_visibility_attr(node)?,
        attr(node, Ns::Xml, b"id").map(str::to_owned),
        cells,
    ))
}

fn project_cell(node: &Node) -> Result<crate::table::Cell> {
    let mut paragraphs = Vec::new();
    let paragraph_count = node
        .children
        .iter()
        .filter(|child| {
            matches!(
                child,
                Child::Element(element)
                    if element.namespace == Ns::Text
                        && matches!(element.local.as_str(), "p" | "h")
            )
        })
        .count();
    reserve_exact(
        &mut paragraphs,
        paragraph_count,
        "OTH table cell paragraphs",
    )?;
    for child in elements(node) {
        if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
            let text = plain_text(child)?;
            paragraphs.push(crate::paragraph::Paragraph::new(text));
        }
    }
    Ok(crate::table::Cell::projected(
        if node.local == "covered-table-cell" {
            crate::table::CellKind::Covered
        } else {
            crate::table::CellKind::Cell
        },
        attr(node, Ns::Table, b"style-name").map(str::to_owned),
        optional_count_attr(node, b"number-columns-repeated")?,
        count_attr(node, b"number-columns-repeated")?,
        optional_count_attr(node, b"number-columns-spanned")?,
        count_attr_default(node, b"number-columns-spanned", 1)?,
        optional_count_attr(node, b"number-rows-spanned")?,
        count_attr_default(node, b"number-rows-spanned", 1)?,
        optional_count_attr(node, b"number-matrix-columns-spanned")?,
        optional_count_attr(node, b"number-matrix-rows-spanned")?,
        attr(node, Ns::Table, b"formula").map(str::to_owned),
        attr(node, Ns::Table, b"content-validation-name").map(str::to_owned),
        optional_schema_bool_attr(node, Ns::Table, b"protect")?,
        optional_schema_bool_attr(node, Ns::Table, b"protected")?,
        attr(node, Ns::Xml, b"id").map(str::to_owned),
        project_in_content_meta(node)?,
        project_cell_value(node)?,
        attr(node, Ns::Office, b"value-type").map(str::to_owned),
        attr(node, Ns::Office, b"value").map(str::to_owned),
        plain_text(node)?,
        paragraphs,
    ))
}

fn collect_sections(
    node: &Node,
    depth: usize,
    output: &mut BodyStructures,
    tracked_scope: bool,
) -> Result<()> {
    if node.namespace == Ns::Text && node.local == "section" {
        let planned = measure_section_node(node)?;
        output.ensure(planned)?;
        let section_len = output.sections.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH section projection length overflow".to_string())
        })?;
        reserve_retained_vec(
            &mut output.sections,
            section_len,
            &mut output.retained_bytes,
            "OTH section projection",
        )?;
        output.account(planned)?;
        let mut paragraphs = Vec::new();
        let paragraph_count = elements(node)
            .filter(|child| {
                child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h")
            })
            .count();
        reserve_exact(&mut paragraphs, paragraph_count, "OTH section paragraphs")?;
        for child in elements(node) {
            if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
                paragraphs.push(crate::paragraph::Paragraph::new(plain_text(child)?));
            }
        }
        let section = crate::section::Section::projected(
            attr(node, Ns::Text, b"name").map(str::to_owned),
            attr(node, Ns::Text, b"style-name").map(str::to_owned),
            bool_attr(node, Ns::Text, b"protected")?,
            optional_bool_attr(node, Ns::Text, b"display")?,
            attr(node, Ns::Text, b"condition").map(str::to_owned),
            depth,
            plain_text(node)?,
            paragraphs,
        );
        output.sections.push(section);
    }
    collect_descendant_roots(
        node,
        depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("OTH section depth overflow".to_string()))?,
        output,
        tracked_scope,
    )?;
    Ok(())
}

fn collect_notes(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Text && node.local == "note" {
        let planned = measure_note_node(node)?;
        output.ensure(planned)?;
        let note_len = output.notes.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH note projection length overflow".to_string())
        })?;
        reserve_retained_vec(
            &mut output.notes,
            note_len,
            &mut output.retained_bytes,
            "OTH note projection",
        )?;
        output.account(planned)?;
        let class = match attr(node, Ns::Text, b"note-class") {
            Some("footnote") => crate::note::NoteClass::Footnote,
            Some("endnote") => crate::note::NoteClass::Endnote,
            Some(value) => crate::note::NoteClass::Other(value.to_owned()),
            None => crate::note::NoteClass::Other(String::new()),
        };
        let citation_node = direct_child(node, Ns::Text, "note-citation");
        let body_node = direct_child(node, Ns::Text, "note-body");
        let mut paragraphs = Vec::new();
        if let Some(body) = body_node {
            let paragraph_count = elements(body)
                .filter(|child| {
                    child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h")
                })
                .count();
            reserve_exact(&mut paragraphs, paragraph_count, "OTH note paragraphs")?;
            for child in elements(body) {
                if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
                    paragraphs.push(crate::paragraph::Paragraph::new(plain_text(child)?));
                }
            }
        }
        let note = crate::note::Note::projected(
            class,
            attr(node, Ns::Text, b"id").map(str::to_owned),
            citation_node.and_then(|value| attr(value, Ns::Text, b"label").map(str::to_owned)),
            citation_node
                .map(plain_text)
                .transpose()?
                .unwrap_or_default(),
            body_node.map(plain_text).transpose()?.unwrap_or_default(),
            paragraphs,
        );
        output.notes.push(note);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn collect_annotations(
    node: &Node,
    output: &mut BodyStructures,
    tracked_scope: bool,
) -> Result<()> {
    if node.namespace == Ns::Office && node.local == "annotation" {
        let planned = measure_annotation_node(node)?;
        output.ensure(planned)?;
        let annotation_len = output.annotations.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH annotation projection length overflow".to_string())
        })?;
        reserve_retained_vec(
            &mut output.annotations,
            annotation_len,
            &mut output.retained_bytes,
            "OTH annotation projection",
        )?;
        output.account(planned)?;
        let creator = direct_child(node, Ns::Dc, "creator")
            .map(plain_text)
            .transpose()?;
        let date = direct_child(node, Ns::Dc, "date")
            .map(plain_text)
            .transpose()?;
        let date_string = direct_child(node, Ns::Meta, "date-string")
            .map(plain_text)
            .transpose()?;
        let initials = direct_child(node, Ns::Meta, "creator-initials")
            .or_else(|| direct_child(node, Ns::Text, "sender-initials"))
            .map(plain_text)
            .transpose()?;
        let annotation = crate::annotation::Annotation::projected(
            attr(node, Ns::Office, b"name").map(str::to_owned),
            creator,
            date,
            date_string,
            initials,
            optional_bool_attr(node, Ns::Office, b"display")?,
            annotation_text(node)?,
        );
        output.annotations.push(annotation);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn annotation_text(node: &Node) -> Result<String> {
    let mut text = String::new();
    let mut text_len = 0usize;
    for child in elements(node) {
        if child.namespace == Ns::Dc
            || child.namespace == Ns::Meta
            || (child.namespace == Ns::Text && child.local == "sender-initials")
        {
            continue;
        }
        text_len = text_len
            .checked_add(plain_text_content_len(child)?)
            .ok_or_else(|| Error::InvalidFormat("OTH annotation text size overflow".to_string()))?;
    }
    reserve_string_exact(&mut text, text_len, "OTH annotation text")?;
    for child in elements(node) {
        if child.namespace == Ns::Dc
            || child.namespace == Ns::Meta
            || (child.namespace == Ns::Text && child.local == "sender-initials")
        {
            continue;
        }
        append_plain_text(&mut text, child)?;
    }
    Ok(text)
}

fn collect_changes(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    let tracked_scope =
        tracked_scope || (node.namespace == Ns::Text && node.local == "tracked-changes");
    let kind_name = match (node.namespace, node.local.as_str()) {
        (Ns::Text, "changed-region") => Some("changed-region"),
        (Ns::Text, "change-start") => Some("change-start"),
        (Ns::Text, "change-end") => Some("change-end"),
        (Ns::Text, "change") => Some("change"),
        _ => None,
    };
    if tracked_scope {
        if let Some(kind_name) = kind_name {
            let planned = measure_change_node(node, kind_name)?;
            output.ensure(planned)?;
            let change_len = output.changes.len().checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("OTH tracked change projection length overflow".to_string())
            })?;
            reserve_retained_vec(
                &mut output.changes,
                change_len,
                &mut output.retained_bytes,
                "OTH tracked change projection",
            )?;
            output.account(planned)?;
            let kind = match node.local.as_str() {
                "change-start" => crate::change::Kind::Start,
                "change-end" => crate::change::Kind::End,
                _ => change_kind(node),
            };
            let (xml_id, region_id, marker_change_id, info) = project_change_metadata(node, &kind)?;
            let author = info.as_ref().map(|value| value.creator().to_owned());
            let date = info.as_ref().map(|value| value.date().to_owned());
            let id = xml_id
                .as_deref()
                .or(region_id.as_deref())
                .or(marker_change_id.as_deref())
                .map(str::to_owned);
            let change = crate::change::Change::projected(
                id,
                kind,
                author,
                date,
                xml_id,
                region_id,
                marker_change_id,
                info,
                change_text(node)?,
            );
            output.changes.push(change);
        }
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn project_change_metadata(
    node: &Node,
    kind: &crate::change::Kind,
) -> Result<(
    Option<String>,
    Option<String>,
    Option<String>,
    Option<crate::change::ChangeInfo>,
)> {
    let is_region = matches!(
        kind,
        crate::change::Kind::Insertion
            | crate::change::Kind::Deletion
            | crate::change::Kind::Format
            | crate::change::Kind::Move
            | crate::change::Kind::Other(_)
    ) && node.namespace == Ns::Text
        && node.local == "changed-region";
    if is_region {
        let xml_id = attr(node, Ns::Xml, b"id").map(str::to_owned);
        if let Some(xml_id) = xml_id.as_deref() {
            validate_ncname(xml_id, "xml:id")?;
        }
        let region_id = attr(node, Ns::Text, b"id").map(str::to_owned);
        if let Some(region_id) = region_id.as_deref() {
            validate_ncname(region_id, "text:id")?;
            if let Some(xml_id) = xml_id.as_deref() {
                if region_id != xml_id {
                    return invalid("OTH changed-region text:id must equal xml:id");
                }
            }
        } else if xml_id.is_none() {
            return invalid("OTH changed-region requires xml:id");
        }
        let content = direct_change_child(node)?;
        let info = validate_change_content(content)?;
        let info = project_change_info(info)?;
        return Ok((xml_id, region_id, None, Some(info)));
    }

    let marker_change_id = attr(node, Ns::Text, b"change-id")
        .ok_or_else(|| {
            Error::InvalidFormat("OTH change marker requires text:change-id".to_string())
        })?
        .to_owned();
    validate_ncname(&marker_change_id, "text:change-id")?;
    Ok((None, None, Some(marker_change_id), None))
}

fn direct_change_child(node: &Node) -> Result<&Node> {
    if node
        .children
        .iter()
        .any(|child| matches!(child, Child::Text(text) if !text.trim().is_empty()))
    {
        return invalid("OTH changed-region has unexpected direct text");
    }
    let mut children = elements(node);
    let Some(child) = children.next() else {
        return invalid("OTH changed-region requires exactly one change kind");
    };
    if children.next().is_some()
        || child.namespace != Ns::Text
        || !matches!(
            child.local.as_str(),
            "insertion" | "deletion" | "format-change"
        )
    {
        return invalid("OTH changed-region requires exactly one change kind");
    }
    Ok(child)
}

fn validate_change_content(node: &Node) -> Result<&Node> {
    let mut children = elements(node);
    let Some(info) = children.next() else {
        return invalid("OTH changed-region requires exactly one office:change-info");
    };
    if info.namespace != Ns::Office || info.local != "change-info" {
        return invalid("OTH changed-region requires office:change-info first");
    }
    match node.local.as_str() {
        "insertion" | "format-change" => {
            if children.next().is_some() {
                return invalid("OTH insertion or format-change cannot contain payload");
            }
        },
        "deletion" => {
            for child in children {
                if !is_text_content_element(child) {
                    return invalid("OTH deletion contains invalid text-content");
                }
            }
        },
        _ => return invalid("OTH changed-region child has invalid kind"),
    }
    if node
        .children
        .iter()
        .any(|child| matches!(child, Child::Text(text) if !text.trim().is_empty()))
    {
        return invalid("OTH change content has unexpected direct text");
    }
    Ok(info)
}

fn is_text_content_element(node: &Node) -> bool {
    match node.namespace {
        Ns::Text => matches!(
            node.local.as_str(),
            "h" | "p"
                | "list"
                | "numbered-paragraph"
                | "soft-page-break"
                | "table-of-content"
                | "illustration-index"
                | "table-index"
                | "object-index"
                | "user-index"
                | "alphabetical-index"
                | "bibliography"
                | "section"
                | "change"
                | "change-start"
                | "change-end"
        ),
        Ns::Table => node.local == "table",
        Ns::Draw => matches!(
            node.local.as_str(),
            "a" | "rect"
                | "line"
                | "polyline"
                | "polygon"
                | "regular-polygon"
                | "path"
                | "circle"
                | "ellipse"
                | "g"
                | "page-thumbnail"
                | "frame"
                | "measure"
                | "caption"
                | "connector"
                | "control"
                | "custom-shape"
        ),
        Ns::Dr3d => node.local == "scene",
        _ => false,
    }
}

fn project_change_info(node: &Node) -> Result<crate::change::ChangeInfo> {
    if node
        .children
        .iter()
        .any(|child| matches!(child, Child::Text(text) if !text.trim().is_empty()))
    {
        return invalid("OTH office:change-info has unexpected direct text");
    }
    let mut children = elements(node);
    let Some(creator_node) = children.next() else {
        return invalid("OTH office:change-info children have invalid order");
    };
    let Some(date_node) = children.next() else {
        return invalid("OTH office:change-info children have invalid order");
    };
    if creator_node.namespace != Ns::Dc
        || creator_node.local != "creator"
        || date_node.namespace != Ns::Dc
        || date_node.local != "date"
    {
        return invalid("OTH office:change-info children have invalid order");
    }
    let creator = plain_text(creator_node)?;
    let date = plain_text(date_node)?;
    validate_date_or_datetime(&date)?;
    let paragraph_count = elements(node)
        .skip(2)
        .try_fold(0usize, |count, paragraph| {
            if paragraph.namespace != Ns::Text || paragraph.local != "p" {
                return Err(Error::InvalidFormat(
                    "OTH office:change-info children have invalid order".to_string(),
                ));
            }
            count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("OTH change-info paragraph count overflow".to_string())
            })
        })?;
    let mut paragraphs = Vec::new();
    reserve_vec_exact_charged(
        &mut paragraphs,
        paragraph_count,
        "OTH change-info paragraphs",
        |_| Ok(()),
    )?;
    for paragraph in elements(node).skip(2) {
        paragraphs.push(crate::paragraph::Paragraph::new(plain_text(paragraph)?));
    }
    Ok(crate::change::ChangeInfo::projected(
        creator, date, paragraphs,
    ))
}

fn change_text(node: &Node) -> Result<String> {
    if node.namespace != Ns::Text || node.local != "changed-region" {
        return plain_text(node);
    }
    let content = direct_change_child(node)?;
    validate_change_content(content)?;
    if content.local != "deletion" {
        return Ok(String::new());
    }
    let mut info_seen = false;
    let mut text_len = 0usize;
    for child in &content.children {
        match child {
            Child::Element(child)
                if child.namespace == Ns::Office && child.local == "change-info" =>
            {
                info_seen = true;
            },
            Child::Element(child) if info_seen => {
                text_len = text_len
                    .checked_add(plain_text_content_len(child)?)
                    .ok_or_else(|| {
                        Error::InvalidFormat("OTH deletion text size overflow".to_string())
                    })?;
            },
            Child::Text(value) if !value.trim().is_empty() => {
                return invalid("OTH deletion contains unexpected direct text");
            },
            Child::Text(_) | Child::Element(_) => {},
        }
    }
    let mut text = String::new();
    reserve_string_exact(&mut text, text_len, "OTH deletion text")?;
    info_seen = false;
    for child in &content.children {
        match child {
            Child::Element(child)
                if child.namespace == Ns::Office && child.local == "change-info" =>
            {
                info_seen = true;
            },
            Child::Element(child) if info_seen => append_plain_text(&mut text, child)?,
            Child::Text(value) if !value.trim().is_empty() => {
                return invalid("OTH deletion contains unexpected direct text");
            },
            Child::Text(_) | Child::Element(_) => {},
        }
    }
    Ok(text)
}

fn change_kind(node: &Node) -> crate::change::Kind {
    if node.namespace == Ns::Text && node.local == "change" {
        return crate::change::Kind::Other("change".to_string());
    }
    if let Some(child) = elements(node).find(|child| {
        child.namespace == Ns::Text
            && matches!(
                child.local.as_str(),
                "insertion" | "deletion" | "format-change"
            )
    }) {
        return match child.local.as_str() {
            "insertion" => crate::change::Kind::Insertion,
            "deletion" => crate::change::Kind::Deletion,
            "format-change" => crate::change::Kind::Format,
            other => crate::change::Kind::Other(other.to_owned()),
        };
    }
    crate::change::Kind::Other("changed-region".to_string())
}

fn collect_indexes(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Text && index_kind(&node.local).is_some() {
        let source_name = index_source_name(&node.local);
        for child in elements(node) {
            if child.namespace == Ns::Text
                && is_known_index_source_name(&child.local)
                && child.local != source_name
            {
                return invalid("OTH index contains a source for another index family");
            }
        }
        let source_node = unique_direct_child(node, Ns::Text, source_name, "index source")?;
        let name = required_attr(node, Ns::Text, b"name", "text:name")?;
        let Some(source_node) = source_node else {
            return invalid("OTH index requires its matching source element");
        };
        let index_body = unique_direct_child(node, Ns::Text, "index-body", "index-body")?;
        let planned = measure_index_node(node, source_node)?;
        output.ensure(planned)?;
        let index_len = output.indexes.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH index projection length overflow".to_string())
        })?;
        reserve_retained_vec(
            &mut output.indexes,
            index_len,
            &mut output.retained_bytes,
            "OTH index projection",
        )?;
        output.account(planned)?;
        let source = Some(plain_text(source_node)?);
        let body = index_body.map(plain_text).transpose()?.unwrap_or_default();
        let index = crate::index::Index::projected(
            index_kind(&node.local)
                .unwrap_or_else(|| crate::index::Kind::Other(node.local.clone())),
            Some(name.to_owned()),
            bool_attr(node, Ns::Text, b"protected")?,
            optional_schema_bool_attr(node, Ns::Text, b"protected")?,
            attr(node, Ns::Text, b"protection-key").map(str::to_owned),
            optional_any_iri_attr(node, Ns::Text, b"protection-key-digest-algorithm")?,
            source,
            project_index_source(node, Some(source_node))?,
            attr(node, Ns::Text, b"style-name").map(str::to_owned),
            attr(node, Ns::Xml, b"id").map(str::to_owned),
            body,
        );
        output.indexes.push(index);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn project_index_source(
    node: &Node,
    source: Option<&Node>,
) -> Result<Option<crate::index::IndexSource>> {
    let Some(source) = source else {
        return Ok(None);
    };
    let scope = attr(source, Ns::Text, b"index-scope")
        .map(|value| match value {
            "document" => Ok(crate::index::IndexScope::Document),
            "chapter" => Ok(crate::index::IndexScope::Chapter),
            _ => invalid("OTH index-scope value is invalid"),
        })
        .transpose()?;
    let relative = optional_schema_bool_attr(source, Ns::Text, b"relative-tab-stop-position")?;
    let options = match node.local.as_str() {
        "table-of-content" => crate::index::IndexSourceOptions::TableOfContents {
            outline_level: optional_positive_attr(source, b"outline-level")?,
            use_outline_level: optional_schema_bool_attr(source, Ns::Text, b"use-outline-level")?,
            use_index_marks: optional_schema_bool_attr(source, Ns::Text, b"use-index-marks")?,
            use_index_source_styles: optional_schema_bool_attr(
                source,
                Ns::Text,
                b"use-index-source-styles",
            )?,
        },
        "illustration-index" | "table-index" => {
            let caption_sequence_format = attr(source, Ns::Text, b"caption-sequence-format")
                .map(|value| match value {
                    "text" => Ok(crate::index::CaptionSequenceFormat::Text),
                    "category-and-value" => {
                        Ok(crate::index::CaptionSequenceFormat::CategoryAndValue)
                    },
                    "caption" => Ok(crate::index::CaptionSequenceFormat::Caption),
                    _ => invalid("OTH caption-sequence-format value is invalid"),
                })
                .transpose()?;
            crate::index::IndexSourceOptions::Illustration {
                use_caption: optional_schema_bool_attr(source, Ns::Text, b"use-caption")?,
                caption_sequence_name: attr(source, Ns::Text, b"caption-sequence-name")
                    .map(str::to_owned),
                caption_sequence_format,
            }
        },
        "object-index" => crate::index::IndexSourceOptions::Object {
            use_spreadsheet_objects: optional_schema_bool_attr(
                source,
                Ns::Text,
                b"use-spreadsheet-objects",
            )?,
            use_math_objects: optional_schema_bool_attr(source, Ns::Text, b"use-math-objects")?,
            use_draw_objects: optional_schema_bool_attr(source, Ns::Text, b"use-draw-objects")?,
            use_chart_objects: optional_schema_bool_attr(source, Ns::Text, b"use-chart-objects")?,
            use_other_objects: optional_schema_bool_attr(source, Ns::Text, b"use-other-objects")?,
        },
        "user-index" => {
            let index_name = required_attr(source, Ns::Text, b"index-name", "text:index-name")?;
            if index_name.is_empty() {
                return invalid("OTH text:index-name must not be empty");
            }
            crate::index::IndexSourceOptions::User {
                use_index_marks: optional_schema_bool_attr(source, Ns::Text, b"use-index-marks")?,
                use_index_source_styles: optional_schema_bool_attr(
                    source,
                    Ns::Text,
                    b"use-index-source-styles",
                )?,
                use_graphics: optional_schema_bool_attr(source, Ns::Text, b"use-graphics")?,
                use_tables: optional_schema_bool_attr(source, Ns::Text, b"use-tables")?,
                use_floating_frames: optional_schema_bool_attr(
                    source,
                    Ns::Text,
                    b"use-floating-frames",
                )?,
                use_objects: optional_schema_bool_attr(source, Ns::Text, b"use-objects")?,
                copy_outline_levels: optional_schema_bool_attr(
                    source,
                    Ns::Text,
                    b"copy-outline-levels",
                )?,
                index_name: index_name.to_owned(),
            }
        },
        "alphabetical-index" => crate::index::IndexSourceOptions::Alphabetical {
            ignore_case: optional_schema_bool_attr(source, Ns::Text, b"ignore-case")?,
            main_entry_style_name: attr(source, Ns::Text, b"main-entry-style-name")
                .map(str::to_owned),
            alphabetical_separators: optional_schema_bool_attr(
                source,
                Ns::Text,
                b"alphabetical-separators",
            )?,
            combine_entries: optional_schema_bool_attr(source, Ns::Text, b"combine-entries")?,
            combine_entries_with_dash: optional_schema_bool_attr(
                source,
                Ns::Text,
                b"combine-entries-with-dash",
            )?,
            combine_entries_with_pp: optional_schema_bool_attr(
                source,
                Ns::Text,
                b"combine-entries-with-pp",
            )?,
            use_keys_as_entries: optional_schema_bool_attr(
                source,
                Ns::Text,
                b"use-keys-as-entries",
            )?,
            capitalize_entries: optional_schema_bool_attr(source, Ns::Text, b"capitalize-entries")?,
            comma_separated: optional_schema_bool_attr(source, Ns::Text, b"comma-separated")?,
            language: attr(source, Ns::Fo, b"language").map(str::to_owned),
            country: attr(source, Ns::Fo, b"country").map(str::to_owned),
            script: attr(source, Ns::Fo, b"script").map(str::to_owned),
            rfc_language_tag: attr(source, Ns::Style, b"rfc-language-tag").map(str::to_owned),
            sort_algorithm: attr(source, Ns::Text, b"sort-algorithm").map(str::to_owned),
        },
        "bibliography" | "bibliography-index" => crate::index::IndexSourceOptions::Bibliography,
        _ => return Ok(None),
    };
    Ok(Some(crate::index::IndexSource::projected(
        scope, relative, options,
    )))
}

fn index_source_name(index_name: &str) -> &str {
    match index_name {
        "table-of-content" => "table-of-content-source",
        "illustration-index" => "illustration-index-source",
        "table-index" => "table-index-source",
        "object-index" => "object-index-source",
        "user-index" => "user-index-source",
        "alphabetical-index" => "alphabetical-index-source",
        "bibliography" | "bibliography-index" => "bibliography-source",
        _ => "",
    }
}

fn is_known_index_source_name(name: &str) -> bool {
    matches!(
        name,
        "table-of-content-source"
            | "illustration-index-source"
            | "table-index-source"
            | "object-index-source"
            | "user-index-source"
            | "alphabetical-index-source"
            | "bibliography-source"
    )
}

fn index_kind(local: &str) -> Option<crate::index::Kind> {
    Some(match local {
        "table-of-content" => crate::index::Kind::TableOfContents,
        "illustration-index" => crate::index::Kind::Illustration,
        "table-index" => crate::index::Kind::Table,
        "object-index" => crate::index::Kind::Object,
        "user-index" => crate::index::Kind::User,
        "alphabetical-index" => crate::index::Kind::Alphabetical,
        "bibliography" | "bibliography-index" => crate::index::Kind::Bibliography,
        _ => return None,
    })
}

fn collect_frames(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Draw && matches!(node.local.as_str(), "frame" | "text-box") {
        let planned = measure_frame_node(node)?;
        output.ensure(planned)?;
        let frame_len = output.frames.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH frame projection length overflow".to_string())
        })?;
        reserve_retained_vec(
            &mut output.frames,
            frame_len,
            &mut output.retained_bytes,
            "OTH frame projection",
        )?;
        output.account(planned)?;
        let kind = frame_kind(node);
        let href = find_descendant(node, Ns::Draw, "image")
            .or_else(|| find_descendant(node, Ns::Draw, "object"))
            .or_else(|| find_descendant(node, Ns::Draw, "object-ole"))
            .or_else(|| find_descendant(node, Ns::Draw, "plugin"))
            .or_else(|| find_descendant(node, Ns::Draw, "floating-frame"))
            .and_then(|value| attr(value, Ns::Xlink, b"href"))
            .map(str::to_owned);
        let frame = crate::frame::Frame::projected(
            kind,
            attr(node, Ns::Draw, b"name").map(str::to_owned),
            attr(node, Ns::Draw, b"style-name").map(str::to_owned),
            attr(node, Ns::Text, b"anchor-type").map(str::to_owned),
            attr(node, Ns::Svg, b"x").map(str::to_owned),
            attr(node, Ns::Svg, b"y").map(str::to_owned),
            attr(node, Ns::Svg, b"width").map(str::to_owned),
            attr(node, Ns::Svg, b"height").map(str::to_owned),
            href,
            plain_text(node)?,
        );
        output.frames.push(frame);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn frame_kind(node: &Node) -> crate::frame::Kind {
    if node.namespace == Ns::Draw && node.local == "text-box"
        || find_descendant(node, Ns::Draw, "text-box").is_some()
    {
        return crate::frame::Kind::TextBox;
    }
    if find_descendant(node, Ns::Draw, "image").is_some() {
        crate::frame::Kind::Image
    } else if find_descendant(node, Ns::Draw, "object-ole").is_some() {
        crate::frame::Kind::OleObject
    } else if find_descendant(node, Ns::Draw, "object").is_some() {
        crate::frame::Kind::Object
    } else if find_descendant(node, Ns::Draw, "plugin").is_some() {
        crate::frame::Kind::Plugin
    } else if find_descendant(node, Ns::Draw, "floating-frame").is_some() {
        crate::frame::Kind::FloatingFrame
    } else {
        crate::frame::Kind::Generic
    }
}

fn collect_rubies(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Text && node.local == "ruby" {
        let planned = measure_ruby_node(node)?;
        output.ensure(planned)?;
        let ruby_len = output.rubies.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH ruby projection length overflow".to_string())
        })?;
        reserve_retained_vec(
            &mut output.rubies,
            ruby_len,
            &mut output.retained_bytes,
            "OTH ruby projection",
        )?;
        output.account(planned)?;
        let base = direct_child(node, Ns::Text, "ruby-base")
            .map(plain_text)
            .transpose()?
            .unwrap_or_default();
        let ruby_text = direct_child(node, Ns::Text, "ruby-text");
        let ruby = crate::ruby::Ruby::projected(
            attr(node, Ns::Text, b"style-name").map(str::to_owned),
            ruby_text.and_then(|value| attr(value, Ns::Text, b"style-name").map(str::to_owned)),
            base,
            ruby_text.map(plain_text).transpose()?.unwrap_or_default(),
        );
        output.rubies.push(ruby);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn elements(node: &Node) -> impl Iterator<Item = &Node> {
    node.children.iter().filter_map(|child| match child {
        Child::Element(element) => Some(element.as_ref()),
        Child::Text(_) => None,
    })
}

fn valid_structure_child(node: &Node) -> bool {
    if node.namespace == Ns::Other && node.namespace_uri.is_none() {
        return false;
    }
    context_kind(node.namespace, node.local.as_bytes()).permits_body_structure()
}

fn direct_child<'a>(node: &'a Node, namespace: Ns, local: &str) -> Option<&'a Node> {
    elements(node).find(|child| child.namespace == namespace && child.local == local)
}

fn unique_direct_child<'a>(
    node: &'a Node,
    namespace: Ns,
    local: &str,
    label: &str,
) -> Result<Option<&'a Node>> {
    let mut found = None;
    for child in elements(node) {
        if child.namespace == namespace && child.local == local {
            if found.is_some() {
                return invalid(&format!("OTH {label} occurs more than once"));
            }
            found = Some(child);
        }
    }
    Ok(found)
}

fn descendant<'a>(node: &'a Node, namespace: Ns, local: &str) -> Option<&'a Node> {
    if node.namespace == namespace && node.local == local {
        return Some(node);
    }
    elements(node).find_map(|child| descendant(child, namespace, local))
}

fn find_descendant<'a>(node: &'a Node, namespace: Ns, local: &str) -> Option<&'a Node> {
    elements(node).find_map(|child| descendant(child, namespace, local))
}

fn attr<'a>(node: &'a Node, namespace: Ns, local: &[u8]) -> Option<&'a str> {
    node.attribute(namespace, local)
}

fn optional_attr_from_start<'a>(
    namespaces: &SemanticNamespaceContext<'_>,
    decoder: quick_xml::encoding::Decoder,
    start: &'a quick_xml::events::BytesStart<'a>,
    namespace: Ns,
    local: &[u8],
    budget: &mut Budget,
) -> Result<Option<String>> {
    for raw in start.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH metadata attribute: {error}"))
        })?;
        if namespace_declaration(attribute.key).is_some() {
            continue;
        }
        let resolved = namespaces.resolve_attribute(attribute.key)?;
        let resolved_namespace = local_namespace(resolved.namespace.as_ref());
        let attribute_local = resolved.local_bytes();
        if resolved_namespace != namespace || attribute_local != local {
            continue;
        }
        budget.text(attribute.value.len())?;
        budget.storage(size_of::<String>())?;
        budget.storage(attribute.value.len())?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid OTH metadata attribute value: {error}"))
            })?;
        budget.text_after_raw(attribute.value.len(), value.len())?;
        if value.len() > attribute.value.len() {
            budget.storage(value.len() - attribute.value.len())?;
        }
        return Ok(Some(value.into_owned()));
    }
    Ok(None)
}

fn parse_odf_bool(value: &str, field: &str) -> Result<bool> {
    Boolean::decode(value)
        .map_err(|_| Error::InvalidFormat(format!("invalid OTH {field} boolean lexical value")))
}

fn bool_attr(node: &Node, namespace: Ns, local: &[u8]) -> Result<bool> {
    Ok(optional_bool_attr(node, namespace, local)?.unwrap_or(false))
}

fn optional_schema_bool_attr(node: &Node, namespace: Ns, local: &[u8]) -> Result<Option<bool>> {
    attr(node, namespace, local)
        .map(|value| parse_odf_bool(value, "schema boolean"))
        .transpose()
}

fn optional_bool_attr(node: &Node, namespace: Ns, local: &[u8]) -> Result<Option<bool>> {
    let Some(value) = attr(node, namespace, local) else {
        return Ok(None);
    };
    match value {
        "true" | "1" => Ok(Some(true)),
        "false" | "0" => Ok(Some(false)),
        _ => invalid("OTH boolean body attribute is invalid"),
    }
}

fn count_attr(node: &Node, local: &[u8]) -> Result<usize> {
    count_attr_default(node, local, 1)
}

fn optional_count_attr(node: &Node, local: &[u8]) -> Result<Option<usize>> {
    let Some(value) = attr(node, Ns::Table, local) else {
        return Ok(None);
    };
    let parsed = value.parse::<usize>().map_err(|_| {
        Error::InvalidFormat("OTH table repetition attribute is invalid".to_string())
    })?;
    if parsed == 0 || parsed > MAX_REPEAT {
        return invalid("OTH table repetition attribute is outside the supported range");
    }
    Ok(Some(parsed))
}

fn optional_positive_attr(node: &Node, local: &[u8]) -> Result<Option<usize>> {
    let Some(value) = attr(node, Ns::Text, local) else {
        return Ok(None);
    };
    let parsed = value
        .parse::<usize>()
        .map_err(|_| Error::InvalidFormat("OTH positive index attribute is invalid".to_string()))?;
    if parsed == 0 || parsed > MAX_REPEAT {
        return invalid("OTH positive index attribute is outside the supported range");
    }
    Ok(Some(parsed))
}

fn count_attr_default(node: &Node, local: &[u8], default: usize) -> Result<usize> {
    Ok(optional_count_attr(node, local)?.unwrap_or(default))
}

fn optional_visibility_attr(node: &Node) -> Result<Option<crate::table::Visibility>> {
    attr(node, Ns::Table, b"visibility")
        .map(|value| match value {
            "visible" => Ok(crate::table::Visibility::Visible),
            "collapse" => Ok(crate::table::Visibility::Collapse),
            "filter" => Ok(crate::table::Visibility::Filter),
            _ => invalid("OTH table:visibility value is invalid"),
        })
        .transpose()
}

fn required_attr<'a>(node: &'a Node, namespace: Ns, local: &[u8], field: &str) -> Result<&'a str> {
    attr(node, namespace, local)
        .ok_or_else(|| Error::InvalidFormat(format!("OTH {field} is required")))
}

fn optional_any_iri_attr(node: &Node, namespace: Ns, local: &[u8]) -> Result<Option<String>> {
    attr(node, namespace, local)
        .map(|value| {
            validate_any_iri(value, "IRI")?;
            Ok(value.to_owned())
        })
        .transpose()
}

fn validate_double(value: &str, field: &str) -> Result<()> {
    if matches!(value, "INF" | "-INF" | "NaN") {
        return Ok(());
    }
    let parsed = value
        .parse::<f64>()
        .map_err(|_| Error::InvalidFormat(format!("OTH {field} numeric value is invalid")))?;
    if !parsed.is_finite() {
        return invalid("OTH numeric value must be finite");
    }
    Ok(())
}

fn validate_date_or_datetime(value: &str) -> Result<()> {
    if date::is_xsd_date_or_datetime(value) {
        Ok(())
    } else {
        invalid("OTH office:date-value is invalid")
    }
}

fn validate_any_iri(value: &str, field: &str) -> Result<()> {
    if value.chars().any(char::is_control) || value.chars().any(char::is_whitespace) {
        return Err(Error::InvalidFormat(format!(
            "OTH {field} anyIRI value is invalid"
        )));
    }
    Ok(())
}

fn validate_uri_or_safe_curie(value: &str, field: &str) -> Result<()> {
    if let Some(inner) = value
        .strip_prefix('[')
        .and_then(|rest| rest.strip_suffix(']'))
    {
        if inner.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "OTH {field} value is invalid"
            )));
        }
        if inner.chars().any(char::is_whitespace) || inner.chars().any(char::is_control) {
            return Err(Error::InvalidFormat(format!(
                "OTH {field} value is invalid"
            )));
        }
        if inner.contains(':') {
            validate_curie(inner, field)
        } else {
            validate_any_iri(inner, field)
        }
    } else {
        validate_any_iri(value, field)
    }
}

fn validate_curies(value: &str, field: &str) -> Result<()> {
    let mut count = 0usize;
    for token in value.split_whitespace() {
        count = count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat(format!("OTH {field} CURIE count overflow")))?;
        validate_curie(token, field)?;
    }
    if count == 0 {
        return Err(Error::InvalidFormat(format!(
            "OTH {field} must not be empty"
        )));
    }
    Ok(())
}

fn validate_curie(value: &str, field: &str) -> Result<()> {
    let (prefix, suffix) = value.split_once(':').unwrap_or(("", value));
    if suffix.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "OTH {field} CURIE is invalid"
        )));
    }
    if !prefix.is_empty() {
        validate_ncname(prefix, field)?;
    }
    if suffix.chars().any(char::is_control)
        || suffix.chars().any(char::is_whitespace)
        || suffix
            .chars()
            .any(|character| matches!(character, '^' | '"' | '<' | '>' | '`'))
    {
        return Err(Error::InvalidFormat(format!(
            "OTH {field} CURIE is invalid"
        )));
    }
    Ok(())
}

fn append_plain_text(output: &mut String, node: &Node) -> Result<()> {
    if matches!(node.namespace, Ns::Other | Ns::None) {
        return Ok(());
    }
    if node.namespace == Ns::Text {
        match node.local.as_str() {
            "s" => {
                let count = node.attribute(Ns::Text, b"c").map_or(Ok(1), |value| {
                    value.parse::<usize>().map_err(|_| {
                        Error::InvalidFormat("OTH text:s count is invalid".to_string())
                    })
                })?;
                if count == 0 || count > MAX_STRUCTURE_TEXT {
                    return invalid("OTH text:s count is outside the supported range");
                }
                checked_append_repeated(output, " ", count)?;
                return Ok(());
            },
            "tab" => {
                checked_append(output, "\t")?;
                return Ok(());
            },
            "line-break" => {
                checked_append(output, "\n")?;
                return Ok(());
            },
            _ => {},
        }
    }
    for child in &node.children {
        match child {
            Child::Element(element) => append_plain_text(output, element)?,
            Child::Text(text) => checked_append(output, text)?,
        }
    }
    Ok(())
}

fn plain_text(node: &Node) -> Result<String> {
    let mut output = String::new();
    reserve_string_exact(
        &mut output,
        plain_text_content_len(node)?,
        "OTH body structure text",
    )?;
    append_plain_text(&mut output, node)?;
    Ok(output)
}

fn checked_append(output: &mut String, value: &str) -> Result<()> {
    let target = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    if target > MAX_STRUCTURE_TEXT {
        return invalid("OTH structure text exceeds the limit");
    }
    reserve_string_charged(output, target, "OTH body structure text", |_| Ok(()))?;
    output.push_str(value);
    Ok(())
}

fn checked_append_with_budget(output: &mut String, value: &str, budget: &mut Budget) -> Result<()> {
    let target = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    if target > MAX_STRUCTURE_TEXT {
        return invalid("OTH structure text exceeds the limit");
    }
    reserve_string_charged(output, target, "OTH structural edit text", |bytes| {
        budget.storage(bytes)
    })?;
    output.push_str(value);
    Ok(())
}

fn checked_append_repeated(output: &mut String, value: &str, count: usize) -> Result<()> {
    let additional = value
        .len()
        .checked_mul(count)
        .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    let target = output
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    if target > MAX_STRUCTURE_TEXT {
        return invalid("OTH structure text exceeds the limit");
    }
    reserve_string_exact(output, target, "OTH body structure text")?;
    for _ in 0..count {
        output.push_str(value);
    }
    Ok(())
}

fn add_string(total: &mut usize, value: Option<&str>) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    add_len(total, value.len())
}

const VECTOR_GROWTH_MIN: usize = 8;

fn geometric_capacity(current: usize, required: usize) -> Result<usize> {
    if required <= current {
        return Ok(current);
    }
    let mut capacity = current.max(VECTOR_GROWTH_MIN);
    while capacity < required {
        capacity = capacity.checked_mul(2).unwrap_or(required).max(required);
    }
    Ok(capacity)
}

fn reserve_vec_charged<T, F>(
    values: &mut Vec<T>,
    required_len: usize,
    resource: &'static str,
    mut charge: F,
) -> Result<()>
where
    F: FnMut(usize) -> Result<()>,
{
    let current_capacity = values.capacity();
    let target_capacity = geometric_capacity(current_capacity, required_len)?;
    if target_capacity <= current_capacity {
        return Ok(());
    }
    let growth = target_capacity
        .checked_sub(current_capacity)
        .ok_or_else(|| Error::InvalidFormat("OTH vector capacity overflow".to_string()))?;
    let bytes = size_of::<T>()
        .checked_mul(growth)
        .ok_or_else(|| Error::InvalidFormat("OTH vector storage size overflow".to_string()))?;
    charge(bytes)?;
    let additional = target_capacity
        .checked_sub(values.len())
        .ok_or_else(|| Error::InvalidFormat("OTH vector reservation underflow".to_string()))?;
    values
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

fn reserve_vec_exact_charged<T, F>(
    values: &mut Vec<T>,
    required_len: usize,
    resource: &'static str,
    mut charge: F,
) -> Result<()>
where
    F: FnMut(usize) -> Result<()>,
{
    let current_capacity = values.capacity();
    if required_len <= current_capacity {
        return Ok(());
    }
    let growth = required_len
        .checked_sub(current_capacity)
        .ok_or_else(|| Error::InvalidFormat("OTH vector capacity overflow".to_string()))?;
    let bytes = size_of::<T>()
        .checked_mul(growth)
        .ok_or_else(|| Error::InvalidFormat("OTH vector storage size overflow".to_string()))?;
    charge(bytes)?;
    let additional = required_len
        .checked_sub(values.len())
        .ok_or_else(|| Error::InvalidFormat("OTH vector reservation underflow".to_string()))?;
    values
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

fn reserve_retained_vec<T>(
    values: &mut Vec<T>,
    required_len: usize,
    retained_bytes: &mut usize,
    resource: &'static str,
) -> Result<()> {
    let current_capacity = values.capacity();
    let target_capacity = geometric_capacity(current_capacity, required_len)?;
    if target_capacity <= current_capacity {
        return Ok(());
    }
    let growth = target_capacity
        .checked_sub(current_capacity)
        .ok_or_else(|| Error::InvalidFormat("OTH retained vector capacity overflow".to_string()))?;
    let bytes = size_of::<T>()
        .checked_mul(growth)
        .ok_or_else(|| Error::InvalidFormat("OTH retained vector storage overflow".to_string()))?;
    let total = retained_bytes
        .checked_add(bytes)
        .ok_or_else(|| Error::InvalidFormat("OTH retained vector budget overflow".to_string()))?;
    if total > MAX_STRUCTURE_TEXT {
        return invalid("OTH projected body structures exceed the aggregate text limit");
    }
    *retained_bytes = total;
    reserve_vec_exact_charged(values, target_capacity, resource, |_| Ok(()))
}

fn reserve_string_charged<F>(
    value: &mut String,
    required_len: usize,
    resource: &'static str,
    mut charge: F,
) -> Result<()>
where
    F: FnMut(usize) -> Result<()>,
{
    let current_capacity = value.capacity();
    let target_capacity = geometric_capacity(current_capacity, required_len)?;
    if target_capacity <= current_capacity {
        return Ok(());
    }
    let growth = target_capacity
        .checked_sub(current_capacity)
        .ok_or_else(|| Error::InvalidFormat("OTH string capacity overflow".to_string()))?;
    charge(growth)?;
    let additional = target_capacity
        .checked_sub(value.len())
        .ok_or_else(|| Error::InvalidFormat("OTH string reservation underflow".to_string()))?;
    value
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

fn reserve_string_exact(
    value: &mut String,
    required_len: usize,
    resource: &'static str,
) -> Result<()> {
    let current_capacity = value.capacity();
    if required_len <= current_capacity {
        return Ok(());
    }
    let additional = required_len
        .checked_sub(value.len())
        .ok_or_else(|| Error::InvalidFormat("OTH string reservation underflow".to_string()))?;
    value
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

fn reserve_exact<T>(values: &mut Vec<T>, additional: usize, resource: &'static str) -> Result<()> {
    let required_len = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("OTH vector length overflow".to_string()))?;
    reserve_vec_exact_charged(values, required_len, resource, |_| Ok(()))
}

fn reference_value(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(value) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!(
            "invalid OTH structure character reference: {error}"
        ))
    })? {
        return Ok(value.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid OTH structure entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "apos" => Ok("'".to_string()),
        "gt" => Ok(">".to_string()),
        "lt" => Ok("<".to_string()),
        "quot" => Ok("\"".to_string()),
        _ => invalid("OTH custom entities are not allowed"),
    }
}

/// Checks references outside an excluded structure using resolved namespaces
/// and XML/URI decoding.  A source-bound removal must make this check before
/// deleting the selected span; it never dereferences or fetches an external
/// URI.
pub(crate) fn has_semantic_reference(
    xml: &str,
    excluded: Range<usize>,
    identity: &str,
) -> Result<bool> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut namespaces = SemanticNamespaceContext::new();
    let mut budget = Budget {
        nodes: 0,
        text_bytes: 0,
        storage_bytes: 0,
    };
    loop {
        let event_start = source_offset(reader.buffer_position())?;
        let event = reader.read_event().map_err(|error| xml_error(&error))?;
        let event_end = source_offset(reader.buffer_position())?;
        let excluded_event = event_start >= excluded.start && event_end <= excluded.end;
        let empty_event = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) if !excluded_event => {
                namespaces.push(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                for raw in start.attributes() {
                    let attribute = raw.map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid OTH structural reference attribute: {error}"
                        ))
                    })?;
                    if namespace_declaration(attribute.key).is_some() {
                        continue;
                    }
                    let resolved = namespaces.resolve_attribute(attribute.key)?;
                    let namespace = local_namespace(resolved.namespace.as_ref());
                    if !reference_attribute(namespace, resolved.local_bytes()) {
                        continue;
                    }
                    let raw_attribute_bytes = attribute
                        .key
                        .as_ref()
                        .len()
                        .checked_add(attribute.value.len())
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "OTH structural reference attribute size overflow".to_string(),
                            )
                        })?;
                    budget.text(raw_attribute_bytes)?;
                    budget.storage(size_of::<String>())?;
                    budget.storage(attribute.value.len())?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid OTH structural reference value: {error}"
                            ))
                        })?;
                    budget.text_after_raw(raw_attribute_bytes, value.len())?;
                    if value.len() > attribute.value.len() {
                        budget.storage(value.len() - attribute.value.len())?;
                    }
                    if semantic_reference_value(value.as_ref(), identity, &mut budget)? {
                        return Ok(true);
                    }
                }
                if empty_event {
                    namespaces.end()?;
                }
            },
            Event::Start(start) | Event::Empty(start) if excluded_event => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                if empty_event {
                    namespaces.end()?;
                }
            },
            Event::End(_) => {
                namespaces.end()?;
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::Eof => {
                if !namespaces.is_empty() {
                    return invalid("OTH structural reference namespace stack is not empty");
                }
                return Ok(false);
            },
            _ => {},
        }
    }
}

fn reference_attribute(namespace: Ns, local: &[u8]) -> bool {
    match namespace {
        Ns::Text => matches!(
            local,
            b"section-name" | b"change-id" | b"id" | b"name" | b"ref-name"
        ),
        Ns::Xlink => local == b"href",
        _ => false,
    }
}

fn semantic_reference_value(value: &str, identity: &str, budget: &mut Budget) -> Result<bool> {
    let Some(value) = percent_decode(value, budget)? else {
        return Ok(false);
    };
    if value == identity {
        return Ok(true);
    }
    Ok(value
        .rsplit_once('#')
        .is_some_and(|(_, fragment)| fragment == identity))
}

fn percent_decode(value: &str, budget: &mut Budget) -> Result<Option<String>> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::new();
    budget.reserve_vec_exact(&mut decoded, bytes.len(), "OTH structural reference value")?;
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            decoded.push(bytes[index]);
            index += 1;
            continue;
        }
        if index + 2 >= bytes.len() {
            return Ok(None);
        }
        let Some(high) = hex_digit(bytes[index + 1]) else {
            return Ok(None);
        };
        let Some(low) = hex_digit(bytes[index + 2]) else {
            return Ok(None);
        };
        decoded.push((high << 4) | low);
        index += 3;
    }
    budget.storage(size_of::<String>())?;
    Ok(String::from_utf8(decoded).ok())
}

fn hex_digit(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

/// Finds the lossless text spans admitted by the first structural write slice.
///
/// A selector is source-order within one namespace-qualified family.  The
/// scanner records every owned section, note, and annotation so empty or rich
/// structures cannot silently change selector numbering.  Only a single plain
/// `text:p`/`text:h` child is editable; all other markup remains untouched and
/// is refused when it would make that span ambiguous.
pub(crate) fn structure_sites(xml: &str) -> Result<Vec<StructureSite>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut namespaces = SemanticNamespaceContext::new();
    let mut contexts = Vec::<ContextKind>::new();
    let mut target_ids = Vec::<Option<usize>>::new();
    let mut targets = Vec::<Option<ActiveSite>>::new();
    let mut counts = [0_usize; 4];
    let mut sites = Vec::new();
    let mut active_block = None::<ActiveBlockSite>;
    let mut budget = Budget {
        nodes: 0,
        text_bytes: 0,
        storage_bytes: 0,
    };

    loop {
        let event_start = source_offset(reader.buffer_position())?;
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                budget.count_node()?;
                if contexts.len() >= MAX_STRUCTURE_SITE_DEPTH {
                    return invalid("OTH structural edit depth exceeds the limit");
                }
                let (_namespace, context, root) = classify_element(&namespaces, start.name())?;
                if let Some(active) = active_block.as_mut() {
                    active.invalid = true;
                }
                let root_id = match editable_kind(root) {
                    Some(kind) if owns_structure_root(&contexts) => {
                        let index = kind.index(&mut counts)?;
                        let target_len = targets.len().checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "OTH structural edit target length overflow".to_string(),
                            )
                        })?;
                        budget.reserve_vec(
                            &mut targets,
                            target_len,
                            "OTH structural edit sites",
                        )?;
                        targets.push(Some(ActiveSite {
                            full_start: event_start,
                            kind,
                            index,
                            identity: structure_identity(
                                &namespaces,
                                reader.decoder(),
                                &start,
                                kind,
                                &mut budget,
                            )?,
                            root_depth: contexts.len(),
                            block_count: 0,
                            invalid: false,
                            text: None,
                            text_value: None,
                        }));
                        Some(targets.len() - 1)
                    },
                    _ => None,
                };
                let block = if matches!(
                    context,
                    ContextKind::TextParagraph | ContextKind::TextHeading
                ) {
                    if let Some(target_id) = block_owner(&targets, &contexts) {
                        if let Some(target) = targets.get_mut(target_id).and_then(Option::as_mut) {
                            target.block_count =
                                target.block_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "OTH structural edit block count overflow".to_string(),
                                    )
                                })?;
                            if target.block_count > 1 {
                                target.invalid = true;
                            }
                        }
                        budget.storage(size_of::<ActiveBlockSite>())?;
                        budget.storage(size_of::<String>())?;
                        Some(ActiveBlockSite {
                            target_id,
                            depth: contexts.len(),
                            content_start: source_offset(reader.buffer_position())?,
                            text: String::new(),
                            invalid: false,
                        })
                    } else {
                        None
                    }
                } else {
                    None
                };
                active_block = block;
                let context_len = contexts.len().checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("OTH structural edit context length overflow".to_string())
                })?;
                budget.reserve_vec(&mut contexts, context_len, "OTH structural edit contexts")?;
                let target_id_len = target_ids.len().checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("OTH structural edit root length overflow".to_string())
                })?;
                budget.reserve_vec(
                    &mut target_ids,
                    target_id_len,
                    "OTH structural edit root stack",
                )?;
                contexts.push(context);
                target_ids.push(root_id);
            },
            Event::Empty(start) => {
                namespaces.start(&start, reader.decoder(), &mut |bytes| budget.storage(bytes))?;
                budget.count_node()?;
                if contexts.len() >= MAX_STRUCTURE_SITE_DEPTH {
                    return invalid("OTH structural edit depth exceeds the limit");
                }
                let (_namespace, context, root) = classify_element(&namespaces, start.name())?;
                if let Some(active) = active_block.as_mut() {
                    active.invalid = true;
                }
                if let Some(kind) = editable_kind(root)
                    && owns_structure_root(&contexts)
                {
                    let index = kind.index(&mut counts)?;
                    let site_len = sites.len().checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("OTH structural edit site length overflow".to_string())
                    })?;
                    budget.reserve_vec(&mut sites, site_len, "OTH structural edit sites")?;
                    sites.push(StructureSite {
                        kind,
                        index,
                        identity: structure_identity(
                            &namespaces,
                            reader.decoder(),
                            &start,
                            kind,
                            &mut budget,
                        )?,
                        full: event_start..source_offset(reader.buffer_position())?,
                        text: None,
                        text_value: None,
                    });
                } else if matches!(
                    context,
                    ContextKind::TextParagraph | ContextKind::TextHeading
                ) && let Some(target_id) = block_owner(&targets, &contexts)
                    && let Some(target) = targets.get_mut(target_id).and_then(Option::as_mut)
                {
                    target.block_count = target.block_count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("OTH structural edit block count overflow".to_string())
                    })?;
                    target.invalid = true;
                    target.text = None;
                    target.text_value = None;
                }
                namespaces.end()?;
            },
            Event::Text(text) => {
                if let Some(active) = active_block.as_mut() {
                    let raw = text.as_ref().len();
                    budget.text(raw)?;
                    budget.storage(raw)?;
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid OTH structural edit text: {error}"))
                    })?;
                    budget.text_after_raw(raw, value.len())?;
                    budget.storage(value.len())?;
                    checked_append_with_budget(&mut active.text, &value, &mut budget)?;
                }
            },
            Event::GeneralRef(reference) if active_block.is_some() => {
                let raw = reference.as_ref().len();
                budget.text(raw)?;
                budget.storage(raw)?;
                let value = reference_value(&reference)?;
                budget.text_after_raw(raw, value.len())?;
                budget.storage(value.len())?;
                checked_append_with_budget(
                    &mut active_block
                        .as_mut()
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "OTH structural block state is missing".to_string(),
                            )
                        })?
                        .text,
                    &value,
                    &mut budget,
                )?;
            },
            Event::End(end) => {
                let (_namespace, context, _root) = classify_element(&namespaces, end.name())?;
                if matches!(
                    context,
                    ContextKind::TextParagraph | ContextKind::TextHeading
                ) && let Some(active) = active_block.take()
                {
                    if active.depth != contexts.len().saturating_sub(1) {
                        return invalid("OTH structural edit text depth is invalid");
                    }
                    if let Some(target) = targets.get_mut(active.target_id).and_then(Option::as_mut)
                    {
                        if !active.invalid && !target.invalid {
                            target.text = Some(crate::codec::ReplacementSite {
                                prefix: String::new(),
                                range: active.content_start..event_start,
                                suffix: String::new(),
                            });
                            target.text_value = Some(active.text);
                        } else {
                            target.text = None;
                            target.text_value = None;
                        }
                    }
                }
                let Some(open_context) = contexts.pop() else {
                    return invalid("OTH structural edit end tag has no matching start tag");
                };
                if open_context != context {
                    return invalid("OTH structural edit end tag does not match its start tag");
                }
                let root_id = target_ids.pop().ok_or_else(|| {
                    Error::InvalidFormat("OTH structural edit root stack is missing".to_string())
                })?;
                if let Some(root_id) = root_id {
                    let target = targets[root_id].take().ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH structural edit root state is missing".to_string(),
                        )
                    })?;
                    let site_len = sites.len().checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("OTH structural edit site length overflow".to_string())
                    })?;
                    budget.reserve_vec(&mut sites, site_len, "OTH structural edit sites")?;
                    sites.push(StructureSite {
                        kind: target.kind,
                        index: target.index,
                        identity: target.identity,
                        full: target.full_start..source_offset(reader.buffer_position())?,
                        text: target.text,
                        text_value: target.text_value,
                    });
                }
                namespaces.end()?;
            },
            Event::CData(_) | Event::Comment(_) | Event::PI(_) if active_block.is_some() => {
                if let Some(active) = active_block.as_mut() {
                    active.invalid = true;
                }
            },
            Event::DocType(_) => return invalid("OTH structural edit source cannot contain a DTD"),
            Event::Decl(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::GeneralRef(_)
            | Event::PI(_) => {},
            Event::Eof => {
                if !contexts.is_empty() || active_block.is_some() {
                    return invalid("OTH structural edit source is not closed");
                }
                if !namespaces.is_empty() {
                    return invalid("OTH structural edit namespace stack is not closed");
                }
                sites.sort_unstable_by_key(|site| site.full.start);
                return Ok(sites);
            },
        }
    }
}

#[derive(Clone)]
struct ActiveBlockSite {
    target_id: usize,
    depth: usize,
    content_start: usize,
    // The text is moved into the target when the block closes.  Keeping this
    // allocation local bounds a rejected rich block to one selected span.
    text: String,
    invalid: bool,
}

struct ActiveSite {
    full_start: usize,
    kind: EditableStructureKind,
    index: usize,
    identity: Option<String>,
    root_depth: usize,
    block_count: usize,
    invalid: bool,
    text: Option<crate::codec::ReplacementSite>,
    text_value: Option<String>,
}

impl EditableStructureKind {
    pub(crate) const fn to_public(self) -> crate::facade::StructureKind {
        match self {
            Self::Section => crate::facade::StructureKind::Section,
            Self::Note => crate::facade::StructureKind::Note,
            Self::Annotation => crate::facade::StructureKind::Annotation,
            Self::Table => crate::facade::StructureKind::Table,
        }
    }

    fn index(self, counts: &mut [usize; 4]) -> Result<usize> {
        let slot = match self {
            Self::Section => &mut counts[0],
            Self::Note => &mut counts[1],
            Self::Annotation => &mut counts[2],
            Self::Table => &mut counts[3],
        };
        let index = *slot;
        *slot = slot.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH structural edit selector overflow".to_string())
        })?;
        if *slot > MAX_STRUCTURE_NODES {
            return invalid("OTH structural edit selector count exceeds the limit");
        }
        Ok(index)
    }
}

fn editable_kind(kind: Option<RootKind>) -> Option<EditableStructureKind> {
    match kind {
        Some(RootKind::Section) => Some(EditableStructureKind::Section),
        Some(RootKind::Note) => Some(EditableStructureKind::Note),
        Some(RootKind::Annotation) => Some(EditableStructureKind::Annotation),
        Some(RootKind::Table) => Some(EditableStructureKind::Table),
        _ => None,
    }
}

fn block_owner(targets: &[Option<ActiveSite>], contexts: &[ContextKind]) -> Option<usize> {
    targets.iter().enumerate().rev().find_map(|(id, target)| {
        let target = target.as_ref()?;
        let direct = match target.kind {
            EditableStructureKind::Section | EditableStructureKind::Annotation => {
                contexts.len() == target.root_depth + 1
            },
            EditableStructureKind::Note => {
                contexts.len() == target.root_depth + 2
                    && contexts.get(target.root_depth + 1) == Some(&ContextKind::TextNoteBody)
            },
            EditableStructureKind::Table => false,
        };
        direct.then_some(id)
    })
}

fn structure_identity(
    namespaces: &SemanticNamespaceContext<'_>,
    decoder: quick_xml::encoding::Decoder,
    start: &quick_xml::events::BytesStart<'_>,
    kind: EditableStructureKind,
    budget: &mut Budget,
) -> Result<Option<String>> {
    let (namespace, local) = match kind {
        EditableStructureKind::Section => (Ns::Text, b"name".as_slice()),
        EditableStructureKind::Note => (Ns::Text, b"id".as_slice()),
        EditableStructureKind::Annotation => (Ns::Office, b"name".as_slice()),
        EditableStructureKind::Table => (Ns::Table, b"name".as_slice()),
    };
    for raw in start.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid OTH structural identity attribute: {error}"
            ))
        })?;
        if namespace_declaration(attribute.key).is_some() {
            continue;
        }
        let resolved = namespaces.resolve_attribute(attribute.key)?;
        if local_namespace(resolved.namespace.as_ref()) != namespace
            || resolved.local_bytes() != local
        {
            continue;
        }
        budget.text(attribute.value.len())?;
        budget.storage(size_of::<String>())?;
        budget.storage(attribute.value.len())?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid OTH structural identity value: {error}"))
            })?;
        budget.text_after_raw(attribute.value.len(), value.len())?;
        if value.len() > attribute.value.len() {
            budget.storage(value.len() - attribute.value.len())?;
        }
        return Ok((!value.is_empty()).then(|| value.into_owned()));
    }
    Ok(None)
}

fn source_offset(offset: u64) -> Result<usize> {
    usize::try_from(offset)
        .map_err(|error| Error::InvalidFormat(format!("OTH source offset overflow: {error}")))
}

fn xml_error(error: &quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("invalid OTH body structure XML: {error}"))
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn retained_budget_checks_one_under_and_one_over_without_partial_accounting() {
        let mut structures = BodyStructures::new();
        assert!(structures.ensure(MAX_STRUCTURE_TEXT - 1).is_ok());
        assert_eq!(structures.retained_bytes, 0);
        assert!(structures.account(MAX_STRUCTURE_TEXT - 1).is_ok());
        assert_eq!(structures.retained_bytes, MAX_STRUCTURE_TEXT - 1);
        assert!(structures.ensure(1).is_ok());
        assert!(structures.account(1).is_ok());
        assert_eq!(structures.retained_bytes, MAX_STRUCTURE_TEXT);
        assert!(structures.ensure(1).is_err());
        assert!(structures.account(1).is_err());
        assert_eq!(structures.retained_bytes, MAX_STRUCTURE_TEXT);
    }

    #[test]
    fn temporary_budget_refuses_the_first_byte_over_the_limit() {
        let mut budget = Budget {
            nodes: 0,
            text_bytes: 0,
            storage_bytes: 0,
        };
        assert!(budget.text(MAX_STRUCTURE_TEXT - 1).is_ok());
        assert!(budget.text(1).is_ok());
        assert!(budget.text(1).is_err());
        assert_eq!(budget.text_bytes, MAX_STRUCTURE_TEXT);
    }

    #[test]
    fn temporary_budget_charges_storage_slots_with_text_atomically() {
        let mut budget = Budget {
            nodes: 0,
            text_bytes: 0,
            storage_bytes: 0,
        };
        assert!(budget.storage(MAX_STRUCTURE_TEXT - 2).is_ok());
        assert!(budget.text(1).is_ok());
        assert!(budget.storage(1).is_ok());
        assert!(budget.text(1).is_err());
        assert_eq!(budget.storage_bytes, MAX_STRUCTURE_TEXT - 1);
        assert_eq!(budget.text_bytes, 1);
    }

    #[test]
    fn vector_growth_uses_bounded_geometric_reservations() {
        let mut values = Vec::<u8>::new();
        let mut growth_events = 0usize;
        for _ in 0..4096 {
            let required_len = values.len() + 1;
            reserve_vec_charged(&mut values, required_len, "OTH vector growth test", |_| {
                growth_events += 1;
                Ok(())
            })
            .unwrap();
            values.push(0);
        }
        assert!(
            growth_events <= 11,
            "unexpected reservation count: {growth_events}"
        );
        assert!(values.capacity() >= values.len());
    }

    #[test]
    fn vector_growth_charges_before_allocation() {
        let mut budget = Budget {
            nodes: 0,
            text_bytes: 0,
            storage_bytes: MAX_STRUCTURE_TEXT - 1,
        };
        let mut values = Vec::<u64>::new();
        assert!(
            budget
                .reserve_vec(&mut values, 1, "OTH vector growth charge test")
                .is_err()
        );
        assert_eq!(values.capacity(), 0);
        assert_eq!(budget.storage_bytes, MAX_STRUCTURE_TEXT - 1);
    }
}
