//! Typed rich-text, table, and inert form-control content for presentation pages.
//!
//! These values deliberately model authored presentation objects rather than
//! executable form behavior. Package mutation is owned by [`crate::edit`].

use litchi_core::{Error, Result, xml::escape_xml};
use litchi_odf_common::package::{is_linked_href, rebuild_package, resolve_package_path, splice};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::cmp::Reverse;

const MAX_NAME_BYTES: usize = 4 * 1024;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_CHILDREN: usize = 65_536;
const MAX_XML_DEPTH: usize = 512;
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const NUMBER_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NS: &[u8] = b"http://www.w3.org/1999/xlink";

#[derive(Clone)]
pub(crate) struct ResourceDependency {
    pub(crate) href: String,
    pub(crate) path: String,
    pub(crate) bytes: Vec<u8>,
    pub(crate) media_type: String,
}

#[derive(Clone)]
pub(crate) struct StyleDependency {
    name: String,
    xml: String,
}

#[derive(Clone)]
pub(crate) struct TransferObject {
    page: usize,
    kind: ObjectKind,
    source_name: String,
    destination_name: String,
    xml: String,
    styles: Vec<StyleDependency>,
    resources: Vec<ResourceDependency>,
}

#[derive(Clone)]
pub(crate) struct TransferControl {
    page: usize,
    source_name: String,
    destination_name: String,
    declaration_xml: String,
    visual_xml: String,
    styles: Vec<StyleDependency>,
    resources: Vec<ResourceDependency>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum ObjectKind {
    TextBox,
    Table,
}

#[derive(Clone)]
pub(crate) enum Operation {
    AddObject {
        page: usize,
        kind: ObjectKind,
        name: String,
        xml: String,
    },
    ReplaceObject {
        kind: ObjectKind,
        name: String,
        new_name: String,
        xml: String,
    },
    RemoveObject {
        kind: ObjectKind,
        name: String,
    },
    AddControl {
        page: usize,
        control: FormControl,
    },
    ReplaceControl {
        name: String,
        control: FormControl,
    },
    RemoveControl {
        name: String,
    },
    TransferObject(TransferObject),
    TransferControl(TransferControl),
    AddResources {
        resources: Vec<ResourceDependency>,
    },
}

#[derive(Clone, Copy)]
struct Span {
    start: usize,
    end: usize,
}

enum FormsSite {
    Content(usize),
    Empty {
        start: usize,
        end: usize,
        name_end: usize,
    },
    Missing(usize),
}

struct ResourceRemap {
    xml: String,
    resources: Vec<ResourceDependency>,
    values: Vec<(String, String)>,
}

/// One styled span in a presentation paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Run {
    text: String,
    style_name: Option<String>,
    href: Option<String>,
}

impl Run {
    /// Create a plain text run.
    ///
    /// # Errors
    ///
    /// Returns an error when the text is not valid bounded XML character data.
    pub fn new(text: impl Into<String>) -> Result<Self> {
        let text_value = text.into();
        validate_text(&text_value)?;
        Ok(Self {
            text: text_value,
            style_name: None,
            href: None,
        })
    }

    /// Assign a named ODF text style dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty, oversized, or XML-invalid style name.
    pub fn with_style(mut self, name: impl Into<String>) -> Result<Self> {
        self.style_name = Some(validate_name(name.into(), "rich-text style")?);
        Ok(self)
    }

    /// Assign an inert hyperlink. The editor never follows or executes it.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty or oversized reference.
    pub fn with_href(mut self, href: impl Into<String>) -> Result<Self> {
        let href_value = href.into();
        if href_value.is_empty() || href_value.len() > MAX_TEXT_BYTES {
            return invalid("ODP rich-text hyperlink is empty or oversized");
        }
        validate_text(&href_value)?;
        self.href = Some(href_value);
        Ok(self)
    }

    /// Borrow the run text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Borrow the optional named text style.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Borrow the optional inert hyperlink.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    fn byte_len(&self) -> usize {
        self.text.len()
    }
}

/// One rich-text paragraph.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Paragraph {
    runs: Vec<Run>,
    style_name: Option<String>,
}

impl Paragraph {
    /// Create a paragraph from checked runs.
    ///
    /// # Errors
    ///
    /// Returns an error when its aggregate text exceeds the bounded owner limit.
    pub fn new(runs: Vec<Run>) -> Result<Self> {
        validate_count(runs.len(), "rich-text run")?;
        validate_aggregate_text(runs.iter().map(Run::byte_len))?;
        Ok(Self {
            runs,
            style_name: None,
        })
    }

    /// Create a one-run plain paragraph.
    ///
    /// # Errors
    ///
    /// Returns an error when the text is invalid or oversized.
    pub fn plain(text: impl Into<String>) -> Result<Self> {
        Self::new(vec![Run::new(text)?])
    }

    /// Assign a named ODF paragraph style dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid style name.
    pub fn with_style(mut self, name: impl Into<String>) -> Result<Self> {
        self.style_name = Some(validate_name(name.into(), "paragraph style")?);
        Ok(self)
    }

    /// Borrow the paragraph runs.
    #[must_use]
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }

    /// Borrow the optional paragraph style name.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }
}

/// Common rich text used by presentation text boxes and table cells.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct RichText {
    paragraphs: Vec<Paragraph>,
}

impl RichText {
    /// Create checked rich text.
    ///
    /// # Errors
    ///
    /// Returns an error when the paragraph count or aggregate text is oversized.
    pub fn new(paragraphs: Vec<Paragraph>) -> Result<Self> {
        validate_count(paragraphs.len(), "rich-text paragraph")?;
        validate_aggregate_text(
            paragraphs
                .iter()
                .flat_map(Paragraph::runs)
                .map(Run::byte_len),
        )?;
        Ok(Self { paragraphs })
    }

    /// Create one plain paragraph.
    ///
    /// # Errors
    ///
    /// Returns an error when the text is invalid or oversized.
    pub fn plain(text: impl Into<String>) -> Result<Self> {
        Self::new(vec![Paragraph::plain(text)?])
    }

    /// Borrow paragraphs in document order.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    pub(crate) fn xml(&self) -> Result<String> {
        let mut output = String::new();
        for paragraph in &self.paragraphs {
            output.push_str("<text:p");
            optional_attribute(&mut output, "text:style-name", paragraph.style_name())?;
            output.push('>');
            for run in paragraph.runs() {
                if let Some(href) = run.href() {
                    output.push_str("<text:a xlink:type=\"simple\" xlink:href=\"");
                    output.push_str(&escape_xml(href));
                    output.push_str("\">");
                }
                if let Some(style) = run.style_name() {
                    output.push_str("<text:span text:style-name=\"");
                    output.push_str(&escape_xml(style));
                    output.push_str("\">");
                }
                output.push_str(&escape_xml(run.text()));
                if run.style_name().is_some() {
                    output.push_str("</text:span>");
                }
                if run.href().is_some() {
                    output.push_str("</text:a>");
                }
            }
            output.push_str("</text:p>");
        }
        Ok(output)
    }
}

/// A named rich-text box placed on a presentation page.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextBox {
    name: String,
    text: RichText,
    style_name: Option<String>,
}

impl TextBox {
    /// Create a named rich-text box.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid object name.
    pub fn new(name: impl Into<String>, text: RichText) -> Result<Self> {
        Ok(Self {
            name: validate_name(name.into(), "rich-text box")?,
            text,
            style_name: None,
        })
    }

    /// Assign a named drawing style dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid style name.
    pub fn with_style(mut self, name: impl Into<String>) -> Result<Self> {
        self.style_name = Some(validate_name(name.into(), "text-box style")?);
        Ok(self)
    }

    /// Borrow the stable object name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the rich text.
    #[must_use]
    pub fn text(&self) -> &RichText {
        &self.text
    }

    pub(crate) fn xml(&self) -> Result<String> {
        let mut output = format!(
            "<draw:frame xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" draw:name=\"{}\"",
            escape_xml(&self.name)
        );
        optional_attribute(&mut output, "draw:style-name", self.style_name.as_deref())?;
        output.push_str("><draw:text-box>");
        output.push_str(&self.text.xml()?);
        output.push_str("</draw:text-box></draw:frame>");
        Ok(output)
    }
}

/// One typed presentation table cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Cell {
    text: RichText,
    style_name: Option<String>,
}

impl Cell {
    /// Create a table cell from common rich text.
    #[must_use]
    pub fn new(text: RichText) -> Self {
        Self {
            text,
            style_name: None,
        }
    }

    /// Assign a named table-cell style dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid style name.
    pub fn with_style(mut self, name: impl Into<String>) -> Result<Self> {
        self.style_name = Some(validate_name(name.into(), "table-cell style")?);
        Ok(self)
    }

    /// Borrow the cell rich text.
    #[must_use]
    pub fn text(&self) -> &RichText {
        &self.text
    }
}

/// A named rectangular presentation table.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Table {
    name: String,
    rows: Vec<Vec<Cell>>,
    style_name: Option<String>,
}

impl Table {
    /// Create a checked rectangular table.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid name, empty/ragged table, or resource limit.
    pub fn new(name: impl Into<String>, rows: Vec<Vec<Cell>>) -> Result<Self> {
        let table_name = validate_name(name.into(), "table")?;
        validate_count(rows.len(), "table row")?;
        let columns = rows.first().map_or(0, Vec::len);
        if columns == 0 || rows.iter().any(|row| row.len() != columns) {
            return invalid("ODP authored table must be non-empty and rectangular");
        }
        validate_count(columns, "table column")?;
        Ok(Self {
            name: table_name,
            rows,
            style_name: None,
        })
    }

    /// Assign a named table style dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid style name.
    pub fn with_style(mut self, name: impl Into<String>) -> Result<Self> {
        self.style_name = Some(validate_name(name.into(), "table style")?);
        Ok(self)
    }

    /// Borrow the stable object name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the rectangular rows.
    #[must_use]
    pub fn rows(&self) -> &[Vec<Cell>] {
        &self.rows
    }

    pub(crate) fn xml(&self) -> Result<String> {
        let mut output = format!(
            "<draw:frame xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" draw:name=\"{}\"><table:table table:name=\"{}\"",
            escape_xml(&self.name),
            escape_xml(&self.name)
        );
        optional_attribute(&mut output, "table:style-name", self.style_name.as_deref())?;
        output.push('>');
        for _ in 0..self.rows[0].len() {
            output.push_str("<table:table-column/>");
        }
        for row in &self.rows {
            output.push_str("<table:table-row>");
            for cell in row {
                output.push_str("<table:table-cell office:value-type=\"string\"");
                optional_attribute(&mut output, "table:style-name", cell.style_name.as_deref())?;
                output.push('>');
                output.push_str(&cell.text.xml()?);
                output.push_str("</table:table-cell>");
            }
            output.push_str("</table:table-row>");
        }
        output.push_str("</table:table></draw:frame>");
        Ok(output)
    }
}

/// Inert form-control kind. No control behavior is executed by this crate.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ControlKind {
    /// Push button.
    Button,
    /// Single-line text input.
    Text,
    /// Boolean checkbox.
    Checkbox,
}

impl ControlKind {
    fn element(self) -> &'static str {
        match self {
            Self::Button => "form:button",
            Self::Text => "form:text",
            Self::Checkbox => "form:checkbox",
        }
    }
}

/// One inert visual form control and its form declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormControl {
    name: String,
    kind: ControlKind,
    label: Option<String>,
    value: Option<String>,
}

impl FormControl {
    /// Create a checked inert form control.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid control name.
    pub fn new(name: impl Into<String>, kind: ControlKind) -> Result<Self> {
        Ok(Self {
            name: validate_name(name.into(), "form control")?,
            kind,
            label: None,
            value: None,
        })
    }

    /// Set the producer-visible label.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid or oversized XML text.
    pub fn with_label(mut self, label: impl Into<String>) -> Result<Self> {
        let label_value = label.into();
        validate_text(&label_value)?;
        self.label = Some(label_value);
        Ok(self)
    }

    /// Set the inert current value.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid or oversized XML text.
    pub fn with_value(mut self, value: impl Into<String>) -> Result<Self> {
        let current_value = value.into();
        validate_text(&current_value)?;
        self.value = Some(current_value);
        Ok(self)
    }

    /// Borrow the stable control name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    pub(crate) fn declaration_xml(&self) -> Result<String> {
        let mut output = format!(
            "<{} form:id=\"{}\" form:name=\"{}\"",
            self.kind.element(),
            escape_xml(&self.name),
            escape_xml(&self.name)
        );
        optional_attribute(&mut output, "form:label", self.label.as_deref())?;
        optional_attribute(&mut output, "form:current-value", self.value.as_deref())?;
        output.push_str("/>");
        Ok(output)
    }

    pub(crate) fn visual_xml(&self) -> String {
        format!(
            "<draw:control xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" draw:name=\"{}\" draw:control=\"{}\"/>",
            escape_xml(&self.name),
            escape_xml(&self.name)
        )
    }
}

fn optional_attribute(output: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(attribute_value) = value {
        validate_text(attribute_value)?;
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(attribute_value));
        output.push('"');
    }
    Ok(())
}

fn validate_name(value: String, kind: &str) -> Result<String> {
    if value.is_empty() || value.len() > MAX_NAME_BYTES {
        return invalid(format!("ODP {kind} name is empty or oversized"));
    }
    validate_text(&value)?;
    Ok(value)
}

fn validate_count(count: usize, kind: &str) -> Result<()> {
    if count > MAX_CHILDREN {
        return invalid(format!("ODP {kind} count exceeds {MAX_CHILDREN}"));
    }
    Ok(())
}

fn validate_aggregate_text(mut lengths: impl Iterator<Item = usize>) -> Result<()> {
    let aggregate = lengths.try_fold(0usize, |accumulated, length| {
        accumulated.checked_add(length)
    });
    if aggregate.is_none_or(|candidate| candidate > MAX_TEXT_BYTES) {
        return invalid("ODP rich text exceeds the 16 MiB limit");
    }
    Ok(())
}

fn validate_text(value: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return invalid("ODP authored text exceeds the 16 MiB limit");
    }
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}'
                | '\u{A}'
                | '\u{D}'
                | '\u{20}'..='\u{D7FF}'
                | '\u{E000}'..='\u{FFFD}'
                | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return invalid("ODP authored text contains a character forbidden by XML 1.0");
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

pub(crate) fn prepare_object_transfer(
    source: &crate::core::OwnedPackage,
    page: usize,
    kind: ObjectKind,
    source_name: &str,
    destination_name: String,
) -> Result<Operation> {
    validate_name(destination_name.clone(), "transferred content object")?;
    let content = package_xml(source, "content.xml")?;
    let span = locate_object(&content, kind, source_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "ODP source content object '{source_name}' was not found"
        ))
    })?;
    let source_xml = content
        .get(span.start..span.end)
        .ok_or_else(|| Error::InvalidFormat("invalid ODP source object span".into()))?
        .to_string();
    let xml = make_fragment_self_contained(&source_xml, &content)?;
    let styles = collect_style_dependencies(source, &content, &xml)?;
    let mut dependency_xml = xml.clone();
    for style in &styles {
        dependency_xml.push_str(&style.xml);
    }
    let resources = collect_resource_dependencies(source, &dependency_xml, "content.xml")?;
    Ok(Operation::TransferObject(TransferObject {
        page,
        kind,
        source_name: source_name.to_string(),
        destination_name,
        xml,
        styles,
        resources,
    }))
}

pub(crate) fn prepare_control_transfer(
    source: &crate::core::OwnedPackage,
    page: usize,
    source_name: &str,
    destination_name: String,
) -> Result<Operation> {
    validate_name(destination_name.clone(), "transferred form control")?;
    let content = package_xml(source, "content.xml")?;
    let visual = locate_named_drawing(&content, source_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "ODP source form control '{source_name}' was not found"
        ))
    })?;
    let declaration = locate_form_declaration(&content, source_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "ODP source form control declaration '{source_name}' was not found"
        ))
    })?;
    let source_visual_xml = content
        .get(visual.start..visual.end)
        .ok_or_else(|| Error::InvalidFormat("invalid ODP source form-control span".into()))?
        .to_string();
    let source_declaration_xml = content
        .get(declaration.start..declaration.end)
        .ok_or_else(|| Error::InvalidFormat("invalid ODP source form declaration span".into()))?
        .to_string();
    let visual_xml = make_fragment_self_contained(&source_visual_xml, &content)?;
    let declaration_xml = make_fragment_self_contained(&source_declaration_xml, &content)?;
    let combined = format!("{declaration_xml}{visual_xml}");
    let styles = collect_style_dependencies(source, &content, &combined)?;
    let mut dependency_xml = combined;
    for style in &styles {
        dependency_xml.push_str(&style.xml);
    }
    let resources = collect_resource_dependencies(source, &dependency_xml, "content.xml")?;
    Ok(Operation::TransferControl(TransferControl {
        page,
        source_name: source_name.to_string(),
        destination_name,
        declaration_xml,
        visual_xml,
        styles,
        resources,
    }))
}

pub(crate) fn prepare_resource_transfer(
    source: &crate::core::OwnedPackage,
    destination: &crate::core::OwnedPackage,
    xml: &str,
    source_base: &str,
    destination_base: &str,
) -> Result<(String, Operation)> {
    let dependencies = collect_resource_dependencies(source, xml, source_base)?;
    let remap = remap_resources(destination, xml, &dependencies, destination_base)?;
    Ok((
        remap.xml,
        Operation::AddResources {
            resources: remap.resources,
        },
    ))
}

pub(crate) fn resource_operation_is_empty(operation: &Operation) -> bool {
    matches!(operation, Operation::AddResources { resources } if resources.is_empty())
}

fn package_xml(source: &crate::core::OwnedPackage, path: &str) -> Result<String> {
    String::from_utf8(source.get_file(path)?).map_err(|cause| {
        Error::InvalidFormat(format!("ODP XML part '{path}' is not UTF-8: {cause}"))
    })
}

fn collect_style_dependencies(
    source: &crate::core::OwnedPackage,
    content: &str,
    object_xml: &str,
) -> Result<Vec<StyleDependency>> {
    let styles_xml = source
        .has_file("styles.xml")?
        .then(|| package_xml(source, "styles.xml"))
        .transpose()?;
    let mut pending = collect_style_attribute_values(object_xml)?;
    let mut visited = std::collections::BTreeSet::new();
    let mut dependencies = Vec::new();
    while let Some(name) = pending.pop() {
        if !visited.insert(name.clone()) {
            continue;
        }
        let mut located_fragment = locate_style_fragment(content, &name)?;
        let mut namespace_owner = content;
        if located_fragment.is_none()
            && let Some(package_styles) = styles_xml.as_deref()
        {
            located_fragment = locate_style_fragment(package_styles, &name)?;
            namespace_owner = package_styles;
        }
        let source_fragment = located_fragment.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "ODP transferred object has dangling style dependency '{name}'"
            ))
        })?;
        for referenced in collect_style_attribute_values(&source_fragment)? {
            if referenced == name {
                continue;
            }
            if !visited.contains(&referenced) {
                pending.push(referenced);
            }
        }
        dependencies.push(StyleDependency {
            name,
            xml: make_fragment_self_contained(&source_fragment, namespace_owner)?,
        });
    }
    Ok(dependencies)
}

fn locate_style_fragment(xml: &str, name: &str) -> Result<Option<String>> {
    let mut span = locate_named_element_local(xml, STYLE_NS, b"style", b"name", name)?;
    if span.is_none() {
        span = locate_named_element_with_attribute(xml, NUMBER_NS, STYLE_NS, b"name", name)?;
    }
    if span.is_none() {
        span = locate_named_element_local_with_attribute(
            xml,
            TEXT_NS,
            b"list-style",
            STYLE_NS,
            b"name",
            name,
        )?;
    }
    let Some(found_span) = span else {
        return Ok(None);
    };
    xml.get(found_span.start..found_span.end)
        .map(str::to_string)
        .ok_or_else(|| Error::InvalidFormat("invalid ODP style dependency span".into()))
        .map(Some)
}

fn make_fragment_self_contained(fragment: &str, owner_xml: &str) -> Result<String> {
    let Some(insertion) = fragment.find(|character: char| {
        character.is_ascii_whitespace() || character == '>' || character == '/'
    }) else {
        return invalid("ODP transferred fragment has no complete root element");
    };
    let root_end = fragment.find('>').unwrap_or(fragment.len());
    let root_start = fragment
        .get(..root_end)
        .ok_or_else(|| Error::InvalidFormat("invalid ODP transferred fragment root span".into()))?;
    let mut declarations = String::new();
    for (qualified_name, uri) in namespace_declarations(owner_xml)? {
        if !root_start.contains(&format!("{qualified_name}=")) {
            declarations.push(' ');
            declarations.push_str(&qualified_name);
            declarations.push_str("=\"");
            declarations.push_str(&escape_xml(&uri));
            declarations.push('"');
        }
    }
    let mut output = fragment.to_string();
    output.insert_str(insertion, &declarations);
    Ok(output)
}

fn namespace_declarations(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(|cause| {
            Error::InvalidFormat(format!("invalid ODP namespace-owner XML: {cause}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let mut declarations = Vec::new();
                for raw_attribute in element.attributes() {
                    let attribute = raw_attribute.map_err(|cause| {
                        Error::InvalidFormat(format!("invalid ODP namespace declaration: {cause}"))
                    })?;
                    let qualified_name = attribute.key.as_ref();
                    if qualified_name == b"xmlns" || qualified_name.starts_with(b"xmlns:") {
                        let uri = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|cause| {
                                Error::InvalidFormat(format!(
                                    "invalid ODP namespace declaration value: {cause}"
                                ))
                            })?;
                        declarations.push((
                            String::from_utf8(qualified_name.to_vec()).map_err(|cause| {
                                Error::InvalidFormat(format!(
                                    "ODP namespace name is not UTF-8: {cause}"
                                ))
                            })?,
                            uri.into_owned(),
                        ));
                    }
                }
                return Ok(declarations);
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in ODP namespace-owner XML"),
            Event::Eof => return invalid("ODP namespace-owner XML has no root element"),
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

fn collect_resource_dependencies(
    source: &crate::core::OwnedPackage,
    xml: &str,
    base_path: &str,
) -> Result<Vec<ResourceDependency>> {
    let hrefs = collect_attribute_values(xml, XLINK_NS, b"href")?;
    let package = source.package()?;
    let mut resources = Vec::new();
    for href in hrefs {
        if is_external_href(&href) {
            continue;
        }
        let path = resolve_package_href(base_path, &href)?;
        if !package.has_file(&path) {
            return invalid(format!(
                "ODP transferred object has dangling package resource '{href}'"
            ));
        }
        let media_type = package
            .manifest()
            .get_entry(&path)
            .map_or_else(String::new, |entry| entry.media_type.clone());
        resources.push(ResourceDependency {
            href,
            bytes: package.get_file(&path)?,
            path,
            media_type,
        });
    }
    Ok(resources)
}

fn is_external_href(href: &str) -> bool {
    href.is_empty() || is_linked_href(href)
}

fn resolve_package_href(base_path: &str, href: &str) -> Result<String> {
    let combined = base_path.rsplit_once('/').map_or_else(
        || href.to_string(),
        |(parent, _)| format!("{parent}/{href}"),
    );
    resolve_package_path(&combined)
}

pub(crate) fn apply(source: &crate::core::OwnedPackage, operation: &Operation) -> Result<Vec<u8>> {
    let content = String::from_utf8(source.get_file("content.xml")?)
        .map_err(|cause| Error::InvalidFormat(format!("ODP content.xml is not UTF-8: {cause}")))?;
    let (updated, resources) = match operation {
        Operation::AddObject {
            page,
            kind,
            name,
            xml,
        } => (add_object(&content, *page, *kind, name, xml)?, Vec::new()),
        Operation::ReplaceObject {
            kind,
            name,
            new_name,
            xml,
        } => (
            replace_object(&content, *kind, name, new_name, xml)?,
            Vec::new(),
        ),
        Operation::RemoveObject { kind, name } => {
            (remove_object(&content, *kind, name)?, Vec::new())
        },
        Operation::AddControl { page, control } => {
            (add_control(&content, *page, control)?, Vec::new())
        },
        Operation::ReplaceControl { name, control } => {
            (replace_control(&content, name, control)?, Vec::new())
        },
        Operation::RemoveControl { name } => (remove_control(&content, name)?, Vec::new()),
        Operation::TransferObject(transfer) => apply_object_transfer(source, &content, transfer)?,
        Operation::TransferControl(transfer) => apply_control_transfer(source, &content, transfer)?,
        Operation::AddResources { resources } => (
            content,
            validate_resource_additions(source, resources.clone())?,
        ),
    };
    let additions = resources
        .into_iter()
        .map(|resource| litchi_odf_common::package::Addition {
            path: resource.path,
            bytes: resource.bytes,
            media_type: resource.media_type,
        })
        .collect();
    rebuild_package(
        source,
        &updated,
        additions,
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}

fn apply_object_transfer(
    destination: &crate::core::OwnedPackage,
    content: &str,
    transfer: &TransferObject,
) -> Result<(String, Vec<ResourceDependency>)> {
    if locate_named_drawing(content, &transfer.destination_name)?.is_some() {
        return invalid(format!(
            "ODP destination content object name '{}' already exists",
            transfer.destination_name
        ));
    }
    let style_names = plan_style_names(destination, content, &transfer.styles)?;
    let mut object_xml = rewrite_known_attributes(
        &transfer.xml,
        &[(
            transfer.source_name.clone(),
            transfer.destination_name.clone(),
        )],
        &["draw:name", "table:name"],
    );
    object_xml = rewrite_known_attributes(
        &object_xml,
        &style_names,
        &[
            "draw:style-name",
            "draw:text-style-name",
            "text:style-name",
            "table:style-name",
            "table:default-cell-style-name",
            "presentation:style-name",
        ],
    );
    let resource_remap =
        remap_resources(destination, &object_xml, &transfer.resources, "content.xml")?;
    object_xml = resource_remap.xml;
    let style_fragments =
        rewrite_style_fragments(&transfer.styles, &style_names, &resource_remap.values);
    let styled_content = inject_automatic_styles(content, &style_fragments)?;
    let updated = add_object(
        &styled_content,
        transfer.page,
        transfer.kind,
        &transfer.destination_name,
        &object_xml,
    )?;
    Ok((updated, resource_remap.resources))
}

fn apply_control_transfer(
    destination: &crate::core::OwnedPackage,
    content: &str,
    transfer: &TransferControl,
) -> Result<(String, Vec<ResourceDependency>)> {
    if locate_named_drawing(content, &transfer.destination_name)?.is_some()
        || locate_form_declaration(content, &transfer.destination_name)?.is_some()
    {
        return invalid(format!(
            "ODP destination form control name '{}' already exists",
            transfer.destination_name
        ));
    }
    let names = [(
        transfer.source_name.clone(),
        transfer.destination_name.clone(),
    )];
    let style_names = plan_style_names(destination, content, &transfer.styles)?;
    let visual_xml = rewrite_known_attributes(
        &rewrite_known_attributes(&transfer.visual_xml, &names, &["draw:name", "draw:control"]),
        &style_names,
        &[
            "draw:style-name",
            "draw:text-style-name",
            "text:style-name",
            "table:style-name",
            "table:default-cell-style-name",
            "presentation:style-name",
        ],
    );
    let declaration_xml =
        rewrite_known_attributes(&transfer.declaration_xml, &names, &["form:id", "form:name"]);
    let combined = format!("{declaration_xml}{visual_xml}");
    let remap = remap_resources(destination, &combined, &transfer.resources, "content.xml")?;
    let remapped_declaration =
        rewrite_known_attributes(&declaration_xml, &remap.values, &["xlink:href"]);
    let remapped_visual = rewrite_known_attributes(&visual_xml, &remap.values, &["xlink:href"]);
    let style_fragments = rewrite_style_fragments(&transfer.styles, &style_names, &remap.values);
    let styled_content = inject_automatic_styles(content, &style_fragments)?;
    let updated = add_control_fragments(
        &styled_content,
        transfer.page,
        &transfer.destination_name,
        &remapped_declaration,
        &remapped_visual,
    )?;
    Ok((updated, remap.resources))
}

fn plan_style_names(
    destination: &crate::core::OwnedPackage,
    content: &str,
    dependencies: &[StyleDependency],
) -> Result<Vec<(String, String)>> {
    let styles_xml = destination
        .has_file("styles.xml")?
        .then(|| package_xml(destination, "styles.xml"))
        .transpose()?;
    let mut reserved_names = std::collections::BTreeSet::new();
    let mut names = Vec::new();
    for dependency in dependencies {
        let exists_in_styles = styles_xml
            .as_deref()
            .map(|package_styles| locate_style_fragment(package_styles, &dependency.name))
            .transpose()?
            .flatten()
            .is_some();
        let destination_name = if locate_style_fragment(content, &dependency.name)?.is_some()
            || exists_in_styles
            || reserved_names.contains(&dependency.name)
        {
            fresh_style_name(
                content,
                styles_xml.as_deref(),
                &dependency.name,
                &reserved_names,
            )?
        } else {
            dependency.name.clone()
        };
        reserved_names.insert(destination_name.clone());
        names.push((dependency.name.clone(), destination_name));
    }
    Ok(names)
}

fn rewrite_style_fragments(
    dependencies: &[StyleDependency],
    style_names: &[(String, String)],
    resource_names: &[(String, String)],
) -> String {
    let mut fragments = String::new();
    for dependency in dependencies {
        let mut fragment = rewrite_known_attributes(
            &dependency.xml,
            style_names,
            &[
                "style:name",
                "style:parent-style-name",
                "style:next-style-name",
                "style:data-style-name",
                "style:list-style-name",
                "style:percentage-data-style-name",
                "draw:style-name",
                "draw:text-style-name",
                "text:style-name",
                "table:style-name",
                "table:default-cell-style-name",
                "presentation:style-name",
            ],
        );
        fragment = rewrite_known_attributes(&fragment, resource_names, &["xlink:href"]);
        fragments.push_str(&fragment);
    }
    fragments
}

fn add_object(
    content: &str,
    page: usize,
    kind: ObjectKind,
    name: &str,
    xml: &str,
) -> Result<String> {
    if locate_object(content, kind, name)?.is_some()
        || locate_named_drawing(content, name)?.is_some()
    {
        return invalid(format!("ODP content object name '{name}' already exists"));
    }
    let pages = crate::charts::locate_pages(content)?;
    let selected_page = pages
        .get(page)
        .ok_or_else(|| Error::InvalidFormat("ODP content page selector is out of bounds".into()))?;
    splice(content, selected_page.end, selected_page.end, xml)
}

fn replace_object(
    content: &str,
    kind: ObjectKind,
    name: &str,
    new_name: &str,
    xml: &str,
) -> Result<String> {
    let span = locate_object(content, kind, name)?.ok_or_else(|| {
        Error::InvalidFormat(format!("ODP content object '{name}' was not found"))
    })?;
    if name != new_name && locate_named_drawing(content, new_name)?.is_some() {
        return invalid(format!(
            "ODP replacement content object name '{new_name}' already exists"
        ));
    }
    splice(content, span.start, span.end, xml)
}

fn remove_object(content: &str, kind: ObjectKind, name: &str) -> Result<String> {
    let span = locate_object(content, kind, name)?.ok_or_else(|| {
        Error::InvalidFormat(format!("ODP content object '{name}' was not found"))
    })?;
    splice(content, span.start, span.end, "")
}

fn add_control(content: &str, page: usize, control: &FormControl) -> Result<String> {
    if locate_named_drawing(content, control.name())?.is_some()
        || locate_form_declaration(content, control.name())?.is_some()
    {
        return invalid(format!(
            "ODP form control name '{}' already exists",
            control.name()
        ));
    }
    let pages = crate::charts::locate_pages(content)?;
    let page_site = pages
        .get(page)
        .ok_or_else(|| Error::InvalidFormat("ODP form page selector is out of bounds".into()))?
        .end;
    let form_xml = format!(
        "<form:form xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" form:name=\"{}-form\">{}</form:form>",
        escape_xml(control.name()),
        control.declaration_xml()?
    );
    let mut edits = vec![(page_site, page_site, control.visual_xml())];
    match locate_forms_site(content)? {
        FormsSite::Content(end) => edits.push((end, end, form_xml)),
        FormsSite::Empty {
            start,
            end,
            name_end,
        } => edits.push((
            start,
            end,
            format!("{}>{}</office:forms>", &content[start..name_end], form_xml),
        )),
        FormsSite::Missing(position) => edits.push((
            position,
            position,
            format!("<office:forms>{form_xml}</office:forms>"),
        )),
    }
    apply_edits(content, edits)
}

fn add_control_fragments(
    content: &str,
    page: usize,
    name: &str,
    declaration_xml: &str,
    visual_xml: &str,
) -> Result<String> {
    let pages = crate::charts::locate_pages(content)?;
    let page_site = pages
        .get(page)
        .ok_or_else(|| Error::InvalidFormat("ODP form page selector is out of bounds".into()))?
        .end;
    let form_xml = format!(
        "<form:form xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" form:name=\"{}-form\">{declaration_xml}</form:form>",
        escape_xml(name)
    );
    let mut edits = vec![(page_site, page_site, visual_xml.to_string())];
    match locate_forms_site(content)? {
        FormsSite::Content(end) => edits.push((end, end, form_xml)),
        FormsSite::Empty {
            start,
            end,
            name_end,
        } => edits.push((
            start,
            end,
            format!("{}>{}</office:forms>", &content[start..name_end], form_xml),
        )),
        FormsSite::Missing(position) => edits.push((
            position,
            position,
            format!("<office:forms>{form_xml}</office:forms>"),
        )),
    }
    apply_edits(content, edits)
}

fn fresh_style_name(
    content: &str,
    styles_xml: Option<&str>,
    requested: &str,
    reserved: &std::collections::BTreeSet<String>,
) -> Result<String> {
    for index in 1..=100_000usize {
        let candidate = format!("{requested}_litchi_{index}");
        let exists_in_styles = styles_xml
            .map(|package_styles| locate_style_fragment(package_styles, &candidate))
            .transpose()?
            .flatten()
            .is_some();
        if !reserved.contains(&candidate)
            && locate_style_fragment(content, &candidate)?.is_none()
            && !exists_in_styles
        {
            return Ok(candidate);
        }
    }
    invalid("ODP style collision remapping exhausted its bounded namespace")
}

fn remap_resources(
    destination: &crate::core::OwnedPackage,
    xml: &str,
    dependencies: &[ResourceDependency],
    destination_base: &str,
) -> Result<ResourceRemap> {
    let package = destination.package()?;
    let mut reserved = destination
        .files()?
        .into_iter()
        .collect::<std::collections::BTreeSet<_>>();
    let mut assigned_paths = std::collections::BTreeMap::<String, String>::new();
    let mut values = Vec::new();
    let mut resources = Vec::new();
    for dependency in dependencies {
        if let Some(destination_path) = assigned_paths.get(&dependency.path) {
            values.push((
                dependency.href.clone(),
                relative_package_href(destination_base, destination_path),
            ));
            continue;
        }
        let mut destination_path = dependency.path.clone();
        if package.has_file(&destination_path) {
            if package.get_file(&destination_path)? == dependency.bytes {
                values.push((
                    dependency.href.clone(),
                    relative_package_href(destination_base, &destination_path),
                ));
                assigned_paths.insert(dependency.path.clone(), destination_path);
                continue;
            }
            destination_path = fresh_resource_path(&dependency.path, &reserved)?;
        }
        reserved.insert(destination_path.clone());
        assigned_paths.insert(dependency.path.clone(), destination_path.clone());
        values.push((
            dependency.href.clone(),
            relative_package_href(destination_base, &destination_path),
        ));
        resources.push(ResourceDependency {
            href: dependency.href.clone(),
            path: destination_path,
            bytes: dependency.bytes.clone(),
            media_type: dependency.media_type.clone(),
        });
    }
    Ok(ResourceRemap {
        xml: rewrite_known_attributes(xml, &values, &["xlink:href"]),
        resources,
        values,
    })
}

fn validate_resource_additions(
    destination: &crate::core::OwnedPackage,
    resources: Vec<ResourceDependency>,
) -> Result<Vec<ResourceDependency>> {
    let package = destination.package()?;
    let mut additions = Vec::new();
    for resource in resources {
        if package.has_file(&resource.path) {
            if package.get_file(&resource.path)? != resource.bytes {
                return invalid(format!(
                    "ODP transferred resource path '{}' collided during replay",
                    resource.path
                ));
            }
            continue;
        }
        additions.push(resource);
    }
    Ok(additions)
}

fn fresh_resource_path(
    requested: &str,
    reserved: &std::collections::BTreeSet<String>,
) -> Result<String> {
    let (parent, file_name) = requested.rsplit_once('/').unwrap_or(("", requested));
    let (stem, extension) = file_name.rsplit_once('.').unwrap_or((file_name, ""));
    for index in 1..=100_000usize {
        let remapped_file = if extension.is_empty() {
            format!("{stem}_litchi_{index}")
        } else {
            format!("{stem}_litchi_{index}.{extension}")
        };
        let candidate = if parent.is_empty() {
            remapped_file
        } else {
            format!("{parent}/{remapped_file}")
        };
        if !reserved.contains(&candidate) {
            return Ok(candidate);
        }
    }
    invalid("ODP resource collision remapping exhausted its bounded namespace")
}

fn relative_package_href(base_path: &str, target_path: &str) -> String {
    let base_depth = base_path
        .rsplit_once('/')
        .map_or(0, |(parent, _)| parent.split('/').count());
    let mut output = "../".repeat(base_depth);
    output.push_str(target_path);
    output
}

fn inject_automatic_styles(content: &str, fragments: &str) -> Result<String> {
    if fragments.is_empty() {
        return Ok(content.to_string());
    }
    if let Some(position) = content.rfind("</office:automatic-styles>") {
        return splice(content, position, position, fragments);
    }
    if let Some(position) = content.find("<office:automatic-styles/>") {
        let end = position + "<office:automatic-styles/>".len();
        return splice(
            content,
            position,
            end,
            &format!("<office:automatic-styles>{fragments}</office:automatic-styles>"),
        );
    }
    Err(Error::Unsupported(
        "ODP content has no transferable automatic-style owner".to_string(),
    ))
}

fn rewrite_known_attributes(
    xml: &str,
    replacements: &[(String, String)],
    qualified_names: &[&str],
) -> String {
    let mut output = xml.to_string();
    for (source, destination) in replacements {
        let escaped_source = escape_xml(source);
        let escaped_destination = escape_xml(destination);
        for qualified_name in qualified_names {
            for quote in ['"', '\''] {
                let needle = format!("{qualified_name}={quote}{escaped_source}{quote}");
                let replacement = format!("{qualified_name}={quote}{escaped_destination}{quote}");
                output = output.replace(&needle, &replacement);
            }
        }
    }
    output
}

fn replace_control(content: &str, name: &str, control: &FormControl) -> Result<String> {
    if name != control.name()
        && (locate_named_drawing(content, control.name())?.is_some()
            || locate_form_declaration(content, control.name())?.is_some())
    {
        return invalid(format!(
            "ODP replacement form control name '{}' already exists",
            control.name()
        ));
    }
    let visual = locate_named_drawing(content, name)?
        .ok_or_else(|| Error::InvalidFormat(format!("ODP form control '{name}' was not found")))?;
    let declaration = locate_form_declaration(content, name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "ODP form control declaration '{name}' was not found"
        ))
    })?;
    apply_edits(
        content,
        vec![
            (visual.start, visual.end, control.visual_xml()),
            (
                declaration.start,
                declaration.end,
                control.declaration_xml()?,
            ),
        ],
    )
}

fn remove_control(content: &str, name: &str) -> Result<String> {
    let visual = locate_named_drawing(content, name)?
        .ok_or_else(|| Error::InvalidFormat(format!("ODP form control '{name}' was not found")))?;
    let declaration = locate_form_declaration(content, name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "ODP form control declaration '{name}' was not found"
        ))
    })?;
    apply_edits(
        content,
        vec![
            (visual.start, visual.end, String::new()),
            (declaration.start, declaration.end, String::new()),
        ],
    )
}

fn apply_edits(content: &str, mut edits: Vec<(usize, usize, String)>) -> Result<String> {
    edits.sort_unstable_by_key(|edit| Reverse(edit.0));
    let mut output = content.to_string();
    for (start, end, replacement) in edits {
        output = splice(&output, start, end, &replacement)?;
    }
    Ok(output)
}

fn locate_object(content: &str, kind: ObjectKind, name: &str) -> Result<Option<Span>> {
    let Some(frame) = locate_named_drawing(content, name)? else {
        return Ok(None);
    };
    let fragment = content
        .get(frame.start..frame.end)
        .ok_or_else(|| Error::InvalidFormat("invalid ODP content object span".into()))?;
    let (namespace, local) = match kind {
        ObjectKind::TextBox => (DRAW_NS, b"text-box".as_slice()),
        ObjectKind::Table => (
            b"urn:oasis:names:tc:opendocument:xmlns:table:1.0".as_slice(),
            b"table".as_slice(),
        ),
    };
    Ok(contains_element(fragment, namespace, local)?.then_some(frame))
}

fn contains_element(content: &str, namespace: &[u8], local: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODP object XML: {error}")))?;
        let matches = resolved_namespace(&resolved) == Some(namespace);
        match event {
            Event::Start(element) | Event::Empty(element)
                if matches && element.local_name().as_ref() == local =>
            {
                return Ok(true);
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in ODP object XML"),
            Event::Eof => return Ok(false),
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

fn locate_named_drawing(content: &str, name: &str) -> Result<Option<Span>> {
    locate_named_element(content, DRAW_NS, b"name", name)
}

fn locate_form_declaration(content: &str, name: &str) -> Result<Option<Span>> {
    locate_named_element(content, FORM_NS, b"id", name)
}

fn locate_named_element(
    content: &str,
    namespace: &[u8],
    attribute: &[u8],
    expected: &str,
) -> Result<Option<Span>> {
    locate_named_element_impl(content, namespace, None, namespace, attribute, expected)
}

fn locate_named_element_local(
    content: &str,
    namespace: &[u8],
    element_local: &[u8],
    attribute: &[u8],
    expected: &str,
) -> Result<Option<Span>> {
    locate_named_element_impl(
        content,
        namespace,
        Some(element_local),
        namespace,
        attribute,
        expected,
    )
}

fn locate_named_element_with_attribute(
    content: &str,
    element_namespace: &[u8],
    attribute_namespace: &[u8],
    attribute: &[u8],
    expected: &str,
) -> Result<Option<Span>> {
    locate_named_element_impl(
        content,
        element_namespace,
        None,
        attribute_namespace,
        attribute,
        expected,
    )
}

fn locate_named_element_local_with_attribute(
    content: &str,
    element_namespace: &[u8],
    element_local: &[u8],
    attribute_namespace: &[u8],
    attribute: &[u8],
    expected: &str,
) -> Result<Option<Span>> {
    locate_named_element_impl(
        content,
        element_namespace,
        Some(element_local),
        attribute_namespace,
        attribute,
        expected,
    )
}

fn locate_named_element_impl(
    content: &str,
    element_namespace: &[u8],
    element_local: Option<&[u8]>,
    attribute_namespace: &[u8],
    attribute: &[u8],
    expected: &str,
) -> Result<Option<Span>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<(usize, usize, Vec<u8>)> = None;
    let mut found = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|error| {
            Error::InvalidFormat(format!("ODP XML position does not fit usize: {error}"))
        })?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODP content XML: {error}")))?;
        let namespace_matches = resolved_namespace(&resolved) == Some(element_namespace);
        drop(resolved);
        let end = usize::try_from(reader.buffer_position()).map_err(|error| {
            Error::InvalidFormat(format!("ODP XML position does not fit usize: {error}"))
        })?;
        match event {
            Event::Start(element) => {
                if namespace_matches
                    && element_local.is_none_or(|local| element.local_name().as_ref() == local)
                    && read_attribute(&reader, &element, attribute_namespace, attribute)?.as_deref()
                        == Some(expected)
                {
                    if active.is_some() || found.is_some() {
                        return invalid(format!("ODP object selector '{expected}' is ambiguous"));
                    }
                    active = Some((depth, start, element.name().as_ref().to_vec()));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODP XML depth overflow".into()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid("ODP content XML exceeds the depth limit");
                }
            },
            Event::Empty(element) => {
                if namespace_matches
                    && element_local.is_none_or(|local| element.local_name().as_ref() == local)
                    && read_attribute(&reader, &element, attribute_namespace, attribute)?.as_deref()
                        == Some(expected)
                {
                    if active.is_some() || found.is_some() {
                        return invalid(format!("ODP object selector '{expected}' is ambiguous"));
                    }
                    found = Some(Span { start, end });
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ODP XML depth underflow".into()))?;
                if active.as_ref().is_some_and(|(active_depth, _, qualified)| {
                    *active_depth == depth && qualified.as_slice() == element.name().as_ref()
                }) {
                    let (_, active_start, _) = active.take().ok_or_else(|| {
                        Error::InvalidFormat("ODP object scan state disappeared".into())
                    })?;
                    found = Some(Span {
                        start: active_start,
                        end,
                    });
                }
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in ODP content XML"),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if active.is_some() || depth != 0 {
        return invalid("unterminated ODP content XML");
    }
    Ok(found)
}

fn locate_forms_site(content: &str) -> Result<FormsSite> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut forms_depth = None;
    let mut presentation_start = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|error| {
            Error::InvalidFormat(format!("ODP XML position does not fit usize: {error}"))
        })?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODP content XML: {error}")))?;
        let namespace = resolved_namespace(&resolved).map(<[u8]>::to_vec);
        drop(resolved);
        let end = usize::try_from(reader.buffer_position()).map_err(|error| {
            Error::InvalidFormat(format!("ODP XML position does not fit usize: {error}"))
        })?;
        match event {
            Event::Start(element) => {
                if namespace.as_deref() == Some(OFFICE_NS)
                    && element.local_name().as_ref() == b"presentation"
                {
                    presentation_start = Some(end);
                } else if namespace.as_deref() == Some(OFFICE_NS)
                    && element.local_name().as_ref() == b"forms"
                {
                    if forms_depth.is_some() {
                        return invalid("ODP content contains nested office:forms");
                    }
                    forms_depth = Some(depth);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODP XML depth overflow".into()))?;
            },
            Event::Empty(element)
                if namespace.as_deref() == Some(OFFICE_NS)
                    && element.local_name().as_ref() == b"forms" =>
            {
                let name_end = start
                    .checked_add(element.name().as_ref().len())
                    .and_then(|value| value.checked_add(1))
                    .ok_or_else(|| Error::InvalidFormat("ODP forms span overflow".into()))?;
                return Ok(FormsSite::Empty {
                    start,
                    end,
                    name_end,
                });
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ODP XML depth underflow".into()))?;
                if namespace.as_deref() == Some(OFFICE_NS)
                    && element.local_name().as_ref() == b"forms"
                    && forms_depth == Some(depth)
                {
                    return Ok(FormsSite::Content(start));
                }
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in ODP content XML"),
            Event::Eof => break,
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        if depth > MAX_XML_DEPTH {
            return invalid("ODP content XML exceeds the depth limit");
        }
        buffer.clear();
    }
    presentation_start
        .map(FormsSite::Missing)
        .ok_or_else(|| Error::InvalidFormat("ODP content has no office:presentation".into()))
}

fn collect_attribute_values(xml: &str, namespace: &[u8], local_name: &[u8]) -> Result<Vec<String>> {
    collect_attributes(xml, |resolved, local| {
        resolved_namespace(resolved) == Some(namespace) && local == local_name
    })
}

fn collect_style_attribute_values(xml: &str) -> Result<Vec<String>> {
    collect_attributes(xml, |_resolved, local| local.ends_with(b"style-name"))
}

fn collect_attributes(
    xml: &str,
    mut selected: impl FnMut(&ResolveResult<'_>, &[u8]) -> bool,
) -> Result<Vec<String>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut values = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|cause| {
                Error::InvalidFormat(format!("invalid ODP dependency XML: {cause}"))
            })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for raw_attribute in element.attributes() {
                    let attribute = raw_attribute.map_err(|cause| {
                        Error::InvalidFormat(format!("invalid ODP dependency attribute: {cause}"))
                    })?;
                    let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
                    if selected(&resolved, local.as_ref()) {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|cause| {
                                Error::InvalidFormat(format!(
                                    "invalid ODP dependency attribute value: {cause}"
                                ))
                            })?;
                        values.push(value.into_owned());
                    }
                }
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in ODP dependency XML"),
            Event::Eof => break,
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    values.sort();
    values.dedup();
    Ok(values)
}

fn read_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute_result in element.attributes() {
        let parsed_attribute = attribute_result
            .map_err(|error| Error::InvalidFormat(format!("invalid ODP XML attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_namespace(&resolved) == Some(namespace) && local.as_ref() == local_name {
            let decoded_value = parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid escaped ODP attribute: {error}"))
                })?;
            return Ok(Some(decoded_value.into_owned()));
        }
    }
    Ok(None)
}

fn resolved_namespace<'a>(resolved: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match resolved {
        ResolveResult::Bound(Namespace(namespace)) => Some(*namespace),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}
