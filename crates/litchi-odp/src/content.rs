//! Typed rich-text, table, and inert form-control content for presentation pages.
//!
//! These values deliberately model authored presentation objects rather than
//! executable form behavior. Package mutation is owned by [`crate::edit`].

use litchi_core::{Error, Result, xml::escape_xml};
use litchi_odf_common::package::{rebuild_package, splice};
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

pub(crate) fn apply(source: &crate::core::OwnedPackage, operation: &Operation) -> Result<Vec<u8>> {
    let content = String::from_utf8(source.get_file("content.xml")?)
        .map_err(|cause| Error::InvalidFormat(format!("ODP content.xml is not UTF-8: {cause}")))?;
    let updated = match operation {
        Operation::AddObject {
            page,
            kind,
            name,
            xml,
        } => add_object(&content, *page, *kind, name, xml)?,
        Operation::ReplaceObject {
            kind,
            name,
            new_name,
            xml,
        } => replace_object(&content, *kind, name, new_name, xml)?,
        Operation::RemoveObject { kind, name } => remove_object(&content, *kind, name)?,
        Operation::AddControl { page, control } => add_control(&content, *page, control)?,
        Operation::ReplaceControl { name, control } => replace_control(&content, name, control)?,
        Operation::RemoveControl { name } => remove_control(&content, name)?,
    };
    rebuild_package(
        source,
        &updated,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
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
        let namespace_matches = resolved_namespace(&resolved) == Some(namespace);
        drop(resolved);
        let end = usize::try_from(reader.buffer_position()).map_err(|error| {
            Error::InvalidFormat(format!("ODP XML position does not fit usize: {error}"))
        })?;
        match event {
            Event::Start(element) => {
                if namespace_matches
                    && read_attribute(&reader, &element, namespace, attribute)?.as_deref()
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
                    && read_attribute(&reader, &element, namespace, attribute)?.as_deref()
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
