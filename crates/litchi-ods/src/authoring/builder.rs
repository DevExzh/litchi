use crate::model::names::{Definition, Expression, Range};
use crate::worksheet::{Cell, Sheet};
use litchi_core::{Error, Result};
use litchi_odf_common::calculation::Settings;
use litchi_odf_common::core::PackageWriter;
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MAX_CONTENT_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_CONTENT_XML_DEPTH: usize = 1024;

/// Minimal package builder; richer sheet authoring is migrated independently.
#[derive(Clone, Debug)]
pub struct Builder {
    content_xml: String,
    definitions: Vec<Definition>,
    metadata: litchi_core::Metadata,
    settings: Option<Settings>,
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

impl Builder {
    #[must_use]
    pub fn new() -> Self {
        Self {
            content_xml: empty_content().to_owned(),
            definitions: Vec::new(),
            metadata: litchi_core::Metadata::default(),
            settings: None,
        }
    }

    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = xml.into();
        self
    }

    /// Borrow the compact metadata value that will be written to `meta.xml`.
    #[must_use]
    pub fn metadata(&self) -> &litchi_core::Metadata {
        &self.metadata
    }

    /// Replace the supported metadata value used for the next build.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_metadata(&mut self, metadata: litchi_core::Metadata) -> Result<&mut Self> {
        let snapshot = crate::metadata::Snapshot::from_source(None)?;
        let mut transaction = snapshot.transaction();
        transaction.replace(metadata.clone())?;
        self.metadata = metadata;
        Ok(self)
    }

    /// Borrow the calculation settings that will be emitted in `content.xml`.
    #[must_use]
    pub fn settings(&self) -> Option<&Settings> {
        self.settings.as_ref()
    }

    /// Set or clear validated calculation settings.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_settings(&mut self, settings: Option<Settings>) -> Result<&mut Self> {
        if let Some(settings) = &settings {
            settings.validate()?;
        }
        self.settings = settings;
        Ok(self)
    }

    /// Decode the builder's current typed worksheet snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn sheets(&self) -> Result<Vec<Sheet>> {
        crate::worksheet::codec::parse(&self.content_xml)
    }

    /// Atomically replace all worksheets in the builder's content snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_sheets(&mut self, sheets: Vec<Sheet>) -> Result<&mut Self> {
        let updated = crate::worksheet::package::replace_tables(&self.content_xml, &sheets)?;
        self.content_xml = updated;
        Ok(self)
    }

    /// Append one validated worksheet while preserving sheet order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_sheet(&mut self, sheet: Sheet) -> Result<&mut Self> {
        let mut sheets = self.sheets()?;
        sheets.push(sheet);
        self.set_sheets(sheets)
    }

    /// Remove one worksheet by its exact ODF name.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn remove_sheet(&mut self, name: &str) -> Result<Sheet> {
        let mut sheets = self.sheets()?;
        let index = sheets
            .iter()
            .position(|sheet| sheet.name == name)
            .ok_or_else(|| Error::InvalidFormat(format!("ODS sheet '{name}' was not found")))?;
        let removed = sheets.remove(index);
        self.set_sheets(sheets)?;
        Ok(removed)
    }

    /// Atomically replace one logical cell in a named worksheet.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell(
        &mut self,
        sheet_name: &str,
        row: usize,
        column: usize,
        cell: Cell,
    ) -> Result<&mut Self> {
        self.edit_sheet(sheet_name, |sheet| sheet.set_cell(row, column, cell))
    }

    /// Clear one logical cell while retaining its direct style, if any.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_cell(&mut self, sheet_name: &str, row: usize, column: usize) -> Result<&mut Self> {
        self.edit_sheet(sheet_name, |sheet| sheet.clear_cell(row, column))
    }

    /// Set an inert formula on one logical cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_formula(
        &mut self,
        sheet_name: &str,
        row: usize,
        column: usize,
        formula: impl Into<String>,
    ) -> Result<&mut Self> {
        let formula = formula.into();
        self.edit_sheet(sheet_name, move |sheet| {
            sheet.set_formula(row, column, formula)
        })
    }

    /// Set a direct cell style reference on one logical cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell_style(
        &mut self,
        sheet_name: &str,
        row: usize,
        column: usize,
        style_name: impl Into<String>,
    ) -> Result<&mut Self> {
        let style_name = style_name.into();
        self.edit_sheet(sheet_name, move |sheet| {
            sheet.set_cell_style(row, column, style_name)
        })
    }

    fn edit_sheet<F>(&mut self, sheet_name: &str, operation: F) -> Result<&mut Self>
    where
        F: FnOnce(&mut Sheet) -> Result<()>,
    {
        let sheets = self.sheets()?;
        let candidate = crate::worksheet::transaction::edit(&sheets, |candidate| {
            let sheet = candidate
                .iter_mut()
                .find(|sheet| sheet.name == sheet_name)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!("ODS sheet '{sheet_name}' was not found"))
                })?;
            operation(sheet)
        })?;
        self.set_sheets(candidate)
    }

    /// Return authored named definitions in their insertion order.
    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        &self.definitions
    }

    /// Append a validated named range to the builder.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_range(&mut self, range: Range) -> Result<&mut Self> {
        self.add_definition(range.into())
    }

    /// Append a validated named expression to the builder.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_expression(&mut self, expression: Expression) -> Result<&mut Self> {
        self.add_definition(expression.into())
    }

    /// Append a named definition while preserving authored order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_definition(&mut self, definition: Definition) -> Result<&mut Self> {
        let mut candidate = self.definitions.clone();
        candidate.push(definition);
        crate::model::names::validate_collection(&candidate)?;
        self.definitions = candidate;
        Ok(self)
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn build(self) -> Result<Vec<u8>> {
        let mut content_xml = if self.definitions.is_empty() {
            self.content_xml
        } else {
            crate::codec::names::replace(&self.content_xml, &self.definitions)?
        };
        if let Some(settings) = &self.settings {
            let snapshot = crate::settings::Snapshot::from_content_xml(&content_xml)?;
            let mut transaction = snapshot.transaction();
            transaction.replace(settings.clone())?;
            content_xml = transaction.commit()?.into_owned();
        }
        validate_content_xml(&content_xml)?;
        crate::worksheet::codec::parse(&content_xml)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIMETYPE)?;
        writer.add_file("content.xml", content_xml.as_bytes())?;
        if self.metadata.has_data() {
            let snapshot = crate::metadata::Snapshot::from_source(None)?;
            let mut transaction = snapshot.transaction();
            transaction.replace(self.metadata)?;
            if let Some(metadata_xml) = transaction.commit()?.into_owned_xml() {
                writer.add_file("meta.xml", metadata_xml.as_bytes())?;
            }
        }
        writer.finish_to_bytes()
    }
}

/// Validate the minimal ODF package-content hierarchy required by an ODS.
///
/// This is intentionally a structural boundary check rather than a complete
/// schema validator. The reader borrows `xml`, so authoring does not create a
/// second content buffer or an intermediate DOM.
pub(crate) fn validate_content_xml(xml: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODS content.xml exceeds {MAX_CONTENT_XML_BYTES} bytes"
        )));
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;
    let mut body_open = false;
    let mut body_seen = false;
    let mut spreadsheet_seen = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS content.xml: {error}")))?;

        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_closed {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml has more than one root element".to_string(),
                        ));
                    }
                    if !is_office_element(
                        &namespace,
                        element.local_name().as_ref(),
                        b"document-content",
                    ) {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml root must be office:document-content".to_string(),
                        ));
                    }
                } else if depth == 1
                    && is_office_element(&namespace, element.local_name().as_ref(), b"body")
                {
                    if body_seen {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml has more than one office:body".to_string(),
                        ));
                    }
                    body_seen = true;
                    body_open = true;
                } else if is_office_element(&namespace, element.local_name().as_ref(), b"body") {
                    return Err(Error::InvalidFormat(
                        "office:body must be a direct child of office:document-content".to_string(),
                    ));
                } else if depth == 2
                    && body_open
                    && is_office_element(&namespace, element.local_name().as_ref(), b"spreadsheet")
                {
                    if spreadsheet_seen {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml has more than one office:spreadsheet".to_string(),
                        ));
                    }
                    spreadsheet_seen = true;
                } else if is_office_element(
                    &namespace,
                    element.local_name().as_ref(),
                    b"spreadsheet",
                ) {
                    return Err(Error::InvalidFormat(
                        "office:spreadsheet must be a direct child of office:body".to_string(),
                    ));
                }

                if depth == MAX_CONTENT_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODS content.xml nesting exceeds {MAX_CONTENT_XML_DEPTH} elements"
                    )));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "ODS content.xml must contain office:body and office:spreadsheet"
                            .to_string(),
                    ));
                }

                if depth == 1
                    && is_office_element(&namespace, element.local_name().as_ref(), b"body")
                {
                    if body_seen {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml has more than one office:body".to_string(),
                        ));
                    }
                    body_seen = true;
                } else if is_office_element(&namespace, element.local_name().as_ref(), b"body") {
                    return Err(Error::InvalidFormat(
                        "office:body must be a direct child of office:document-content".to_string(),
                    ));
                } else if depth == 2
                    && body_open
                    && is_office_element(&namespace, element.local_name().as_ref(), b"spreadsheet")
                {
                    if spreadsheet_seen {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml has more than one office:spreadsheet".to_string(),
                        ));
                    }
                    spreadsheet_seen = true;
                } else if is_office_element(
                    &namespace,
                    element.local_name().as_ref(),
                    b"spreadsheet",
                ) {
                    return Err(Error::InvalidFormat(
                        "office:spreadsheet must be a direct child of office:body".to_string(),
                    ));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "ODS content.xml has an unexpected closing element".to_string(),
                    ));
                }

                if depth == 1 {
                    if !is_office_element(
                        &namespace,
                        element.local_name().as_ref(),
                        b"document-content",
                    ) {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml root must close with office:document-content"
                                .to_string(),
                        ));
                    }
                    if body_open {
                        return Err(Error::InvalidFormat(
                            "ODS content.xml closes its root before office:body".to_string(),
                        ));
                    }
                    depth = 0;
                    root_closed = true;
                } else {
                    if depth == 2
                        && body_open
                        && is_office_element(&namespace, element.local_name().as_ref(), b"body")
                    {
                        body_open = false;
                        if !spreadsheet_seen {
                            return Err(Error::InvalidFormat(
                                "office:body must contain office:spreadsheet".to_string(),
                            ));
                        }
                    }
                    depth -= 1;
                }
            },
            Event::Text(text) if depth == 0 && has_non_whitespace(text.as_ref()) => {
                return Err(Error::InvalidFormat(
                    "ODS content.xml has non-whitespace text outside its root".to_string(),
                ));
            },
            Event::CData(text) if depth == 0 && has_non_whitespace(text.as_ref()) => {
                return Err(Error::InvalidFormat(
                    "ODS content.xml has non-whitespace text outside its root".to_string(),
                ));
            },
            Event::Eof => {
                if !root_closed || depth != 0 {
                    return Err(Error::InvalidFormat(
                        "ODS content.xml ended before its root element was closed".to_string(),
                    ));
                }
                if !body_seen || !spreadsheet_seen {
                    return Err(Error::InvalidFormat(
                        "ODS content.xml must contain office:body with office:spreadsheet"
                            .to_string(),
                    ));
                }
                return Ok(());
            },
            Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Text(_)
            | Event::CData(_) => {},
        }
        buffer.clear();
    }
}

fn is_office_element(namespace: &ResolveResult<'_>, local_name: &[u8], expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE)
        && local_name == expected
}

fn has_non_whitespace(value: &[u8]) -> bool {
    !value.iter().all(u8::is_ascii_whitespace)
}

fn empty_content() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3"><office:body><office:spreadsheet/></office:body></office:document-content>"#
}

#[cfg(test)]
mod tests {
    use super::Builder;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    #[test]
    fn build_accepts_namespace_aliases_and_empty_spreadsheet() {
        let content = format!(
            r#"<o:document-content xmlns:o="{OFFICE}"><o:body><o:spreadsheet/></o:body></o:document-content>"#
        );

        assert!(Builder::new().content_xml(content).build().is_ok());
    }

    #[test]
    fn build_rejects_malformed_or_incomplete_ods_content() {
        let malformed = Builder::new()
            .content_xml(format!(
                r#"<office:document-content xmlns:office="{OFFICE}"><office:body>"#
            ))
            .build()
            .unwrap_err()
            .to_string();
        assert!(malformed.contains("ended before its root element was closed"));

        let incomplete = Builder::new()
            .content_xml(format!(
                r#"<office:document-content xmlns:office="{OFFICE}"><office:body/></office:document-content>"#
            ))
            .build()
            .unwrap_err()
            .to_string();
        assert!(incomplete.contains("office:spreadsheet"));
    }
}
