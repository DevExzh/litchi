//! ODS-specific parsing utilities.

use super::{
    Cell, CellValue, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange,
    NamedRangeUsage, Row, Sheet,
    annotation::{AnnotationBuilder, decode_reference},
};
use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// Parser for ODS-specific structures.
///
/// This provides parsing logic specific to spreadsheets,
/// including sheet, row, and cell parsing with proper type detection.
pub(crate) struct OdsParser;

impl OdsParser {
    /// Parse all sheets from ODS content.xml
    pub fn parse_sheets(xml_content: &str) -> Result<Vec<Sheet>> {
        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut sheets = Vec::new();

        // Parser state
        let mut current_sheet: Option<SheetBuilder> = None;
        let mut current_row: Option<RowBuilder> = None;
        let mut current_cell: Option<CellBuilder> = None;
        let mut text_element_depth = 0usize;
        let mut text_content = String::new();
        let mut annotation_builder: Option<AnnotationBuilder> = None;
        let mut annotation_depth = 0usize;
        let mut document_namespaces = BTreeMap::new();
        let mut namespace_scopes = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let namespace_scope =
                        Self::push_namespace_scope(e, reader.decoder(), &mut document_namespaces)?;
                    namespace_scopes.push(namespace_scope);

                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.start(e, reader.decoder())?;
                        annotation_depth += 1;
                    } else if current_cell.is_some()
                        && Self::is_office_annotation(e, &document_namespaces)
                    {
                        annotation_builder = Some(AnnotationBuilder::new(
                            e,
                            reader.decoder(),
                            document_namespaces.clone(),
                        )?);
                    } else if text_element_depth > 0 {
                        text_element_depth += 1;
                    } else {
                        match e.name().as_ref() {
                            b"table:table" => {
                                let name = Self::extract_table_name(e)?;
                                current_sheet = Some(SheetBuilder::new(name));
                            },
                            b"table:table-row" if current_sheet.is_some() => {
                                current_row = Some(RowBuilder::new());
                            },
                            b"table:table-cell" if current_row.is_some() => {
                                let cell_builder =
                                    Self::parse_cell_attributes(e, reader.decoder())?;
                                current_cell = Some(cell_builder);
                                text_content.clear();
                            },
                            b"text:p" if current_cell.is_some() => {
                                if !text_content.is_empty() {
                                    text_content.push('\n');
                                }
                                text_element_depth = 1;
                            },
                            _ => {},
                        }
                    }
                },
                Ok(Event::Empty(ref e)) => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.empty(e, reader.decoder())?;
                    } else if current_cell.is_some()
                        && Self::is_office_annotation(e, &document_namespaces)
                    {
                        let annotation = AnnotationBuilder::new(
                            e,
                            reader.decoder(),
                            document_namespaces.clone(),
                        )?
                        .finish()?;
                        if let Some(cell) = current_cell.as_mut() {
                            cell.annotation = Some(annotation);
                        }
                    } else if text_element_depth > 0 {
                        Self::push_text_empty_element(e, reader.decoder(), &mut text_content)?;
                    } else if e.name().as_ref() == b"table:table-cell" && current_row.is_some() {
                        let cell_builder = Self::parse_cell_attributes(e, reader.decoder())?;
                        if let Some(row) = current_row.as_mut() {
                            for _ in 0..cell_builder.repeated {
                                row.add_cell(cell_builder.build(""));
                            }
                        }
                    }
                },
                Ok(Event::Text(ref t)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.text(t)?;
                    }
                },
                Ok(Event::CData(ref t)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.cdata(t)?;
                    }
                },
                Ok(Event::GeneralRef(ref reference)) if annotation_builder.is_some() => {
                    if let Some(builder) = annotation_builder.as_mut() {
                        builder.reference(reference)?;
                    }
                },
                Ok(Event::Text(ref t)) if text_element_depth > 0 && current_cell.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell text: {error}"))
                    })?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell character reference: {error}"))
                    })?;
                    text_content.push_str(&decoded);
                },
                Ok(Event::CData(ref t)) if text_element_depth > 0 && current_cell.is_some() => {
                    let decoded = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid cell CDATA: {error}"))
                    })?;
                    text_content.push_str(&decoded);
                },
                Ok(Event::GeneralRef(ref reference))
                    if text_element_depth > 0 && current_cell.is_some() =>
                {
                    text_content.push_str(&decode_reference(reference)?);
                },
                Ok(Event::End(ref e)) => {
                    if annotation_builder.is_some() {
                        if annotation_depth == 0 {
                            let annotation = annotation_builder
                                .take()
                                .expect("annotation builder was checked")
                                .finish()?;
                            if let Some(cell) = current_cell.as_mut() {
                                cell.annotation = Some(annotation);
                            }
                        } else {
                            annotation_builder
                                .as_mut()
                                .expect("annotation builder was checked")
                                .end_element()?;
                            annotation_depth -= 1;
                        }
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    if text_element_depth > 0 {
                        text_element_depth -= 1;
                        Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                        buf.clear();
                        continue;
                    }

                    match e.name().as_ref() {
                        b"table:table-cell" => {
                            if let Some(cell_builder) = current_cell.take() {
                                let repeated = cell_builder.repeated;
                                if let Some(ref mut row_builder) = current_row {
                                    // Handle repeated cells - build cell once for first occurrence
                                    if repeated > 0 {
                                        let cell = cell_builder.build(&text_content);
                                        row_builder.add_cell(cell);
                                        // Clone only for additional repetitions
                                        for _ in 1..repeated {
                                            row_builder.add_cell(cell_builder.build(&text_content));
                                        }
                                    }
                                }
                            }
                        },
                        b"table:table-row" => {
                            if let Some(row_builder) = current_row.take() {
                                let row = row_builder.build();
                                if let Some(ref mut sheet_builder) = current_sheet {
                                    sheet_builder.add_row(row);
                                }
                            }
                        },
                        b"table:table" => {
                            if let Some(sheet_builder) = current_sheet.take() {
                                let sheet = sheet_builder.build();
                                sheets.push(sheet);
                            }
                        },
                        _ => {},
                    }
                    Self::pop_namespace_scope(&mut document_namespaces, namespace_scopes.pop());
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(Error::InvalidFormat(format!("XML parsing error: {}", e)));
                },
                _ => {},
            }
            buf.clear();
        }

        Ok(sheets)
    }

    fn is_office_annotation(
        element: &BytesStart<'_>,
        namespaces: &BTreeMap<String, String>,
    ) -> bool {
        let qualified_name = element.name();
        let Ok(name) = std::str::from_utf8(qualified_name.as_ref()) else {
            return false;
        };
        let Some((prefix, local)) = name.split_once(':') else {
            return false;
        };
        local == "annotation"
            && namespaces
                .get(prefix)
                .is_some_and(|namespace| namespace == OFFICE_NAMESPACE)
    }

    fn push_namespace_scope(
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: &mut BTreeMap<String, String>,
    ) -> Result<Vec<(String, Option<String>)>> {
        let mut previous_bindings = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                Error::InvalidFormat(format!("invalid XML namespace declaration: {error}"))
            })?;
            let name = std::str::from_utf8(attribute.key.as_ref()).map_err(|_| {
                Error::InvalidFormat("invalid UTF-8 in XML namespace declaration".to_string())
            })?;
            let Some(prefix) = name.strip_prefix("xmlns:") else {
                continue;
            };
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML namespace URI: {error}"))
                })?
                .into_owned();
            let prefix = prefix.to_string();
            let previous = namespaces.insert(prefix.clone(), value);
            previous_bindings.push((prefix, previous));
        }
        Ok(previous_bindings)
    }

    fn pop_namespace_scope(
        namespaces: &mut BTreeMap<String, String>,
        scope: Option<Vec<(String, Option<String>)>>,
    ) {
        let Some(scope) = scope else { return };
        for (prefix, previous) in scope.into_iter().rev() {
            if let Some(previous) = previous {
                namespaces.insert(prefix, previous);
            } else {
                namespaces.remove(&prefix);
            }
        }
    }

    fn push_text_empty_element(
        element: &BytesStart<'_>,
        decoder: Decoder,
        text: &mut String,
    ) -> Result<()> {
        match element.name().as_ref() {
            b"text:line-break" => text.push('\n'),
            b"text:tab" => text.push('\t'),
            b"text:s" => {
                let mut count = 1usize;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid text:s attribute: {error}"))
                    })?;
                    if attribute.key.as_ref() == b"text:c" {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                            .map_err(|error| {
                                Error::InvalidFormat(format!("invalid text:s count: {error}"))
                            })?;
                        count = value.parse::<usize>().map_err(|_| {
                            Error::InvalidFormat(format!("invalid text:s count '{value}'"))
                        })?;
                    }
                }
                if count > 1_000_000 {
                    return Err(Error::InvalidFormat(
                        "text:s count exceeds the supported safety limit".to_string(),
                    ));
                }
                text.extend(std::iter::repeat_n(' ', count));
            },
            _ => {},
        }
        Ok(())
    }

    /// Parse document-global and sheet-local named ranges and expressions.
    ///
    /// Namespace URIs are resolved rather than assuming the conventional
    /// `table` prefix, because XML namespace prefixes are freely replaceable.
    pub fn parse_named_definitions(xml_content: &str) -> Result<Vec<NamedDefinition>> {
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut definitions = Vec::new();
        let mut table_stack: Vec<Option<String>> = Vec::new();
        let mut active_scope: Option<NamedDefinitionScope> = None;

        loop {
            let event = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;

            match event {
                (namespace, Event::Start(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"table" => {
                            let name = Self::table_attribute(
                                reader.resolver(),
                                reader.decoder(),
                                &element,
                                b"name",
                            )?;
                            table_stack.push(name);
                        },
                        b"named-expressions" => {
                            if active_scope.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested table:named-expressions element".to_string(),
                                ));
                            }
                            active_scope = Some(match table_stack.last() {
                                Some(Some(sheet)) => NamedDefinitionScope::Sheet(sheet.clone()),
                                Some(None) => {
                                    return Err(Error::InvalidFormat(
                                        "sheet-local named definitions require table:name"
                                            .to_string(),
                                    ));
                                },
                                None => NamedDefinitionScope::Global,
                            });
                        },
                        b"named-range" | b"named-expression" => {
                            if let Some(scope) = &active_scope {
                                definitions.push(Self::parse_named_definition(
                                    reader.resolver(),
                                    reader.decoder(),
                                    &element,
                                    scope.clone(),
                                )?);
                            }
                        },
                        _ => {},
                    }
                },
                (namespace, Event::Empty(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"named-range" | b"named-expression" => {
                            if let Some(scope) = &active_scope {
                                definitions.push(Self::parse_named_definition(
                                    reader.resolver(),
                                    reader.decoder(),
                                    &element,
                                    scope.clone(),
                                )?);
                            }
                        },
                        _ => {},
                    }
                },
                (namespace, Event::End(element)) if Self::is_table_namespace(&namespace) => {
                    match element.local_name().as_ref() {
                        b"named-expressions" => active_scope = None,
                        b"table" => {
                            table_stack.pop();
                        },
                        _ => {},
                    }
                },
                (_, Event::Eof) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(definitions)
    }

    fn is_table_namespace(namespace: &ResolveResult<'_>) -> bool {
        matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TABLE_NAMESPACE)
    }

    fn parse_named_definition(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        scope: NamedDefinitionScope,
    ) -> Result<NamedDefinition> {
        let name = Self::required_table_attribute(resolver, decoder, element, b"name")?;
        let base_cell_address =
            Self::table_attribute(resolver, decoder, element, b"base-cell-address")?;

        match element.local_name().as_ref() {
            b"named-range" => {
                let cell_range_address = Self::required_table_attribute(
                    resolver,
                    decoder,
                    element,
                    b"cell-range-address",
                )?;
                let mut range = NamedRange::new(name, cell_range_address, scope)?;
                range.base_cell_address = base_cell_address;
                if let Some(usable_as) =
                    Self::table_attribute(resolver, decoder, element, b"range-usable-as")?
                {
                    if usable_as.is_empty() {
                        return Err(Error::InvalidFormat(
                            "table:range-usable-as must not be empty".to_string(),
                        ));
                    }
                    if usable_as != "none" {
                        for token in usable_as.split_whitespace() {
                            let usage = NamedRangeUsage::parse(token)?;
                            if !range.usable_as.contains(&usage) {
                                range.usable_as.push(usage);
                            }
                        }
                    }
                }
                range.validate()?;
                Ok(range.into())
            },
            b"named-expression" => {
                let expression =
                    Self::required_table_attribute(resolver, decoder, element, b"expression")?;
                let formula_namespace_uri = Self::formula_namespace_uri(resolver, &expression)?;
                let mut expression = if let Some(uri) = formula_namespace_uri {
                    NamedExpression::new_with_namespace(name, expression, uri, scope)?
                } else {
                    NamedExpression::new(name, expression, scope)?
                };
                expression.base_cell_address = base_cell_address;
                expression.validate()?;
                Ok(expression.into())
            },
            _ => Err(Error::InvalidFormat(
                "unexpected named definition element".to_string(),
            )),
        }
    }

    fn required_table_attribute(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        local_name: &[u8],
    ) -> Result<String> {
        Self::table_attribute(resolver, decoder, element, local_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{} is missing required table:{} attribute",
                String::from_utf8_lossy(element.local_name().as_ref()),
                String::from_utf8_lossy(local_name)
            ))
        })
    }

    fn table_attribute(
        resolver: &NamespaceResolver,
        decoder: Decoder,
        element: &BytesStart<'_>,
        local_name: &[u8],
    ) -> Result<Option<String>> {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let (namespace, local) = resolver.resolve_attribute(attribute.key);
            if Self::is_table_namespace(&namespace) && local.as_ref() == local_name {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                    })?;
                return Ok(Some(value.into_owned()));
            }
        }
        Ok(None)
    }

    fn formula_namespace_uri(
        resolver: &NamespaceResolver,
        expression: &str,
    ) -> Result<Option<String>> {
        let Some((prefix, remainder)) = expression.split_once(':') else {
            return Ok(None);
        };
        if prefix.is_empty() || !remainder.starts_with('=') {
            return Ok(None);
        }

        for (declaration, namespace) in resolver.bindings() {
            if let PrefixDeclaration::Named(candidate) = declaration
                && candidate == prefix.as_bytes()
            {
                return String::from_utf8(namespace.as_ref().to_vec())
                    .map(Some)
                    .map_err(|_| {
                        Error::InvalidFormat(format!(
                            "formula namespace for prefix '{prefix}' is not UTF-8"
                        ))
                    });
            }
        }

        Err(Error::InvalidFormat(format!(
            "formula prefix '{prefix}' is not bound to a namespace"
        )))
    }

    /// Extract table name from table:table element
    fn extract_table_name(e: &quick_xml::events::BytesStart) -> Result<String> {
        for attr_result in e.attributes() {
            let attr =
                attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
            if attr.key.as_ref() == b"table:name" {
                return String::from_utf8(attr.value.to_vec())
                    .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in table name".to_string()));
            }
        }
        Ok("Sheet1".to_string()) // Default name
    }

    /// Parse cell attributes and create a CellBuilder
    fn parse_cell_attributes(
        e: &quick_xml::events::BytesStart,
        decoder: Decoder,
    ) -> Result<CellBuilder> {
        let mut value_type = None;
        let mut value_str = None;
        let mut currency = None;
        let mut formula = None;
        let mut validation_name = None;
        let mut protect = None;
        let mut protected = None;
        let mut repeated = 1;

        for attr_result in e.attributes() {
            let attr =
                attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
            match attr.key.as_ref() {
                b"office:value-type" => {
                    value_type = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"office:value" => {
                    value_str = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"office:currency" => {
                    currency = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"table:formula" => {
                    formula = Some(
                        String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
                    );
                },
                b"table:content-validation-name" => {
                    validation_name = Some(
                        attr.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                            .map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid content-validation name: {error}"
                                ))
                            })?
                            .into_owned(),
                    );
                },
                b"table:protect" => {
                    protect = Some(Self::parse_bool_attribute(&attr, decoder)?);
                },
                b"table:protected" => {
                    protected = Some(Self::parse_bool_attribute(&attr, decoder)?);
                },
                b"table:number-columns-repeated" => {
                    if let Ok(rep) = String::from_utf8(attr.value.to_vec())
                        .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?
                        .parse::<usize>()
                    {
                        repeated = rep;
                    }
                },
                _ => {},
            }
        }

        Ok(CellBuilder {
            value_type,
            value_str,
            currency,
            formula,
            validation_name,
            protect,
            protected,
            repeated,
            annotation: None,
        })
    }

    fn parse_bool_attribute(
        attribute: &quick_xml::events::attributes::Attribute<'_>,
        decoder: Decoder,
    ) -> Result<bool> {
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid Boolean attribute: {error}")))?;
        match value.as_ref() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid Boolean attribute value '{value}'"
            ))),
        }
    }
}

/// Builder for constructing Sheet during parsing
pub(crate) struct SheetBuilder {
    name: String,
    rows: Vec<Row>,
}

impl SheetBuilder {
    pub fn new(name: String) -> Self {
        Self {
            name,
            rows: Vec::new(),
        }
    }

    pub fn add_row(&mut self, mut row: Row) {
        let row_index = self.rows.len();
        row.index = row_index;
        // Update row index for all cells in this row
        for cell in &mut row.cells {
            cell.row = row_index;
        }
        self.rows.push(row);
    }

    pub fn build(self) -> Sheet {
        Sheet {
            name: self.name,
            rows: self.rows,
            protection: super::SheetProtection::default(),
        }
    }
}

/// Builder for constructing Row during parsing
pub(crate) struct RowBuilder {
    cells: Vec<Cell>,
}

impl RowBuilder {
    pub fn new() -> Self {
        Self { cells: Vec::new() }
    }

    pub fn add_cell(&mut self, mut cell: Cell) {
        cell.col = self.cells.len();
        self.cells.push(cell);
    }

    pub fn build(mut self) -> Row {
        // Row index will be set by the parent SheetBuilder
        // For now, set to 0 and update cells
        for cell in &mut self.cells {
            cell.row = 0; // Will be updated by parent
        }

        Row {
            cells: self.cells,
            index: 0, // Will be set by parent
        }
    }
}

/// Builder for constructing Cell during parsing
pub(crate) struct CellBuilder {
    value_type: Option<String>,
    value_str: Option<String>,
    currency: Option<String>,
    formula: Option<String>,
    validation_name: Option<String>,
    protect: Option<bool>,
    protected: Option<bool>,
    repeated: usize,
    annotation: Option<super::CellAnnotation>,
}

impl CellBuilder {
    pub fn build(&self, text_content: &str) -> Cell {
        let value = self.parse_value(text_content);

        Cell {
            value,
            text: text_content.to_string(),
            // Clone necessary: formula may be reused for repeated cells
            formula: self.formula.clone(),
            annotation: self.annotation.clone(),
            validation_name: self.validation_name.clone(),
            protect: self.protect,
            protected: self.protected,
            row: 0, // Will be set by parent
            col: 0, // Will be set by parent
        }
    }

    fn parse_value(&self, text_content: &str) -> CellValue {
        match self.value_type.as_deref() {
            Some("float") | Some("double") | Some("decimal") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Number(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("currency") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        let currency_code = self.currency.as_deref().unwrap_or("USD").to_string();
                        CellValue::Currency(num, currency_code)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("percentage") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Percentage(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("boolean") => {
                if let Some(ref val_str) = self.value_str {
                    match val_str.as_str() {
                        "true" => CellValue::Boolean(true),
                        "false" => CellValue::Boolean(false),
                        _ => CellValue::Text(text_content.to_string()),
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("date") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Date(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("time") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Time(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            _ => {
                if text_content.trim().is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEST_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Hello</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="float" office:value="42">
                        <text:p>42</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_MULTIPLE_SHEETS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Sheet1">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>First Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
            <table:table table:name="Sheet2">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Second Sheet</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_CELL_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="TypesTest">
                <table:table-row>
                    <table:table-cell office:value-type="string"><text:p>Text</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="3.14"><text:p>3.14</text:p></table:table-cell>
                    <table:table-cell office:value-type="currency" office:value="100" office:currency="EUR"><text:p>€100</text:p></table:table-cell>
                    <table:table-cell office:value-type="percentage" office:value="0.5"><text:p>50%</text:p></table:table-cell>
                    <table:table-cell office:value-type="boolean" office:value="true"><text:p>TRUE</text:p></table:table-cell>
                    <table:table-cell office:value-type="date" office:value="2024-03-15"><text:p>2024-03-15</text:p></table:table-cell>
                    <table:table-cell office:value-type="time" office:value="PT12H30M00S"><text:p>12:30:00</text:p></table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_FORMULA_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="FormulaTest">
                <table:table-row>
                    <table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="20"><text:p>20</text:p></table:table-cell>
                    <table:table-cell table:formula="=SUM([.A1]:[.B1])" office:value-type="float" office:value="30">
                        <text:p>30</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_REPEATED_CELLS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="RepeatedTest">
                <table:table-row>
                    <table:table-cell table:number-columns-repeated="3" office:value-type="string">
                        <text:p>Repeated</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>Single</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_EMPTY_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="EmptySheet">
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    const TEST_SPAN_TEXT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table table:name="SpanTest">
                <table:table-row>
                    <table:table-cell office:value-type="string">
                        <text:p>Normal text <text:span>spanned text</text:span> more text</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>"#;

    #[test]
    fn test_parse_sheets_basic() {
        let sheets = OdsParser::parse_sheets(TEST_SHEETS_XML).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[0].rows.len(), 1);
    }

    #[test]
    fn test_parse_multiple_sheets() {
        let sheets = OdsParser::parse_sheets(TEST_MULTIPLE_SHEETS_XML).unwrap();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[1].name, "Sheet2");
    }

    #[test]
    fn test_parse_cell_types() {
        let sheets = OdsParser::parse_sheets(TEST_CELL_TYPES_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 7);

        // Text cell
        match &row.cells[0].value {
            CellValue::Text(t) => assert_eq!(t, "Text"),
            _ => panic!("Expected Text"),
        }

        // Float/Number cell
        match &row.cells[1].value {
            CellValue::Number(n) => {
                let expected = (std::f64::consts::PI * 100.0).trunc() / 100.0;
                assert!((n - expected).abs() < f64::EPSILON);
            },
            _ => panic!("Expected Number"),
        }

        // Currency cell
        match &row.cells[2].value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 100.0).abs() < f64::EPSILON);
                assert_eq!(currency, "EUR");
            },
            _ => panic!("Expected Currency"),
        }

        // Percentage cell
        match &row.cells[3].value {
            CellValue::Percentage(p) => assert!((p - 0.5).abs() < f64::EPSILON),
            _ => panic!("Expected Percentage"),
        }

        // Boolean cell
        match &row.cells[4].value {
            CellValue::Boolean(b) => assert!(*b),
            _ => panic!("Expected Boolean"),
        }

        // Date cell
        match &row.cells[5].value {
            CellValue::Date(d) => assert_eq!(d, "2024-03-15"),
            _ => panic!("Expected Date"),
        }

        // Time cell
        match &row.cells[6].value {
            CellValue::Time(t) => assert_eq!(t, "PT12H30M00S"),
            _ => panic!("Expected Time"),
        }
    }

    #[test]
    fn test_parse_formula() {
        let sheets = OdsParser::parse_sheets(TEST_FORMULA_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 3);

        // Cell with formula
        assert_eq!(row.cells[2].formula, Some("=SUM([.A1]:[.B1])".to_string()));
        match &row.cells[2].value {
            CellValue::Number(n) => assert!((n - 30.0).abs() < f64::EPSILON),
            _ => panic!("Expected Number for formula result"),
        }
    }

    #[test]
    fn test_parse_repeated_cells() {
        let sheets = OdsParser::parse_sheets(TEST_REPEATED_CELLS_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        // 3 repeated cells + 1 single = 4 cells
        assert_eq!(row.cells.len(), 4);

        for i in 0..3 {
            match &row.cells[i].value {
                CellValue::Text(t) => assert_eq!(t, "Repeated"),
                _ => panic!("Expected Text for repeated cell {i}"),
            }
        }

        match &row.cells[3].value {
            CellValue::Text(t) => assert_eq!(t, "Single"),
            _ => panic!("Expected Text for single cell"),
        }
    }

    #[test]
    fn test_parse_empty_sheet() {
        let sheets = OdsParser::parse_sheets(TEST_EMPTY_SHEET_XML).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "EmptySheet");
        assert_eq!(sheets[0].rows.len(), 0);
    }

    #[test]
    fn test_parse_span_text() {
        let sheets = OdsParser::parse_sheets(TEST_SPAN_TEXT_XML).unwrap();
        assert_eq!(sheets.len(), 1);

        let row = &sheets[0].rows[0];
        assert_eq!(row.cells.len(), 1);

        // Text should include content from both text:p and text:span
        assert!(row.cells[0].text.contains("Normal text"));
        assert!(row.cells[0].text.contains("spanned text"));
    }

    #[test]
    fn parses_rich_annotations_without_mixing_them_into_cell_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <o:body><o:spreadsheet><table:table table:name="Notes"><table:table-row>
    <table:table-cell table:number-columns-repeated="2" o:value-type="string">
      <o:annotation o:display="true" draw:style-name="gr1" svg:width="3.2cm">
        <dc:creator>A &amp; B</dc:creator><dc:date>2026-07-13T12:34:56Z</dc:date>
        <text:p text:style-name="P1"><text:span text:style-name="T1">first</text:span><text:line-break/>second</text:p>
        <text:list><text:list-item><text:p>item</text:p></text:list-item></text:list>
      </o:annotation>
      <text:p>cell <text:span>value</text:span></text:p><text:p>line two</text:p>
    </table:table-cell>
  </table:table-row></table:table></o:spreadsheet></o:body>
</o:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        let cells = &sheets[0].rows[0].cells;
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].text, "cell value\nline two");
        assert_eq!(cells[1].text, "cell value\nline two");

        for cell in cells {
            let annotation = cell.annotation().unwrap();
            assert_eq!(annotation.creator().as_deref(), Some("A & B"));
            assert_eq!(annotation.date().as_deref(), Some("2026-07-13T12:34:56Z"));
            assert_eq!(annotation.display(), Some(true));
            assert_eq!(annotation.attribute("draw:style-name"), Some("gr1"));
            assert_eq!(annotation.attribute("svg:width"), Some("3.2cm"));
            assert_eq!(annotation.text(), "first\nsecond\nitem");
            assert_eq!(annotation.children()[2].name(), "text:p");
            assert_eq!(annotation.children()[3].name(), "text:list");
        }
    }

    #[test]
    fn test_extract_table_name_default() {
        // XML without table:name attribute
        let xml = r#"<?xml version="1.0"?>
<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
    <table:table>
    </table:table>
</office:document-content>"#;

        let sheets = OdsParser::parse_sheets(xml).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1"); // Default name
    }

    #[test]
    fn test_sheet_builder() {
        let mut builder = SheetBuilder::new("TestSheet".to_string());

        let row1 = Row {
            cells: vec![],
            index: 0,
        };
        builder.add_row(row1);

        let row2 = Row {
            cells: vec![Cell {
                value: CellValue::Text("A1".to_string()),
                text: "A1".to_string(),
                formula: None,
                annotation: None,
                validation_name: None,
                protect: None,
                protected: None,
                row: 0,
                col: 0,
            }],
            index: 0,
        };
        builder.add_row(row2);

        let sheet = builder.build();
        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(sheet.rows.len(), 2);
        assert_eq!(sheet.rows[0].index, 0);
        assert_eq!(sheet.rows[1].index, 1);
    }

    #[test]
    fn test_row_builder() {
        let mut builder = RowBuilder::new();

        let cell1 = Cell {
            value: CellValue::Text("A".to_string()),
            text: "A".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        builder.add_cell(cell1);

        let cell2 = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        builder.add_cell(cell2);

        let row = builder.build();
        assert_eq!(row.cells.len(), 2);
        assert_eq!(row.cells[0].col, 0);
        assert_eq!(row.cells[1].col, 1);
    }

    #[test]
    fn test_cell_builder_float_types() {
        // Test "float" value type
        let builder = CellBuilder {
            value_type: Some("float".to_string()),
            value_str: Some("123.45".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("123.45");
        match cell.value {
            CellValue::Number(n) => assert!((n - 123.45).abs() < f64::EPSILON),
            _ => panic!("Expected Number for float"),
        }

        // Test "double" value type
        let builder = CellBuilder {
            value_type: Some("double".to_string()),
            value_str: Some("99.99".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("99.99");
        match cell.value {
            CellValue::Number(n) => assert!((n - 99.99).abs() < f64::EPSILON),
            _ => panic!("Expected Number for double"),
        }

        // Test "decimal" value type
        let builder = CellBuilder {
            value_type: Some("decimal".to_string()),
            value_str: Some("0.001".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("0.001");
        match cell.value {
            CellValue::Number(n) => assert!((n - 0.001).abs() < f64::EPSILON),
            _ => panic!("Expected Number for decimal"),
        }
    }

    #[test]
    fn test_cell_builder_invalid_number_fallback() {
        let builder = CellBuilder {
            value_type: Some("float".to_string()),
            value_str: Some("not-a-number".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("some text");
        match cell.value {
            CellValue::Text(t) => assert_eq!(t, "some text"),
            _ => panic!("Expected Text fallback for invalid number"),
        }
    }

    #[test]
    fn test_cell_builder_boolean_variations() {
        // Test "false" boolean
        let builder = CellBuilder {
            value_type: Some("boolean".to_string()),
            value_str: Some("false".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("FALSE");
        match cell.value {
            CellValue::Boolean(b) => assert!(!b),
            _ => panic!("Expected Boolean false"),
        }

        // Test invalid boolean value (falls back to text)
        let builder = CellBuilder {
            value_type: Some("boolean".to_string()),
            value_str: Some("maybe".to_string()),
            currency: None,
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("maybe");
        match cell.value {
            CellValue::Text(t) => assert_eq!(t, "maybe"),
            _ => panic!("Expected Text for invalid boolean"),
        }
    }

    #[test]
    fn test_cell_builder_empty_text() {
        let builder = CellBuilder {
            value_type: None,
            value_str: None,
            currency: None,
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("   ");
        match cell.value {
            CellValue::Empty => {},
            _ => panic!("Expected Empty for whitespace-only text"),
        }
    }

    #[test]
    fn test_cell_builder_currency_default() {
        let builder = CellBuilder {
            value_type: Some("currency".to_string()),
            value_str: Some("50".to_string()),
            currency: None, // No currency specified
            formula: None,
            annotation: None,
            validation_name: None,
            protect: None,
            protected: None,
            repeated: 1,
        };
        let cell = builder.build("$50");
        match cell.value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 50.0).abs() < f64::EPSILON);
                assert_eq!(currency, "USD"); // Default
            },
            _ => panic!("Expected Currency with default USD"),
        }
    }

    #[test]
    fn test_parse_invalid_xml() {
        let invalid_xml = "<invalid>unclosed tag";
        let result = OdsParser::parse_sheets(invalid_xml);
        // The parser may return Ok with empty sheets or Err depending on implementation
        // Either behavior is acceptable - we just verify it doesn't panic
        match result {
            Ok(sheets) => {
                // If parsing succeeds, we should get 0 sheets
                assert_eq!(sheets.len(), 0);
            },
            Err(_) => {
                // Error is also acceptable
            },
        }
    }

    #[test]
    fn parses_global_and_sheet_local_named_definitions_with_namespace_aliases() {
        let xml = r#"<?xml version="1.0"?>
            <o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                xmlns:f="urn:example:formula">
              <o:body><o:spreadsheet>
                <t:table t:name="Sales &amp; Tax">
                  <t:table-column/><t:table-row><t:table-cell/></t:table-row>
                  <t:named-expressions>
                    <t:named-range t:name="LocalRange"
                      t:cell-range-address="$'Sales &amp; Tax'.$A$1:.$B$2"
                      t:base-cell-address="$'Sales &amp; Tax'.$A$1"
                      t:range-usable-as="print-range filter repeat-row repeat-column"/>
                  </t:named-expressions>
                </t:table>
                <t:named-expressions>
                  <t:named-expression t:name="TaxRate"
                    t:expression="f:=0.2" t:base-cell-address="$'Sales &amp; Tax'.$A$1"/>
                </t:named-expressions>
              </o:spreadsheet></o:body>
            </o:document-content>"#;

        let definitions = OdsParser::parse_named_definitions(xml).unwrap();
        assert_eq!(definitions.len(), 2);
        let NamedDefinition::Range(range) = &definitions[0] else {
            panic!("expected named range");
        };
        assert_eq!(range.name, "LocalRange");
        assert_eq!(
            range.scope,
            NamedDefinitionScope::Sheet("Sales & Tax".to_string())
        );
        assert_eq!(range.usable_as.len(), 4);
        assert_eq!(
            range.base_cell_address.as_deref(),
            Some("$'Sales & Tax'.$A$1")
        );

        let NamedDefinition::Expression(expression) = &definitions[1] else {
            panic!("expected named expression");
        };
        assert_eq!(expression.name, "TaxRate");
        assert_eq!(expression.expression, "f:=0.2");
        assert_eq!(
            expression.formula_namespace.as_ref().unwrap().uri,
            "urn:example:formula"
        );
        assert_eq!(expression.scope, NamedDefinitionScope::Global);
    }

    #[test]
    fn named_definition_parser_rejects_missing_attributes_and_invalid_usage() {
        let missing_address = r#"<office:spreadsheet
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:named-expressions><table:named-range table:name="Broken"/>
            </table:named-expressions></office:spreadsheet>"#;
        assert!(OdsParser::parse_named_definitions(missing_address).is_err());

        let invalid_usage = r#"<office:spreadsheet
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
            <table:named-expressions><table:named-range table:name="Broken"
              table:cell-range-address="$Sheet1.$A$1" table:range-usable-as="chart"/>
            </table:named-expressions></office:spreadsheet>"#;
        assert!(OdsParser::parse_named_definitions(invalid_usage).is_err());
    }
}
