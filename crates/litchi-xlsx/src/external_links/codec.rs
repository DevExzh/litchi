//! Bounded SpreadsheetML external-link XML codec.

use crate::error::{Error, Result};
use crate::raw::namespace::{is_spreadsheetml_name, relationship_attribute_value};
use litchi_ooxml_common::external_link::EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES;
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use litchi_opc::constants::relationship_type as rt;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::*;
use super::{invalid, limit};

impl Link {
    /// Serialize this link as a canonical transitional SpreadsheetML external-link part.
    ///
    /// External targets are represented only as OPC relationship metadata and are never opened.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_xml_with_conformance(Conformance::Transitional)
    }

    /// Serialize this link without dereferencing or executing any target metadata.
    pub fn to_xml_with_conformance(&self, conformance: Conformance) -> Result<Vec<u8>> {
        let has_x14 = matches!(self, Self::Ole(link) if link.items.iter().any(|item| item.source == ItemSource::Office2010));
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="{}""#,
            conformance.sml()
        );
        if matches!(self, Self::Workbook(_) | Self::Ole(_)) {
            xml.push_str(r#" xmlns:r="#);
            xml.push('"');
            xml.push_str(conformance.rel());
            xml.push('"');
        }
        if has_x14 {
            xml.push_str(
                r#" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main""#,
            );
        }
        xml.push('>');
        match self {
            Self::Workbook(link) => write_external_workbook(&mut xml, link)?,
            Self::Dde(link) => write_dde_link(&mut xml, link)?,
            Self::Ole(link) => write_ole_link(&mut xml, link)?,
        }
        xml.push_str("</externalLink>");
        if xml.len() > MAX_CACHE_TEXT_BYTES {
            return Err(limit("serialized XML"));
        }
        Ok(xml.into_bytes())
    }
}

fn validate_external_target(
    target: &Target,
    allowed_types: &[&str],
    description: &str,
) -> Result<()> {
    if target.relationship_id.is_empty() {
        return Err(invalid(format!(
            "{description} relationship ID must not be empty"
        )));
    }
    if target.target.is_empty() {
        return Err(invalid(format!("{description} target must not be empty")));
    }
    if target.target.len() > MAX_EXTERNAL_TARGET_BYTES {
        return Err(limit(&format!("{description} target URI")));
    }
    if target.target.chars().any(|character| {
        character.is_control() || character == '\u{fffe}' || character == '\u{ffff}'
    }) {
        return Err(invalid(format!(
            "{description} target URI contains an invalid character"
        )));
    }
    if target.relationship_id.len() > 1024 || target.relationship_id.chars().any(char::is_control) {
        return Err(invalid(format!("{description} relationship ID is invalid")));
    }
    if !allowed_types.contains(&target.relationship_type.as_str()) {
        return Err(invalid(format!(
            "{description} has invalid relationship type '{}'",
            target.relationship_type
        )));
    }
    Ok(())
}

fn write_external_workbook(xml: &mut String, link: &Workbook) -> Result<()> {
    validate_external_target(
        &link.target,
        EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES,
        "external workbook",
    )?;
    xml.push_str("<externalBook");
    push_xml_attr(xml, "r:id", &link.target.relationship_id)?;
    xml.push('>');
    if link.sheet_names.len() > MAX_SHEET_NAMES {
        return Err(limit("sheet names"));
    }
    if !link.sheet_names.is_empty() {
        xml.push_str("<sheetNames>");
        for name in &link.sheet_names {
            xml.push_str("<sheetName");
            push_xml_attr(xml, "val", name)?;
            xml.push_str("/>");
        }
        xml.push_str("</sheetNames>");
    }
    if link.defined_names.len() > MAX_DEFINED_NAMES {
        return Err(limit("defined names"));
    }
    if !link.defined_names.is_empty() {
        xml.push_str("<definedNames>");
        for name in &link.defined_names {
            xml.push_str("<definedName");
            push_xml_attr(xml, "name", &name.name)?;
            if let Some(value) = &name.refers_to {
                push_xml_attr(xml, "refersTo", value)?;
            }
            if let Some(sheet_id) = name.sheet_id {
                push_u32_attr(xml, "sheetId", sheet_id);
            }
            xml.push_str("/>");
        }
        xml.push_str("</definedNames>");
    }
    if link.cached_sheets.len() > MAX_CACHED_SHEETS {
        return Err(limit("cached sheets"));
    }
    if !link.cached_sheets.is_empty() {
        let mut sheet_ids = std::collections::HashSet::with_capacity(link.cached_sheets.len());
        let mut row_count = 0usize;
        let mut cell_count = 0usize;
        xml.push_str("<sheetDataSet>");
        for sheet in &link.cached_sheets {
            if !sheet_ids.insert(sheet.sheet_id) {
                return Err(invalid(format!(
                    "duplicate external cached sheetId {}",
                    sheet.sheet_id
                )));
            }
            xml.push_str("<sheetData");
            push_u32_attr(xml, "sheetId", sheet.sheet_id);
            if sheet.refresh_error {
                xml.push_str(" refreshError=\"1\"");
            }
            xml.push('>');
            let mut rows = std::collections::HashSet::with_capacity(sheet.rows.len());
            for row in &sheet.rows {
                if row.row == 0 || row.row > 1_048_576 {
                    return Err(invalid("external cached row is outside worksheet bounds"));
                }
                if !rows.insert(row.row) {
                    return Err(invalid(format!(
                        "duplicate external cached row {}",
                        row.row
                    )));
                }
                row_count = row_count
                    .checked_add(1)
                    .ok_or_else(|| limit("cached rows"))?;
                if row_count > MAX_CACHED_ROWS {
                    return Err(limit("cached rows"));
                }
                xml.push_str("<row");
                push_u32_attr(xml, "r", row.row);
                xml.push('>');
                let mut references = std::collections::HashSet::with_capacity(row.cells.len());
                for cell in &row.cells {
                    cell_count = cell_count
                        .checked_add(1)
                        .ok_or_else(|| limit("cached cells"))?;
                    if cell_count > MAX_CACHED_CELLS {
                        return Err(limit("cached cells"));
                    }
                    if let Some(reference) = &cell.reference {
                        validate_cell_reference(reference)?;
                        if !references.insert(reference.as_str()) {
                            return Err(invalid(format!(
                                "duplicate external cached cell '{reference}'"
                            )));
                        }
                    }
                    xml.push_str("<cell");
                    if let Some(reference) = &cell.reference {
                        push_xml_attr(xml, "r", reference)?;
                    }
                    if cell.cell_type != CellType::Number {
                        push_xml_attr(xml, "t", external_cell_type_token(cell.cell_type))?;
                    }
                    if cell.value_metadata_index != 0 {
                        push_u32_attr(xml, "vm", cell.value_metadata_index);
                    }
                    if let Some(value) = &cell.raw_value {
                        xml.push_str("><v>");
                        push_xml_text(xml, value)?;
                        xml.push_str("</v></cell>");
                    } else {
                        xml.push_str("/>");
                    }
                }
                xml.push_str("</row>");
            }
            xml.push_str("</sheetData>");
        }
        xml.push_str("</sheetDataSet>");
    }
    xml.push_str("</externalBook>");
    Ok(())
}

fn write_dde_link(xml: &mut String, link: &Dde) -> Result<()> {
    xml.push_str("<ddeLink");
    push_xml_attr(xml, "ddeService", &link.service)?;
    push_xml_attr(xml, "ddeTopic", &link.topic)?;
    if link.items.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }
    if link.items.len() > MAX_LINK_ITEMS {
        return Err(limit("DDE items"));
    }
    xml.push_str("><ddeItems>");
    for item in &link.items {
        xml.push_str("<ddeItem");
        if let Some(name) = &item.name {
            push_xml_attr(xml, "name", name)?;
        }
        push_true_attr(xml, "ole", item.use_ole);
        push_true_attr(xml, "advise", item.advise);
        push_true_attr(xml, "preferPic", item.prefer_picture);
        if let Some(values) = &item.values {
            xml.push('>');
            write_dde_values(xml, "values", values)?;
            xml.push_str("</ddeItem>");
        } else {
            xml.push_str("/>");
        }
    }
    xml.push_str("</ddeItems></ddeLink>");
    Ok(())
}

fn write_ole_link(xml: &mut String, link: &Ole) -> Result<()> {
    validate_external_target(
        &link.target,
        &[rt::OLE_OBJECT, rt::STRICT_OLE_OBJECT],
        "OLE",
    )?;
    if link.program_id.is_empty() {
        return Err(invalid("OLE program ID must not be empty"));
    }
    xml.push_str("<oleLink");
    push_xml_attr(xml, "r:id", &link.target.relationship_id)?;
    push_xml_attr(xml, "progId", &link.program_id)?;
    if link.items.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }
    if link.items.len() > MAX_LINK_ITEMS {
        return Err(limit("OLE items"));
    }
    xml.push_str("><oleItems>");
    for item in &link.items {
        if item.name.is_empty() {
            return Err(invalid("OLE item name must not be empty"));
        }
        let element = if item.source == ItemSource::Office2010 {
            "x14:oleItem"
        } else {
            "oleItem"
        };
        xml.push('<');
        xml.push_str(element);
        push_xml_attr(xml, "name", &item.name)?;
        push_true_attr(xml, "icon", item.icon);
        push_true_attr(xml, "advise", item.advise);
        push_true_attr(xml, "preferPic", item.prefer_picture);
        match (&item.values, item.source) {
            (Some(values), ItemSource::Office2010) => {
                xml.push('>');
                write_dde_values(xml, "x14:values", values)?;
                xml.push_str("</x14:oleItem>");
            },
            (Some(_), ItemSource::SpreadsheetMl) => {
                return Err(invalid("cached OLE values require an Office 2010 oleItem"));
            },
            (None, _) => xml.push_str("/>"),
        }
    }
    xml.push_str("</oleItems></oleLink>");
    Ok(())
}

fn write_dde_values(xml: &mut String, element: &str, values: &DdeValues) -> Result<()> {
    let expected = u64::from(values.rows)
        .checked_mul(u64::from(values.columns))
        .ok_or_else(|| limit("DDE/OLE matrix dimensions"))?;
    if expected == 0 || expected > MAX_CACHED_CELLS as u64 {
        return Err(limit("DDE/OLE matrix dimensions"));
    }
    if expected != values.values.len() as u64 {
        return Err(invalid(format!(
            "DDE/OLE matrix declares {expected} values but contains {}",
            values.values.len()
        )));
    }
    xml.push('<');
    xml.push_str(element);
    push_u32_attr(xml, "rows", values.rows);
    push_u32_attr(xml, "cols", values.columns);
    xml.push('>');
    for value in &values.values {
        xml.push_str("<value");
        if value.value_type != DdeValueType::Nil {
            push_xml_attr(xml, "t", dde_value_type_token(value.value_type))?;
        }
        xml.push_str("><val>");
        push_xml_text(xml, &value.raw_value)?;
        xml.push_str("</val></value>");
    }
    xml.push_str("</");
    xml.push_str(element);
    xml.push('>');
    Ok(())
}

fn external_cell_type_token(value: CellType) -> &'static str {
    match value {
        CellType::Number => "n",
        CellType::Boolean => "b",
        CellType::Date => "d",
        CellType::Error => "e",
        CellType::InlineString => "inlineStr",
        CellType::SharedString => "s",
        CellType::String => "str",
    }
}

fn dde_value_type_token(value: DdeValueType) -> &'static str {
    match value {
        DdeValueType::Nil => "nil",
        DdeValueType::Boolean => "b",
        DdeValueType::Number => "n",
        DdeValueType::Error => "e",
        DdeValueType::String => "str",
    }
}

fn push_true_attr(xml: &mut String, name: &str, value: bool) {
    if value {
        xml.push(' ');
        xml.push_str(name);
        xml.push_str("=\"1\"");
    }
}

fn push_u32_attr(xml: &mut String, name: &str, value: u32) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    xml.push_str(&value.to_string());
    xml.push('"');
}

fn push_xml_attr(xml: &mut String, name: &str, value: &str) -> Result<()> {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    push_xml_text(xml, value)?;
    xml.push('"');
    Ok(())
}

fn push_xml_text(xml: &mut String, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            '\t' | '\n' | '\r' => xml.push(character),
            value if value >= '\u{20}' && value != '\u{fffe}' && value != '\u{ffff}' => {
                xml.push(value)
            },
            value => {
                return Err(invalid(format!(
                    "invalid XML character U+{:04X} in external link",
                    value as u32
                )));
            },
        }
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    ExternalBook,
    DdeLink,
    DdeItems,
    DdeItem(usize),
    DdeValues(usize),
    DdeValue(usize, usize),
    DdeValueText(usize, usize),
    OleLink,
    OleItems,
    OleItem(usize),
    OleValues(usize),
    OleValue(usize, usize),
    OleValueText(usize, usize),
    SheetNames,
    DefinedNames,
    SheetDataSet,
    SheetData(usize),
    Row(usize, usize),
    Cell(usize, usize, usize),
    Value(usize, usize, usize),
    Other,
}

enum ParsedLink {
    Workbook(ParsedExternalBook),
    Dde(ParsedDdeLink),
    Ole(ParsedOleLink),
}

struct ParsedDdeLink {
    service: String,
    topic: String,
    items: Vec<ParsedDdeItem>,
    saw_items: bool,
}

struct ParsedOleLink {
    target_relationship_id: String,
    program_id: String,
    items: Vec<ParsedOleItem>,
    saw_items: bool,
}

struct ParsedDdeItem {
    name: Option<String>,
    use_ole: bool,
    advise: bool,
    prefer_picture: bool,
    values: Option<ParsedDdeValues>,
}

struct ParsedOleItem {
    source: ItemSource,
    name: String,
    icon: bool,
    advise: bool,
    prefer_picture: bool,
    values: Option<ParsedDdeValues>,
}

struct ParsedDdeValues {
    rows: u32,
    columns: u32,
    values: Vec<ParsedDdeValue>,
}

struct ParsedDdeValue {
    value_type: DdeValueType,
    raw_value: Option<String>,
}

struct ParsedExternalBook {
    target_relationship_id: String,
    sheet_names: Vec<String>,
    defined_names: Vec<DefinedName>,
    cached_sheets: Vec<SheetData>,
    saw_sheet_names: bool,
    saw_defined_names: bool,
    saw_sheet_data_set: bool,
}

struct Parser {
    kind: Option<ParsedLink>,
    cached_rows: usize,
    cached_cells: usize,
    text_bytes: usize,
}

impl Parser {
    fn new() -> Self {
        Self {
            kind: None,
            cached_rows: 0,
            cached_cells: 0,
            text_bytes: 0,
        }
    }

    fn book_mut(&mut self) -> Result<&mut ParsedExternalBook> {
        match self.kind.as_mut() {
            Some(ParsedLink::Workbook(book)) => Ok(book),
            _ => Err(invalid("external-link content is outside externalBook")),
        }
    }

    fn dde_mut(&mut self) -> Result<&mut ParsedDdeLink> {
        match self.kind.as_mut() {
            Some(ParsedLink::Dde(link)) => Ok(link),
            _ => Err(invalid("DDE content is outside ddeLink")),
        }
    }

    fn ole_mut(&mut self) -> Result<&mut ParsedOleLink> {
        match self.kind.as_mut() {
            Some(ParsedLink::Ole(link)) => Ok(link),
            _ => Err(invalid("OLE content is outside oleLink")),
        }
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Context> {
        if parent == Context::Root
            && is_spreadsheetml_name(namespace, element.name(), b"externalBook")
        {
            if self.kind.is_some() {
                return Err(invalid("externalLink has multiple link kinds"));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| invalid("externalBook is missing relationship ID"))?;
            self.kind = Some(ParsedLink::Workbook(ParsedExternalBook {
                target_relationship_id: relationship_id,
                sheet_names: Vec::new(),
                defined_names: Vec::new(),
                cached_sheets: Vec::new(),
                saw_sheet_names: false,
                saw_defined_names: false,
                saw_sheet_data_set: false,
            }));
            return Ok(Context::ExternalBook);
        }
        if parent == Context::Root && is_spreadsheetml_name(namespace, element.name(), b"ddeLink") {
            if self.kind.is_some() {
                return Err(invalid("externalLink has multiple link kinds"));
            }
            let service = required_attr(element, b"ddeService", decoder, "DDE service")?;
            let topic = required_attr(element, b"ddeTopic", decoder, "DDE topic")?;
            self.add_text(service.len() + topic.len())?;
            self.kind = Some(ParsedLink::Dde(ParsedDdeLink {
                service,
                topic,
                items: Vec::new(),
                saw_items: false,
            }));
            return Ok(Context::DdeLink);
        }
        if parent == Context::Root && is_spreadsheetml_name(namespace, element.name(), b"oleLink") {
            if self.kind.is_some() {
                return Err(invalid("externalLink has multiple link kinds"));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| invalid("oleLink is missing relationship ID"))?;
            let program_id = required_attr(element, b"progId", decoder, "OLE program ID")?;
            if program_id.is_empty() {
                return Err(invalid("OLE program ID must not be empty"));
            }
            self.add_text(relationship_id.len() + program_id.len())?;
            self.kind = Some(ParsedLink::Ole(ParsedOleLink {
                target_relationship_id: relationship_id,
                program_id,
                items: Vec::new(),
                saw_items: false,
            }));
            return Ok(Context::OleLink);
        }
        if parent == Context::DdeLink
            && is_spreadsheetml_name(namespace, element.name(), b"ddeItems")
        {
            let link = self.dde_mut()?;
            mark_once(&mut link.saw_items, "DDE items")?;
            return Ok(Context::DdeItems);
        }
        if parent == Context::DdeItems
            && is_spreadsheetml_name(namespace, element.name(), b"ddeItem")
        {
            let name = unqualified_attribute_value(element, b"name", decoder)?;
            let use_ole = optional_bool(element, b"ole", decoder, "DDE item ole")?.unwrap_or(false);
            let advise =
                optional_bool(element, b"advise", decoder, "DDE item advise")?.unwrap_or(false);
            let prefer_picture =
                optional_bool(element, b"preferPic", decoder, "DDE item preferPic")?
                    .unwrap_or(false);
            self.add_text(name.as_ref().map_or(0, String::len))?;
            let link = self.dde_mut()?;
            if link.items.len() >= MAX_LINK_ITEMS {
                return Err(limit("DDE items"));
            }
            let index = link.items.len();
            link.items.push(ParsedDdeItem {
                name,
                use_ole,
                advise,
                prefer_picture,
                values: None,
            });
            return Ok(Context::DdeItem(index));
        }
        if let Context::DdeItem(item) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"values")
        {
            let rows = optional_u32(element, b"rows", decoder, "DDE value rows")?.unwrap_or(1);
            let columns =
                optional_u32(element, b"cols", decoder, "DDE value columns")?.unwrap_or(1);
            let target = &mut self
                .dde_mut()?
                .items
                .get_mut(item)
                .ok_or_else(|| invalid("invalid DDE item context"))?
                .values;
            if target
                .replace(ParsedDdeValues {
                    rows,
                    columns,
                    values: Vec::new(),
                })
                .is_some()
            {
                return Err(invalid("duplicate DDE item values"));
            }
            return Ok(Context::DdeValues(item));
        }
        if let Context::DdeValues(item) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"value")
        {
            let value_type = parse_dde_value_type(
                unqualified_attribute_value(element, b"t", decoder)?.as_deref(),
            )?;
            self.cached_cells = self
                .cached_cells
                .checked_add(1)
                .ok_or_else(|| limit("DDE values"))?;
            if self.cached_cells > MAX_CACHED_CELLS {
                return Err(limit("DDE values"));
            }
            let values = &mut self
                .dde_mut()?
                .items
                .get_mut(item)
                .and_then(|item| item.values.as_mut())
                .ok_or_else(|| invalid("invalid DDE values context"))?
                .values;
            let index = values.len();
            values.push(ParsedDdeValue {
                value_type,
                raw_value: None,
            });
            return Ok(Context::DdeValue(item, index));
        }
        if let Context::DdeValue(item, value) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"val")
        {
            let parsed_value = self
                .dde_mut()?
                .items
                .get_mut(item)
                .and_then(|item| item.values.as_mut())
                .and_then(|values| values.values.get_mut(value))
                .ok_or_else(|| invalid("invalid DDE value context"))?;
            if parsed_value.raw_value.replace(String::new()).is_some() {
                return Err(invalid("duplicate DDE val element"));
            }
            return Ok(Context::DdeValueText(item, value));
        }
        if parent == Context::OleLink
            && is_spreadsheetml_name(namespace, element.name(), b"oleItems")
        {
            let link = self.ole_mut()?;
            mark_once(&mut link.saw_items, "OLE items")?;
            return Ok(Context::OleItems);
        }
        if parent == Context::OleItems
            && (is_spreadsheetml_name(namespace, element.name(), b"oleItem")
                || is_exact_name(namespace, element.name(), X14, b"oleItem"))
        {
            let source = if is_exact_name(namespace, element.name(), X14, b"oleItem") {
                ItemSource::Office2010
            } else {
                ItemSource::SpreadsheetMl
            };
            let name = required_attr(element, b"name", decoder, "OLE item name")?;
            if name.is_empty() {
                return Err(invalid("OLE item name must not be empty"));
            }
            let icon = optional_bool(element, b"icon", decoder, "OLE item icon")?.unwrap_or(false);
            let advise =
                optional_bool(element, b"advise", decoder, "OLE item advise")?.unwrap_or(false);
            let prefer_picture =
                optional_bool(element, b"preferPic", decoder, "OLE item preferPic")?
                    .unwrap_or(false);
            self.add_text(name.len())?;
            let link = self.ole_mut()?;
            if link.items.len() >= MAX_LINK_ITEMS {
                return Err(limit("OLE items"));
            }
            let index = link.items.len();
            link.items.push(ParsedOleItem {
                source,
                name,
                icon,
                advise,
                prefer_picture,
                values: None,
            });
            return Ok(Context::OleItem(index));
        }
        if let Context::OleItem(item) = parent
            && is_exact_name(namespace, element.name(), X14, b"values")
        {
            if self.ole_mut()?.items[item].source != ItemSource::Office2010 {
                return Err(invalid("cached OLE values require an Office 2010 oleItem"));
            }
            let rows = optional_u32(element, b"rows", decoder, "OLE value rows")?.unwrap_or(1);
            let columns =
                optional_u32(element, b"cols", decoder, "OLE value columns")?.unwrap_or(1);
            let target = &mut self.ole_mut()?.items[item].values;
            if target
                .replace(ParsedDdeValues {
                    rows,
                    columns,
                    values: Vec::new(),
                })
                .is_some()
            {
                return Err(invalid("duplicate OLE item values"));
            }
            return Ok(Context::OleValues(item));
        }
        if let Context::OleValues(item) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"value")
        {
            let value_type = parse_dde_value_type(
                unqualified_attribute_value(element, b"t", decoder)?.as_deref(),
            )?;
            self.cached_cells = self
                .cached_cells
                .checked_add(1)
                .ok_or_else(|| limit("OLE values"))?;
            if self.cached_cells > MAX_CACHED_CELLS {
                return Err(limit("OLE values"));
            }
            let values = &mut self
                .ole_mut()?
                .items
                .get_mut(item)
                .and_then(|item| item.values.as_mut())
                .ok_or_else(|| invalid("invalid OLE values context"))?
                .values;
            let index = values.len();
            values.push(ParsedDdeValue {
                value_type,
                raw_value: None,
            });
            return Ok(Context::OleValue(item, index));
        }
        if let Context::OleValue(item, value) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"val")
        {
            let parsed_value = self
                .ole_mut()?
                .items
                .get_mut(item)
                .and_then(|item| item.values.as_mut())
                .and_then(|values| values.values.get_mut(value))
                .ok_or_else(|| invalid("invalid OLE value context"))?;
            if parsed_value.raw_value.replace(String::new()).is_some() {
                return Err(invalid("duplicate OLE val element"));
            }
            return Ok(Context::OleValueText(item, value));
        }
        if parent == Context::ExternalBook
            && is_spreadsheetml_name(namespace, element.name(), b"sheetNames")
        {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_sheet_names, "external sheetNames")?;
            return Ok(Context::SheetNames);
        }
        if parent == Context::ExternalBook
            && is_spreadsheetml_name(namespace, element.name(), b"definedNames")
        {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_defined_names, "external definedNames")?;
            return Ok(Context::DefinedNames);
        }
        if parent == Context::ExternalBook
            && is_spreadsheetml_name(namespace, element.name(), b"sheetDataSet")
        {
            let book = self.book_mut()?;
            mark_once(&mut book.saw_sheet_data_set, "external sheetDataSet")?;
            return Ok(Context::SheetDataSet);
        }
        if parent == Context::SheetNames
            && is_spreadsheetml_name(namespace, element.name(), b"sheetName")
        {
            let value = unqualified_attribute_value(element, b"val", decoder)?.unwrap_or_default();
            self.add_text(value.len())?;
            let book = self.book_mut()?;
            if book.sheet_names.len() >= MAX_SHEET_NAMES {
                return Err(limit("sheet names"));
            }
            book.sheet_names.push(value);
            return Ok(Context::Other);
        }
        if parent == Context::DefinedNames
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            let name = required_attr(element, b"name", decoder, "external defined name")?;
            let refers_to = unqualified_attribute_value(element, b"refersTo", decoder)?;
            let sheet_id = optional_u32(
                element,
                b"sheetId",
                decoder,
                "external defined-name sheetId",
            )?;
            self.add_text(name.len() + refers_to.as_ref().map_or(0, String::len))?;
            let book = self.book_mut()?;
            if book.defined_names.len() >= MAX_DEFINED_NAMES {
                return Err(limit("defined names"));
            }
            book.defined_names.push(DefinedName {
                name,
                refers_to,
                sheet_id,
            });
            return Ok(Context::Other);
        }
        if parent == Context::SheetDataSet
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            let sheet_id = required_u32(element, b"sheetId", decoder, "external cached sheetId")?;
            let refresh_error =
                optional_bool(element, b"refreshError", decoder, "external refreshError")?
                    .unwrap_or(false);
            let book = self.book_mut()?;
            if book.cached_sheets.len() >= MAX_CACHED_SHEETS {
                return Err(limit("cached sheets"));
            }
            if book
                .cached_sheets
                .iter()
                .any(|sheet| sheet.sheet_id == sheet_id)
            {
                return Err(invalid(format!(
                    "duplicate external cached sheetId {sheet_id}"
                )));
            }
            let index = book.cached_sheets.len();
            book.cached_sheets.push(SheetData {
                sheet_id,
                refresh_error,
                rows: Vec::new(),
            });
            return Ok(Context::SheetData(index));
        }
        if let Context::SheetData(sheet) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            let row = required_u32(element, b"r", decoder, "external cached row")?;
            if row == 0 {
                return Err(invalid("external cached row must be positive"));
            }
            self.cached_rows = self
                .cached_rows
                .checked_add(1)
                .ok_or_else(|| limit("cached rows"))?;
            if self.cached_rows > MAX_CACHED_ROWS {
                return Err(limit("cached rows"));
            }
            let book = self.book_mut()?;
            let rows = &mut book.cached_sheets[sheet].rows;
            if rows.iter().any(|item| item.row == row) {
                return Err(invalid(format!("duplicate external cached row {row}")));
            }
            let index = rows.len();
            rows.push(Row {
                row,
                cells: Vec::new(),
            });
            return Ok(Context::Row(sheet, index));
        }
        if let Context::Row(sheet, row) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"cell")
        {
            let reference = unqualified_attribute_value(element, b"r", decoder)?;
            if let Some(value) = reference.as_deref() {
                validate_cell_reference(value)?;
            }
            let cell_type =
                parse_cell_type(unqualified_attribute_value(element, b"t", decoder)?.as_deref())?;
            let value_metadata_index =
                optional_u32(element, b"vm", decoder, "external value metadata index")?
                    .unwrap_or(0);
            self.cached_cells = self
                .cached_cells
                .checked_add(1)
                .ok_or_else(|| limit("cached cells"))?;
            if self.cached_cells > MAX_CACHED_CELLS {
                return Err(limit("cached cells"));
            }
            let book = self.book_mut()?;
            let cells = &mut book.cached_sheets[sheet].rows[row].cells;
            if let Some(reference) = reference.as_deref()
                && cells
                    .iter()
                    .any(|cell| cell.reference.as_deref() == Some(reference))
            {
                return Err(invalid(format!(
                    "duplicate external cached cell '{reference}'"
                )));
            }
            let index = cells.len();
            cells.push(Cell {
                reference,
                cell_type,
                raw_value: None,
                value_metadata_index,
            });
            return Ok(Context::Cell(sheet, row, index));
        }
        if let Context::Cell(sheet, row, cell) = parent
            && is_spreadsheetml_name(namespace, element.name(), b"v")
        {
            return Ok(Context::Value(sheet, row, cell));
        }
        if matches!(
            parent,
            Context::DdeLink
                | Context::DdeItems
                | Context::DdeItem(_)
                | Context::DdeValues(_)
                | Context::DdeValue(_, _)
                | Context::OleLink
                | Context::OleItems
                | Context::OleItem(_)
                | Context::OleValues(_)
                | Context::OleValue(_, _)
        ) {
            return Err(invalid("unexpected child in DDE/OLE external link"));
        }
        Ok(Context::Other)
    }

    fn push_value(&mut self, sheet: usize, row: usize, cell: usize, value: &str) -> Result<()> {
        self.add_text(value.len())?;
        let target = &mut self.book_mut()?.cached_sheets[sheet].rows[row].cells[cell].raw_value;
        target.get_or_insert_with(String::new).push_str(value);
        Ok(())
    }

    fn push_dde_value(&mut self, item: usize, value: usize, text: &str) -> Result<()> {
        self.add_text(text.len())?;
        let value = self
            .dde_mut()?
            .items
            .get_mut(item)
            .and_then(|item| item.values.as_mut())
            .and_then(|values| values.values.get_mut(value))
            .ok_or_else(|| invalid("invalid DDE value context"))?;
        value
            .raw_value
            .as_mut()
            .ok_or_else(|| invalid("invalid DDE value text context"))?
            .push_str(text);
        Ok(())
    }

    fn push_ole_value(&mut self, item: usize, value: usize, text: &str) -> Result<()> {
        self.add_text(text.len())?;
        let value = self
            .ole_mut()?
            .items
            .get_mut(item)
            .and_then(|item| item.values.as_mut())
            .and_then(|values| values.values.get_mut(value))
            .ok_or_else(|| invalid("invalid OLE value context"))?;
        value
            .raw_value
            .as_mut()
            .ok_or_else(|| invalid("invalid OLE value text context"))?
            .push_str(text);
        Ok(())
    }

    fn add_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self
            .text_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("cache text"))?;
        if self.text_bytes > MAX_CACHE_TEXT_BYTES {
            return Err(limit("cache text"));
        }
        Ok(())
    }
}

pub fn parse_external_link(xml: &[u8]) -> Result<Link> {
    let xml = litchi_ooxml_common::mce::process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut parser = Parser::new();
    let mut stack = Vec::new();
    let mut closed_root = false;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Invalid(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if stack.is_empty() => {
                if closed_root
                    || !is_spreadsheetml_name(&namespace, element.name(), b"externalLink")
                {
                    return Err(invalid(
                        "external-link XML must have one SpreadsheetML externalLink root",
                    ));
                }
                stack.push(Context::Root);
            },
            Event::Empty(element) if stack.is_empty() => {
                if closed_root
                    || !is_spreadsheetml_name(&namespace, element.name(), b"externalLink")
                {
                    return Err(invalid(
                        "external-link XML must have one SpreadsheetML externalLink root",
                    ));
                }
                return Err(invalid("externalLink must contain a link kind"));
            },
            Event::Start(element) => {
                let parent = *stack
                    .last()
                    .ok_or_else(|| invalid("external-link XML is missing its root"))?;
                let context = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                stack.push(context);
            },
            Event::Empty(element) => {
                let parent = *stack
                    .last()
                    .ok_or_else(|| invalid("external-link XML is missing its root"))?;
                parser.start(parent, &namespace, &element, decoder, &resolver)?;
            },
            Event::Text(text) => {
                let text = text.decode().map_err(|e| Error::Invalid(e.to_string()))?;
                push_context_text(&mut parser, stack.last().copied(), &text)?;
            },
            Event::CData(text) => {
                let text = text.decode().map_err(|e| Error::Invalid(e.to_string()))?;
                push_context_text(&mut parser, stack.last().copied(), &text)?;
            },
            Event::GeneralRef(reference) => {
                let text = decode_xml_reference(&reference)?;
                push_context_text(&mut parser, stack.last().copied(), &text)?;
            },
            Event::End(element) => {
                let context = stack
                    .pop()
                    .ok_or_else(|| invalid("external-link closing element outside root"))?;
                if context == Context::Root {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"externalLink") {
                        return Err(invalid(
                            "external-link XML has invalid root closing element",
                        ));
                    }
                    closed_root = true;
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => {
                return Err(invalid("external-link XML has an unterminated root"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    match parser
        .kind
        .ok_or_else(|| invalid("externalLink must contain a link kind"))?
    {
        ParsedLink::Workbook(book) => Ok(Link::Workbook(Workbook {
            target: Target {
                relationship_id: book.target_relationship_id,
                target: String::new(),
                relationship_type: String::new(),
            },
            sheet_names: book.sheet_names,
            defined_names: book.defined_names,
            cached_sheets: book.cached_sheets,
        })),
        ParsedLink::Dde(link) => {
            if link.saw_items && link.items.is_empty() {
                return Err(invalid("ddeItems must contain at least one ddeItem"));
            }
            Ok(Link::Dde(Dde {
                service: link.service,
                topic: link.topic,
                items: link
                    .items
                    .into_iter()
                    .map(finalize_dde_item)
                    .collect::<Result<_>>()?,
            }))
        },
        ParsedLink::Ole(link) => {
            if link.saw_items && link.items.is_empty() {
                return Err(invalid("oleItems must contain at least one oleItem"));
            }
            Ok(Link::Ole(Ole {
                target: Target {
                    relationship_id: link.target_relationship_id,
                    target: String::new(),
                    relationship_type: String::new(),
                },
                program_id: link.program_id,
                items: link
                    .items
                    .into_iter()
                    .map(finalize_ole_item)
                    .collect::<Result<_>>()?,
            }))
        },
    }
}

fn push_context_text(parser: &mut Parser, context: Option<Context>, text: &str) -> Result<()> {
    match context {
        Some(Context::Value(sheet, row, cell)) => parser.push_value(sheet, row, cell, text),
        Some(Context::DdeValueText(item, value)) => parser.push_dde_value(item, value, text),
        Some(Context::OleValueText(item, value)) => parser.push_ole_value(item, value, text),
        Some(
            Context::DdeLink
            | Context::DdeItems
            | Context::DdeItem(_)
            | Context::DdeValues(_)
            | Context::DdeValue(_, _)
            | Context::OleLink
            | Context::OleItems
            | Context::OleItem(_)
            | Context::OleValues(_)
            | Context::OleValue(_, _),
        ) if !text.trim().is_empty() => Err(invalid("unexpected text in DDE/OLE external link")),
        _ => Ok(()),
    }
}

fn finalize_dde_item(item: ParsedDdeItem) -> Result<DdeItem> {
    Ok(DdeItem {
        name: item.name,
        use_ole: item.use_ole,
        advise: item.advise,
        prefer_picture: item.prefer_picture,
        values: item.values.map(finalize_dde_values).transpose()?,
    })
}

fn finalize_ole_item(item: ParsedOleItem) -> Result<OleItem> {
    Ok(OleItem {
        source: item.source,
        name: item.name,
        icon: item.icon,
        advise: item.advise,
        prefer_picture: item.prefer_picture,
        values: item.values.map(finalize_dde_values).transpose()?,
    })
}

fn finalize_dde_values(values: ParsedDdeValues) -> Result<DdeValues> {
    let ParsedDdeValues {
        rows,
        columns,
        values,
    } = values;
    let expected = u64::from(rows)
        .checked_mul(u64::from(columns))
        .ok_or_else(|| limit("DDE/OLE matrix dimensions"))?;
    if expected > MAX_CACHED_CELLS as u64 {
        return Err(limit("DDE/OLE matrix dimensions"));
    }
    if expected != values.len() as u64 {
        return Err(invalid(format!(
            "DDE/OLE matrix declares {expected} values but contains {}",
            values.len()
        )));
    }
    let values = values
        .into_iter()
        .map(|value| {
            Ok(DdeValue {
                value_type: value.value_type,
                raw_value: value
                    .raw_value
                    .ok_or_else(|| invalid("DDE/OLE value is missing val child"))?,
            })
        })
        .collect::<Result<_>>()?;
    Ok(DdeValues {
        rows,
        columns,
        values,
    })
}

fn parse_dde_value_type(value: Option<&str>) -> Result<DdeValueType> {
    match value.unwrap_or("nil") {
        "nil" => Ok(DdeValueType::Nil),
        "b" => Ok(DdeValueType::Boolean),
        "n" => Ok(DdeValueType::Number),
        "e" => Ok(DdeValueType::Error),
        "str" => Ok(DdeValueType::String),
        value => Err(invalid(format!("invalid DDE value type '{value}'"))),
    }
}

fn is_exact_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    uri: &[u8],
    local: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == uri)
        && name.local_name().as_ref() == local
}

fn parse_cell_type(value: Option<&str>) -> Result<CellType> {
    match value.unwrap_or("n") {
        "n" => Ok(CellType::Number),
        "b" => Ok(CellType::Boolean),
        "d" => Ok(CellType::Date),
        "e" => Ok(CellType::Error),
        "inlineStr" => Ok(CellType::InlineString),
        "s" => Ok(CellType::SharedString),
        "str" => Ok(CellType::String),
        value => Err(invalid(format!(
            "invalid external cached cell type '{value}'"
        ))),
    }
}

/// Validate an `A1`-style cached cell reference.
///
/// Accepts the full SpreadsheetML address space, so multi-letter columns
/// (`AA1` through `XFD1048576`) are valid. The column prefix is bounded to
/// [`MAX_COLUMN_LETTERS`] so a long letter run cannot drive the accumulator.
fn validate_cell_reference(value: &str) -> Result<()> {
    let malformed = || invalid(format!("invalid external cached cell reference '{value}'"));

    let mut column = 0u32;
    let mut split = 0usize;
    for byte in value.bytes() {
        if !byte.is_ascii_alphabetic() {
            break;
        }
        if split == MAX_COLUMN_LETTERS {
            return Err(malformed());
        }
        column = column
            .checked_mul(26)
            .and_then(|n| n.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid("external cell column overflow"))?;
        split += 1;
    }
    if split == 0 || split == value.len() || column > MAX_CELL_COLUMN {
        return Err(malformed());
    }
    let row = value[split..].parse::<u32>().map_err(|_| malformed())?;
    if row == 0 || row > MAX_CELL_ROW {
        return Err(malformed());
    }
    Ok(())
}

fn required_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}
fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}
fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}
fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" | "on" => Ok(true),
            "0" | "false" | "off" => Ok(false),
            _ => Err(invalid(format!("invalid {description} '{value}'"))),
        })
        .transpose()
}
fn mark_once(seen: &mut bool, description: &str) -> Result<()> {
    if std::mem::replace(seen, true) {
        Err(invalid(format!("duplicate {description}")))
    } else {
        Ok(())
    }
}
