//! Compatibility facade for the canonical XLSX external-link owner.
//!
//! The typed SpreadsheetML external-link model, bounded externalLink XML
//! codec, cached DDE/OLE/workbook values, and inert relationship validation
//! live in litchi_xlsx::external_links. This module retains the historical
//! OOXML-host data names and OoxmlError boundary while delegating all codec
//! work to that owner. External targets remain opaque metadata and are never
//! opened, fetched, executed, or dereferenced.

use crate::error::{OoxmlError, Result};
use litchi_opc::part::BlobPart;
use litchi_opc::{PackURI, Part};
use litchi_xlsx::external_links as owner;

#[cfg(test)]
use litchi_opc::constants::relationship_type as rt;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ExternalLinkConformance {
    #[default]
    Transitional,
    Strict,
}

impl ExternalLinkConformance {
    fn into_owner(self) -> owner::ExternalLinkConformance {
        match self {
            Self::Transitional => owner::ExternalLinkConformance::Transitional,
            Self::Strict => owner::ExternalLinkConformance::Strict,
        }
    }

    pub(crate) fn external_link_relationship(self) -> &'static str {
        self.into_owner().external_link_relationship()
    }
}

impl From<owner::ExternalLinkConformance> for ExternalLinkConformance {
    fn from(value: owner::ExternalLinkConformance) -> Self {
        match value {
            owner::ExternalLinkConformance::Transitional => Self::Transitional,
            owner::ExternalLinkConformance::Strict => Self::Strict,
        }
    }
}

impl From<ExternalLinkConformance> for owner::ExternalLinkConformance {
    fn from(value: ExternalLinkConformance) -> Self {
        value.into_owner()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalLinkEntry {
    pub index: u32,
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub kind: ExternalLinkKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExternalLinkKind {
    Workbook(ExternalWorkbookLink),
    Dde(ExternalDdeLink),
    Ole(ExternalOleLink),
}

impl ExternalLinkKind {
    fn into_owner(self) -> owner::ExternalLinkKind {
        self.into()
    }

    /// Serialize this link as a canonical transitional SpreadsheetML part.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        owner::ExternalLinkKind::to_xml(&self.clone().into_owner()).map_err(map_owner_error)
    }

    /// Serialize this link using the requested SpreadsheetML conformance.
    pub fn to_xml_with_conformance(&self, conformance: ExternalLinkConformance) -> Result<Vec<u8>> {
        owner::ExternalLinkKind::to_xml_with_conformance(
            &self.clone().into_owner(),
            conformance.into_owner(),
        )
        .map_err(map_owner_error)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeLink {
    pub service: String,
    pub topic: String,
    pub items: Vec<ExternalDdeItem>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalOleLink {
    pub target: ExternalOleTarget,
    pub program_id: String,
    pub items: Vec<ExternalOleItem>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalOleTarget {
    pub relationship_id: String,
    pub target: String,
    pub relationship_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeItem {
    pub name: Option<String>,
    pub use_ole: bool,
    pub advise: bool,
    pub prefer_picture: bool,
    pub values: Option<ExternalDdeValues>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalOleItemSource {
    SpreadsheetMl,
    Office2010,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalOleItem {
    pub source: ExternalOleItemSource,
    pub name: String,
    pub icon: bool,
    pub advise: bool,
    pub prefer_picture: bool,
    pub values: Option<ExternalDdeValues>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeValues {
    pub rows: u32,
    pub columns: u32,
    pub values: Vec<ExternalDdeValue>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalDdeValueType {
    Nil,
    Boolean,
    Number,
    Error,
    String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDdeValue {
    pub value_type: ExternalDdeValueType,
    pub raw_value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalWorkbookLink {
    pub target: ExternalWorkbookTarget,
    pub sheet_names: Vec<String>,
    pub defined_names: Vec<ExternalDefinedName>,
    pub cached_sheets: Vec<ExternalSheetData>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalWorkbookTarget {
    pub relationship_id: String,
    pub target: String,
    pub relationship_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalDefinedName {
    pub name: String,
    pub refers_to: Option<String>,
    pub sheet_id: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalSheetData {
    pub sheet_id: u32,
    pub refresh_error: bool,
    pub rows: Vec<ExternalRow>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalRow {
    pub row: u32,
    pub cells: Vec<ExternalCell>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExternalCellType {
    Number,
    Boolean,
    Date,
    Error,
    InlineString,
    SharedString,
    String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalCell {
    pub reference: Option<String>,
    pub cell_type: ExternalCellType,
    pub raw_value: Option<String>,
    pub value_metadata_index: u32,
}

impl From<ExternalLinkKind> for owner::ExternalLinkKind {
    fn from(value: ExternalLinkKind) -> Self {
        match value {
            ExternalLinkKind::Workbook(value) => Self::Workbook(value.into()),
            ExternalLinkKind::Dde(value) => Self::Dde(value.into()),
            ExternalLinkKind::Ole(value) => Self::Ole(value.into()),
        }
    }
}

impl From<owner::ExternalLinkKind> for ExternalLinkKind {
    fn from(value: owner::ExternalLinkKind) -> Self {
        match value {
            owner::ExternalLinkKind::Workbook(value) => Self::Workbook(value.into()),
            owner::ExternalLinkKind::Dde(value) => Self::Dde(value.into()),
            owner::ExternalLinkKind::Ole(value) => Self::Ole(value.into()),
        }
    }
}

impl From<ExternalLinkEntry> for owner::ExternalLinkEntry {
    fn from(value: ExternalLinkEntry) -> Self {
        Self {
            index: value.index,
            relationship_id: value.relationship_id,
            part_uri: value.part_uri,
            kind: value.kind.into(),
        }
    }
}

impl From<owner::ExternalLinkEntry> for ExternalLinkEntry {
    fn from(value: owner::ExternalLinkEntry) -> Self {
        Self {
            index: value.index,
            relationship_id: value.relationship_id,
            part_uri: value.part_uri,
            kind: value.kind.into(),
        }
    }
}

impl From<ExternalDdeLink> for owner::ExternalDdeLink {
    fn from(value: ExternalDdeLink) -> Self {
        Self {
            service: value.service,
            topic: value.topic,
            items: value.items.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<owner::ExternalDdeLink> for ExternalDdeLink {
    fn from(value: owner::ExternalDdeLink) -> Self {
        Self {
            service: value.service,
            topic: value.topic,
            items: value.items.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<ExternalOleLink> for owner::ExternalOleLink {
    fn from(value: ExternalOleLink) -> Self {
        Self {
            target: value.target.into(),
            program_id: value.program_id,
            items: value.items.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<owner::ExternalOleLink> for ExternalOleLink {
    fn from(value: owner::ExternalOleLink) -> Self {
        Self {
            target: value.target.into(),
            program_id: value.program_id,
            items: value.items.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<ExternalOleTarget> for owner::ExternalOleTarget {
    fn from(value: ExternalOleTarget) -> Self {
        Self {
            relationship_id: value.relationship_id,
            target: value.target,
            relationship_type: value.relationship_type,
        }
    }
}

impl From<owner::ExternalOleTarget> for ExternalOleTarget {
    fn from(value: owner::ExternalOleTarget) -> Self {
        Self {
            relationship_id: value.relationship_id,
            target: value.target,
            relationship_type: value.relationship_type,
        }
    }
}

impl From<ExternalDdeItem> for owner::ExternalDdeItem {
    fn from(value: ExternalDdeItem) -> Self {
        Self {
            name: value.name,
            use_ole: value.use_ole,
            advise: value.advise,
            prefer_picture: value.prefer_picture,
            values: value.values.map(Into::into),
        }
    }
}

impl From<owner::ExternalDdeItem> for ExternalDdeItem {
    fn from(value: owner::ExternalDdeItem) -> Self {
        Self {
            name: value.name,
            use_ole: value.use_ole,
            advise: value.advise,
            prefer_picture: value.prefer_picture,
            values: value.values.map(Into::into),
        }
    }
}

impl From<ExternalOleItemSource> for owner::ExternalOleItemSource {
    fn from(value: ExternalOleItemSource) -> Self {
        match value {
            ExternalOleItemSource::SpreadsheetMl => Self::SpreadsheetMl,
            ExternalOleItemSource::Office2010 => Self::Office2010,
        }
    }
}

impl From<owner::ExternalOleItemSource> for ExternalOleItemSource {
    fn from(value: owner::ExternalOleItemSource) -> Self {
        match value {
            owner::ExternalOleItemSource::SpreadsheetMl => Self::SpreadsheetMl,
            owner::ExternalOleItemSource::Office2010 => Self::Office2010,
        }
    }
}

impl From<ExternalOleItem> for owner::ExternalOleItem {
    fn from(value: ExternalOleItem) -> Self {
        Self {
            source: value.source.into(),
            name: value.name,
            icon: value.icon,
            advise: value.advise,
            prefer_picture: value.prefer_picture,
            values: value.values.map(Into::into),
        }
    }
}

impl From<owner::ExternalOleItem> for ExternalOleItem {
    fn from(value: owner::ExternalOleItem) -> Self {
        Self {
            source: value.source.into(),
            name: value.name,
            icon: value.icon,
            advise: value.advise,
            prefer_picture: value.prefer_picture,
            values: value.values.map(Into::into),
        }
    }
}

impl From<ExternalDdeValues> for owner::ExternalDdeValues {
    fn from(value: ExternalDdeValues) -> Self {
        Self {
            rows: value.rows,
            columns: value.columns,
            values: value.values.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<owner::ExternalDdeValues> for ExternalDdeValues {
    fn from(value: owner::ExternalDdeValues) -> Self {
        Self {
            rows: value.rows,
            columns: value.columns,
            values: value.values.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<ExternalDdeValueType> for owner::ExternalDdeValueType {
    fn from(value: ExternalDdeValueType) -> Self {
        match value {
            ExternalDdeValueType::Nil => Self::Nil,
            ExternalDdeValueType::Boolean => Self::Boolean,
            ExternalDdeValueType::Number => Self::Number,
            ExternalDdeValueType::Error => Self::Error,
            ExternalDdeValueType::String => Self::String,
        }
    }
}

impl From<owner::ExternalDdeValueType> for ExternalDdeValueType {
    fn from(value: owner::ExternalDdeValueType) -> Self {
        match value {
            owner::ExternalDdeValueType::Nil => Self::Nil,
            owner::ExternalDdeValueType::Boolean => Self::Boolean,
            owner::ExternalDdeValueType::Number => Self::Number,
            owner::ExternalDdeValueType::Error => Self::Error,
            owner::ExternalDdeValueType::String => Self::String,
        }
    }
}

impl From<ExternalDdeValue> for owner::ExternalDdeValue {
    fn from(value: ExternalDdeValue) -> Self {
        Self {
            value_type: value.value_type.into(),
            raw_value: value.raw_value,
        }
    }
}

impl From<owner::ExternalDdeValue> for ExternalDdeValue {
    fn from(value: owner::ExternalDdeValue) -> Self {
        Self {
            value_type: value.value_type.into(),
            raw_value: value.raw_value,
        }
    }
}

impl From<ExternalWorkbookLink> for owner::ExternalWorkbookLink {
    fn from(value: ExternalWorkbookLink) -> Self {
        Self {
            target: value.target.into(),
            sheet_names: value.sheet_names,
            defined_names: value.defined_names.into_iter().map(Into::into).collect(),
            cached_sheets: value.cached_sheets.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<owner::ExternalWorkbookLink> for ExternalWorkbookLink {
    fn from(value: owner::ExternalWorkbookLink) -> Self {
        Self {
            target: value.target.into(),
            sheet_names: value.sheet_names,
            defined_names: value.defined_names.into_iter().map(Into::into).collect(),
            cached_sheets: value.cached_sheets.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<ExternalWorkbookTarget> for owner::ExternalWorkbookTarget {
    fn from(value: ExternalWorkbookTarget) -> Self {
        Self {
            relationship_id: value.relationship_id,
            target: value.target,
            relationship_type: value.relationship_type,
        }
    }
}

impl From<owner::ExternalWorkbookTarget> for ExternalWorkbookTarget {
    fn from(value: owner::ExternalWorkbookTarget) -> Self {
        Self {
            relationship_id: value.relationship_id,
            target: value.target,
            relationship_type: value.relationship_type,
        }
    }
}

impl From<ExternalDefinedName> for owner::ExternalDefinedName {
    fn from(value: ExternalDefinedName) -> Self {
        Self {
            name: value.name,
            refers_to: value.refers_to,
            sheet_id: value.sheet_id,
        }
    }
}

impl From<owner::ExternalDefinedName> for ExternalDefinedName {
    fn from(value: owner::ExternalDefinedName) -> Self {
        Self {
            name: value.name,
            refers_to: value.refers_to,
            sheet_id: value.sheet_id,
        }
    }
}

impl From<ExternalSheetData> for owner::ExternalSheetData {
    fn from(value: ExternalSheetData) -> Self {
        Self {
            sheet_id: value.sheet_id,
            refresh_error: value.refresh_error,
            rows: value.rows.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<owner::ExternalSheetData> for ExternalSheetData {
    fn from(value: owner::ExternalSheetData) -> Self {
        Self {
            sheet_id: value.sheet_id,
            refresh_error: value.refresh_error,
            rows: value.rows.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<ExternalRow> for owner::ExternalRow {
    fn from(value: ExternalRow) -> Self {
        Self {
            row: value.row,
            cells: value.cells.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<owner::ExternalRow> for ExternalRow {
    fn from(value: owner::ExternalRow) -> Self {
        Self {
            row: value.row,
            cells: value.cells.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<ExternalCellType> for owner::ExternalCellType {
    fn from(value: ExternalCellType) -> Self {
        match value {
            ExternalCellType::Number => Self::Number,
            ExternalCellType::Boolean => Self::Boolean,
            ExternalCellType::Date => Self::Date,
            ExternalCellType::Error => Self::Error,
            ExternalCellType::InlineString => Self::InlineString,
            ExternalCellType::SharedString => Self::SharedString,
            ExternalCellType::String => Self::String,
        }
    }
}

impl From<owner::ExternalCellType> for ExternalCellType {
    fn from(value: owner::ExternalCellType) -> Self {
        match value {
            owner::ExternalCellType::Number => Self::Number,
            owner::ExternalCellType::Boolean => Self::Boolean,
            owner::ExternalCellType::Date => Self::Date,
            owner::ExternalCellType::Error => Self::Error,
            owner::ExternalCellType::InlineString => Self::InlineString,
            owner::ExternalCellType::SharedString => Self::SharedString,
            owner::ExternalCellType::String => Self::String,
        }
    }
}

impl From<ExternalCell> for owner::ExternalCell {
    fn from(value: ExternalCell) -> Self {
        Self {
            reference: value.reference,
            cell_type: value.cell_type.into(),
            raw_value: value.raw_value,
            value_metadata_index: value.value_metadata_index,
        }
    }
}

impl From<owner::ExternalCell> for ExternalCell {
    fn from(value: owner::ExternalCell) -> Self {
        Self {
            reference: value.reference,
            cell_type: value.cell_type.into(),
            raw_value: value.raw_value,
            value_metadata_index: value.value_metadata_index,
        }
    }
}

#[cfg(test)]
pub(crate) fn build_external_link_part(
    part_uri: PackURI,
    kind: &ExternalLinkKind,
) -> Result<BlobPart> {
    owner::build_external_link_part(part_uri, &kind.clone().into_owner()).map_err(map_owner_error)
}

pub(crate) fn build_external_link_part_with_conformance(
    part_uri: PackURI,
    kind: &ExternalLinkKind,
    conformance: ExternalLinkConformance,
) -> Result<BlobPart> {
    owner::build_external_link_part_with_conformance(
        part_uri,
        &kind.clone().into_owner(),
        conformance.into_owner(),
    )
    .map_err(map_owner_error)
}

pub(crate) fn load_external_link(
    part: &dyn Part,
    workbook_relationship_id: String,
    index: u32,
) -> Result<ExternalLinkEntry> {
    owner::load_external_link(part, workbook_relationship_id, index)
        .map(Into::into)
        .map_err(map_owner_error)
}

fn map_owner_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_xlsx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Xlsx(other),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    #[test]
    fn parses_sparse_lexical_cache_without_dereferencing_target() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Data"/></sheetNames><definedNames><definedName name="Rate" refersTo="Data!$A$1" sheetId="1"/></definedNames><sheetDataSet><sheetData sheetId="1"><row r="1"><cell r="A1" t="str"><v>001.2300</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to_ext(
            "https://127.0.0.1:9/never-open.xlsx",
            rt::EXTERNAL_LINK_PATH,
        );
        let link = load_external_link(&part, "bookRel".into(), 1).unwrap();
        let ExternalLinkKind::Workbook(book) = link.kind else {
            panic!("expected workbook link")
        };
        assert_eq!(book.target.target, "https://127.0.0.1:9/never-open.xlsx");
        assert_eq!(book.sheet_names, ["Data"]);
        assert_eq!(
            book.cached_sheets[0].rows[0].cells[0].raw_value.as_deref(),
            Some("001.2300")
        );
    }

    #[test]
    fn parses_typed_dde_and_ole_links_without_dereferencing_targets() {
        let dde = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="excel" ddeTopic="[source.xlsx]Sheet1"><ddeItems><ddeItem name="R1C1:R1C2" advise="1"><values cols="2"><value t="n"><val>001.20</val></value><value t="str"><val>A&amp;B</val></value></values></ddeItem></ddeItems></ddeLink></externalLink>"#;
        let part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            dde.to_vec(),
        );
        let ExternalLinkKind::Dde(link) = load_external_link(&part, "rId1".into(), 1).unwrap().kind
        else {
            panic!("expected DDE link")
        };
        assert_eq!(link.service, "excel");
        assert!(link.items[0].advise);
        let values = link.items[0].values.as_ref().unwrap();
        assert_eq!((values.rows, values.columns), (1, 2));
        assert_eq!(values.values[0].raw_value, "001.20");
        assert_eq!(values.values[1].raw_value, "A&B");

        let ole = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><oleLink r:id="rId1" progId="Excel.Sheet.12"><oleItems><oleItem name="Core" icon="1"/><x14:oleItem name="Cached" preferPic="true"><x14:values><value t="b"><val>1</val></value></x14:values></x14:oleItem></oleItems></oleLink></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink2.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            ole.to_vec(),
        );
        part.relate_to_ext("https://127.0.0.1:9/never-open.bin", rt::OLE_OBJECT);
        let ExternalLinkKind::Ole(link) = load_external_link(&part, "rId2".into(), 2).unwrap().kind
        else {
            panic!("expected OLE link")
        };
        assert_eq!(link.target.target, "https://127.0.0.1:9/never-open.bin");
        assert_eq!(link.items[0].source, ExternalOleItemSource::SpreadsheetMl);
        assert_eq!(link.items[1].source, ExternalOleItemSource::Office2010);
        assert_eq!(
            link.items[1].values.as_ref().unwrap().values[0].value_type,
            ExternalDdeValueType::Boolean
        );
    }

    #[test]
    fn rejects_malformed_dde_ole_links_and_matrices() {
        let sml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        for child in [
            "<ddeLink ddeService=\"x\"/>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems/></ddeLink>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems><ddeItem><values rows=\"2\"><value><val>x</val></value></values></ddeItem></ddeItems></ddeLink>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems><ddeItem><values><value t=\"future\"><val>x</val></value></values></ddeItem></ddeItems></ddeLink>",
            "<ddeLink ddeService=\"x\" ddeTopic=\"y\"><ddeItems><ddeItem><values><value/></values></ddeItem></ddeItems></ddeLink>",
            "<oleLink xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/>",
        ] {
            let xml = format!("<externalLink xmlns=\"{sml}\">{child}</externalLink>");
            let part = BlobPart::new(
                PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
                litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
                xml.into_bytes(),
            );
            assert!(
                load_external_link(&part, "rId1".into(), 1).is_err(),
                "accepted {child}"
            );
        }
    }

    #[test]
    fn canonical_writer_round_trips_typed_dde_and_ole_links() {
        let dde = ExternalLinkKind::Dde(ExternalDdeLink {
            service: "x&y".into(),
            topic: "topic".into(),
            items: vec![ExternalDdeItem {
                name: Some("R1C1".into()),
                use_ole: false,
                advise: true,
                prefer_picture: false,
                values: Some(ExternalDdeValues {
                    rows: 1,
                    columns: 1,
                    values: vec![ExternalDdeValue {
                        value_type: ExternalDdeValueType::String,
                        raw_value: "<&>".into(),
                    }],
                }),
            }],
        });
        let xml = dde.to_xml().unwrap();
        let part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml,
        );
        assert_eq!(
            load_external_link(&part, "rId1".into(), 1).unwrap().kind,
            dde
        );

        let ole = ExternalLinkKind::Ole(ExternalOleLink {
            target: ExternalOleTarget {
                relationship_id: "rId1".into(),
                target: "file.bin".into(),
                relationship_type: rt::OLE_OBJECT.into(),
            },
            program_id: "Excel.Sheet.12".into(),
            items: vec![ExternalOleItem {
                source: ExternalOleItemSource::Office2010,
                name: "Item".into(),
                icon: false,
                advise: false,
                prefer_picture: true,
                values: Some(ExternalDdeValues {
                    rows: 1,
                    columns: 1,
                    values: vec![ExternalDdeValue {
                        value_type: ExternalDdeValueType::Number,
                        raw_value: "1.00".into(),
                    }],
                }),
            }],
        });
        let part = build_external_link_part(
            PackURI::new("/xl/externalLinks/externalLink2.xml").unwrap(),
            &ole,
        )
        .unwrap();
        assert_eq!(
            load_external_link(&part, "rId2".into(), 2).unwrap().kind,
            ole
        );
    }

    #[test]
    fn workbook_add_and_replace_external_link_survive_save() {
        let first = ExternalLinkKind::Dde(ExternalDdeLink {
            service: "one".into(),
            topic: "topic".into(),
            items: Vec::new(),
        });
        let replacement = ExternalLinkKind::Dde(ExternalDdeLink {
            service: "two".into(),
            topic: "updated".into(),
            items: vec![ExternalDdeItem {
                name: Some("R1C1".into()),
                use_ole: false,
                advise: true,
                prefer_picture: false,
                values: None,
            }],
        });
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        assert_eq!(workbook.add_external_link(first).unwrap(), 1);
        workbook
            .replace_external_link(1, replacement.clone())
            .unwrap();
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("typed-external-link.xlsx");
        workbook.save(&path).unwrap();
        let package = litchi_opc::OpcPackage::from_bytes(&std::fs::read(path).unwrap()).unwrap();
        let reopened = crate::xlsx::Workbook::new(package).unwrap();
        assert_eq!(reopened.external_links().len(), 1);
        assert_eq!(reopened.external_link(1).unwrap().kind, replacement);
    }

    #[test]
    fn mce_fallback_is_semantic_input() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:no"><mc:AlternateContent><mc:Choice Requires="x"><x:no/></mc:Choice><mc:Fallback><externalBook r:id="rId1"/></mc:Fallback></mc:AlternateContent></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to_ext("opaque.xlsx", rt::EXTERNAL_LINK_PATH);
        assert!(matches!(
            load_external_link(&part, "rId1".into(), 1).unwrap().kind,
            ExternalLinkKind::Workbook(_)
        ));
    }

    #[test]
    fn accepts_strict_external_workbook_namespaces_and_relationships() {
        let xml = br#"<externalLink xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Strict"/></sheetNames></externalBook></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to_ext("strict-target.xlsx", rt::STRICT_EXTERNAL_LINK_PATH);
        let link = load_external_link(&part, "rId1".into(), 1).unwrap();
        let ExternalLinkKind::Workbook(book) = link.kind else {
            panic!("expected workbook link")
        };
        assert_eq!(book.sheet_names, ["Strict"]);
        assert_eq!(book.target.relationship_type, rt::STRICT_EXTERNAL_LINK_PATH);
    }

    #[test]
    fn rejects_duplicate_and_over_limit_workbook_external_references() {
        const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let duplicate = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="sheet"/></sheets><externalReferences><externalReference r:id="link"/><externalReference r:id="link"/></externalReferences></workbook>"#
        );
        assert!(crate::xlsx::parsers::workbook_parser::parse_workbook_details(&duplicate).is_err());

        let mut oversized = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="sheet"/></sheets><externalReferences>"#
        );
        for index in 0..=4096 {
            oversized.push_str(&format!(r#"<externalReference r:id="link{index}"/>"#));
        }
        oversized.push_str("</externalReferences></workbook>");
        assert!(crate::xlsx::parsers::workbook_parser::parse_workbook_details(&oversized).is_err());
    }

    #[test]
    fn rejects_internal_targets_and_malformed_caches() {
        let xml = br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetDataSet><sheetData sheetId="0"/></sheetDataSet></externalBook></externalLink>"#;
        let mut part = BlobPart::new(
            PackURI::new("/xl/externalLinks/externalLink1.xml").unwrap(),
            litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
            xml.to_vec(),
        );
        part.relate_to("../workbook.xml", rt::EXTERNAL_LINK_PATH);
        assert!(load_external_link(&part, "rId1".into(), 1).is_err());
    }

    #[test]
    fn loads_poi_ordered_external_workbook_reference() {
        let package = litchi_opc::OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/poi/test-data/spreadsheet/link-external-workbook-b.xlsx"
        ))
        .unwrap();
        let workbook = crate::xlsx::Workbook::new(package).unwrap();
        assert_eq!(workbook.external_links().len(), 1);
        let link = workbook.external_link(1).unwrap();
        assert_eq!(link.index, 1);
        assert_eq!(link.relationship_id, "rId4");
        let ExternalLinkKind::Workbook(book) = &link.kind else {
            panic!("expected external workbook")
        };
        assert_eq!(book.target.target, "link-external-workbook-a.xlsx");
    }

    #[test]
    fn loads_libreoffice_sparse_external_cache() {
        let package = litchi_opc::OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/external-refs.xlsx"
        ))
        .unwrap();
        let workbook = crate::xlsx::Workbook::new(package).unwrap();
        let values: Vec<&str> = workbook
            .external_links()
            .iter()
            .filter_map(|link| match &link.kind {
                ExternalLinkKind::Workbook(book) => Some(book),
                _ => None,
            })
            .flat_map(|book| &book.cached_sheets)
            .flat_map(|sheet| &sheet.rows)
            .flat_map(|row| &row.cells)
            .filter_map(|cell| cell.raw_value.as_deref())
            .collect();
        for expected in ["Name", "Andy", "Bruce", "Charlie"] {
            assert!(
                values.contains(&expected),
                "missing cached value {expected}"
            );
        }
    }

    #[test]
    fn unrelated_modified_save_preserves_external_part_and_relationship() {
        const FIXTURE: &[u8] = include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/external-refs.xlsx"
        );
        let original = litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap();
        let original_workbook =
            crate::xlsx::Workbook::new(litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap())
                .unwrap();
        let original_link = original_workbook.external_links().first().unwrap();
        let external_uri = original_link.part_uri.clone();
        let workbook_relationship_id = original_link.relationship_id.clone();
        let original_external = original.get_part(&external_uri).unwrap().blob().to_vec();
        let original_workbook_part = original.main_document_part().unwrap();
        let original_relationship = original_workbook_part
            .rels()
            .get(&workbook_relationship_id)
            .unwrap();
        let original_target = original_relationship.target_ref().to_string();
        let original_type = original_relationship.reltype().to_string();

        let mut workbook =
            crate::xlsx::Workbook::new(litchi_opc::OpcPackage::from_bytes(FIXTURE).unwrap())
                .unwrap();
        workbook.props_mut().unwrap().title = Some("Unrelated edit".to_owned());
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("external-link-roundtrip.xlsx");
        workbook.save(&path).unwrap();

        let bytes = std::fs::read(&path).unwrap();
        let saved = litchi_opc::OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            saved.get_part(&external_uri).unwrap().blob(),
            original_external
        );
        let workbook_part = saved.main_document_part().unwrap();
        let relationship = workbook_part.rels().get(&workbook_relationship_id).unwrap();
        assert_eq!(relationship.reltype(), original_type);
        assert_eq!(relationship.target_ref(), original_target);
        let workbook_xml = std::str::from_utf8(workbook_part.blob()).unwrap();
        assert!(workbook_xml.contains("<externalReferences>"));
        assert!(workbook_xml.contains(&format!(
            "<externalReference r:id=\"{workbook_relationship_id}\"/>"
        )));
        assert_eq!(
            crate::xlsx::Workbook::new(saved)
                .unwrap()
                .external_links()
                .len(),
            original_workbook.external_links().len()
        );
    }
}
