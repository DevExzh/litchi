//! Table types and implementation for DOCX documents.
use crate::docx::namespace::normalize_xml_integer;
use crate::docx::table::VMergeState;
use crate::error::{OoxmlError, Result};
use litchi_core::xml::escape_xml;
use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::TableBorderStyle;
// Import paragraph types
use super::paragraph::MutableParagraph;
use super::revision::{
    CellRevisionKind, RevisionMetadata, RowRevisionKind, TableCellMergeRevisionState,
    TableRevisionKind,
};

/// Border definition for table or cell.
#[derive(Debug, Clone)]
pub struct TableBorder {
    /// Border style
    pub style: TableBorderStyle,
    /// Border width in eighths of a point (e.g., 8 = 1pt, 24 = 3pt)
    pub size: u32,
    /// Border color in hex RGB format (e.g., "FF0000" for red)
    pub color: String,
}

impl Default for TableBorder {
    fn default() -> Self {
        Self {
            style: TableBorderStyle::Single,
            size: 4,
            color: "000000".to_string(),
        }
    }
}

/// Table borders (all sides).
#[derive(Debug, Clone, Default)]
pub struct TableBorders {
    pub top: Option<TableBorder>,
    pub left: Option<TableBorder>,
    pub bottom: Option<TableBorder>,
    pub right: Option<TableBorder>,
    pub inside_h: Option<TableBorder>,
    pub inside_v: Option<TableBorder>,
}

/// Table properties.
#[derive(Debug, Default, Clone)]
pub(crate) struct TableProperties {
    pub(crate) borders: TableBorders,
    pub(crate) width_pct: Option<u32>,
}

/// Cell properties.
#[derive(Debug, Default, Clone)]
pub struct CellProperties {
    /// Cell background color in hex RGB format
    pub background_color: Option<String>,
    /// Cell borders (if different from table borders)
    pub borders: Option<TableBorders>,
    /// Cell width in DXA units (twentieth of a point)
    pub width_dxa: Option<u32>,
    /// Normal vertical-merge state applied to this cell.
    pub vertical_merge: Option<VMergeState>,
}

#[derive(Debug, Clone)]
struct PropertyChange<T> {
    metadata: RevisionMetadata,
    previous: T,
}

#[derive(Debug, Clone)]
struct CellMergeChange {
    metadata: RevisionMetadata,
    original: TableCellMergeRevisionState,
    current: TableCellMergeRevisionState,
}

/// A mutable table.
#[derive(Debug)]
pub struct MutableTable {
    /// Table rows
    pub(crate) rows: Vec<MutableRow>,
    /// Table properties
    pub(crate) properties: TableProperties,
    revision: Option<(TableRevisionKind, RevisionMetadata)>,
    property_change: Option<PropertyChange<TableProperties>>,
}

impl MutableTable {
    pub(crate) fn new(rows: usize, cols: usize) -> Self {
        let mut table = Self {
            rows: Vec::with_capacity(rows),
            properties: TableProperties::default(),
            revision: None,
            property_change: None,
        };
        for _ in 0..rows {
            table.add_row(cols);
        }
        table
    }

    /// Add a new row with specified column count.
    pub fn add_row(&mut self, cols: usize) -> &mut MutableRow {
        self.rows.push(MutableRow::new(cols));
        self.rows.last_mut().unwrap()
    }

    /// Set table width as percentage (100-500 where 100=20%, 500=100%).
    pub fn set_width_percent(&mut self, percent: u32) {
        self.properties.width_pct = Some(percent * 50);
    }

    /// Set all table borders at once.
    pub fn set_borders(&mut self, border: TableBorder) {
        self.properties.borders.top = Some(border.clone());
        self.properties.borders.left = Some(border.clone());
        self.properties.borders.bottom = Some(border.clone());
        self.properties.borders.right = Some(border.clone());
        self.properties.borders.inside_h = Some(border.clone());
        self.properties.borders.inside_v = Some(border);
    }

    /// Get a cell by row and column index.
    pub fn cell(&mut self, row: usize, col: usize) -> Option<&mut MutableCell> {
        self.rows.get_mut(row)?.cell(col)
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get a row by index.
    pub fn row(&mut self, index: usize) -> Option<&mut MutableRow> {
        self.rows.get_mut(index)
    }

    /// Mark this whole table as inserted or deleted.
    pub fn set_revision(
        &mut self,
        kind: TableRevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<&mut Self> {
        if self.revision.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "table insertion and deletion revisions conflict".into(),
            ));
        }
        self.ensure_local_id_available(metadata.id())?;
        self.revision = Some((kind, metadata));
        Ok(self)
    }

    /// Record the table properties which existed before a tracked formatting change.
    pub fn set_property_change(
        &mut self,
        metadata: RevisionMetadata,
        previous: &MutableTable,
    ) -> Result<&mut Self> {
        if self.property_change.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "table property revision already exists".into(),
            ));
        }
        if previous.revision.is_some() || previous.property_change.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "previous table properties must not contain revision metadata".into(),
            ));
        }
        self.ensure_local_id_available(metadata.id())?;
        self.property_change = Some(PropertyChange {
            metadata,
            previous: previous.properties.clone(),
        });
        Ok(self)
    }

    pub fn revision_kind(&self) -> Option<TableRevisionKind> {
        self.revision.as_ref().map(|(kind, _)| *kind)
    }

    fn ensure_local_id_available(&self, id: u32) -> Result<()> {
        if self.revision.as_ref().is_some_and(|(_, value)| value.id() == id)
            || self.property_change.as_ref().is_some_and(|value| value.metadata.id() == id)
        {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate table revision ID {id}"
            )));
        }
        Ok(())
    }

    fn write_border(xml: &mut String, name: &str, border: &TableBorder) -> Result<()> {
        write!(
            xml,
            "<w:{} w:val=\"{}\" w:sz=\"{}\" w:space=\"0\" w:color=\"{}\"/>",
            name,
            border.style.as_str(),
            border.size,
            escape_xml(&border.color)
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))
    }

    fn write_borders(
        xml: &mut String,
        wrapper: &str,
        borders: &TableBorders,
        default_missing: bool,
    ) -> Result<()> {
        write!(xml, "<w:{wrapper}>")?;
        for (name, border) in [
            ("top", &borders.top),
            ("left", &borders.left),
            ("bottom", &borders.bottom),
            ("right", &borders.right),
            ("insideH", &borders.inside_h),
            ("insideV", &borders.inside_v),
        ] {
            if let Some(border) = border {
                Self::write_border(xml, name, border)?;
            } else if default_missing {
                write!(xml, "<w:{name} w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>")?;
            }
        }
        write!(xml, "</w:{wrapper}>")?;
        Ok(())
    }

    fn write_property_values(xml: &mut String, properties: &TableProperties) -> Result<()> {
        let width = properties.width_pct.unwrap_or(5000);
        write!(xml, "<w:tblW w:w=\"{width}\" w:type=\"pct\"/>")?;
        Self::write_borders(xml, "tblBorders", &properties.borders, true)
    }

    fn validate_revision_ids(&self) -> Result<()> {
        let mut ids = HashSet::new();
        let mut insert = |metadata: &RevisionMetadata| {
            if ids.insert(metadata.id()) {
                Ok(())
            } else {
                Err(OoxmlError::InvalidFormat(format!(
                    "duplicate table subtree revision ID {}",
                    metadata.id()
                )))
            }
        };
        if let Some((_, metadata)) = &self.revision { insert(metadata)?; }
        if let Some(change) = &self.property_change { insert(&change.metadata)?; }
        for row in &self.rows {
            if let Some((_, metadata)) = &row.revision { insert(metadata)?; }
            if let Some(change) = &row.property_change { insert(&change.metadata)?; }
            for cell in &row.cells {
                if let Some((_, metadata)) = &cell.revision { insert(metadata)?; }
                if let Some(change) = &cell.merge_change { insert(&change.metadata)?; }
                if let Some(change) = &cell.property_change { insert(&change.metadata)?; }
            }
        }
        Ok(())
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        self.validate_revision_ids()?;
        xml.push_str("<w:tbl>");

        // Write table properties
        xml.push_str("<w:tblPr>");

        Self::write_property_values(xml, &self.properties)?;
        if let Some((kind, metadata)) = &self.revision {
            write!(xml, "<w:{}", kind.element())?;
            metadata.write_attributes(xml)?;
            xml.push_str("/>");
        }
        if let Some(change) = &self.property_change {
            xml.push_str("<w:tblPrChange");
            change.metadata.write_attributes(xml)?;
            xml.push_str("><w:tblPr>");
            Self::write_property_values(xml, &change.previous)?;
            xml.push_str("</w:tblPr></w:tblPrChange>");
        }
        xml.push_str("</w:tblPr>");

        // Write grid
        if let Some(first_row) = self.rows.first() {
            xml.push_str("<w:tblGrid>");
            for _ in 0..first_row.cell_count() {
                xml.push_str("<w:gridCol/>");
            }
            xml.push_str("</w:tblGrid>");
        }

        // Write rows
        for row in &self.rows {
            row.to_xml(xml)?;
        }

        xml.push_str("</w:tbl>");

        Ok(())
    }
}

/// A mutable table row.
#[derive(Debug)]
pub struct MutableRow {
    /// Table cells in this row
    pub(crate) cells: Vec<MutableCell>,
    /// HTML division ID referenced by this row
    pub(crate) division_id: Option<String>,
    revision: Option<(RowRevisionKind, RevisionMetadata)>,
    property_change: Option<PropertyChange<Option<String>>>,
}

impl MutableRow {
    pub(crate) fn new(cols: usize) -> Self {
        let mut row = Self {
            cells: Vec::with_capacity(cols),
            division_id: None,
            revision: None,
            property_change: None,
        };
        for _ in 0..cols {
            row.cells.push(MutableCell::new());
        }
        row
    }

    /// Get a cell by index.
    pub fn cell(&mut self, index: usize) -> Option<&mut MutableCell> {
        self.cells.get_mut(index)
    }

    /// Add a new cell.
    pub fn add_cell(&mut self) -> &mut MutableCell {
        self.cells.push(MutableCell::new());
        self.cells.last_mut().unwrap()
    }

    /// Get the number of cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Set the HTML division ID referenced by this row.
    ///
    /// Division IDs are XML Schema integers and are kept as strings so values
    /// larger than the native integer types can be written without truncation.
    pub fn set_division_id(&mut self, id: impl Into<String>) -> Result<&mut Self> {
        self.division_id = Some(normalize_xml_integer(
            id.into(),
            "Word table-row division ID",
        )?);
        Ok(self)
    }

    /// Return the HTML division ID referenced by this row, if set.
    pub fn division_id(&self) -> Option<&str> {
        self.division_id.as_deref()
    }

    /// Remove the HTML division reference from this row.
    pub fn clear_division_id(&mut self) -> &mut Self {
        self.division_id = None;
        self
    }

    pub fn set_revision(
        &mut self,
        kind: RowRevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<&mut Self> {
        if self.revision.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "row insertion and deletion revisions conflict".into(),
            ));
        }
        self.ensure_local_id_available(metadata.id())?;
        self.revision = Some((kind, metadata));
        Ok(self)
    }

    pub fn set_property_change(
        &mut self,
        metadata: RevisionMetadata,
        previous: &MutableRow,
    ) -> Result<&mut Self> {
        if self.property_change.is_some() {
            return Err(OoxmlError::InvalidFormat("row property revision already exists".into()));
        }
        if previous.revision.is_some() || previous.property_change.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "previous row properties must not contain revision metadata".into(),
            ));
        }
        self.ensure_local_id_available(metadata.id())?;
        self.property_change = Some(PropertyChange { metadata, previous: previous.division_id.clone() });
        Ok(self)
    }

    pub fn revision_kind(&self) -> Option<RowRevisionKind> {
        self.revision.as_ref().map(|(kind, _)| *kind)
    }

    fn ensure_local_id_available(&self, id: u32) -> Result<()> {
        if self.revision.as_ref().is_some_and(|(_, value)| value.id() == id)
            || self.property_change.as_ref().is_some_and(|value| value.metadata.id() == id)
        {
            return Err(OoxmlError::InvalidFormat(format!("duplicate row revision ID {id}")));
        }
        Ok(())
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:tr>");

        if self.division_id.is_some() || self.revision.is_some() || self.property_change.is_some() {
            xml.push_str("<w:trPr>");
            if let Some(division_id) = &self.division_id {
                write!(xml, "<w:divId w:val=\"{}\"/>", escape_xml(division_id))?;
            }
            if let Some((kind, metadata)) = &self.revision {
                write!(xml, "<w:{}", kind.element())?;
                metadata.write_attributes(xml)?;
                xml.push_str("/>");
            }
            if let Some(change) = &self.property_change {
                xml.push_str("<w:trPrChange");
                change.metadata.write_attributes(xml)?;
                xml.push_str("><w:trPr>");
                if let Some(division_id) = &change.previous {
                    write!(xml, "<w:divId w:val=\"{}\"/>", escape_xml(division_id))?;
                }
                xml.push_str("</w:trPr></w:trPrChange>");
            }
            xml.push_str("</w:trPr>");
        }

        for cell in &self.cells {
            cell.to_xml(xml)?;
        }

        xml.push_str("</w:tr>");

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::MutableRow;
    use crate::docx::table::Row;

    #[test]
    fn writes_row_division_id_for_reader_round_trip() {
        let mut row = MutableRow::new(1);
        row.set_division_id("-123456789012345678901234567890")
            .unwrap();

        assert_eq!(row.division_id(), Some("-123456789012345678901234567890"));

        let mut xml = String::new();
        row.to_xml(&mut xml).unwrap();
        let parsed = Row::new(xml.into_bytes());
        assert_eq!(
            parsed.division_id().unwrap().as_deref(),
            Some("-123456789012345678901234567890")
        );

        row.clear_division_id();
        let mut cleared_xml = String::new();
        row.to_xml(&mut cleared_xml).unwrap();
        assert!(!cleared_xml.contains("<w:divId"));
    }

    #[test]
    fn rejects_invalid_row_division_id() {
        let mut row = MutableRow::new(0);
        assert!(row.set_division_id("+").is_err());
        assert_eq!(row.division_id(), None);
    }
}

/// A mutable table cell.
#[derive(Debug)]
pub struct MutableCell {
    /// Paragraphs in this cell
    pub(crate) paragraphs: Vec<MutableParagraph>,
    /// Cell properties
    pub(crate) properties: CellProperties,
    revision: Option<(CellRevisionKind, RevisionMetadata)>,
    merge_change: Option<CellMergeChange>,
    property_change: Option<PropertyChange<CellProperties>>,
}

impl MutableCell {
    pub(crate) fn new() -> Self {
        Self {
            paragraphs: vec![MutableParagraph::new()],
            properties: CellProperties::default(),
            revision: None,
            merge_change: None,
            property_change: None,
        }
    }

    /// Add a new paragraph to the cell.
    pub fn add_paragraph(&mut self) -> &mut MutableParagraph {
        self.paragraphs.push(MutableParagraph::new());
        self.paragraphs.last_mut().unwrap()
    }

    /// Get the number of paragraphs.
    pub fn paragraph_count(&self) -> usize {
        self.paragraphs.len()
    }

    /// Get a paragraph by index.
    pub fn paragraph(&mut self, index: usize) -> Option<&mut MutableParagraph> {
        self.paragraphs.get_mut(index)
    }

    /// Set text in the first paragraph.
    pub fn set_text(&mut self, text: &str) {
        self.paragraphs.clear();
        let para = self.add_paragraph();
        para.add_run_with_text(text);
    }

    /// Set cell background color in hex RGB format (e.g., "FFFF00" for yellow).
    pub fn set_background_color(&mut self, color: &str) {
        self.properties.background_color = Some(color.to_string());
    }

    /// Set cell width in DXA units (twentieth of a point).
    pub fn set_width_dxa(&mut self, width: u32) {
        self.properties.width_dxa = Some(width);
    }

    pub fn set_borders(&mut self, borders: TableBorders) {
        self.properties.borders = Some(borders);
    }

    pub fn set_vertical_merge(&mut self, state: Option<VMergeState>) {
        self.properties.vertical_merge = state;
    }

    pub fn set_revision(
        &mut self,
        kind: CellRevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<&mut Self> {
        if self.revision.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "cell insertion and deletion revisions conflict".into(),
            ));
        }
        self.ensure_local_id_available(metadata.id())?;
        self.revision = Some((kind, metadata));
        Ok(self)
    }

    pub fn set_merge_revision(
        &mut self,
        metadata: RevisionMetadata,
        original: TableCellMergeRevisionState,
        current: TableCellMergeRevisionState,
    ) -> Result<&mut Self> {
        if original == current {
            return Err(OoxmlError::InvalidFormat(
                "cell merge revision must change the vertical merge state".into(),
            ));
        }
        if self.merge_change.is_some() {
            return Err(OoxmlError::InvalidFormat("cell merge revision already exists".into()));
        }
        self.ensure_local_id_available(metadata.id())?;
        self.merge_change = Some(CellMergeChange { metadata, original, current });
        Ok(self)
    }

    pub fn set_property_change(
        &mut self,
        metadata: RevisionMetadata,
        previous: &MutableCell,
    ) -> Result<&mut Self> {
        if self.property_change.is_some() {
            return Err(OoxmlError::InvalidFormat("cell property revision already exists".into()));
        }
        if previous.revision.is_some() || previous.merge_change.is_some() || previous.property_change.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "previous cell properties must not contain revision metadata".into(),
            ));
        }
        self.ensure_local_id_available(metadata.id())?;
        self.property_change = Some(PropertyChange { metadata, previous: previous.properties.clone() });
        Ok(self)
    }

    pub fn revision_kind(&self) -> Option<CellRevisionKind> {
        self.revision.as_ref().map(|(kind, _)| *kind)
    }

    fn ensure_local_id_available(&self, id: u32) -> Result<()> {
        if self.revision.as_ref().is_some_and(|(_, value)| value.id() == id)
            || self.merge_change.as_ref().is_some_and(|value| value.metadata.id() == id)
            || self.property_change.as_ref().is_some_and(|value| value.metadata.id() == id)
        {
            return Err(OoxmlError::InvalidFormat(format!("duplicate cell revision ID {id}")));
        }
        Ok(())
    }

    fn has_properties(&self) -> bool {
        self.properties.background_color.is_some()
            || self.properties.borders.is_some()
            || self.properties.width_dxa.is_some()
            || self.properties.vertical_merge.is_some()
    }

    fn write_property_values(xml: &mut String, properties: &CellProperties) -> Result<()> {
        if let Some(width) = properties.width_dxa {
            write!(xml, "<w:tcW w:w=\"{width}\" w:type=\"dxa\"/>")?;
        }
        if let Some(state) = properties.vertical_merge {
            match state {
                VMergeState::Restart => xml.push_str("<w:vMerge w:val=\"restart\"/>"),
                VMergeState::Continue => xml.push_str("<w:vMerge/>")
            }
        }
        if let Some(borders) = &properties.borders {
            MutableTable::write_borders(xml, "tcBorders", borders, false)?;
        }
        if let Some(color) = &properties.background_color {
            write!(xml, "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{}\"/>", escape_xml(color))?;
        }
        Ok(())
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:tc>");

        // Write cell properties if any
        if self.has_properties() || self.revision.is_some() || self.merge_change.is_some() || self.property_change.is_some() {
            xml.push_str("<w:tcPr>");
            Self::write_property_values(xml, &self.properties)?;
            if let Some((kind, metadata)) = &self.revision {
                write!(xml, "<w:{}", kind.element())?;
                metadata.write_attributes(xml)?;
                xml.push_str("/>");
            }
            if let Some(change) = &self.merge_change {
                xml.push_str("<w:cellMerge");
                change.metadata.write_attributes(xml)?;
                write!(xml, " w:vMerge=\"{}\" w:vMergeOrig=\"{}\"/>", change.current.token(), change.original.token())?;
            }
            if let Some(change) = &self.property_change {
                xml.push_str("<w:tcPrChange");
                change.metadata.write_attributes(xml)?;
                xml.push_str("><w:tcPr>");
                Self::write_property_values(xml, &change.previous)?;
                xml.push_str("</w:tcPr></w:tcPrChange>");
            }
            xml.push_str("</w:tcPr>");
        }

        for para in &self.paragraphs {
            para.to_xml(xml)?;
        }

        xml.push_str("</w:tc>");

        Ok(())
    }
}

#[cfg(test)]
mod revision_tests {
    use super::*;
    use crate::docx::{RevisionType, Table};

    fn metadata(id: &str) -> RevisionMetadata {
        let mut metadata = RevisionMetadata::new(id, "表 & \"作者\"").unwrap();
        metadata
            .set_date(Some("2026-07-19T10:30:00+08:00"))
            .unwrap();
        metadata
    }

    #[test]
    fn writes_and_parses_all_table_row_and_cell_revision_types() {
        let mut previous_table = MutableTable::new(1, 1);
        previous_table.set_width_percent(42);
        let mut table = MutableTable::new(2, 2);
        table.set_width_percent(80);
        table
            .set_revision(TableRevisionKind::Insert, metadata("1"))
            .unwrap();
        table
            .set_property_change(metadata("2"), &previous_table)
            .unwrap();

        let mut previous_row = MutableRow::new(1);
        previous_row.set_division_id("17").unwrap();
        let row = table.row(0).unwrap();
        row.set_division_id("18").unwrap();
        row.set_revision(RowRevisionKind::Insert, metadata("3"))
            .unwrap();
        row.set_property_change(metadata("4"), &previous_row)
            .unwrap();

        let mut previous_cell = MutableCell::new();
        previous_cell.set_width_dxa(720);
        previous_cell.set_background_color("00FF00");
        previous_cell.set_vertical_merge(Some(VMergeState::Restart));
        let cell = table.cell(0, 0).unwrap();
        cell.set_width_dxa(1440);
        cell.set_background_color("F0F0F0");
        cell.set_vertical_merge(Some(VMergeState::Continue));
        cell.set_revision(CellRevisionKind::Insert, metadata("5"))
            .unwrap();
        cell.set_merge_revision(
            metadata("6"),
            TableCellMergeRevisionState::Rest,
            TableCellMergeRevisionState::Continue,
        )
        .unwrap();
        cell.set_property_change(metadata("7"), &previous_cell)
            .unwrap();

        table
            .row(1)
            .unwrap()
            .set_revision(RowRevisionKind::Delete, metadata("8"))
            .unwrap();
        table
            .cell(1, 0)
            .unwrap()
            .set_revision(CellRevisionKind::Delete, metadata("9"))
            .unwrap();

        let mut xml = String::new();
        table.to_xml(&mut xml).unwrap();

        assert!(xml.contains("<w:tblIns w:id=\"1\" w:author=\"表 &amp; &quot;作者&quot;\""));
        assert!(xml.contains("<w:ins w:id=\"3\""));
        assert!(xml.contains("<w:del w:id=\"8\""));
        assert!(xml.contains("<w:cellIns w:id=\"5\""));
        assert!(xml.contains("<w:cellDel w:id=\"9\""));
        assert!(xml.contains("<w:cellMerge w:id=\"6\""));
        assert!(xml.contains("w:vMergeOrig=\"rest\""));
        assert!(xml.contains("w:vMerge=\"cont\""));
        assert!(xml.contains("<w:vMerge/>"));
        assert!(xml.contains("<w:tblPr><w:tblW w:w=\"2100\""));
        assert!(xml.contains("<w:trPr><w:divId w:val=\"17\"/></w:trPr></w:trPrChange>"));
        assert!(xml.contains("<w:tcPr><w:tcW w:w=\"720\""));

        let tbl_change = xml.find("<w:tblPrChange").unwrap();
        let tbl_pr_end = xml.find("</w:tblPr>").unwrap();
        assert!(tbl_change < tbl_pr_end);
        let row_change = xml.find("<w:trPrChange").unwrap();
        let row_pr_end = xml.find("</w:trPr>").unwrap();
        assert!(row_change < row_pr_end);
        let cell_marker = xml.find("<w:cellIns").unwrap();
        let cell_merge = xml.find("<w:cellMerge").unwrap();
        let cell_change = xml.find("<w:tcPrChange").unwrap();
        assert!(cell_marker < cell_merge && cell_merge < cell_change);

        let parsed = Table::new(xml.into_bytes()).revisions().unwrap();
        let kinds: Vec<_> = parsed.iter().map(|revision| revision.revision_type()).collect();
        for kind in [
            RevisionType::TableInsert,
            RevisionType::TablePropertiesChange,
            RevisionType::RowInsert,
            RevisionType::RowDelete,
            RevisionType::RowPropertiesChange,
            RevisionType::CellInsert,
            RevisionType::CellDelete,
            RevisionType::CellMerge,
            RevisionType::CellPropertiesChange,
        ] {
            assert!(kinds.contains(&kind), "missing parsed revision {kind:?}");
        }
        assert_eq!(parsed[0].author(), "表 & \"作者\"");
    }

    #[test]
    fn writes_table_delete_and_parses_strict_namespace() {
        let mut table = MutableTable::new(1, 1);
        table
            .set_revision(TableRevisionKind::Delete, metadata("10"))
            .unwrap();
        let mut body = String::new();
        table.to_xml(&mut body).unwrap();
        assert!(body.contains("<w:tblDel w:id=\"10\""));

        let strict = body.replacen(
            "<w:tbl>",
            "<w:tbl xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\">",
            1,
        );
        let revisions = Table::new(strict.into_bytes()).revisions().unwrap();
        assert_eq!(revisions.len(), 1);
        assert_eq!(revisions[0].revision_type(), RevisionType::TableDelete);
    }

    #[test]
    fn rejects_conflicts_duplicates_nested_snapshots_and_noop_merges_atomically() {
        let mut table = MutableTable::new(1, 1);
        table
            .set_revision(TableRevisionKind::Insert, metadata("20"))
            .unwrap();
        assert!(
            table
                .set_revision(TableRevisionKind::Delete, metadata("21"))
                .is_err()
        );
        assert_eq!(table.revision_kind(), Some(TableRevisionKind::Insert));

        let mut nested_previous = MutableTable::new(1, 1);
        nested_previous
            .set_revision(TableRevisionKind::Delete, metadata("22"))
            .unwrap();
        assert!(
            table
                .set_property_change(metadata("23"), &nested_previous)
                .is_err()
        );

        let row = table.row(0).unwrap();
        row.set_revision(RowRevisionKind::Insert, metadata("24"))
            .unwrap();
        assert!(
            row.set_revision(RowRevisionKind::Delete, metadata("25"))
                .is_err()
        );
        assert_eq!(row.revision_kind(), Some(RowRevisionKind::Insert));

        let cell = table.cell(0, 0).unwrap();
        cell.set_revision(CellRevisionKind::Insert, metadata("26"))
            .unwrap();
        assert!(
            cell.set_revision(CellRevisionKind::Delete, metadata("27"))
                .is_err()
        );
        assert!(
            cell.set_merge_revision(
                metadata("28"),
                TableCellMergeRevisionState::Rest,
                TableCellMergeRevisionState::Rest,
            )
            .is_err()
        );
        assert_eq!(cell.revision_kind(), Some(CellRevisionKind::Insert));

        let mut xml = String::new();
        table.to_xml(&mut xml).unwrap();
        assert!(xml.contains("<w:tblIns w:id=\"20\""));
        assert!(!xml.contains("<w:tblDel"));
        assert!(!xml.contains("<w:tblPrChange"));
        assert!(!xml.contains("<w:cellMerge"));

        let mut duplicate = MutableTable::new(1, 1);
        duplicate
            .set_revision(TableRevisionKind::Insert, metadata("30"))
            .unwrap();
        duplicate
            .row(0)
            .unwrap()
            .set_revision(RowRevisionKind::Insert, metadata("30"))
            .unwrap();
        let mut unchanged = String::from("sentinel");
        assert!(duplicate.to_xml(&mut unchanged).is_err());
        assert_eq!(unchanged, "sentinel");
    }
}
