//! Exact row-visibility state over the source-backed scalar worksheet closure.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::Row;

use super::rewrite;
use crate::Selector;
use crate::cell_values;
use crate::error::Result;

/// Exact source-bound row-owner visibility state.
#[derive(Clone, Debug)]
pub struct Snapshot {
    inner: cell_values::Snapshot,
    rows: Arc<[(Row, bool)]>,
}

impl Snapshot {
    /// Load the conservative row-visibility closure from an owning package.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        Self::from_inner(cell_values::Snapshot::load(package, selector)?)
    }

    pub(crate) fn from_inner(inner: cell_values::Snapshot) -> Result<Self> {
        let rows = rewrite::scan(inner.source_xml())?;
        inner.check_execution()?;
        Ok(Self {
            inner,
            rows: Arc::from(rows),
        })
    }

    pub(crate) fn from_rewritten_source<'source>(
        source: &'source Self,
        rewrite: rewrite::VisibilityRewrite<'source>,
    ) -> Result<Self> {
        source.check_execution()?;
        Self::from_inner(cell_values::Snapshot::from_visibility_rewrite(
            &source.inner,
            rewrite,
        )?)
    }

    /// Selected worksheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        self.inner.sheet_name()
    }

    /// Selected zero-based sheet position.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.inner.sheet_position()
    }

    /// Selected worksheet Part URI.
    #[must_use]
    pub const fn worksheet_part_name(&self) -> &PackURI {
        self.inner.worksheet_part_name()
    }

    /// Exact source worksheet XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.inner.source_xml()
    }

    /// Whether an explicit row owner exists at this coordinate.
    #[must_use]
    pub fn contains_row(&self, row: Row) -> bool {
        self.rows
            .binary_search_by_key(&row, |(row, _)| *row)
            .is_ok()
    }

    /// Effective direct visibility of an existing row owner.
    ///
    /// `None` means no explicit `<row>` owner exists at this coordinate.
    #[must_use]
    pub fn is_hidden(&self, row: Row) -> Option<bool> {
        self.rows
            .binary_search_by_key(&row, |(row, _)| *row)
            .ok()
            .map(|index| self.rows[index].1)
    }

    pub(crate) const fn inner(&self) -> &cell_values::Snapshot {
        &self.inner
    }

    pub(crate) fn check_execution(&self) -> Result<()> {
        self.inner.check_execution()
    }
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, PackURI, TargetMode};

    use super::Snapshot;
    use crate::cell::Value;
    use crate::cell_values;
    use crate::row_visibility::rewrite;

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn fixture(worksheet: &str) -> litchi_opc::OpcPackage {
        let workbook = format!(
            r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
        );
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        let mut package = litchi_opc::OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                workbook_uri.clone(),
                ct::SML_SHEET_MAIN.to_owned(),
                workbook.into_bytes(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                worksheet_uri,
                ct::SML_WORKSHEET.to_owned(),
                worksheet.as_bytes().to_vec(),
            )))
            .unwrap();
        package
            .get_part_mut(&workbook_uri)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::WORKSHEET.to_owned(),
                "worksheets/sheet1.xml".to_owned(),
                "rIdSheet".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package
    }

    #[test]
    fn visibility_rewrite_reuses_cells_and_matches_full_candidate_parse() {
        let worksheet = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="2" hidden="true"><c r="B2" t="b"><v>1</v></c></row></sheetData></worksheet>"#
        );
        let source = Snapshot::load(&fixture(&worksheet), "Sheet1").unwrap();
        let first_row = litchi_sheet::Row::new(0).unwrap();
        let first_cell = litchi_sheet::Cell::from_a1("A1").unwrap();
        let second_cell = litchi_sheet::Cell::from_a1("B2").unwrap();
        let mut actions = BTreeMap::new();
        actions.insert(first_row, true);
        let (rewrite, changed) = rewrite::rewrite(source.source_xml(), &actions).unwrap();
        assert_eq!(changed, 1);

        let candidate = Snapshot::from_rewritten_source(&source, rewrite).unwrap();
        assert!(candidate.inner.shares_cell_store_with(&source.inner));
        assert!(matches!(
            candidate.inner.value(first_cell),
            Some(Value::Number(_))
        ));
        assert_eq!(
            candidate.inner.value(first_cell),
            source.inner.value(first_cell)
        );
        assert_eq!(candidate.inner.value(second_cell), Some(&Value::Bool(true)));
        assert_eq!(
            candidate.inner.value(second_cell),
            source.inner.value(second_cell)
        );
        assert_eq!(candidate.is_hidden(first_row), Some(true));

        let fully_parsed = Snapshot::from_inner(
            cell_values::Snapshot::from_rewritten_source(
                &source.inner,
                candidate.source_xml().to_vec(),
            )
            .unwrap(),
        )
        .unwrap();
        assert_eq!(candidate.rows, fully_parsed.rows);
        assert_eq!(
            candidate.inner.value(first_cell),
            fully_parsed.inner.value(first_cell)
        );
        assert_eq!(
            candidate.inner.value(second_cell),
            fully_parsed.inner.value(second_cell)
        );
        assert_eq!(candidate.source_xml(), fully_parsed.source_xml());
    }

    #[test]
    fn visibility_rewrite_proof_is_bound_to_its_exact_source() {
        let source_xml = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#
        );
        let foreign_xml = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>8</v></c></row></sheetData></worksheet>"#
        );
        let source = Snapshot::load(&fixture(&source_xml), "Sheet1").unwrap();
        let foreign = Snapshot::load(&fixture(&foreign_xml), "Sheet1").unwrap();
        let first_row = litchi_sheet::Row::new(0).unwrap();
        let mut actions = BTreeMap::new();
        actions.insert(first_row, true);
        let (rewrite, changed) = rewrite::rewrite(source.source_xml(), &actions).unwrap();
        assert_eq!(changed, 1);

        let error = Snapshot::from_rewritten_source(&foreign, rewrite).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("rewrite belongs to another worksheet source")
        );
    }
}
