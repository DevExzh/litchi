//! Worksheet feature reads of a source-backed workbook (unified-ops record
//! 0322): only the selected worksheet stream and the parts its features
//! live in are read, and the worksheet's cells are not decoded.

use super::{
    SourceBackedWorkbook, SourceBackedWorksheet, XLSB_COMMENTS_CONTENT_TYPE, postflight_package,
    preflight_package, require_content_type,
};
use crate::comments::Record;
use crate::conditional_formatting::edit::DifferentialStyle;
use crate::drawing::edit::Element;
use crate::package::error::{Error, Result};
use crate::sheet::{FeatureSet, Features, Kind};

impl SourceBackedWorksheet {
    fn fenced<T>(&self, read: impl FnOnce() -> Result<T>) -> Result<T> {
        let _ = preflight_package(&self.inner.package)?;
        if !matches!(self.metadata().kind, Kind::Worksheet) {
            return Err(Error::UnsupportedFeature(
                "worksheet features belong to worksheets",
            ));
        }
        let result = read();
        postflight_package(&self.inner.package)?;
        result
    }

    /// The `select`ed hyperlinks, auto-filter, classic validations and
    /// conditional formats, and page breaks of this worksheet, read without
    /// decoding its cells.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-worksheet tab, a changed source, or a
    /// malformed stream.
    pub fn features(&self, select: FeatureSet) -> Result<Features> {
        self.fenced(|| {
            let partname = self.physical_partname()?.clone();
            let mut formula_context = self.inner.formula_context.clone();
            formula_context.external_books = self.inner.external_books()?;
            formula_context.tables = self.inner.tables()?;
            let part = self.inner.package.part(&partname)?;
            let data = part.data()?;
            crate::sheet::features::read(
                data.as_bytes(),
                part.rels(),
                &formula_context,
                self.metadata().workbook_position,
                self.inner.leniency,
                select,
                self.inner.raw_limits,
            )
        })
    }

    /// The notes of this worksheet, in stream order.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-worksheet tab, a changed source, or a
    /// malformed comments part.
    pub fn comments(&self) -> Result<Vec<Record>> {
        self.fenced(|| {
            let partname = self.physical_partname()?.clone();
            let part = self.inner.package.part(&partname)?;
            let Some((_, name)) = crate::comments::edit::comments_part(part.rels())? else {
                return Ok(Vec::new());
            };
            let comments = self.inner.package.part(&name)?;
            require_content_type(&comments, XLSB_COMMENTS_CONTENT_TYPE)?;
            crate::comments::edit::read(comments.data()?.as_bytes(), self.inner.raw_limits)
        })
    }

    /// The top-level objects of this worksheet's drawing, in drawing order.
    /// Only the worksheet's relationships, its drawing part and the image
    /// parts' entries are read.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-worksheet tab, a changed source, or a
    /// malformed drawing graph.
    pub fn drawing_elements(&self) -> Result<Vec<Element>> {
        self.fenced(|| {
            let partname = self.physical_partname()?.clone();
            crate::drawing::edit::source_elements(&self.inner.package, &partname)
        })
    }

    /// The charts of this worksheet's drawing, in drawing order, with their
    /// chart part markup.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-worksheet tab, a changed source, or a
    /// malformed drawing graph.
    pub fn drawing_charts(&self) -> Result<Vec<crate::drawing::edit::ChartPart>> {
        self.fenced(|| {
            let partname = self.physical_partname()?.clone();
            crate::drawing::edit::source_charts(&self.inner.package, &partname)
        })
    }
}

impl SourceBackedWorkbook {
    /// The differential formats (`BrtDXF`) of the styles part.
    ///
    /// # Errors
    ///
    /// Returns an error for a changed source or a malformed styles part.
    pub fn differential_formats(&self) -> Result<Vec<DifferentialStyle>> {
        let _ = preflight_package(&self.inner.package)?;
        let result = (|| {
            let Some(styles) = self.inner.styles_part.clone() else {
                return Ok(Vec::new());
            };
            let part = self.inner.package.part(&styles)?;
            crate::conditional_formatting::edit::read_differentials(
                part.data()?.as_bytes(),
                self.inner.raw_limits,
            )
        })();
        postflight_package(&self.inner.package)?;
        result
    }
}

#[cfg(test)]
mod tests {
    use crate::Package;
    use crate::SourceBackedWorkbook;
    use crate::conditional_formatting::edit::DifferentialStyle;
    use crate::conditional_formatting::{Formatting, Rule, RuleType};
    use crate::data_validation::Validation;
    use crate::drawing::edit::{Marker, Placement, ShapeSpec};
    use crate::hyperlinks::Hyperlink;
    use crate::page_breaks::{Break, PageBreaks};
    use crate::sheet::FeatureSet;

    fn featured() -> Package {
        let package = Package::new().unwrap();
        let mut edit = package.edit_hyperlinks().unwrap();
        edit.add(
            0,
            Hyperlink::new_external(0, 0, 0, 0, "https://a.test/".into()),
        )
        .unwrap();
        let package = edit.commit().unwrap().into_package();
        let mut edit = package.edit_data_validations().unwrap();
        let mut validation = Validation::new(1, "A1:A3".to_owned());
        validation.operator = 4;
        validation.formula1 = Some("1".to_owned());
        edit.add(0, validation).unwrap();
        let package = edit.commit().unwrap().into_package();
        let mut edit = package.edit_conditional_formats().unwrap();
        let dxf = edit
            .add_differential(&DifferentialStyle {
                bold: Some(true),
                ..DifferentialStyle::default()
            })
            .unwrap();
        let mut rule = Rule::new(RuleType::CellIs, 1);
        rule.operator = Some(5);
        rule.formula_texts = vec!["3".to_owned()];
        rule.dxf_id = Some(dxf);
        let mut format = Formatting::new(vec!["B1:B4".to_owned()]);
        format.add_rule(rule);
        edit.add(0, format).unwrap();
        let package = edit.commit().unwrap().into_package();
        let mut edit = package.edit_page_breaks().unwrap();
        edit.set_page_breaks(
            0,
            &PageBreaks {
                rows: vec![Break::row(5)],
                columns: vec![],
            },
        )
        .unwrap();
        let package = edit.commit().unwrap().into_package();
        let mut edit = package.edit_comments().unwrap();
        edit.add(0, 2, 2, "Ann", "note").unwrap();
        let package = edit.commit().unwrap().into_package();
        let mut edit = package.edit_drawings().unwrap();
        let placement = Placement::TwoCell {
            from: Marker::cell(3, 3),
            to: Marker::cell(5, 6),
        };
        edit.add_shape(0, &ShapeSpec::new(placement, "Box"))
            .unwrap();
        edit.commit().unwrap().into_package()
    }

    #[test]
    fn source_reads_equal_eager_reads() {
        let package = featured();
        let bytes = package.to_bytes().unwrap();
        let workbook = package.workbook().unwrap();
        let source = SourceBackedWorkbook::from_reader(bytes.as_slice()).unwrap();
        let sheet = source.sheet_by_index(0).unwrap().unwrap();
        let eager = workbook.sheet_features(0, FeatureSet::all()).unwrap();
        let lazy = sheet.features(FeatureSet::all()).unwrap();
        assert_eq!(lazy.hyperlinks, eager.hyperlinks);
        assert_eq!(
            lazy.hyperlinks[0].target.as_deref(),
            Some("https://a.test/")
        );
        assert_eq!(lazy.validations, eager.validations);
        assert_eq!(lazy.formats.len(), 1);
        assert_eq!(lazy.formats[0].rules[0].formula_texts, ["3"]);
        assert_eq!(lazy.page_breaks, eager.page_breaks);
        assert_eq!(
            sheet.comments().unwrap(),
            workbook.sheet_comments(0).unwrap()
        );
        assert_eq!(
            sheet.drawing_elements().unwrap(),
            workbook.drawing_elements(0).unwrap()
        );
        assert_eq!(
            source.differential_formats().unwrap(),
            workbook.differential_formats().unwrap()
        );
        assert_eq!(workbook.differential_formats().unwrap()[0].bold, Some(true));
    }

    #[test]
    fn charts_read_on_both_owners() {
        let bytes = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/testVarious.xlsb"
        ))
        .unwrap();
        let workbook = Package::from_bytes(&bytes)
            .unwrap()
            .materialize_workbook()
            .unwrap();
        let eager = workbook.drawing_charts(0).unwrap();
        assert_eq!(eager.len(), 1);
        assert!(eager[0].xml.starts_with(b"<?xml"));
        let source = SourceBackedWorkbook::from_reader(bytes.as_slice()).unwrap();
        let sheet = source.sheet_by_index(0).unwrap().unwrap();
        assert_eq!(sheet.drawing_charts().unwrap(), eager);
    }
}
