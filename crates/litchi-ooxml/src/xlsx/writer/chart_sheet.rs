//! Chart sheet authoring for the mutable XLSX workbook writer.
//!
//! A chartsheet is a workbook sheet whose part lives under
//! `/xl/chartsheets/` (content type `...spreadsheetml.chartsheet+xml`,
//! relationship type `.../relationships/chartsheet`) and whose root
//! `chartsheet` element (ECMA-376 part 1, `CT_ChartSheet`) references a
//! drawing part that anchors exactly one chart. This module holds the
//! mutable model and the part XML emission; the save pipeline in
//! `xlsx::workbook` wires the parts and relationships into the package.

use super::sheet::WorksheetChart;
use crate::xlsx::chartsheet::{
    ChartSheet, ChartSheetConformance, ChartSheetMargins, ChartSheetView, write_chartsheet,
};
use litchi_core::sheet::Result as SheetResult;

/// Relationship type linking a workbook to a chartsheet part.
pub(crate) const CHARTSHEET_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
/// Content type of a chartsheet part.
pub(crate) const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
/// Relationship ID used for the chartsheet's drawing and the drawing's
/// chart (each part is written fresh with a single relationship).
pub(crate) const CHARTSHEET_DRAWING_REL_ID: &str = "rId1";
const MAX_SHEET_NAME_CHARS: usize = 31;
pub(crate) const MAX_CHART_SHEETS: usize = 65_536;
const DEFAULT_CHART_EXTENT_X: u64 = 8_582_025;
const DEFAULT_CHART_EXTENT_Y: u64 = 5_838_825;
const CHART_NAME_ID: u32 = 1;

/// A chartsheet being authored in a mutable workbook.
///
/// A chartsheet hosts exactly one chart, anchored through its drawing part
/// with an absolute anchor (the only anchor kind chartsheet drawings use).
#[derive(Debug, Clone)]
pub struct MutableChartSheet {
    name: String,
    sheet_id: u32,
    chart: WorksheetChart,
}

impl MutableChartSheet {
    pub(crate) fn new(name: String, sheet_id: u32, chart: WorksheetChart) -> Self {
        Self {
            name,
            sheet_id,
            chart,
        }
    }

    /// Chartsheet name as it appears in the workbook's sheet list.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Unique sheet ID recorded in the workbook's sheet entry.
    pub fn sheet_id(&self) -> u32 {
        self.sheet_id
    }

    /// The chart hosted on this chartsheet.
    pub fn chart(&self) -> &WorksheetChart {
        &self.chart
    }

    /// Mutable access to the hosted chart, for example to bind it to a
    /// pivot table with `WorksheetChart::into_pivot_chart`.
    pub fn chart_mut(&mut self) -> &mut WorksheetChart {
        &mut self.chart
    }
}

/// Validate a chartsheet chart for the dedicated chartsheet emission path.
///
/// Chart external data, user shapes, and additional relationships are only
/// wired for worksheet drawings today and are rejected here.
pub(crate) fn validate_chart_sheet_chart(chart: &WorksheetChart) -> SheetResult<()> {
    if chart.external_data_part.is_some() {
        return Err("chartsheet charts do not support external data parts".into());
    }
    if chart.user_shapes_part.is_some() {
        return Err("chartsheet charts do not support user-shapes parts".into());
    }
    if !chart.additional_relationships.is_empty() {
        return Err("chartsheet charts do not support additional relationships".into());
    }
    Ok(())
}

/// Validate an Excel sheet name: 1-31 characters, none of `: \ / ? * [ ]`,
/// not the reserved name `History` (any letter case), and not starting or
/// ending with an apostrophe (Excel quotes sheet names in references with
/// `'` and forbids it at the edges of the name itself).
pub(crate) fn validate_sheet_name(name: &str) -> SheetResult<()> {
    let length = name.chars().count();
    if length == 0 || length > MAX_SHEET_NAME_CHARS {
        return Err(
            format!("sheet name '{name}' must be 1-{MAX_SHEET_NAME_CHARS} characters").into(),
        );
    }
    if let Some(bad) = name
        .chars()
        .find(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(format!("sheet name '{name}' contains forbidden character '{bad}'").into());
    }
    if name.eq_ignore_ascii_case("history") {
        return Err("sheet name 'History' is reserved by Excel".into());
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(format!("sheet name '{name}' must not start or end with an apostrophe").into());
    }
    Ok(())
}

/// Serialize the chartsheet part XML through the typed chartsheet writer.
pub(crate) fn chart_sheet_part_xml() -> SheetResult<Vec<u8>> {
    let chartsheet = ChartSheet {
        properties: None,
        views: vec![ChartSheetView {
            tab_selected: None,
            zoom_scale: None,
            workbook_view_id: 0,
            zoom_to_fit: None,
        }],
        protection: None,
        custom_views: None,
        margins: Some(ChartSheetMargins {
            left: 0.75,
            right: 0.75,
            top: 1.0,
            bottom: 1.0,
            header: 0.5,
            footer: 0.5,
        }),
        page_setup: None,
        header_footer: None,
        drawing_relationship_id: CHARTSHEET_DRAWING_REL_ID.to_string(),
        legacy_drawing_relationship_id: None,
        legacy_header_footer_drawing_relationship_id: None,
        background_picture_relationship_id: None,
        web_publish_items: None,
        extension_list: None,
    };
    write_chartsheet(&chartsheet, ChartSheetConformance::Transitional).map_err(Into::into)
}

/// Serialize the chartsheet drawing part: one absolute-anchored graphic
/// frame referencing the chart through `rId1`.
pub(crate) fn chart_sheet_drawing_xml(chart_title: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="{DEFAULT_CHART_EXTENT_X}" cy="{DEFAULT_CHART_EXTENT_Y}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{CHART_NAME_ID}" name="{}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{CHARTSHEET_DRAWING_REL_ID}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#,
        litchi_core::xml::escape_xml(chart_title)
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::chartsheet::parse_chartsheet;

    #[test]
    fn validates_sheet_names() {
        assert!(validate_sheet_name("Chart 1").is_ok());
        // Interior apostrophes are legal in Excel sheet names.
        assert!(validate_sheet_name("Sheet'1").is_ok());
        for name in [
            "",
            "This chartsheet name is way too long to be valid",
            "a/b",
            "a\\b",
            "a:b",
            "a?b",
            "a*b",
            "a[b",
            "a]b",
            "History",
            "HISTORY",
            "'Quoted",
            "Quoted'",
        ] {
            assert!(validate_sheet_name(name).is_err(), "accepted '{name}'");
        }
    }

    #[test]
    fn emitted_chartsheet_part_parses_through_typed_reader() {
        let xml = chart_sheet_part_xml().unwrap();
        let (conformance, chartsheet) = parse_chartsheet(&xml).unwrap();
        assert_eq!(conformance, ChartSheetConformance::Transitional);
        assert_eq!(chartsheet.views.len(), 1);
        assert_eq!(
            chartsheet.drawing_relationship_id,
            CHARTSHEET_DRAWING_REL_ID
        );
        assert!(chartsheet.margins.is_some());
    }

    #[test]
    fn emitted_drawing_has_one_absolute_anchor_chart_reference() {
        let xml = chart_sheet_drawing_xml("Chart 1");
        let drawing = crate::xlsx::drawing::parse_drawing_xml(&xml)
            .unwrap()
            .expect("drawing has a wsDr root");
        assert!(drawing.pictures.is_empty());
        assert_eq!(drawing.charts.len(), 1);
        assert_eq!(drawing.charts[0].relationship_id, CHARTSHEET_DRAWING_REL_ID);
    }

    fn temp_xlsx_path(tag: &str) -> std::path::PathBuf {
        std::env::temp_dir().join(format!(
            "litchi-chartsheet-{tag}-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ))
    }

    fn bar_chart() -> WorksheetChart {
        crate::xlsx::WorksheetChart::bar_chart(
            "Sales",
            "Sheet1!$A$2:$A$3",
            "Sheet1!$B$2:$B$3",
            crate::xlsx::ChartAnchor::new(0, 0, 10, 15),
        )
        .unwrap()
    }

    #[test]
    fn authored_chartsheet_round_trips_through_save() {
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        {
            let worksheet = workbook.worksheet_mut(0).unwrap();
            worksheet.set_cell_value(1, 1, "Region");
            worksheet.set_cell_value(1, 2, "Sales");
            worksheet.set_cell_value(2, 1, "North");
            worksheet.set_cell_value(2, 2, 10.0);
            worksheet.set_cell_value(3, 1, "South");
            worksheet.set_cell_value(3, 2, 20.0);
        }
        // Interleave: chartsheet inserted between two worksheets.
        workbook
            .add_chart_sheet("Sales Chart", bar_chart())
            .unwrap();
        workbook.add_worksheet("Sheet2");
        let path = temp_xlsx_path("round-trip");
        workbook.save(&path).unwrap();

        let reopened = crate::xlsx::Workbook::open(&path).unwrap();
        std::fs::remove_file(&path).ok();
        // Sheet order is preserved across worksheets and chartsheets.
        assert_eq!(
            litchi_core::sheet::WorkbookTrait::worksheet_names(&reopened),
            ["Sheet1", "Sales Chart", "Sheet2"]
        );

        // The chartsheet resolves through the typed read-side graph.
        let package = reopened.package();
        let workbook_part = package.main_document_part().unwrap();
        let relationship_id = workbook_part
            .rels()
            .iter()
            .find(|rel| rel.reltype() == CHARTSHEET_RELATIONSHIP_TYPE)
            .expect("workbook has a chartsheet relationship")
            .r_id()
            .to_string();
        let loaded = crate::xlsx::chartsheet::load_chartsheet(
            package,
            workbook_part.partname(),
            &relationship_id,
        )
        .unwrap();
        assert_eq!(loaded.entry.name, "Sales Chart");
        assert_eq!(loaded.entry.part_name, "/xl/chartsheets/sheet1.xml");
        assert_eq!(loaded.chartsheet.views.len(), 1);
        assert_eq!(loaded.drawing.charts.len(), 1);
        let chart_resource = &loaded.drawing.charts[0];
        assert!(chart_resource.part_name.starts_with("/xl/charts/"));
        assert!(chart_resource.part_name.ends_with(".xml"));
        let chart = crate::charts::reader::parse_chart(chart_resource.data.as_slice()).unwrap();
        assert!(chart.pivot_source.is_none());
        assert_eq!(chart.plot_area.type_groups.len(), 1);
    }

    #[test]
    fn rejects_duplicate_and_invalid_chartsheet_names() {
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        // Duplicate of an existing worksheet name (any case).
        assert!(workbook.add_chart_sheet("sheet1", bar_chart()).is_err());
        workbook.add_chart_sheet("Chart1", bar_chart()).unwrap();
        // Duplicate of an existing chartsheet name (any case).
        assert!(workbook.add_chart_sheet("CHART1", bar_chart()).is_err());
        // Invalid names.
        assert!(workbook.add_chart_sheet("", bar_chart()).is_err());
        assert!(workbook.add_chart_sheet("a/b", bar_chart()).is_err());
        assert!(
            workbook
                .add_chart_sheet(
                    "This chartsheet name is way too long to be valid",
                    bar_chart()
                )
                .is_err()
        );
    }
}
