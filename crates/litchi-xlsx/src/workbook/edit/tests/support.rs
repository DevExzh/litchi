//! Shared workbook fixtures for the semantic edit test modules.

use litchi_ooxml_common::web as common_web;
use litchi_opc::{BlobPart, PackURI, TargetMode};

use super::super::{Workbook, WorksheetKind};

pub(super) fn task_panes(app_ref: &str) -> common_web::Panes {
    let reference = common_web::Reference::new("test-add-in", "1.0", common_web::Store::Omex)
        .expect("reference");
    let add_in = common_web::AddIn::new("test-add-in", reference)
        .and_then(|add_in| add_in.bind(common_web::Binding::new("table", "table", app_ref)?))
        .expect("add-in");
    let mut panes = common_web::Panes::new();
    panes
        .push(common_web::Pane::new(add_in))
        .expect("task pane");
    panes
}

pub(super) fn rename_reference_workbook() -> Workbook {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rIdTab2"/></sheets><definedNames><definedName name="Source">Data!$A$1</definedName></definedNames></workbook>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("Data worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[1].part_uri)
            .expect("Calc worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><f>Data!A1</f><v>7</v></c></row></sheetData><dataValidations count="1"><dataValidation type="custom" sqref="B1"><formula1>Data!A1&gt;0</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="C1" location="Data!$A$1"/></hyperlinks></worksheet>"#.to_vec(),
            );
    for (uri, content_type, content) in [
            (
                "/xl/tables/table1.xml",
                litchi_opc::constants::content_type::SML_TABLE,
                br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><tableColumns count="1"><tableColumn id="1" name="Value"><calculatedColumnFormula>Data!A1</calculatedColumnFormula></tableColumn></tableColumns></table>"#.as_slice(),
            ),
            (
                "/xl/charts/chart1.xml",
                litchi_opc::constants::content_type::DML_CHART,
                br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$A$1</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
            ),
            (
                "/xl/pivotCache/pivotCacheDefinition1.xml",
                litchi_opc::constants::content_type::SML_PIVOT_CACHE_DEFINITION,
                br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1"/></cacheSource></pivotCacheDefinition>"#.as_slice(),
            ),
            (
                "/docProps/app.xml",
                litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES,
                br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Calc</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#.as_slice(),
            ),
            (
                "/xl/externalLinks/externalLink1.xml",
                litchi_opc::constants::content_type::SML_EXTERNAL_LINK,
                br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook><definedNames><definedName name="External" refersTo="[1]Data!A1"/></definedNames></externalBook></externalLink>"#.as_slice(),
            ),
        ] {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(uri).expect("part URI"),
                    content_type.to_owned(),
                    content.to_vec(),
                )))
                .expect("reference part");
        }
    Workbook::from_package(package).expect("rename reference workbook")
}

pub(super) fn part_text<'a>(workbook: &'a Workbook, uri: &str) -> &'a str {
    let uri = PackURI::new(uri).expect("part URI");
    let bytes = workbook
        .inner
        .package
        .get_part(&uri)
        .expect("package part")
        .blob();
    std::str::from_utf8(bytes).expect("XML part")
}

pub(super) fn styled_workbook() -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
            .expect("styles part")
            .set_blob(
                br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#.to_vec(),
            );
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("styled workbook")
}

pub(super) fn styled_column_workbook() -> Workbook {
    let baseline = styled_workbook();
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="3" max="3" style="1"/></cols><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("styled column workbook")
}

pub(super) fn styled_row_workbook() -> Workbook {
    let baseline = styled_workbook();
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row><row r="2" s="1" customFormat="1"/></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("styled row workbook")
}

pub(super) fn defaults_workbook() -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="ac"><sheetFormatPr baseColWidth="10" defaultColWidth="12" defaultRowHeight="15" customHeight="0" zeroHeight="1" thickTop="1" ac:dyDescent="0.1"/><sheetData><row r="2" customHeight="0" ac:dyDescent="0.2"/></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("defaults workbook")
}

pub(super) fn merged_workbook() -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>anchor</t></is></c></row><row r="2"><c r="E2" t="inlineStr"><is><t>keep</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("merged workbook")
}

pub(super) fn two_sheet_workbook(second_kind: WorksheetKind) -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    let (relationship_type, content_type, part_xml) = match second_kind {
            WorksheetKind::Worksheet => (
                litchi_opc::constants::relationship_type::WORKSHEET,
                litchi_opc::constants::content_type::SML_WORKSHEET,
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/></worksheet>"#.as_slice(),
            ),
            WorksheetKind::Chart => (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
                br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.as_slice(),
            ),
            _ => panic!("test helper only models worksheet and chart tabs"),
        };
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            match second_kind {
                WorksheetKind::Worksheet => "worksheets/sheet2.xml",
                WorksheetKind::Chart => "chartsheets/sheet2.xml",
                _ => unreachable!("guarded above"),
            }
            .to_owned(),
            "rIdTab2".to_owned(),
            TargetMode::Internal,
        )
        .expect("second sheet relationship");
    let part_uri = match second_kind {
        WorksheetKind::Worksheet => "/xl/worksheets/sheet2.xml",
        WorksheetKind::Chart => "/xl/chartsheets/sheet2.xml",
        _ => unreachable!("guarded above"),
    };
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(part_uri).expect("second sheet URI"),
            content_type.to_owned(),
            part_xml.to_vec(),
        )))
        .expect("second sheet part");
    Workbook::from_package(package).expect("two-sheet workbook")
}

pub(super) fn three_sheet_workbook() -> Workbook {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1" firstSheet="0"/><workbookView activeTab="2" firstSheet="1"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/><sheet name="Sheet3" sheetId="3" r:id="rIdTab3"/></sheets><definedNames><definedName name="FirstLocal" localSheetId="0">Sheet1!$A$1</definedName><definedName name="ThirdLocal" localSheetId="2">Sheet3!$A$1</definedName><definedName name="Global">1</definedName></definedNames></workbook>"#.to_vec(),
            );
    package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
            "worksheets/sheet3.xml".to_owned(),
            "rIdTab3".to_owned(),
            TargetMode::Internal,
        )
        .expect("third sheet relationship");
    package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/worksheets/sheet3.xml").expect("third sheet URI"),
                litchi_opc::constants::content_type::SML_WORKSHEET.to_owned(),
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>three</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            )))
            .expect("third sheet part");
    package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("first sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[1].part_uri)
            .expect("second sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>two</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("three-sheet workbook")
}

pub(super) fn active_second_sheet_workbook(second_kind: WorksheetKind) -> Workbook {
    let source = two_sheet_workbook(second_kind);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("first sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#.to_vec(),
            );
    let second_xml = match second_kind {
            WorksheetKind::Worksheet => br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#.as_slice(),
            WorksheetKind::Chart => br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews></chartsheet>"#.as_slice(),
            _ => unreachable!("test helper only models worksheet and chart tabs"),
        };
    package
        .get_part_mut(&source.inner.sheets[1].part_uri)
        .expect("second sheet part")
        .set_blob(second_xml.to_vec());
    Workbook::from_package(package).expect("active second sheet workbook")
}
