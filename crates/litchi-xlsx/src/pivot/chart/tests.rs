//! Focused pivot-chart parser and package-graph tests.

use super::{C14_CHART_NAMESPACE, DEFAULT_FORMAT_ID};
use super::{FieldType, OPTIONS_EXTENSION_URI, parse_binding};

const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_C: &str = "http://purl.oclc.org/ooxml/drawingml/chart";

#[test]
fn parses_transitional_source_and_options() {
    let xml = format!(
        r#"<c:chartSpace xmlns:c="{C}" xmlns:c14="{C14_CHART_NAMESPACE}">
            <c:pivotSource>
                <c:name>Pivot!PivotOne</c:name>
                <c:fmtId val="7"/>
                <c:extLst><c:ext uri="{{unknown-source}}"/></c:extLst>
            </c:pivotSource>
            <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>
                <c:extLst><c:ext uri="{OPTIONS_EXTENSION_URI}">
                    <c14:pivotOptions>
                        <c14:dropZoneVisible val="0"/>
                        <c14:dropZoneData val="1"/>
                    </c14:pivotOptions>
                </c:ext></c:extLst>
            </c:ser></c:barChart></c:plotArea></c:chart>
        </c:chartSpace>"#
    );
    let binding = parse_binding(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(binding.source.name, "Pivot!PivotOne");
    assert_eq!(binding.source.format_id, 7);
    assert_eq!(binding.source.extension_uris, ["{unknown-source}"]);
    assert_eq!(binding.series.len(), 1);
    let options = binding.series[0].options.as_ref().unwrap();
    assert_eq!(options.drop_zone_visible, Some(false));
    assert_eq!(options.visibility(FieldType::DataFields), Some(true));
    assert_eq!(options.visibility(FieldType::AxisRow), None);
}

#[test]
fn parses_strict_chart_namespace_and_ignores_unknown_extensions() {
    let xml = format!(
        r#"<c:chartSpace xmlns:c="{STRICT_C}" xmlns:c14="{C14_CHART_NAMESPACE}">
            <c:pivotSource><c:name>PivotOne</c:name><c:fmtId val="0"/></c:pivotSource>
            <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="3"/>
                <c:extLst>
                    <c:ext uri="{{unknown-series}}"><x:payload xmlns:x="urn:inert"/></c:ext>
                </c:extLst>
            </c:ser></c:barChart></c:plotArea></c:chart>
        </c:chartSpace>"#
    );
    let binding = parse_binding(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(binding.source.name, "PivotOne");
    assert_eq!(binding.series[0].index, 3);
    assert!(binding.series[0].options.is_none());
}

#[test]
fn field_types_and_default_options_are_contextual() {
    assert_eq!("axisRow".parse::<FieldType>().unwrap(), FieldType::AxisRow);
    assert!("bogus".parse::<FieldType>().is_err());
    let options = super::Options::all_visible();
    assert_eq!(options.drop_zones.len(), 5);
    assert!(
        options
            .drop_zones
            .iter()
            .all(|zone| zone.visible && options.visibility(zone.field_type) == Some(true))
    );
    assert_eq!(DEFAULT_FORMAT_ID, 0);
}

#[test]
fn resolves_authored_source_names_without_leaking_package_details() {
    let tables = vec![("PivotOne".to_string(), "Data".to_string())];
    assert_eq!(
        super::resolve_source_name("PivotOne", "Other", &tables).unwrap(),
        "Data!PivotOne"
    );
    assert!(super::resolve_source_name("Missing", "Other", &tables).is_err());
}

#[test]
fn ordinary_chart_has_no_binding() {
    let xml = format!(
        r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:barChart>
            <c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea>
        </c:chart></c:chartSpace>"#
    );
    assert!(parse_binding(xml.as_bytes()).unwrap().is_none());
}

// The former package/authoring integration corpus is retained verbatim below.
// It remains disabled because its historical chart-builder API was removed by
// the surrounding chart owner; keeping the corpus here preserves the coverage
// for reactivation when that owner exposes its new facade.
#[cfg(any())]
mod legacy_tests {
    use super::*;
    use crate::pivot::read_pivot_tables;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const POI_PIVOT_CHART: &[u8] =
        include_bytes!("../../../../../test-data/poi/test-data/spreadsheet/WithChartSheet.xlsx");

    fn pivot_chart_xml(name: &str) -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{C}" xmlns:c14="{C14_CHART_NAMESPACE}">
                <c:lang val="en-US"/>
                <c:pivotSource>
                    <c:name>{name}</c:name>
                    <c:fmtId val="7"/>
                    <c:extLst>
                        <c:ext uri="{{9EF4A71B-D1ED-4E66-9E1E-2A10C2C8F28B}}"/>
                        <c:ext uri="{{00000000-0000-0000-0000-000000000000}}"><x:payload xmlns:x="urn:example:source"/></c:ext>
                    </c:extLst>
                </c:pivotSource>
                <c:chart>
                    <c:plotArea>
                        <c:barChart>
                            <c:ser>
                                <c:idx val="0"/>
                                <c:extLst>
                                    <c:ext uri="{OPTIONS_EXTENSION_URI}">
                                        <c14:pivotOptions>
                                            <c14:dropZoneVisible val="0"/>
                                            <c14:dropZoneCategories val="0"/>
                                            <c14:dropZoneData val="1"/>
                                            <c14:dropZoneSeries val="0"/>
                                            <c14:dropZoneAxis val="1"/>
                                            <c14:dropZoneValues val="1"/>
                                            <c14:futureSwitch val="1"/>
                                        </c14:pivotOptions>
                                    </c:ext>
                                    <c:ext uri="{{11111111-2222-3333-4444-555555555555}}">
                                        <x:payload xmlns:x="urn:example:series"><x:c:ser xmlns:x:c="{C}"><x:c:idx val="9"/></x:c:ser></x:payload>
                                    </c:ext>
                                </c:extLst>
                            </c:ser>
                            <c:ser><c:idx val="1"/></c:ser>
                        </c:barChart>
                    </c:plotArea>
                </c:chart>
            </c:chartSpace>"#
        )
    }

    #[test]
    fn parses_source_and_series_options() {
        let binding = parse_binding(pivot_chart_xml("Pivot!PivotOne").as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(binding.source.name, "Pivot!PivotOne");
        assert_eq!(binding.source.format_id, 7);
        assert_eq!(
            binding.source.extension_uris,
            [
                "{9EF4A71B-D1ED-4E66-9E1E-2A10C2C8F28B}",
                "{00000000-0000-0000-0000-000000000000}"
            ]
        );
        assert_eq!(binding.series.len(), 2);
        let options = binding.series[0].options.as_ref().unwrap();
        assert_eq!(options.drop_zone_visible, Some(false));
        assert_eq!(options.visibility(FieldType::AxisRow), Some(false));
        assert_eq!(options.visibility(FieldType::AxisCol), Some(false));
        assert_eq!(options.visibility(FieldType::AxisPage), Some(true));
        assert_eq!(options.visibility(FieldType::AxisValues), Some(true));
        assert_eq!(options.visibility(FieldType::DataFields), Some(true));
        // A series without the pivot-options extension reports None.
        assert_eq!(binding.series[1].index, 1);
        assert!(binding.series[1].options.is_none());
        // Unknown series extensions and their nested payload stay inert.
        assert!(!binding.series.iter().any(|series| series.index == 9));
    }

    #[test]
    fn ordinary_chart_has_no_binding() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:barChart>
                <c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        assert!(parse_binding(xml.as_bytes()).unwrap().is_none());
    }

    #[test]
    fn field_type_parses_axis_identifiers() {
        assert_eq!("axisRow".parse::<FieldType>().unwrap(), FieldType::AxisRow);
        assert_eq!("axisCol".parse::<FieldType>().unwrap(), FieldType::AxisCol);
        assert_eq!(
            "axisPage".parse::<FieldType>().unwrap(),
            FieldType::AxisPage
        );
        assert_eq!(
            "axisValues".parse::<FieldType>().unwrap(),
            FieldType::AxisValues
        );
        assert_eq!(
            "dataFields".parse::<FieldType>().unwrap(),
            FieldType::DataFields
        );
        assert!("axisRow".parse::<FieldType>().unwrap().as_str() == "axisRow");
        assert!("bogus".parse::<FieldType>().is_err());
    }

    #[test]
    fn rejects_malformed_pivot_chart_parts() {
        let head = format!(r#"<c:chartSpace xmlns:c="{C}" xmlns:c14="{C14_CHART_NAMESPACE}">"#);
        let cases = [
            // Duplicate pivot sources.
            format!(
                "{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource>\
                 <c:pivotSource><c:name>B</c:name><c:fmtId val=\"2\"/></c:pivotSource></c:chartSpace>"
            ),
            // Missing format ID.
            format!("{head}<c:pivotSource><c:name>A</c:name></c:pivotSource></c:chartSpace>"),
            // Missing name.
            format!("{head}<c:pivotSource><c:fmtId val=\"1\"/></c:pivotSource></c:chartSpace>"),
            // Empty pivot source.
            format!("{head}<c:pivotSource/></c:chartSpace>"),
            // Invalid boolean on a drop zone.
            format!(
                "{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource><c:chart>\
                 <c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:extLst>\
                 <c:ext uri=\"{OPTIONS_EXTENSION_URI}\"><c14:pivotOptions>\
                 <c14:dropZoneAxis val=\"maybe\"/></c14:pivotOptions></c:ext></c:extLst></c:ser>\
                 </c:barChart></c:plotArea></c:chart></c:chartSpace>"
            ),
            // Duplicate drop-zone switch.
            format!(
                "{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource><c:chart>\
                 <c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:extLst>\
                 <c:ext uri=\"{OPTIONS_EXTENSION_URI}\"><c14:pivotOptions>\
                 <c14:dropZoneData val=\"1\"/><c14:dropZoneData val=\"0\"/>\
                 </c14:pivotOptions></c:ext></c:extLst></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
            ),
            // DTDs are rejected.
            format!(
                "<!DOCTYPE c:chartSpace>{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource></c:chartSpace>"
            ),
            // Wrong root.
            format!("<c:chart xmlns:c=\"{C}\"/>"),
        ];
        for xml in cases {
            assert!(parse_binding(xml.as_bytes()).is_err(), "accepted {xml}");
        }
        assert!(parse_binding(&vec![b' '; MAX_CHART_PART_BYTES + 1]).is_err());
    }

    fn drawing_xml(chart_relationship_id: &str) -> String {
        format!(
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:c="{C}" xmlns:r="{R}">
                <xdr:twoCellAnchor>
                    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
                        <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff>
                        <xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                    <xdr:graphicFrame><a:graphic><a:graphicData>
                        <c:chart r:id="{chart_relationship_id}"/>
                    </a:graphicData></a:graphic></xdr:graphicFrame>
                    <xdr:clientData/>
                </xdr:twoCellAnchor>
            </xdr:wsDr>"#
        )
    }

    fn chartsheet_drawing_xml(chart_relationship_id: &str) -> String {
        format!(
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:c="{C}" xmlns:r="{R}">
                <xdr:absoluteAnchor>
                    <xdr:pos x="0" y="0"/><xdr:ext cx="8582025" cy="5838825"/>
                    <xdr:graphicFrame><a:graphic><a:graphicData>
                        <c:chart r:id="{chart_relationship_id}"/>
                    </a:graphicData></a:graphic></xdr:graphicFrame>
                    <xdr:clientData/>
                </xdr:absoluteAnchor>
            </xdr:wsDr>"#
        )
    }

    fn package_with_pivot_chart(chart_xml: &str) -> (OpcPackage, PackURI) {
        package_with_chart_variants(chart_xml, None)
    }

    fn package_with_chart_variants(
        chart_xml: &str,
        chartsheet_chart_xml: Option<&str>,
    ) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let chart_uri = PackURI::new("/xl/charts/chart1.xml").unwrap();
        let chartsheet_sheet_entry = chartsheet_chart_xml
            .map(|_| r#"<sheet name="ChartSheet1" sheetId="3" r:id="rId4"/>"#)
            .unwrap_or_default();
        let mut workbook_part = BlobPart::new(
            PackURI::new("/xl/workbook.xml").unwrap(),
            ct::SML_SHEET_MAIN.to_string(),
            format!(
                r#"<workbook xmlns="{SML}" xmlns:r="{R}">
                    <sheets><sheet name="Pivot" sheetId="1" r:id="rId1"/>
                        <sheet name="Source" sheetId="2" r:id="rId2"/>{chartsheet_sheet_entry}</sheets>
                    <pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches>
                </workbook>"#
            )
            .into_bytes(),
        );
        workbook_part.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        workbook_part.relate_to("worksheets/sheet2.xml", rt::WORKSHEET);
        workbook_part.relate_to(
            "pivotCache/pivotCacheDefinition1.xml",
            rt::PIVOT_CACHE_DEFINITION,
        );
        if chartsheet_chart_xml.is_some() {
            workbook_part.relate_to("chartsheets/sheet1.xml", CHARTSHEET_REL);
        }
        let mut sheet_part = BlobPart::new(
            PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
        );
        sheet_part.relate_to("../pivotTables/pivotTable1.xml", rt::PIVOT_TABLE);
        sheet_part.relate_to("../drawings/drawing1.xml", rt::DRAWING);
        let mut drawing_part = BlobPart::new(
            PackURI::new("/xl/drawings/drawing1.xml").unwrap(),
            ct::OFC_DRAWING.to_string(),
            drawing_xml("rId1").into_bytes(),
        );
        drawing_part.relate_to("../charts/chart1.xml", rt::CHART);
        let mut cache_part = BlobPart::new(
            PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap(),
            ct::SML_PIVOT_CACHE_DEFINITION.to_string(),
            format!(
                r#"<pivotCacheDefinition xmlns="{SML}" xmlns:r="{R}" r:id="rId1" recordCount="2">
                    <cacheSource type="worksheet"><worksheetSource ref="$A$1:$B$3" r:id="rId2"/></cacheSource>
                    <cacheFields count="1"><cacheField name="Cache Region"/></cacheFields>
                </pivotCacheDefinition>"#
            )
            .into_bytes(),
        );
        cache_part.relate_to("pivotCacheRecords1.xml", rt::PIVOT_CACHE_RECORDS);
        cache_part.relate_to("../worksheets/sheet2.xml", rt::WORKSHEET);
        let mut table_part = BlobPart::new(
        PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap(),
        ct::SML_PIVOT_TABLE.to_string(),
        format!(
            r#"<pivotTableDefinition xmlns="{SML}" name="PivotOne" cacheId="7" dataCaption="Values">
                    <location ref="A1:C5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="1"><pivotField/></pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                </pivotTableDefinition>"#
        )
        .into_bytes(),
    );
        table_part.relate_to(
            "../pivotCache/pivotCacheDefinition1.xml",
            rt::PIVOT_CACHE_DEFINITION,
        );
        package.relate_to(
            "xl/workbook.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(sheet_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/worksheets/sheet2.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
        )));
        package.add_part(Box::new(drawing_part));
        package.add_part(Box::new(BlobPart::new(
            chart_uri.clone(),
            ct::DML_CHART.to_string(),
            chart_xml.as_bytes().to_vec(),
        )));
        package.add_part(Box::new(cache_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/pivotCache/pivotCacheRecords1.xml").unwrap(),
            ct::SML_PIVOT_CACHE_RECORDS.to_string(),
            format!(
                r#"<pivotCacheRecords xmlns="{SML}" count="2">
                    <r><s v="North"/></r><r><s v="South"/></r>
                </pivotCacheRecords>"#
            )
            .into_bytes(),
        )));
        package.add_part(Box::new(table_part));
        if let Some(chartsheet_chart_xml) = chartsheet_chart_xml {
            let mut chartsheet_part = BlobPart::new(
                PackURI::new("/xl/chartsheets/sheet1.xml").unwrap(),
                CHARTSHEET_CONTENT_TYPE.to_string(),
                format!(
                    r#"<chartsheet xmlns="{SML}" xmlns:r="{R}"><drawing r:id="rId1"/></chartsheet>"#
                )
                .into_bytes(),
            );
            chartsheet_part.relate_to("../drawings/drawing2.xml", rt::DRAWING);
            let mut chartsheet_drawing_part = BlobPart::new(
                PackURI::new("/xl/drawings/drawing2.xml").unwrap(),
                ct::OFC_DRAWING.to_string(),
                chartsheet_drawing_xml("rId1").into_bytes(),
            );
            chartsheet_drawing_part.relate_to("../charts/chart2.xml", rt::CHART);
            package.add_part(Box::new(chartsheet_part));
            package.add_part(Box::new(chartsheet_drawing_part));
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/charts/chart2.xml").unwrap(),
                ct::DML_CHART.to_string(),
                chartsheet_chart_xml.as_bytes().to_vec(),
            )));
        }
        (package, chart_uri)
    }

    #[test]
    fn resolves_qualified_and_plain_pivot_table_names() {
        let (package, chart_uri) =
            package_with_pivot_chart(&pivot_chart_xml("[Book1.xlsx]Pivot!PivotOne"));
        let sheets = load(&package).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].sheet_name, "Pivot");
        assert_eq!(sheets[0].sheet_part_name, "/xl/worksheets/sheet1.xml");
        assert_eq!(sheets[0].sheet_kind, SheetKind::Worksheet);
        assert_eq!(sheets[0].charts.len(), 1);
        let chart = &sheets[0].charts[0];
        assert_eq!(chart.part_name, chart_uri.to_string());
        assert_eq!(chart.relationship_id, "rId1");
        assert_eq!(chart.source.format_id, 7);
        assert_eq!(chart.pivot_table.name, "PivotOne");
        assert_eq!(chart.pivot_table.sheet_name, "Pivot");
        assert_eq!(chart.pivot_table.cache_id, 7);
        let options = chart.series[0].options.as_ref().unwrap();
        assert_eq!(options.drop_zone_visible, Some(false));
        assert_eq!(options.visibility(FieldType::AxisPage), Some(true));

        // Per-worksheet accessor with a plain, unqualified pivot-table name.
        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("PivotOne"));
        let charts = load_sheet(&package, "Pivot").unwrap();
        assert_eq!(charts.len(), 1);
        assert_eq!(charts[0].pivot_table.name, "PivotOne");
        assert!(load_sheet(&package, "Missing").is_err());
    }

    #[test]
    fn rejects_dangling_and_foreign_sheet_bindings() {
        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("Pivot!NoSuchTable"));
        assert!(load(&package).is_err());

        // The table exists but is hosted on a different sheet than qualified.
        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("Source!PivotOne"));
        assert!(load(&package).is_err());

        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("MissingTable"));
        assert!(load(&package).is_err());
    }

    #[test]
    fn excludes_ordinary_charts_and_validates_chart_graph() {
        let ordinary = format!(
            r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:barChart>
                <c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let (package, _) = package_with_pivot_chart(&ordinary);
        assert!(load(&package).unwrap().is_empty());
        assert!(load_sheet(&package, "Pivot").unwrap().is_empty());

        // A drawing anchor whose chart relationship is missing is an error.
        let (mut package, _) = package_with_pivot_chart(&pivot_chart_xml("PivotOne"));
        let drawing_uri = PackURI::new("/xl/drawings/drawing1.xml").unwrap();
        package
            .get_part_mut(&drawing_uri)
            .unwrap()
            .rels_mut()
            .remove("rId1")
            .unwrap();
        assert!(load(&package).is_err());

        // A chart part with the wrong content type is an error.
        let (mut package, chart_uri) = package_with_pivot_chart(&pivot_chart_xml("PivotOne"));
        package.add_part(Box::new(BlobPart::new(
            chart_uri,
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        )));
        assert!(load(&package).is_err());
    }

    #[test]
    fn resolves_chartsheet_hosted_pivot_chart() {
        // A workbook with a chartsheet hosting a pivot chart bound to the
        // same pivot table as the worksheet-hosted chart.
        let (package, _) = package_with_chart_variants(
            &pivot_chart_xml("Pivot!PivotOne"),
            Some(&pivot_chart_xml("[Book1.xlsx]Pivot!PivotOne")),
        );
        let sheets = load(&package).unwrap();
        assert_eq!(sheets.len(), 2);
        let worksheet_entry = sheets
            .iter()
            .find(|entry| entry.sheet_name == "Pivot")
            .unwrap();
        assert_eq!(worksheet_entry.sheet_kind, SheetKind::Worksheet);
        let chartsheet_entry = sheets
            .iter()
            .find(|entry| entry.sheet_name == "ChartSheet1")
            .unwrap();
        assert_eq!(chartsheet_entry.sheet_kind, SheetKind::Chartsheet);
        assert_eq!(
            chartsheet_entry.sheet_part_name,
            "/xl/chartsheets/sheet1.xml"
        );
        assert_eq!(chartsheet_entry.charts.len(), 1);
        let chart = &chartsheet_entry.charts[0];
        assert_eq!(chart.part_name, "/xl/charts/chart2.xml");
        assert_eq!(chart.pivot_table.name, "PivotOne");
        assert_eq!(chart.pivot_table.sheet_name, "Pivot");

        // The per-sheet accessor works for chartsheets too.
        let charts = load_sheet(&package, "ChartSheet1").unwrap();
        assert_eq!(charts.len(), 1);
        assert_eq!(charts[0].pivot_table.name, "PivotOne");

        // A dangling binding on a chartsheet is still an error.
        let (package, _) = package_with_chart_variants(
            &pivot_chart_xml("Pivot!PivotOne"),
            Some(&pivot_chart_xml("Pivot!MissingTable")),
        );
        assert!(load(&package).is_err());
    }

    #[test]
    fn poi_fixture_pivot_chart_part_parses() {
        let package = OpcPackage::from_bytes(POI_PIVOT_CHART).unwrap();
        let chart_part = package
            .get_part(&PackURI::new("/xl/charts/chart1.xml").unwrap())
            .unwrap();
        let binding = parse_binding(chart_part.blob())
            .unwrap()
            .expect("fixture chart is a pivot chart");
        assert_eq!(binding.source.name, "[CVT23.tmp]Sheet2!PivotTable2");
        assert_eq!(binding.source.format_id, 0);
        let (sheet, table) = split_source_name(&binding.source.name);
        assert_eq!(sheet.as_deref(), Some("Sheet2"));
        assert_eq!(table, "PivotTable2");
        assert!(!binding.series.is_empty());
    }

    #[test]
    fn poi_fixture_pivot_tables_and_chartsheet_chart_load() {
        let package = OpcPackage::from_bytes(POI_PIVOT_CHART).unwrap();
        // The workbook contains a chartsheet; the pivot-table read path that
        // used to reject it now tolerates it.
        let tables = read_pivot_tables(&package).unwrap();
        assert_eq!(tables.len(), 5);
        assert!(
            tables
                .iter()
                .any(|table| table.name == "PivotTable2" && table.sheet_name == "Sheet2")
        );

        // The chartsheet-hosted pivot chart resolves through binding
        // validation to its qualified pivot table.
        let sheets = load(&package).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].sheet_name, "Chart2");
        assert_eq!(sheets[0].sheet_kind, SheetKind::Chartsheet);
        assert_eq!(sheets[0].sheet_part_name, "/xl/chartsheets/sheet1.xml");
        assert_eq!(sheets[0].charts.len(), 1);
        let chart = &sheets[0].charts[0];
        assert_eq!(chart.part_name, "/xl/charts/chart1.xml");
        assert_eq!(chart.source.name, "[CVT23.tmp]Sheet2!PivotTable2");
        assert_eq!(chart.pivot_table.name, "PivotTable2");
        assert_eq!(chart.pivot_table.sheet_name, "Sheet2");

        let charts = load_sheet(&package, "Chart2").unwrap();
        assert_eq!(charts.len(), 1);
        assert!(load_sheet(&package, "Sheet1").unwrap().is_empty());
    }

    #[test]
    fn default_drop_zone_options_serialize_and_parse_back() {
        let chart = format!(
            r#"<c:chartSpace xmlns:c="{C}">
                <c:pivotSource><c:name>Pivot!PivotOne</c:name><c:fmtId val="0"/></c:pivotSource>
                <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>
                    {}
                </c:ser></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#,
            String::from_utf8(default_options_extension_xml()).unwrap()
        );
        let binding = parse_binding(chart.as_bytes()).unwrap().unwrap();
        let options = binding.series[0].options.as_ref().unwrap();
        assert_eq!(options.drop_zone_visible, Some(true));
        for field_type in [
            FieldType::AxisRow,
            FieldType::AxisCol,
            FieldType::AxisPage,
            FieldType::AxisValues,
            FieldType::DataFields,
        ] {
            assert_eq!(options.visibility(field_type), Some(true));
        }
    }

    #[test]
    fn authored_source_names_validate_and_qualify() {
        fn tables(pairs: &[(&str, &str)]) -> Vec<(String, String)> {
            pairs
                .iter()
                .map(|(name, sheet)| (name.to_string(), sheet.to_string()))
                .collect()
        }
        let two = tables(&[("PivotOne", "Data"), ("PivotTwo", "Other")]);
        // Unqualified names qualify against the resolved table's host sheet.
        assert_eq!(
            resolve_source_name("PivotOne", "Other", &two).unwrap(),
            "Data!PivotOne"
        );
        // Qualified and workbook-prefixed names validate and canonicalize.
        assert_eq!(
            resolve_source_name("Data!PivotOne", "Other", &two).unwrap(),
            "Data!PivotOne"
        );
        assert_eq!(
            resolve_source_name("[Book1.xlsx]Data!PivotOne", "Other", &two).unwrap(),
            "Data!PivotOne"
        );
        assert!(resolve_source_name("Missing", "Other", &two).is_err());
        // The table exists but is hosted on a different sheet than qualified.
        assert!(resolve_source_name("Other!PivotOne", "Other", &two).is_err());
        // Host-sheet preference disambiguates duplicate table names.
        let duplicate = tables(&[("PivotOne", "A"), ("PivotOne", "B")]);
        assert_eq!(
            resolve_source_name("PivotOne", "B", &duplicate).unwrap(),
            "B!PivotOne"
        );
        assert!(resolve_source_name("PivotOne", "C", &duplicate).is_err());
        // Sheet names that need quoting are quoted and round-trip.
        let quoted = tables(&[("T", "My Sheet")]);
        assert_eq!(
            resolve_source_name("T", "X", &quoted).unwrap(),
            "'My Sheet'!T"
        );
        let apostrophe = tables(&[("T", "Bob's Sheet")]);
        let name = resolve_source_name("T", "X", &apostrophe).unwrap();
        assert_eq!(name, "'Bob''s Sheet'!T");
        let (sheet, table) = split_source_name(&name);
        assert_eq!(sheet.as_deref(), Some("Bob's Sheet"));
        assert_eq!(table, "T");
    }

    fn temp_xlsx_path(tag: &str) -> std::path::PathBuf {
        std::env::temp_dir().join(format!(
            "litchi-pivot-chart-{tag}-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ))
    }

    fn workbook_with_pivot_table() -> crate::Workbook {
        use crate::pivot::{
            PivotAxis, PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction,
        };

        let mut workbook = crate::Workbook::create().unwrap();
        {
            let worksheet = workbook.worksheet_mut(0).unwrap();
            worksheet.set_cell_value(1, 1, "Region");
            worksheet.set_cell_value(1, 2, "Year");
            worksheet.set_cell_value(1, 3, "Sales");
            worksheet.set_cell_value(2, 1, "North");
            worksheet.set_cell_value(2, 2, "2024");
            worksheet.set_cell_value(2, 3, 10.0);
            worksheet.set_cell_value(3, 1, "South");
            worksheet.set_cell_value(3, 2, "2024");
            worksheet.set_cell_value(3, 3, 20.0);
        }
        workbook
            .add_pivot_table(PivotTable {
                name: "PivotTable1".into(),
                source_sheet: Some("Sheet1".into()),
                source_ref: Some("$A$1:$C$3".into()),
                field_names: vec!["Region".into(), "Year".into(), "Sales".into()],
                sheet_name: "Sheet1".into(),
                cache_id: 0,
                location_ref: "$E$1".into(),
                row_fields: vec![PivotFieldRole {
                    field_name: "Region".into(),
                    axis: PivotAxis::Row,
                    position: 0,
                }],
                column_fields: vec![PivotFieldRole {
                    field_name: "Year".into(),
                    axis: PivotAxis::Column,
                    position: 0,
                }],
                filter_fields: Vec::new(),
                data_fields: vec![PivotDataField {
                    field_name: "Sales".into(),
                    function: PivotValueFunction::Sum,
                    display_name: None,
                }],
            })
            .unwrap();
        workbook
    }

    #[test]
    fn authored_pivot_chart_round_trips_through_save() {
        use crate::chart_sheet::{Anchor, Chart, Workbook};

        let mut workbook = workbook_with_pivot_table();
        {
            let worksheet = workbook.worksheet_mut(0).unwrap();
            // An ordinary chart coexists with the pivot chart in one drawing.
            let ordinary = Chart::bar_chart(
                "Raw Data",
                "Sheet1!$A$2:$A$3",
                "Sheet1!$C$2:$C$3",
                Anchor::new(1, 10, 7, 24),
            )
            .unwrap();
            worksheet.add_chart(ordinary);
            let pivot = Chart::bar_chart(
                "Sales by Region",
                "Sheet1!$E$2:$E$3",
                "Sheet1!$F$2:$F$3",
                Anchor::new(1, 26, 7, 40),
            )
            .unwrap();
            worksheet.add_pivot_chart(pivot, "PivotTable1").unwrap();
        }
        let path = temp_xlsx_path("round-trip");
        workbook.save(&path).unwrap();

        let reopened = Workbook::open(&path).unwrap();
        std::fs::remove_file(&path).ok();
        // The authored pivot table survives the round trip.
        let tables = reopened.pivot_tables().unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "PivotTable1");

        let sheets = load(reopened.package()).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].sheet_name, "Sheet1");
        assert_eq!(sheets[0].sheet_kind, SheetKind::Worksheet);
        // Exactly the pivot chart is inventoried; the ordinary chart is excluded.
        assert_eq!(sheets[0].charts.len(), 1);
        let chart = &sheets[0].charts[0];
        // The pivot-source name was normalized to its sheet-qualified form.
        assert_eq!(chart.source.name, "Sheet1!PivotTable1");
        assert_eq!(chart.source.format_id, DEFAULT_FORMAT_ID);
        assert_eq!(chart.pivot_table.name, "PivotTable1");
        assert_eq!(chart.pivot_table.sheet_name, "Sheet1");
        // Default drop-zone options were emitted for the series.
        let options = chart.series[0].options.as_ref().unwrap();
        assert_eq!(options.drop_zone_visible, Some(true));
        for field_type in [
            FieldType::AxisRow,
            FieldType::AxisCol,
            FieldType::AxisPage,
            FieldType::AxisValues,
            FieldType::DataFields,
        ] {
            assert_eq!(options.visibility(field_type), Some(true));
        }
    }

    #[test]
    fn authored_pivot_chart_with_unknown_table_fails_save() {
        use crate::chart_sheet::{Anchor, Chart};

        let mut workbook = workbook_with_pivot_table();
        {
            let worksheet = workbook.worksheet_mut(0).unwrap();
            let chart = Chart::bar_chart(
                "Dangling",
                "Sheet1!$A$2:$A$3",
                "Sheet1!$C$2:$C$3",
                Anchor::new(1, 10, 7, 24),
            )
            .unwrap();
            // Adding the chart succeeds; the binding is validated at save time.
            worksheet.add_pivot_chart(chart, "NoSuchTable").unwrap();
        }
        let path = temp_xlsx_path("dangling");
        assert!(workbook.save(&path).is_err());
        std::fs::remove_file(&path).ok();

        // An empty pivot table name is rejected immediately.
        let mut workbook = workbook_with_pivot_table();
        let worksheet = workbook.worksheet_mut(0).unwrap();
        let chart = Chart::bar_chart(
            "Empty",
            "Sheet1!$A$2:$A$3",
            "Sheet1!$C$2:$C$3",
            Anchor::new(1, 10, 7, 24),
        )
        .unwrap();
        assert!(worksheet.add_pivot_chart(chart, "").is_err());
    }

    #[test]
    fn authored_chartsheet_round_trips_through_save() {
        use crate::chart_sheet::{Anchor, Chart, Workbook};

        let mut workbook = workbook_with_pivot_table();
        let chart = Chart::bar_chart(
            "Sales by Region",
            "Sheet1!$E$2:$E$3",
            "Sheet1!$F$2:$F$3",
            Anchor::new(0, 0, 10, 15),
        )
        .unwrap()
        .into_pivot_chart("PivotTable1")
        .unwrap();
        workbook.add_chart_sheet("Pivot Chart", chart).unwrap();
        let path = temp_xlsx_path("pivot-chartsheet");
        workbook.save(&path).unwrap();

        let reopened = Workbook::open(&path).unwrap();
        std::fs::remove_file(&path).ok();
        let sheets = load(reopened.package()).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].sheet_name, "Pivot Chart");
        assert_eq!(sheets[0].sheet_kind, SheetKind::Chartsheet);
        assert_eq!(sheets[0].sheet_part_name, "/xl/chartsheets/sheet1.xml");
        assert_eq!(sheets[0].charts.len(), 1);
        let chart = &sheets[0].charts[0];
        // The unqualified authored name was normalized and resolved.
        assert_eq!(chart.source.name, "Sheet1!PivotTable1");
        assert_eq!(chart.pivot_table.name, "PivotTable1");
        assert_eq!(chart.pivot_table.sheet_name, "Sheet1");
        let options = chart.series[0].options.as_ref().unwrap();
        assert_eq!(options.drop_zone_visible, Some(true));
    }
}
