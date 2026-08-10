//! Regression coverage for the chartsheet package graph.

use super::codec::{
    DrawingChartKind, DrawingChartReference, drawing_chart_references,
    validate_chart_companion_xml, validate_chart_ex_relationships,
};
use super::model::{
    BackgroundImageContentType, BackgroundPicture, ChartCompanionResource,
    ChartEmbeddedPackageContentType, ChartEmbeddedPackageResource, ChartOutboundResource,
    ChartResource, ChartResourceKind, ChartThemeOverrideResource, ChartUserShapesResource,
    DrawingResource, Entry, ExtensionRelationship, ExtensionRelationshipTarget, ImageContentType,
    ImageResource, Package, PrinterSettings, VmlDrawingResource,
};
use super::operations::{load_chartsheet, store_chartsheet, validate_package};
use super::{
    CHART, CHART_COLOR_STYLE_CT, CHART_CT, CHART_EX, CHART_EX_CHOICE, CHART_EX_CT, CHART_EX_REL,
    CHART_STYLE, CHART_STYLE_CT, CHART_USER_SHAPES_CT, CHARTSHEET_REL, DRAWING_CT, DRAWING_MAIN,
    IMAGE_REL, MAX_BACKGROUND_IMAGE_BYTES, MAX_CHART_EMBEDDED_PACKAGE_BYTES, MAX_CHART_EX_BYTES,
    MAX_CHART_STYLE_BYTES, MAX_CHART_USER_SHAPE_IMAGE_BYTES, MAX_CHARTS,
    MAX_EXTENSION_PAYLOAD_BYTES, MAX_EXTENSION_RELATIONSHIP_STRING_BYTES, MAX_EXTENSION_URI_BYTES,
    MAX_EXTENSIONS, MAX_VML_DRAWING_BYTES, MAX_WEB_PUBLISH_ITEMS, MAX_WEB_PUBLISH_STRING_BYTES,
    MAX_XML_BYTES, REL, SML, STRICT_CHART, STRICT_DRAWING_MAIN, STRICT_REL, STRICT_SML, STRICT_XDR,
    THEME_OVERRIDE_CT, VML_DRAWING_CT, VML_DRAWING_REL, XDR,
};
use crate::chart_sheet::{
    Chart, Color, Conformance, CustomView, Extension, ExtensionList, HeaderFooter, Margins,
    PageOrientation, PageSetup, Properties, Protection, State, View, WebPublishItem,
    WebPublishItems, WebSourceType, parse_chartsheet, write_chartsheet,
};
use crate::package::printer_settings::{
    MAX_SETTINGS_BYTES, PRINTER_CT, PRINTER_REL, PrinterSettingsResource,
};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI};

const POI_ONE: &[u8] =
    include_bytes!("../../../../../test-data/poi/test-data/spreadsheet/WithChartSheet.xlsx");
const POI_TWO: &[u8] =
    include_bytes!("../../../../../test-data/poi/test-data/spreadsheet/chart_sheet.xlsx");
const LO_CHART_EX: &[u8] = include_bytes!(
    "../../../../../test-data/libreoffice-core/chart2/qa/extras/data/xlsx/boxWhisker.xlsx"
);
pub(super) const LO_USER_SHAPES_IMAGES: &[u8] = include_bytes!(
    "../../../../../test-data/libreoffice-core/chart2/qa/extras/data/xlsx/tdf143127.xlsx"
);
fn sheet() -> Chart {
    Chart {
        properties: Some(Properties {
            published: Some(true),
            code_name: Some("ChartCode".into()),
            tab_color: Some(Color {
                automatic: None,
                indexed: None,
                rgb: Some("FF336699".into()),
                theme: None,
                tint: Some(0.25),
            }),
        }),
        views: vec![View {
            tab_selected: Some(true),
            zoom_scale: Some(125),
            workbook_view_id: 0,
            zoom_to_fit: Some(false),
        }],
        protection: Some(Protection {
            password_hash: Some("ABCD".into()),
            content: Some(true),
            objects: Some(false),
        }),
        custom_views: Some(vec![
            CustomView {
                guid: "{00112233-4455-6677-8899-AABBCCDDEEFF}".into(),
                scale: Some(175),
                state: Some(State::Hidden),
                zoom_to_fit: Some(true),
            },
            CustomView {
                guid: "{10213243-5465-7687-98A9-BACBDCEDFE0F}".into(),
                scale: None,
                state: None,
                zoom_to_fit: Some(false),
            },
        ]),
        margins: Some(Margins {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }),
        page_setup: Some(PageSetup {
            paper_size: Some(1),
            first_page_number: Some(1),
            orientation: Some(PageOrientation::Landscape),
            use_printer_defaults: Some(true),
            black_and_white: Some(false),
            draft: Some(false),
            use_first_page_number: Some(true),
            horizontal_dpi: Some(600),
            vertical_dpi: Some(600),
            copies: Some(1),
            printer_settings_relationship_id: Some("rIdPrinter".into()),
        }),
        header_footer: Some(HeaderFooter {
            align_with_margins: Some(false),
            odd_header: Some("&CChart & Report".into()),
            ..Default::default()
        }),
        drawing_relationship_id: "rIdDrawing".into(),
        legacy_drawing_relationship_id: Some("rIdLegacy".into()),
        legacy_header_footer_drawing_relationship_id: Some("rIdLegacyHF".into()),
        background_picture_relationship_id: Some("rIdBackground".into()),
        web_publish_items: Some(WebPublishItems {
            count: Some(2),
            items: vec![
                WebPublishItem {
                    id: 11289,
                    div_id: "Views_11289".into(),
                    source_type: WebSourceType::Range,
                    source_ref: Some("A6:C6".into()),
                    source_object: None,
                    destination_file: "file:///definitely/not/accessed/Publish.htm".into(),
                    title: Some("Range & title".into()),
                    auto_republish: Some(false),
                },
                WebPublishItem {
                    id: 6433,
                    div_id: "Views_6433".into(),
                    source_type: WebSourceType::Chart,
                    source_ref: None,
                    source_object: Some("https://example.invalid/Chart 1".into()),
                    destination_file: "https://example.invalid/Publish.mht".into(),
                    title: None,
                    auto_republish: None,
                },
            ],
        }),
        extension_list: None,
    }
}
fn drawing(conformance: Conformance) -> Vec<u8> {
    format!("<xdr:wsDr xmlns:xdr=\"{}\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><xdr:absoluteAnchor><a:graphic><a:graphicData><c:chart xmlns:c=\"{}\" xmlns:r=\"{}\" r:id=\"rIdChart\"/></a:graphicData></a:graphic></xdr:absoluteAnchor></xdr:wsDr>", conformance.xdr(), conformance.chart(), conformance.rel()).into_bytes()
}
fn chart(conformance: Conformance) -> Vec<u8> {
    format!(
        "<c:chartSpace xmlns:c=\"{}\"><c:chart/></c:chartSpace>",
        conformance.chart()
    )
    .into_bytes()
}
fn vml(id: &str, name: &str) -> VmlDrawingResource {
    VmlDrawingResource{relationship_id:id.into(),part_name:format!("/xl/drawings/{name}.vml"),content_type:VML_DRAWING_CT.into(),data:format!("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape href=\"https://example.invalid/{name}\"/></xml>").into_bytes()}
}
fn value(conformance: Conformance) -> Package {
    Package {
        entry: Entry {
            name: "Chart 1".into(),
            sheet_id: 2,
            state: State::Visible,
            workbook_relationship_id: "rIdChartSheet".into(),
            part_name: "/xl/chartsheets/sheet1.xml".into(),
        },
        chartsheet: sheet(),
        drawing: DrawingResource {
            part_name: "/xl/drawings/drawing1.xml".into(),
            content_type: DRAWING_CT.into(),
            data: drawing(conformance),
            charts: vec![ChartResource {
                relationship_id: "rIdChart".into(),
                part_name: "/xl/charts/chart1.xml".into(),
                content_type: CHART_CT.into(),
                data: chart(conformance),
                kind: ChartResourceKind::Classic,
            }],
        },
        legacy_drawing: Some(vml("rIdLegacy", "vmlDrawing1")),
        legacy_header_footer_drawing: Some(vml("rIdLegacyHF", "vmlDrawing2")),
        background_picture: Some(BackgroundPicture {
            relationship_id: "rIdBackground".into(),
            part_name: "/xl/media/background1.png".into(),
            content_type: BackgroundImageContentType::Png,
            data: vec![0, 255, 1, 254],
        }),
        printer_settings: Some(PrinterSettings {
            relationship_id: "rIdPrinter".into(),
            resource: PrinterSettingsResource {
                part_name: "/xl/printerSettings/printerSettings1.bin".into(),
                data: vec![0x44, 0x45, 0x56, 0x4d, 0x4f, 0x44, 0x45, 0, 255],
            },
        }),
        extension_relationships: vec![],
    }
}

#[test]
fn validates_typed_package_graph_without_store_side_effects() {
    let conformance = Conformance::Transitional;
    let expected = value(conformance);
    validate_package(&expected, conformance).unwrap();

    let mut invalid = expected.clone();
    invalid.chartsheet.drawing_relationship_id = "rId drawing".into();
    assert!(validate_package(&invalid, conformance).is_err());
}

pub(super) fn base_package(conformance: Conformance) -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let uri = PackURI::new("/xl/workbook.xml").unwrap();
    let xml = format!(
        "<x:workbook xmlns:x=\"{}\" xmlns:r=\"{}\"><x:sheets><x:sheet name=\"Data\" sheetId=\"1\" r:id=\"rIdData\"/></x:sheets></x:workbook>",
        conformance.sml(),
        conformance.rel()
    );
    package.add_part(Box::new(BlobPart::new(
        uri.clone(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
        xml.into_bytes(),
    )));
    (package, uri)
}

fn ext(uri: &str, payload: &str) -> Extension {
    Extension {
        uri: uri.into(),
        payload_xml: payload.as_bytes().to_vec(),
    }
}
fn companion(id: &str, path: &str, content_type: &str, data: &[u8]) -> ChartCompanionResource {
    ChartCompanionResource {
        relationship_id: id.into(),
        part_name: path.into(),
        content_type: content_type.into(),
        data: data.to_vec(),
    }
}

#[test]
fn real_libreoffice_chart_ex_fixture_round_trips_as_inert_resources() {
    let source = OpcPackage::from_bytes(LO_CHART_EX).unwrap();
    let blob = |path: &str| {
        source
            .get_part(&PackURI::new(path).unwrap())
            .unwrap()
            .blob()
            .to_vec()
    };
    let mut expected = value(Conformance::Transitional);
    expected.drawing.data = blob("/xl/drawings/drawing1.xml");
    expected.drawing.charts = vec![ChartResource {
        relationship_id: "rId1".into(),
        part_name: "/xl/charts/chartEx1.xml".into(),
        content_type: CHART_EX_CT.into(),
        data: blob("/xl/charts/chartEx1.xml"),
        kind: ChartResourceKind::Extended {
            styles: vec![companion(
                "rId1",
                "/xl/charts/style1.xml",
                CHART_STYLE_CT,
                &blob("/xl/charts/style1.xml"),
            )],
            color_styles: vec![companion(
                "rId2",
                "/xl/charts/colors1.xml",
                CHART_COLOR_STYLE_CT,
                &blob("/xl/charts/colors1.xml"),
            )],
            user_shapes: None,
            outbound_resources: vec![],
        },
    }];
    let (mut package, workbook) = base_package(Conformance::Transitional);
    store_chartsheet(
        &mut package,
        &workbook,
        &expected,
        Conformance::Transitional,
    )
    .unwrap();
    let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
    assert_eq!(loaded, expected);
    assert!(matches!(
        loaded.drawing.charts[0].kind,
        ChartResourceKind::Extended { .. }
    ));
}

#[test]
fn chart_ex_strict_mce_selects_extended_choice_and_preserves_classic_fallback_behavior() {
    let strict = format!(
        "<xdr:wsDr xmlns:xdr=\"{STRICT_XDR}\" xmlns:a=\"{STRICT_DRAWING_MAIN}\" xmlns:c=\"{STRICT_CHART}\" xmlns:r=\"{STRICT_REL}\" xmlns:cx=\"{CHART_EX}\" xmlns:cx1=\"{CHART_EX_CHOICE}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:AlternateContent><mc:Choice Requires=\"cx1\"><a:graphic><a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rIdExtended\"/></a:graphicData></a:graphic></mc:Choice><mc:Fallback><a:graphic><a:graphicData uri=\"{STRICT_CHART}\"><c:chart r:id=\"rIdClassic\"/></a:graphicData></a:graphic></mc:Fallback></mc:AlternateContent></xdr:wsDr>"
    );
    let references = drawing_chart_references(strict.as_bytes(), Conformance::Strict).unwrap();
    assert_eq!(
        references,
        vec![DrawingChartReference {
            relationship_id: "rIdExtended".into(),
            kind: DrawingChartKind::Extended
        }]
    );
    let fallback = strict.replace(
        &format!("xmlns:cx1=\"{CHART_EX_CHOICE}\""),
        "xmlns:cx1=\"urn:unsupported-chart-version\"",
    );
    let references = drawing_chart_references(fallback.as_bytes(), Conformance::Strict).unwrap();
    assert_eq!(
        references,
        vec![DrawingChartReference {
            relationship_id: "rIdClassic".into(),
            kind: DrawingChartKind::Classic
        }]
    );
}

#[test]
fn rejects_chart_ex_drawing_shape_roots_cardinality_relationships_and_caps() {
    let drawing = |body: &str| {
        format!(
            "<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:a=\"{DRAWING_MAIN}\" xmlns:r=\"{REL}\" xmlns:cx=\"{CHART_EX}\">{body}</xdr:wsDr>"
        )
    };
    for body in [
        "<a:graphicData uri=\"urn:wrong\"><cx:chart r:id=\"rId1\"/></a:graphicData>".to_string(),
        format!("<a:graphicData uri=\"{CHART_EX}\"><cx:chart/></a:graphicData>"),
        format!(
            "<a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rId1\"/><cx:chart r:id=\"rId2\"/></a:graphicData>"
        ),
        format!("<a:graphicData uri=\"{CHART_EX}\"><cx:wrong r:id=\"rId1\"/></a:graphicData>"),
        format!(
            "<a:graphicData uri=\"{CHART_EX}\" bad=\"1\"><cx:chart r:id=\"rId1\"/></a:graphicData>"
        ),
    ] {
        assert!(
            drawing_chart_references(drawing(&body).as_bytes(), Conformance::Transitional).is_err(),
            "accepted {body}"
        );
    }
    assert!(
        validate_chart_ex_relationships(
            b"<cx:chartSpace xmlns:cx=\"urn:wrong\"/>",
            Conformance::Transitional
        )
        .is_err()
    );
    assert!(validate_chart_companion_xml(b"<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\"/>","chartStyle",MAX_CHART_STYLE_BYTES).is_err());
    assert!(
        validate_chart_ex_relationships(&[b' '; MAX_CHART_EX_BYTES + 1], Conformance::Transitional)
            .is_err()
    );
    let mut bad = value(Conformance::Transitional);
    bad.drawing.data = drawing(&format!(
        "<a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rIdEx\"/></a:graphicData>"
    ))
    .into_bytes();
    bad.drawing.charts = vec![ChartResource {
        relationship_id: "rIdEx".into(),
        part_name: "/xl/charts/chartEx1.xml".into(),
        content_type: CHART_EX_CT.into(),
        data: format!("<cx:chartSpace xmlns:cx=\"{CHART_EX}\"/>").into_bytes(),
        kind: ChartResourceKind::Extended {
            styles: vec![companion(
                "rIdSame",
                "/xl/charts/style1.xml",
                CHART_STYLE_CT,
                &format!("<cs:chartStyle xmlns:cs=\"{CHART_STYLE}\"/>").into_bytes(),
            )],
            color_styles: vec![companion(
                "rIdSame",
                "/xl/charts/colors1.xml",
                CHART_COLOR_STYLE_CT,
                &format!("<cs:colorStyle xmlns:cs=\"{CHART_STYLE}\"/>").into_bytes(),
            )],
            user_shapes: None,
            outbound_resources: vec![],
        },
    }];
    let (mut package, workbook) = base_package(Conformance::Transitional);
    assert!(store_chartsheet(&mut package, &workbook, &bad, Conformance::Transitional).is_err());
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_err()
    );
}

#[test]
fn rejects_chart_ex_wrong_type_orphan_escape_outbound_and_companion_graphs() {
    let conformance = Conformance::Transitional;
    let source = OpcPackage::from_bytes(LO_CHART_EX).unwrap();
    let blob = |path: &str| {
        source
            .get_part(&PackURI::new(path).unwrap())
            .unwrap()
            .blob()
            .to_vec()
    };
    let mut expected = value(conformance);
    expected.drawing.data = blob("/xl/drawings/drawing1.xml");
    expected.drawing.charts = vec![ChartResource {
        relationship_id: "rId1".into(),
        part_name: "/xl/charts/chartEx1.xml".into(),
        content_type: CHART_EX_CT.into(),
        data: blob("/xl/charts/chartEx1.xml"),
        kind: ChartResourceKind::Extended {
            styles: vec![companion(
                "rId1",
                "/xl/charts/style1.xml",
                CHART_STYLE_CT,
                &blob("/xl/charts/style1.xml"),
            )],
            color_styles: vec![companion(
                "rId2",
                "/xl/charts/colors1.xml",
                CHART_COLOR_STYLE_CT,
                &blob("/xl/charts/colors1.xml"),
            )],
            user_shapes: None,
            outbound_resources: vec![],
        },
    }];
    for (kind, target, external) in [
        (rt::CHART, "../charts/chartEx1.xml", false),
        (CHART_EX_REL, "../../../evil.xml", false),
        (CHART_EX_REL, "https://example.invalid/chartEx.xml", true),
    ] {
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let drawing = package
            .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap())
            .unwrap();
        drawing.rels_mut().remove("rId1");
        drawing
            .rels_mut()
            .add_relationship(kind.into(), target.into(), "rId1".into(), external);
        assert!(
            load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
            "accepted {kind} {target}"
        );
    }
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            CHART_EX_REL.into(),
            "../charts/chartEx1.xml".into(),
            "rIdOrphan".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/charts/chartEx1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            IMAGE_REL.into(),
            "../media/image1.png".into(),
            "rIdImage".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/charts/style1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            IMAGE_REL.into(),
            "../media/image1.png".into(),
            "rIdOutbound".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
}

pub(super) fn chart_ex_user_shapes(conformance: Conformance) -> Package {
    let source = OpcPackage::from_bytes(LO_USER_SHAPES_IMAGES).unwrap();
    let blob = |path: &str| {
        source
            .get_part(&PackURI::new(path).unwrap())
            .unwrap()
            .blob()
            .to_vec()
    };
    let mut value = value(conformance);
    value.drawing.data=format!("<xdr:wsDr xmlns:xdr=\"{}\" xmlns:a=\"{}\" xmlns:r=\"{}\" xmlns:cx=\"{CHART_EX}\"><a:graphic><a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rIdChartEx\"/></a:graphicData></a:graphic></xdr:wsDr>",conformance.xdr(),if conformance==Conformance::Strict{STRICT_DRAWING_MAIN}else{DRAWING_MAIN},conformance.rel()).into_bytes();
    let user_shapes_data = if conformance == Conformance::Transitional {
        blob("/xl/drawings/drawing2.xml")
    } else {
        String::from_utf8(blob("/xl/drawings/drawing2.xml"))
            .unwrap()
            .replace(CHART, STRICT_CHART)
            .replace(
                "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                "http://purl.oclc.org/ooxml/drawingml/chartDrawing",
            )
            .replace(DRAWING_MAIN, STRICT_DRAWING_MAIN)
            .replace(REL, STRICT_REL)
            .into_bytes()
    };
    value.drawing.charts = vec![ChartResource {
        relationship_id: "rIdChartEx".into(),
        part_name: "/xl/charts/chartEx1.xml".into(),
        content_type: CHART_EX_CT.into(),
        data: format!("<cx:chartSpace xmlns:cx=\"{CHART_EX}\"/>").into_bytes(),
        kind: ChartResourceKind::Extended {
            styles: vec![],
            color_styles: vec![],
            user_shapes: Some(ChartUserShapesResource {
                relationship_id: "rIdUserShapes".into(),
                part_name: "/xl/drawings/chartDrawing1.xml".into(),
                content_type: CHART_USER_SHAPES_CT.into(),
                data: user_shapes_data,
                images: vec![
                    ImageResource {
                        relationship_id: "rId1".into(),
                        part_name: "/xl/media/image1.png".into(),
                        content_type: ImageContentType::Png,
                        data: blob("/xl/media/image1.png"),
                    },
                    ImageResource {
                        relationship_id: "rId2".into(),
                        part_name: "/xl/media/image2.svg".into(),
                        content_type: ImageContentType::Svg,
                        data: blob("/xl/media/image2.svg"),
                    },
                ],
            }),
            outbound_resources: vec![],
        },
    }];
    value
}
#[test]
fn chart_ex_user_shapes_png_svg_transitional_and_strict_round_trip() {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let expected = chart_ex_user_shapes(conformance);
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
        assert_eq!(loaded, expected);
        let chart = package
            .get_part(&PackURI::new("/xl/charts/chartEx1.xml").unwrap())
            .unwrap();
        assert_eq!(
            chart.rels().get("rIdUserShapes").unwrap().reltype(),
            conformance.chart_user_shapes_rel()
        );
        let shapes = package
            .get_part(&PackURI::new("/xl/drawings/chartDrawing1.xml").unwrap())
            .unwrap();
        assert!(
            shapes
                .rels()
                .iter()
                .all(|relationship| relationship.reltype() == conformance.image_rel())
        );
    }
}
#[test]
fn chart_ex_user_shapes_rejects_graph_mime_namespace_collision_and_caps() {
    let conformance = Conformance::Transitional;
    let mut bad = chart_ex_user_shapes(conformance);
    if let ChartResourceKind::Extended {
        user_shapes: Some(shapes),
        ..
    } = &mut bad.drawing.charts[0].kind
    {
        shapes.images[0].content_type = ImageContentType::Gif;
    }
    let (mut package, workbook) = base_package(conformance);
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let mut bad = chart_ex_user_shapes(conformance);
    if let ChartResourceKind::Extended {
        user_shapes: Some(shapes),
        ..
    } = &mut bad.drawing.charts[0].kind
    {
        shapes.images.pop();
    }
    let (mut package, workbook) = base_package(conformance);
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let mut bad = chart_ex_user_shapes(Conformance::Strict);
    if let ChartResourceKind::Extended {
        user_shapes: Some(shapes),
        ..
    } = &mut bad.drawing.charts[0].kind
    {
        shapes.data = String::from_utf8(std::mem::take(&mut shapes.data))
            .unwrap()
            .replace(STRICT_CHART, CHART)
            .into_bytes();
    }
    let (mut package, workbook) = base_package(Conformance::Strict);
    assert!(store_chartsheet(&mut package, &workbook, &bad, Conformance::Strict).is_err());
    let mut bad = chart_ex_user_shapes(conformance);
    if let ChartResourceKind::Extended {
        styles,
        user_shapes: Some(shapes),
        ..
    } = &mut bad.drawing.charts[0].kind
    {
        styles.push(companion(
            &shapes.relationship_id,
            "/xl/charts/style1.xml",
            CHART_STYLE_CT,
            format!("<cs:chartStyle xmlns:cs=\"{CHART_STYLE}\"/>").as_bytes(),
        ));
    }
    let (mut package, workbook) = base_package(conformance);
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let mut bad = chart_ex_user_shapes(conformance);
    if let ChartResourceKind::Extended {
        user_shapes: Some(shapes),
        ..
    } = &mut bad.drawing.charts[0].kind
    {
        shapes.images[0].data = vec![0; MAX_CHART_USER_SHAPE_IMAGE_BYTES + 1];
    }
    let (mut package, workbook) = base_package(conformance);
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
}
fn with_extension_relationships(conformance: Conformance) -> Package {
    let mut value = value(conformance);
    value.chartsheet.extension_list = Some(ExtensionList {
        extensions: vec![
            ext(
                "urn:duplicate",
                &format!(
                    "<u:payload xmlns:u=\"urn:vendor\" xmlns:r=\"{}\" r:id=\"rIdExtInternal\">before<u:child/>after</u:payload>",
                    conformance.rel()
                ),
            ),
            ext(
                "urn:duplicate",
                &format!(
                    "<v:external xmlns:v=\"urn:vendor-two\" xmlns:r=\"{}\" r:link=\"rIdExtExternal\"/>",
                    conformance.rel()
                ),
            ),
        ],
    });
    value.extension_relationships = vec![
        ExtensionRelationship {
            relationship_id: "rIdExtInternal".into(),
            relationship_type: "urn:relationship:internal".into(),
            target: ExtensionRelationshipTarget::Internal {
                part_name: "/xl/custom/ext.bin".into(),
            },
        },
        ExtensionRelationship {
            relationship_id: "rIdExtExternal".into(),
            relationship_type: "urn:relationship:external".into(),
            target: ExtensionRelationshipTarget::External {
                target: "https://example.invalid/not-fetched".into(),
            },
        },
    ];
    value
        .extension_relationships
        .sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    let xml = write_chartsheet(&value.chartsheet, conformance).unwrap();
    value.chartsheet = parse_chartsheet(&xml).unwrap().1;
    value
}

#[test]
fn ext_list_strict_mce_duplicate_uri_and_deterministic_round_trip() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{STRICT_SML}\" xmlns:r=\"{STRICT_REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:z=\"urn:unsupported-choice\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdD\"/><mc:AlternateContent><mc:Choice Requires=\"z\"><z:ignored/></mc:Choice><mc:Fallback><x:extLst><x:ext uri=\"urn:same\"><u:payload xmlns:u=\"urn:vendor\" r:id=\"rIdExt\">before<u:child a=\"1\"/>after</u:payload></x:ext><x:ext uri=\"urn:same\"><v:other xmlns:v=\"urn:vendor-two\"/></x:ext></x:extLst></mc:Fallback></mc:AlternateContent></x:chartsheet>"
    );
    let (kind, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
    assert_eq!(kind, Conformance::Strict);
    let extensions = &parsed.extension_list.as_ref().unwrap().extensions;
    assert_eq!(extensions.len(), 2);
    assert_eq!(extensions[0].uri, extensions[1].uri);
    let payload = std::str::from_utf8(&extensions[0].payload_xml).unwrap();
    assert!(payload.contains("before<e0:child a=\"1\"/>after"));
    assert!(payload.contains("r:id=\"rIdExt\""));
    let first = write_chartsheet(&parsed, kind).unwrap();
    let reparsed = parse_chartsheet(&first).unwrap().1;
    let second = write_chartsheet(&reparsed, kind).unwrap();
    assert_eq!(first, second);
    assert_eq!(parsed, reparsed);
}

#[test]
fn ext_list_package_round_trip_preserves_inert_internal_and_external_relationships() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let expected = with_extension_relationships(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
    assert_eq!(loaded, expected);
    let part = package
        .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
        .unwrap();
    assert_eq!(
        part.rels()
            .get("rIdExtInternal")
            .unwrap()
            .target_partname()
            .unwrap()
            .as_str(),
        "/xl/custom/ext.bin"
    );
    assert!(part.rels().get("rIdExtExternal").unwrap().is_external());
    assert_eq!(
        part.rels().get("rIdExtExternal").unwrap().target_ref(),
        "https://example.invalid/not-fetched"
    );
}

#[test]
fn rejects_ext_list_schema_order_uri_payload_and_caps() {
    let wrap = |body: &str| {
        format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
        )
    };
    for body in [
        "<extLst/>",
        "<extLst bad=\"1\"><ext uri=\"u\"><a/></ext></extLst>",
        "<extLst><ext><a/></ext></extLst>",
        "<extLst><ext uri=\"\"><a/></ext></extLst>",
        "<extLst><ext uri=\"bad uri\"><a/></ext></extLst>",
        "<extLst><ext uri=\"u\" bad=\"1\"><a/></ext></extLst>",
        "<extLst><ext uri=\"u\"/></extLst>",
        "<extLst><ext uri=\"u\"><a/><b/></ext></extLst>",
        "<extLst><bad uri=\"u\"><a/></bad></extLst>",
        "<u:extLst xmlns:u=\"urn:foreign\"><u:ext uri=\"u\"><a/></u:ext></u:extLst>",
    ] {
        let xml = wrap(body);
        assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
    }
    let out_of_order = format!(
        "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><extLst><ext uri=\"u\"><a/></ext></extLst><drawing r:id=\"rIdD\"/></chartsheet>"
    );
    assert!(parse_chartsheet(out_of_order.as_bytes()).is_err());
    let mut value = sheet();
    value.extension_list = Some(ExtensionList {
        extensions: vec![ext("u", "<?run?><a/>")],
    });
    assert!(write_chartsheet(&value, Conformance::Transitional).is_err());
    value.extension_list = Some(ExtensionList {
        extensions: vec![ext(&"u".repeat(MAX_EXTENSION_URI_BYTES + 1), "<a/>")],
    });
    assert!(write_chartsheet(&value, Conformance::Transitional).is_err());
    value.extension_list = Some(ExtensionList {
        extensions: vec![ext(
            "u",
            &format!("<a>{}</a>", "x".repeat(MAX_EXTENSION_PAYLOAD_BYTES)),
        )],
    });
    assert!(write_chartsheet(&value, Conformance::Transitional).is_err());
    value.extension_list = Some(ExtensionList {
        extensions: vec![ext("u", "<a/>"); MAX_EXTENSIONS + 1],
    });
    assert!(write_chartsheet(&value, Conformance::Transitional).is_err());
}

#[test]
fn rejects_extension_relationship_missing_orphan_mismatch_duplicate_escape_caps_and_wrong_namespace()
 {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let mut missing = with_extension_relationships(conformance);
    missing.extension_relationships.clear();
    assert!(store_chartsheet(&mut package, &workbook, &missing, conformance).is_err());
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_err()
    );
    let (mut package, workbook) = base_package(conformance);
    let mut mismatched = with_extension_relationships(conformance);
    mismatched.extension_relationships[0].relationship_id = "rIdWrong".into();
    assert!(store_chartsheet(&mut package, &workbook, &mismatched, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut duplicate = with_extension_relationships(conformance);
    duplicate.extension_relationships[1] = duplicate.extension_relationships[0].clone();
    assert!(store_chartsheet(&mut package, &workbook, &duplicate, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut escaped = with_extension_relationships(conformance);
    escaped.extension_relationships[0].target = ExtensionRelationshipTarget::Internal {
        part_name: "../../../evil.bin".into(),
    };
    assert!(store_chartsheet(&mut package, &workbook, &escaped, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut oversized = with_extension_relationships(conformance);
    oversized.extension_relationships[0].relationship_type =
        "x".repeat(MAX_EXTENSION_RELATIONSHIP_STRING_BYTES + 1);
    assert!(store_chartsheet(&mut package, &workbook, &oversized, conformance).is_err());
    let (mut package, workbook) = base_package(Conformance::Strict);
    let mut wrong_namespace = with_extension_relationships(Conformance::Strict);
    wrong_namespace.chartsheet.extension_list = Some(ExtensionList {
        extensions: vec![ext(
            "u",
            &format!("<u:a xmlns:u=\"urn:v\" xmlns:r=\"{REL}\" r:id=\"rIdExtInternal\"/>"),
        )],
    });
    wrong_namespace.extension_relationships.truncate(1);
    assert!(
        store_chartsheet(
            &mut package,
            &workbook,
            &wrong_namespace,
            Conformance::Strict
        )
        .is_err()
    );
    let (mut package, workbook) = base_package(conformance);
    let expected = with_extension_relationships(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:orphan".into(),
            "https://example.invalid/orphan".into(),
            "rIdOrphan".into(),
            true,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .remove("rIdExtInternal");
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
}

#[test]
fn strict_typed_xml_round_trip() {
    let expected = sheet();
    let xml = write_chartsheet(&expected, Conformance::Strict).unwrap();
    let (kind, parsed) = parse_chartsheet(&xml).unwrap();
    assert_eq!(kind, Conformance::Strict);
    assert_eq!(parsed, expected);
}
#[test]
fn transitional_custom_view_reference_round_trip() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:customSheetViews><x:customSheetView guid=\"{{00112233-4455-6677-8899-AABBCCDDEEFF}}\" scale=\"10\" state=\"veryHidden\" zoomToFit=\"1\"/><x:customSheetView guid=\"{{10213243-5465-7687-98A9-BACBDCEDFE0F}}\"/></x:customSheetViews><x:drawing r:id=\"rId1\"/></x:chartsheet>"
    );
    let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
    let views = parsed.custom_views.as_ref().unwrap();
    assert_eq!(views.len(), 2);
    assert_eq!(views[0].state, Some(State::VeryHidden));
    assert_eq!(views[0].scale, Some(10));
    let written = write_chartsheet(&parsed, Conformance::Transitional).unwrap();
    assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
}
#[test]
fn mce_fallback_selects_chartsheet_views() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:views/></mc:Choice><mc:Fallback><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rId1\"/></x:chartsheet>"
    );
    assert_eq!(parse_chartsheet(xml.as_bytes()).unwrap().1.views.len(), 1);
}
#[test]
fn mce_fallback_selects_custom_chartsheet_views() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><mc:AlternateContent><mc:Choice Requires=\"u\"><u:customViews/></mc:Choice><mc:Fallback><x:customSheetViews><x:customSheetView guid=\"{{00112233-4455-6677-8899-AABBCCDDEEFF}}\" scale=\"200\"/></x:customSheetViews></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rId1\"/></x:chartsheet>"
    );
    let parsed = parse_chartsheet(xml.as_bytes()).unwrap().1;
    assert_eq!(parsed.custom_views.unwrap()[0].scale, Some(200));
}
#[test]
fn loads_both_poi_chartsheet_graphs() {
    for (bytes, name, zoom) in [(POI_ONE, "Chart2", 131), (POI_TWO, "Chart1", 84)] {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let workbook_part = package.get_part(&workbook).unwrap();
        let id = workbook_part
            .rels()
            .iter()
            .find(|rel| rel.reltype() == CHARTSHEET_REL)
            .unwrap()
            .r_id()
            .to_owned();
        let loaded = load_chartsheet(&package, &workbook, &id).unwrap();
        assert_eq!(loaded.entry.name, name);
        assert_eq!(loaded.chartsheet.views[0].zoom_scale, Some(zoom));
        assert_eq!(loaded.drawing.charts.len(), 1);
        assert!(loaded.drawing.charts[0].data.starts_with(b"<?xml"));
    }
}
#[test]
fn strict_package_writer_round_trips_complete_leaf_graph() {
    let conformance = Conformance::Strict;
    let (mut package, workbook) = base_package(conformance);
    let expected = value(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    assert_eq!(
        load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap(),
        expected
    );
}

#[test]
fn classic_chart_embedded_workbook_relationship_round_trips() {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut expected = value(conformance);
        let chart = &mut expected.drawing.charts[0];
        chart.data = format!(
            "<c:chartSpace xmlns:c=\"{}\" xmlns:r=\"{}\"><c:chart><c:plotArea/></c:chart><c:externalData r:id=\"rIdWorkbook\"><c:autoUpdate val=\"0\"/></c:externalData></c:chartSpace>",
            conformance.chart(),
            conformance.rel(),
        )
        .into_bytes();
        chart.kind = ChartResourceKind::ClassicWithRelationships {
            user_shapes: None,
            outbound_resources: vec![ChartOutboundResource::EmbeddedPackage(
                ChartEmbeddedPackageResource {
                    relationship_id: "rIdWorkbook".into(),
                    part_name: "/xl/embeddings/embeddedWorkbook1.xlsx".into(),
                    content_type: ChartEmbeddedPackageContentType::Xlsx,
                    data: vec![0x50, 0x4b, 0x03, 0x04, 0, 255],
                },
            )],
        };

        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
        assert_eq!(loaded, expected);

        let (mut second, second_workbook) = base_package(conformance);
        store_chartsheet(&mut second, &second_workbook, &loaded, conformance).unwrap();
        assert_eq!(
            load_chartsheet(&second, &second_workbook, "rIdChartSheet").unwrap(),
            expected
        );
    }
}

#[test]
fn classic_chart_relationship_metadata_must_match_xml() {
    let conformance = Conformance::Transitional;
    let mut missing = value(conformance);
    missing.drawing.charts[0].data = format!(
        "<c:chartSpace xmlns:c=\"{}\" xmlns:r=\"{}\"><c:chart><c:plotArea/></c:chart><c:externalData r:id=\"rIdWorkbook\"/></c:chartSpace>",
        conformance.chart(),
        conformance.rel(),
    )
    .into_bytes();
    let (mut package, workbook) = base_package(conformance);
    assert!(store_chartsheet(&mut package, &workbook, &missing, conformance).is_err());

    let mut stale = value(conformance);
    stale.drawing.charts[0].kind = ChartResourceKind::ClassicWithRelationships {
        user_shapes: None,
        outbound_resources: vec![ChartOutboundResource::EmbeddedPackage(
            ChartEmbeddedPackageResource {
                relationship_id: "rIdWorkbook".into(),
                part_name: "/xl/embeddings/embeddedWorkbook1.xlsx".into(),
                content_type: ChartEmbeddedPackageContentType::Xlsx,
                data: vec![0x50, 0x4b, 0x03, 0x04],
            },
        )],
    };
    let (mut package, workbook) = base_package(conformance);
    assert!(store_chartsheet(&mut package, &workbook, &stale, conformance).is_err());
}
#[test]
fn printer_settings_page_setup_strict_mce_and_schema_order() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{STRICT_SML}\" xmlns:r=\"{STRICT_REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><mc:AlternateContent><mc:Choice Requires=\"u\"><u:pageSetup/></mc:Choice><mc:Fallback><x:pageSetup orientation=\"landscape\" r:id=\"rIdPrinter\"/></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rIdDrawing\"/></x:chartsheet>"
    );
    let (kind, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
    assert_eq!(kind, Conformance::Strict);
    assert_eq!(
        parsed
            .page_setup
            .as_ref()
            .unwrap()
            .printer_settings_relationship_id
            .as_deref(),
        Some("rIdPrinter")
    );
    let written = write_chartsheet(&parsed, kind).unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.find("pageSetup").unwrap() < text.find("drawing").unwrap());
    assert!(text.contains("r:id=\"rIdPrinter\""));
    assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
    for body in [
        format!("<pageSetup xmlns:r=\"{REL}\" r:id=\"rIdP\"/><pageSetup/>"),
        format!("<drawing xmlns:r=\"{REL}\" r:id=\"rIdD\"/><pageSetup r:id=\"rIdP\"/>"),
        format!(
            "<pageSetup xmlns:r=\"{STRICT_REL}\" xmlns:t=\"{REL}\" t:id=\"rIdP\"/><drawing xmlns:r=\"{STRICT_REL}\" r:id=\"rIdD\"/>"
        ),
    ] {
        let xml = format!(
            "<chartsheet xmlns=\"{STRICT_SML}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>{body}</chartsheet>"
        );
        assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
    }
}
#[test]
fn printer_settings_package_round_trip_preserves_opaque_bytes() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let expected = value(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
    assert_eq!(loaded.printer_settings, expected.printer_settings);
    let part = package
        .get_part(&PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap())
        .unwrap();
    assert_eq!(part.content_type(), PRINTER_CT);
    assert_eq!(
        part.blob(),
        [0x44, 0x45, 0x56, 0x4d, 0x4f, 0x44, 0x45, 0, 255]
    );
}
#[test]
fn rejects_printer_settings_pairing_paths_collisions_and_caps_before_mutation() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.printer_settings.as_mut().unwrap().relationship_id = "rIdOther".into();
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    for path in [
        "/xl/printerSettings/sub/settings.bin",
        "/xl/printerSettings/settings.dat",
        "/xl/media/settings.bin",
    ] {
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.printer_settings.as_mut().unwrap().resource.part_name = path.into();
        assert!(
            store_chartsheet(&mut package, &workbook, &bad, conformance).is_err(),
            "accepted {path}"
        );
    }
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.printer_settings.as_mut().unwrap().resource.data = vec![0; MAX_SETTINGS_BYTES + 1];
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_err()
    );
    let (mut package, workbook) = base_package(conformance);
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap(),
        PRINTER_CT.into(),
        vec![1],
    )));
    assert!(store_chartsheet(&mut package, &workbook, &value(conformance), conformance).is_err());
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_err()
    );
}
#[test]
fn rejects_printer_settings_external_wrong_type_escape_orphan_content_type_and_outbound_graphs() {
    for (kind, target, external) in [
        (PRINTER_REL, "https://example.invalid/settings.bin", true),
        (IMAGE_REL, "../printerSettings/printerSettings1.bin", false),
        (PRINTER_REL, "../../../evil.bin", false),
    ] {
        let conformance = Conformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        let chartsheet = package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap();
        chartsheet.rels_mut().remove("rIdPrinter");
        chartsheet.rels_mut().add_relationship(
            kind.into(),
            target.into(),
            "rIdPrinter".into(),
            external,
        );
        assert!(
            load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
            "accepted {kind} {target}"
        );
    }
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            PRINTER_REL.into(),
            "../printerSettings/printerSettings1.bin".into(),
            "rIdOrphan".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            IMAGE_REL.into(),
            "../media/evil.png".into(),
            "rIdOutbound".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap(),
        "application/octet-stream".into(),
        vec![1],
    )));
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
}
#[test]
fn web_publish_schema_enum_and_deterministic_round_trip() {
    let mut body = String::from("<webPublishItems count=\"8\">");
    for (index, kind) in [
        "sheet",
        "printArea",
        "autoFilter",
        "range",
        "chart",
        "pivotTable",
        "query",
        "label",
    ]
    .iter()
    .enumerate()
    {
        let source_ref = if *kind == "range" {
            " sourceRef=\"A1:B2\""
        } else {
            ""
        };
        let source_object = if matches!(*kind, "pivotTable" | "query" | "label") {
            " sourceObject=\"OpaqueName\""
        } else {
            ""
        };
        body.push_str(&format!("<webPublishItem id=\"{index}\" divId=\"Div{index}\" sourceType=\"{kind}\"{source_ref}{source_object} destinationFile=\"opaque:{index}\" autoRepublish=\"{}\"/>",if index%2==0{"true"}else{"0"}));
    }
    body.push_str("</webPublishItems>");
    let xml = format!(
        "<chartsheet xmlns=\"{STRICT_SML}\" xmlns:r=\"{STRICT_REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
    );
    let (kind, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
    assert_eq!(kind, Conformance::Strict);
    let items = &parsed.web_publish_items.as_ref().unwrap().items;
    assert_eq!(
        items
            .iter()
            .map(|item| item.source_type)
            .collect::<Vec<_>>(),
        vec![
            WebSourceType::Sheet,
            WebSourceType::PrintArea,
            WebSourceType::AutoFilter,
            WebSourceType::Range,
            WebSourceType::Chart,
            WebSourceType::PivotTable,
            WebSourceType::Query,
            WebSourceType::Label
        ]
    );
    let first = write_chartsheet(&parsed, kind).unwrap();
    let reparsed = parse_chartsheet(&first).unwrap().1;
    let second = write_chartsheet(&reparsed, kind).unwrap();
    assert_eq!(first, second);
    assert_eq!(parsed, reparsed);
}
#[test]
fn web_publish_mce_fallback_and_inert_references() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdD\"/><mc:AlternateContent><mc:Choice Requires=\"u\"><u:run href=\"https://example.invalid/execute\"/></mc:Choice><mc:Fallback><x:webPublishItems><x:webPublishItem id=\"0\" divId=\"D\" sourceType=\"chart\" sourceObject=\"file:///not/read\" destinationFile=\"/tmp/not-written\" title=\"$(never-execute)\"/></x:webPublishItems></mc:Fallback></mc:AlternateContent></x:chartsheet>"
    );
    let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
    let item = &parsed.web_publish_items.as_ref().unwrap().items[0];
    assert_eq!(item.destination_file, "/tmp/not-written");
    assert_eq!(item.auto_republish, None);
    assert_eq!(
        parse_chartsheet(&write_chartsheet(&parsed, Conformance::Transitional).unwrap())
            .unwrap()
            .1,
        parsed
    );
}
#[test]
fn web_publish_package_load_store_preserves_metadata_only() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let expected = value(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
    assert_eq!(
        loaded.chartsheet.web_publish_items,
        expected.chartsheet.web_publish_items
    );
}
#[test]
fn rejects_web_publish_malformed_duplicates_cardinality_and_order() {
    let wrap = |body: &str| {
        format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
        )
    };
    for body in [
        "<webPublishItems/>",
        "<webPublishItems count=\"2\"><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/><webPublishItem id=\"1\" divId=\"B\" sourceType=\"sheet\" destinationFile=\"y\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/><webPublishItem id=\"2\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"y\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"4294967296\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"bad\" destinationFile=\"x\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"range\" destinationFile=\"x\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"query\" destinationFile=\"x\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\" autoRepublish=\"on\"/></webPublishItems>",
        "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\" extra=\"1\"/></webPublishItems>",
    ] {
        let xml = wrap(body);
        assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
    }
    let out_of_order = format!(
        "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/></webPublishItems><drawing r:id=\"rIdD\"/></chartsheet>"
    );
    assert!(parse_chartsheet(out_of_order.as_bytes()).is_err());
}
#[test]
fn rejects_web_publish_count_and_string_caps() {
    let mut body = String::from("<webPublishItems>");
    for index in 0..=MAX_WEB_PUBLISH_ITEMS {
        body.push_str(&format!("<webPublishItem id=\"{index}\" divId=\"D{index}\" sourceType=\"sheet\" destinationFile=\"x\"/>"));
    }
    body.push_str("</webPublishItems>");
    let xml = format!(
        "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
    );
    assert!(parse_chartsheet(xml.as_bytes()).is_err());
    let mut value = sheet();
    value.web_publish_items.as_mut().unwrap().items[0].title =
        Some("x".repeat(MAX_WEB_PUBLISH_STRING_BYTES + 1));
    assert!(write_chartsheet(&value, Conformance::Transitional).is_err());
}
#[test]
fn picture_mce_schema_and_inert_round_trip() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdDrawing\"/><mc:AlternateContent><mc:Choice Requires=\"u\"><u:picture/></mc:Choice><mc:Fallback><x:picture r:id=\"rIdBackground\"/></mc:Fallback></mc:AlternateContent></x:chartsheet>"
    );
    let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
    assert_eq!(
        parsed.background_picture_relationship_id.as_deref(),
        Some("rIdBackground")
    );
    let written = write_chartsheet(&parsed, Conformance::Transitional).unwrap();
    assert!(
        String::from_utf8(written.clone())
            .unwrap()
            .contains("<x:drawing r:id=\"rIdDrawing\"/><x:picture r:id=\"rIdBackground\"/>")
    );
    assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
}
#[test]
fn transitional_picture_package_round_trip_preserves_opaque_bytes() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let expected = value(conformance);
    store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
    let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
    assert_eq!(
        loaded.background_picture.as_ref().unwrap().data,
        vec![0, 255, 1, 254]
    );
    assert_eq!(loaded, expected);
}
#[test]
fn vml_mce_schema_order_and_inert_round_trip() {
    let xml = format!(
        "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdDrawing\"/><mc:AlternateContent><mc:Choice Requires=\"u\"><u:vml/></mc:Choice><mc:Fallback><x:legacyDrawing r:id=\"rIdLegacy\"/><x:legacyDrawingHF r:id=\"rIdLegacyHF\"/></mc:Fallback></mc:AlternateContent><x:picture r:id=\"rIdBackground\"/></x:chartsheet>"
    );
    let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
    assert_eq!(
        parsed.legacy_drawing_relationship_id.as_deref(),
        Some("rIdLegacy")
    );
    assert_eq!(
        parsed
            .legacy_header_footer_drawing_relationship_id
            .as_deref(),
        Some("rIdLegacyHF")
    );
    let written = write_chartsheet(&parsed, Conformance::Transitional).unwrap();
    let text = String::from_utf8(written.clone()).unwrap();
    assert!(text.contains("<x:drawing r:id=\"rIdDrawing\"/><x:legacyDrawing r:id=\"rIdLegacy\"/><x:legacyDrawingHF r:id=\"rIdLegacyHF\"/><x:picture"));
    assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
}
#[test]
fn rejects_vml_schema_duplicates_order_and_missing_ids() {
    for body in [
        "<legacyDrawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdL\"/><drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/>",
        "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/><legacyDrawing r:id=\"rIdL\"/><legacyDrawing r:id=\"rIdL2\"/>",
        "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/><legacyDrawingHF/>",
        "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/><legacyDrawingHF r:id=\"rIdHF\"/><legacyDrawing r:id=\"rIdL\"/>",
    ] {
        let xml = format!(
            "<chartsheet xmlns=\"{SML}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>{body}</chartsheet>"
        );
        assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
    }
}
#[test]
fn rejects_vml_pairing_content_type_collision_and_caps() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.legacy_drawing.as_mut().unwrap().relationship_id = "rIdOther".into();
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.legacy_drawing.as_mut().unwrap().content_type = "application/xml".into();
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.legacy_header_footer_drawing.as_mut().unwrap().part_name =
        bad.legacy_drawing.as_ref().unwrap().part_name.clone();
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.legacy_drawing.as_mut().unwrap().data = vec![0; MAX_VML_DRAWING_BYTES + 1];
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/drawings/vmlDrawing1.vml").unwrap())
        .unwrap()
        .set_blob(vec![0; MAX_VML_DRAWING_BYTES + 1]);
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
}
#[test]
fn rejects_external_wrong_type_escaped_orphan_and_outbound_vml_graphs() {
    for (kind, target, external) in [
        (
            VML_DRAWING_REL,
            "https://example.invalid/vmlDrawing1.vml",
            true,
        ),
        (IMAGE_REL, "../drawings/vmlDrawing1.vml", false),
        (VML_DRAWING_REL, "../../../evil.vml", false),
    ] {
        let conformance = Conformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        let chartsheet = package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap();
        chartsheet.rels_mut().remove("rIdLegacy");
        chartsheet.rels_mut().add_relationship(
            kind.into(),
            target.into(),
            "rIdLegacy".into(),
            external,
        );
        assert!(
            load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
            "accepted {kind} {target}"
        );
    }
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            VML_DRAWING_REL.into(),
            "../drawings/vmlDrawing1.vml".into(),
            "rIdOrphan".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/drawings/vmlDrawing1.vml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            IMAGE_REL.into(),
            "../media/evil.png".into(),
            "rIdOutbound".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
}
#[test]
fn rejects_picture_cardinality_order_metadata_and_caps() {
    for xml in [
        format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><picture r:id=\"rIdP\"/><drawing r:id=\"rIdD\"/></chartsheet>"
        ),
        format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/><picture r:id=\"rIdP\"/><picture r:id=\"rIdQ\"/></chartsheet>"
        ),
        format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/><picture/></chartsheet>"
        ),
    ] {
        assert!(parse_chartsheet(xml.as_bytes()).is_err(), "{xml}");
    }
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.background_picture.as_mut().unwrap().relationship_id = "different".into();
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.background_picture.as_mut().unwrap().part_name = "/xl/charts/chart1.xml".into();
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.background_picture.as_mut().unwrap().data = vec![0; MAX_BACKGROUND_IMAGE_BYTES + 1];
    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
}
#[test]
fn rejects_external_wrong_type_escaped_and_unreferenced_picture_relationships() {
    for (kind, target, external) in [
        (IMAGE_REL, "https://example.invalid/background.png", true),
        (rt::CHART, "../media/background1.png", false),
        (IMAGE_REL, "../../../evil.png", false),
    ] {
        let conformance = Conformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        let chartsheet = package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap();
        chartsheet.rels_mut().remove("rIdBackground");
        chartsheet.rels_mut().add_relationship(
            kind.into(),
            target.into(),
            "rIdBackground".into(),
            external,
        );
        assert!(
            load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
            "accepted {kind} {target}"
        );
    }
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            IMAGE_REL.into(),
            "../media/background1.png".into(),
            "rIdExtra".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
}
#[test]
fn rejects_existing_background_part_collision_before_mutation() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/media/background1.png").unwrap(),
        "image/png".into(),
        vec![9],
    )));
    assert!(store_chartsheet(&mut package, &workbook, &value(conformance), conformance).is_err());
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_err()
    );
}
#[test]
fn store_is_atomic_when_new_candidate_parts_conflict_case_insensitively() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let original_workbook = package.get_part(&workbook).unwrap().blob().to_vec();
    let mut bad = value(conformance);
    bad.drawing.data = format!(
            "<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:c=\"{CHART}\" xmlns:r=\"{REL}\"><c:chart r:id=\"rIdChart\"/><c:chart r:id=\"rIdChart2\"/></xdr:wsDr>"
        )
        .into_bytes();
    let mut second = bad.drawing.charts[0].clone();
    second.relationship_id = "rIdChart2".into();
    second.part_name = "/xl/charts/CHART1.xml".into();
    bad.drawing.charts.push(second);

    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        store_chartsheet(&mut package, &workbook, &bad, conformance)
    }));
    assert!(result.is_ok(), "store_chartsheet panicked");
    assert!(result.unwrap().is_err());
    assert_eq!(package.part_count(), 1);
    assert_eq!(
        package.get_part(&workbook).unwrap().blob(),
        original_workbook
    );
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_err()
    );
}
#[test]
fn store_rejects_non_bijective_drawing_chart_resources() {
    let conformance = Conformance::Transitional;
    let (mut package, workbook) = base_package(conformance);
    let mut bad = value(conformance);
    bad.drawing.data = format!(
            "<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:c=\"{CHART}\" xmlns:r=\"{REL}\"><c:chart r:id=\"rIdChart\"/><c:chart r:id=\"rIdChart2\"/></xdr:wsDr>"
        )
        .into_bytes();
    let mut second = bad.drawing.charts[0].clone();
    second.part_name = "/xl/charts/chart2.xml".into();
    bad.drawing.charts.push(second);

    assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    assert_eq!(package.part_count(), 1);
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_err()
    );
}
#[test]
fn drawing_chart_reference_cap_is_checked_before_retention_grows() {
    let conformance = Conformance::Transitional;
    let mut xml = format!("<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:c=\"{CHART}\" xmlns:r=\"{REL}\">");
    for index in 0..=MAX_CHARTS {
        xml.push_str(&format!("<c:chart r:id=\"rIdChart{index}\"/>"));
    }
    xml.push_str("</xdr:wsDr>");
    assert!(drawing_chart_references(xml.as_bytes(), conformance).is_err());
}
#[test]
fn rejects_malformed_caps_and_graphs() {
    assert!(parse_chartsheet(b"<!DOCTYPE x><chartsheet/>").is_err());
    assert!(parse_chartsheet(format!("<chartsheet xmlns=\"{SML}\"><sheetViews><sheetView workbookViewId=\"0\" zoomScale=\"401\"/></sheetViews><drawing xmlns:r=\"{REL}\" r:id=\"rId1\"/></chartsheet>").as_bytes()).is_err());
    for custom in [
        "<customSheetViews/>",
        "<customSheetViews><customSheetView guid=\"bad\"/></customSheetViews>",
        "<customSheetViews><customSheetView guid=\"{00112233-4455-6677-8899-AABBCCDDEEFF}\" scale=\"401\"/></customSheetViews>",
        "<customSheetViews><customSheetView guid=\"{00112233-4455-6677-8899-AABBCCDDEEFF}\"/><customSheetView guid=\"{00112233-4455-6677-8899-aabbccddeeff}\"/></customSheetViews>",
    ] {
        let xml = format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>{custom}<drawing r:id=\"rId1\"/></chartsheet>"
        );
        assert!(parse_chartsheet(xml.as_bytes()).is_err(), "{custom}");
    }
    assert!(parse_chartsheet(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
    let (mut package, workbook) = base_package(Conformance::Transitional);
    let expected = value(Conformance::Transitional);
    store_chartsheet(
        &mut package,
        &workbook,
        &expected,
        Conformance::Transitional,
    )
    .unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::IMAGE.into(),
            "../media/x.png".into(),
            "rIdBad".into(),
            false,
        );
    assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
}

mod chart_outbound_tests {
    use super::{
        CHART_EX, ChartEmbeddedPackageContentType, ChartEmbeddedPackageResource,
        ChartOutboundResource, ChartResourceKind, ChartThemeOverrideResource, Conformance,
        DRAWING_MAIN, ImageContentType, ImageResource, LO_USER_SHAPES_IMAGES,
        MAX_CHART_EMBEDDED_PACKAGE_BYTES, OpcPackage, PackURI, Package, STRICT_DRAWING_MAIN,
        THEME_OVERRIDE_CT, base_package, chart_ex_user_shapes, load_chartsheet, store_chartsheet,
    };

    fn outbound_value(conformance: Conformance) -> Package {
        let source = OpcPackage::from_bytes(LO_USER_SHAPES_IMAGES).unwrap();
        let blob = |path: &str| {
            source
                .get_part(&PackURI::new(path).unwrap())
                .unwrap()
                .blob()
                .to_vec()
        };
        let mut value = chart_ex_user_shapes(conformance);
        let drawing_main = if conformance == Conformance::Strict {
            STRICT_DRAWING_MAIN
        } else {
            DRAWING_MAIN
        };
        let chart = &mut value.drawing.charts[0];
        chart.data = format!(
            "<cx:chartSpace xmlns:cx=\"{CHART_EX}\" xmlns:a=\"{drawing_main}\" xmlns:r=\"{}\"><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series><cx:spPr><a:blipFill><a:blip r:embed=\"rIdDirectImage\"/></a:blipFill></cx:spPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart><cx:externalData r:id=\"rIdPackage\" cx:autoUpdate=\"0\"/></cx:chartSpace>",
            conformance.rel()
        )
        .into_bytes();
        let theme_data = format!(
            "<a:themeOverride xmlns:a=\"{drawing_main}\" xmlns:r=\"{}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:active/></mc:Choice><mc:Fallback><a:fmtScheme name=\"Inert\"><a:fillStyleLst><a:blipFill><a:blip r:embed=\"rIdThemeImage\"/></a:blipFill></a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></mc:Fallback></mc:AlternateContent></a:themeOverride>",
            conformance.rel()
        )
        .into_bytes();
        let ChartResourceKind::Extended {
            outbound_resources, ..
        } = &mut chart.kind
        else {
            unreachable!()
        };
        *outbound_resources = vec![
            ChartOutboundResource::Image(ImageResource {
                relationship_id: "rIdDirectImage".into(),
                part_name: "/xl/media/chartDirect1.png".into(),
                content_type: ImageContentType::Png,
                data: blob("/xl/media/image1.png"),
            }),
            ChartOutboundResource::EmbeddedPackage(ChartEmbeddedPackageResource {
                relationship_id: "rIdPackage".into(),
                part_name: "/xl/embeddings/Microsoft_Excel_Worksheet1.xlsx".into(),
                content_type: ChartEmbeddedPackageContentType::Xlsx,
                data: LO_USER_SHAPES_IMAGES.to_vec(),
            }),
            ChartOutboundResource::ThemeOverride(ChartThemeOverrideResource {
                relationship_id: "rIdTheme".into(),
                part_name: "/xl/theme/themeOverride1.xml".into(),
                content_type: THEME_OVERRIDE_CT.into(),
                data: theme_data,
                images: vec![ImageResource {
                    relationship_id: "rIdThemeImage".into(),
                    part_name: "/xl/media/themeImage1.svg".into(),
                    content_type: ImageContentType::Svg,
                    data: blob("/xl/media/image2.svg"),
                }],
            }),
        ];
        value
    }

    #[test]
    fn chart_ex_complete_outbound_family_round_trips_strict_and_transitional() {
        for conformance in [Conformance::Transitional, Conformance::Strict] {
            let expected = outbound_value(conformance);
            let (mut package, workbook) = base_package(conformance);
            store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
            let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
            assert_eq!(loaded, expected);
            let chart_uri = PackURI::new("/xl/charts/chartEx1.xml").unwrap();
            let chart = package.get_part(&chart_uri).unwrap();
            assert_eq!(
                chart.rels().get("rIdDirectImage").unwrap().reltype(),
                conformance.image_rel()
            );
            assert_eq!(
                chart.rels().get("rIdTheme").unwrap().reltype(),
                conformance.theme_override_rel()
            );
            assert_eq!(
                chart.rels().get("rIdPackage").unwrap().reltype(),
                conformance.package_rel()
            );
            let theme = package
                .get_part(&PackURI::new("/xl/theme/themeOverride1.xml").unwrap())
                .unwrap();
            assert_eq!(
                theme.rels().get("rIdThemeImage").unwrap().reltype(),
                conformance.image_rel()
            );
            let (mut second, second_workbook) = base_package(conformance);
            store_chartsheet(&mut second, &second_workbook, &loaded, conformance).unwrap();
            assert_eq!(
                load_chartsheet(&second, &second_workbook, "rIdChartSheet").unwrap(),
                loaded
            );
        }
    }

    #[test]
    fn chart_ex_outbound_rejects_active_external_mismatch_collision_roots_and_caps() {
        let conformance = Conformance::Transitional;
        let mut bad = outbound_value(conformance);
        bad.drawing.charts[0].data = String::from_utf8(bad.drawing.charts[0].data.clone())
            .unwrap()
            .replace("autoUpdate=\"0\"", "autoUpdate=\"true\"")
            .into_bytes();
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );

        let mut bad = outbound_value(conformance);
        bad.drawing.charts[0].data = String::from_utf8(bad.drawing.charts[0].data.clone())
            .unwrap()
            .replace(" r:embed=\"rIdDirectImage\"", "")
            .into_bytes();
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());

        let mut bad = outbound_value(conformance);
        if let ChartResourceKind::Extended {
            outbound_resources, ..
        } = &mut bad.drawing.charts[0].kind
        {
            let theme = outbound_resources
                .iter_mut()
                .find_map(|resource| match resource {
                    ChartOutboundResource::ThemeOverride(theme) => Some(theme),
                    _ => None,
                })
                .unwrap();
            theme.data = String::from_utf8(theme.data.clone())
                .unwrap()
                .replace("themeOverride", "theme")
                .into_bytes();
        }
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());

        let mut bad = outbound_value(conformance);
        if let ChartResourceKind::Extended {
            user_shapes: Some(shapes),
            outbound_resources,
            ..
        } = &mut bad.drawing.charts[0].kind
        {
            outbound_resources[0] = match outbound_resources[0].clone() {
                ChartOutboundResource::Image(mut image) => {
                    image.relationship_id = shapes.relationship_id.clone();
                    ChartOutboundResource::Image(image)
                },
                _ => unreachable!(),
            };
        }
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());

        let expected = outbound_value(conformance);
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let chart = package
            .get_part_mut(&PackURI::new("/xl/charts/chartEx1.xml").unwrap())
            .unwrap();
        chart.rels_mut().remove("rIdPackage");
        chart.rels_mut().add_relationship(
            conformance.package_rel().into(),
            "https://example.invalid/active.xlsx".into(),
            "rIdPackage".into(),
            true,
        );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());

        let mut bad = outbound_value(conformance);
        if let ChartResourceKind::Extended {
            outbound_resources, ..
        } = &mut bad.drawing.charts[0].kind
        {
            let embedded = outbound_resources
                .iter_mut()
                .find_map(|resource| match resource {
                    ChartOutboundResource::EmbeddedPackage(embedded) => Some(embedded),
                    _ => None,
                })
                .unwrap();
            embedded.data = vec![0; MAX_CHART_EMBEDDED_PACKAGE_BYTES + 1];
        }
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );
    }
}
