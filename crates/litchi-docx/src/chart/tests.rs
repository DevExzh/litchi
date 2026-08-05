use super::model::{
    A, C, CHART_CT, COLOR_STYLE_CT, COLOR_STYLE_REL, Conformance, DOCUMENT_CT, MAX_WORKBOOK_BYTES,
    STYLE_CT, STYLE_REL, WORKBOOK_CT,
};
use super::{load, store};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};

const STYLE_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const POI: &[u8] = include_bytes!("../../../../test-data/poi/test-data/document/61745.docx");
const LO_INTERNAL: &[u8] = include_bytes!(
    "../../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx"
);
const LO_EXTERNAL: &[u8] = include_bytes!(
    "../../../../test-data/libreoffice-core/oox/qa/unit/data/chart-data-label-char-color.docx"
);

fn document() -> PackURI {
    PackURI::new("/word/document.xml").unwrap()
}

fn style_entry(name: &str) -> String {
    format!(
        r#"<cs:{name}><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"/></cs:{name}>"#
    )
}

fn style_xml() -> String {
    let names = [
        "axisTitle",
        "categoryAxis",
        "chartArea",
        "dataLabel",
        "dataLabelCallout",
        "dataPoint",
        "dataPoint3D",
        "dataPointLine",
        "dataPointMarker",
        "dataPointWireframe",
        "dataTable",
        "downBar",
        "dropLine",
        "errorBar",
        "floor",
        "gridlineMajor",
        "gridlineMinor",
        "hiLoLine",
        "leaderLine",
        "legend",
        "plotArea",
        "plotArea3D",
        "seriesAxis",
        "seriesLine",
        "title",
        "trendline",
        "trendlineLabel",
        "upBar",
        "valueAxis",
        "wall",
    ];
    let body = names
        .iter()
        .map(|name| style_entry(name))
        .collect::<String>();
    format!(r#"<cs:chartStyle xmlns:cs="{STYLE_NS}">{body}</cs:chartStyle>"#)
}

fn color_style_xml() -> String {
    format!(
        r#"<cs:colorStyle xmlns:cs="{STYLE_NS}" xmlns:a="{A}" meth="cycle"><a:srgbClr val="FF0000"/></cs:colorStyle>"#
    )
}

#[test]
fn poi_and_libreoffice_internal_charts_round_trip_deterministically() {
    for (bytes, count) in [(POI, 2usize), (LO_INTERNAL, 1usize)] {
        let mut package = OpcPackage::from_bytes(bytes).unwrap();
        let name = document();
        let graph = load(&package, &name).unwrap();
        assert_eq!(graph.charts.len(), count);
        assert!(graph.charts.iter().all(|chart| chart.workbook.is_some()));
        store(&mut package, &name, &graph).unwrap();
        assert_eq!(load(&package, &name).unwrap(), graph);
        store(&mut package, &name, &graph).unwrap();
        assert_eq!(load(&package, &name).unwrap(), graph);
    }
}

fn synthetic(conformance: Conformance) -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let name = document();
    let mut document_part = BlobPart::new(
        name.clone(),
        DOCUMENT_CT.into(),
        format!(
            r#"<w:document xmlns:w="{}" xmlns:a="{}" xmlns:c="{}" xmlns:r="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="u"><u:active/></mc:Choice><mc:Fallback><w:p><w:r><w:drawing><a:graphic><a:graphicData uri="{}"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></w:drawing></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#,
            conformance.w(),
            conformance.a(),
            conformance.c(),
            conformance.r(),
            conformance.c()
        )
        .into_bytes(),
    );
    document_part.rels_mut().add_relationship(
        conformance.chart_rel().into(),
        "charts/chart1.xml".into(),
        "rIdChart".into(),
        false,
    );
    package.add_part(Box::new(document_part));
    let chart_uri = PackURI::new("/word/charts/chart1.xml").unwrap();
    let mut chart = BlobPart::new(
        chart_uri,
        CHART_CT.into(),
        format!(
            r#"<c:chartSpace xmlns:c="{}" xmlns:r="{}"><c:chart></c:chart><c:externalData r:id="rIdWorkbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>"#,
            conformance.c(),
            conformance.r()
        )
        .into_bytes(),
    );
    chart.rels_mut().add_relationship(
        STYLE_REL.into(),
        "style1.xml".into(),
        "rIdStyle".into(),
        false,
    );
    chart.rels_mut().add_relationship(
        COLOR_STYLE_REL.into(),
        "colors1.xml".into(),
        "rIdColors".into(),
        false,
    );
    chart.rels_mut().add_relationship(
        conformance.package_rel().into(),
        "../embeddings/data1.xlsx".into(),
        "rIdWorkbook".into(),
        false,
    );
    package.add_part(Box::new(chart));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/charts/style1.xml").unwrap(),
        STYLE_CT.into(),
        style_xml().into_bytes(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/charts/colors1.xml").unwrap(),
        COLOR_STYLE_CT.into(),
        color_style_xml().into_bytes(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/embeddings/data1.xlsx").unwrap(),
        WORKBOOK_CT.into(),
        b"PK opaque workbook".to_vec(),
    )));
    (package, name)
}

#[test]
fn strict_mce_graph_round_trips_without_opening_workbook() {
    let (mut package, name) = synthetic(Conformance::Strict);
    let graph = load(&package, &name).unwrap();
    assert_eq!(
        graph.charts[0].workbook.as_ref().unwrap().data,
        b"PK opaque workbook"
    );
    store(&mut package, &name, &graph).unwrap();
    assert_eq!(load(&package, &name).unwrap(), graph);
}

#[test]
fn rejects_external_ole_malformed_orphan_unsupported_and_caps_before_mutation() {
    let package = OpcPackage::from_bytes(LO_EXTERNAL).unwrap();
    assert!(load(&package, &document()).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    package
        .get_part_mut(&PackURI::new("/word/charts/chart1.xml").unwrap())
        .unwrap()
        .set_blob(format!(r#"<c:wrong xmlns:c="{C}"/>"#).into_bytes());
    assert!(load(&package, &name).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/charts/orphan.xml").unwrap(),
        CHART_CT.into(),
        format!(r#"<c:chartSpace xmlns:c="{C}"><c:chart></c:chart></c:chartSpace>"#).into_bytes(),
    )));
    assert!(load(&package, &name).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    package
        .get_part_mut(&PackURI::new("/word/charts/chart1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::IMAGE.into(),
            "../media/image1.png".into(),
            "rIdImage".into(),
            false,
        );
    assert!(load(&package, &name).is_err());
    let (mut package, name) = synthetic(Conformance::Transitional);
    let mut graph = load(&package, &name).unwrap();
    graph.charts[0].workbook.as_mut().unwrap().data = vec![0; MAX_WORKBOOK_BYTES + 1];
    let before = package.get_part(&name).unwrap().blob().to_vec();
    assert!(store(&mut package, &name, &graph).is_err());
    assert_eq!(package.get_part(&name).unwrap().blob(), before);
}
