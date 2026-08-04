use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI};

use super::animations::Sequence;

#[allow(dead_code)]
pub(super) fn parse_package_slide(
    package: &OpcPackage,
    slide_part_name: &PackURI,
) -> Result<Sequence> {
    litchi_pptx::animations::parse_package_slide(package, slide_part_name).map_err(OoxmlError::from)
}

#[cfg(test)]
const MAX_SLIDE_XML: usize = 64 * 1024 * 1024;

#[cfg(test)]
const P_NS: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
#[cfg(test)]
const P_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
#[cfg(test)]
const A_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
#[cfg(test)]
const A_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
#[cfg(test)]
const C_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
#[cfg(test)]
const C_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
#[cfg(test)]
const DGM_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
#[cfg(test)]
const R_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
#[cfg(test)]
const R_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
#[cfg(test)]
const MC_NS: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
#[cfg(test)]
const CHARTEX_URI: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";

#[cfg(test)]
const SLIDE_CT: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
#[cfg(test)]
const CHART_CT: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
#[cfg(test)]
const CHARTEX_CT: &str = "application/vnd.ms-office.chartex+xml";
#[cfg(test)]
const DIAGRAM_DATA_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
#[cfg(test)]
const DIAGRAM_LAYOUT_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
#[cfg(test)]
const DIAGRAM_STYLE_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";
#[cfg(test)]
const DIAGRAM_COLORS_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
#[cfg(test)]
const OLE_CT: &str = "application/vnd.openxmlformats-officedocument.oleObject";

#[cfg(test)]
const CHART_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
];
#[cfg(test)]
const CHARTEX_REL: [&str; 1] = ["http://schemas.microsoft.com/office/2014/relationships/chartEx"];
#[cfg(test)]
const DIAGRAM_DATA_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData",
];
#[cfg(test)]
const DIAGRAM_LAYOUT_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramLayout",
];
#[cfg(test)]
const DIAGRAM_STYLE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramQuickStyle",
];
#[cfg(test)]
const DIAGRAM_COLORS_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramColors",
];
#[cfg(test)]
const OLE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject",
];

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::animations::{
        DiagramBuild, Effect, EffectInstance, GraphicBuild, GroupId, OleChartBuild,
    };
    use litchi_opc::{Part, part::BlobPart};

    const SLIDE: &str = "/ppt/slides/slide1.xml";

    fn timing(kind: &str, shape_id: u32, group_id: u32) -> String {
        let mut sequence = Sequence::new();
        sequence.add(EffectInstance::new(shape_id, Effect::Fade).with_group_id(group_id));
        match kind {
            "chart" => {
                sequence.add_graphic_build(GraphicBuild::chart(shape_id, GroupId::new(group_id)))
            },
            "diagram" => {
                sequence.add_graphic_build(GraphicBuild::diagram(shape_id, GroupId::new(group_id)))
            },
            "ole-diagram" => {
                sequence.add_diagram_build(DiagramBuild::new(shape_id, GroupId::new(group_id)))
            },
            "ole-chart" => {
                sequence.add_ole_chart_build(OleChartBuild::new(shape_id, GroupId::new(group_id)))
            },
            _ => unreachable!(),
        }
        sequence.to_xml()
    }

    fn slide(host: &str, timing: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="{p}" xmlns:a="{a}" xmlns:c="{c}" xmlns:dgm="{dgm}" xmlns:r="{r}"><p:cSld><p:spTree>{host}</p:spTree></p:cSld>{timing}</p:sld>"#,
            p = std::str::from_utf8(P_NS).unwrap(),
            a = std::str::from_utf8(A_NS).unwrap(),
            c = std::str::from_utf8(C_NS).unwrap(),
            dgm = std::str::from_utf8(DGM_NS).unwrap(),
            r = std::str::from_utf8(R_NS).unwrap(),
        )
    }

    fn host(child: &str) -> String {
        host_with_uri(child, std::str::from_utf8(C_NS).unwrap())
    }

    fn host_with_uri(child: &str, uri: &str) -> String {
        format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Host"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="{uri}">{child}</a:graphicData></a:graphic></p:graphicFrame>"#
        )
    }

    fn package(
        xml: String,
        relationships: &[(&str, &str, &str, bool)],
        parts: &[(&str, &str)],
    ) -> OpcPackage {
        let slide_uri = PackURI::new(SLIDE).unwrap();
        let mut slide = BlobPart::new(slide_uri, SLIDE_CT.into(), xml.into_bytes());
        for (id, rel_type, target, external) in relationships {
            slide.rels_mut().add_relationship(
                (*rel_type).into(),
                (*target).into(),
                (*id).into(),
                *external,
            );
        }
        let mut package = OpcPackage::new();
        package.add_part(Box::new(slide));
        for (name, content_type) in parts {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new(*name).unwrap(),
                (*content_type).into(),
                Vec::new(),
            )));
        }
        package
    }

    #[test]
    fn validates_real_producer_chart_diagram_and_ole_hosts_without_reading_targets() {
        let chart_xml = slide(
            &host(r#"<c:chart r:id="rIdChart"/>"#),
            &timing("chart", 5, 1),
        );
        let chart = package(
            chart_xml,
            &[("rIdChart", CHART_REL[0], "../charts/chart1.xml", false)],
            &[("/ppt/charts/chart1.xml", CHART_CT)],
        );
        assert!(Sequence::from_package_slide(&chart, &PackURI::new(SLIDE).unwrap()).is_ok());

        let diagram_xml = slide(
            &host(
                r#"<dgm:relIds r:dm="rIdData" r:lo="rIdLayout" r:qs="rIdStyle" r:cs="rIdColors"/>"#,
            ),
            &timing("diagram", 5, 2),
        );
        let diagram = package(
            diagram_xml,
            &[
                (
                    "rIdData",
                    DIAGRAM_DATA_REL[0],
                    "../diagrams/data1.xml",
                    false,
                ),
                (
                    "rIdLayout",
                    DIAGRAM_LAYOUT_REL[0],
                    "../diagrams/layout1.xml",
                    false,
                ),
                (
                    "rIdStyle",
                    DIAGRAM_STYLE_REL[0],
                    "../diagrams/quickStyle1.xml",
                    false,
                ),
                (
                    "rIdColors",
                    DIAGRAM_COLORS_REL[0],
                    "../diagrams/colors1.xml",
                    false,
                ),
            ],
            &[
                ("/ppt/diagrams/data1.xml", DIAGRAM_DATA_CT),
                ("/ppt/diagrams/layout1.xml", DIAGRAM_LAYOUT_CT),
                ("/ppt/diagrams/quickStyle1.xml", DIAGRAM_STYLE_CT),
                ("/ppt/diagrams/colors1.xml", DIAGRAM_COLORS_CT),
            ],
        );
        assert!(Sequence::from_package_slide(&diagram, &PackURI::new(SLIDE).unwrap()).is_ok());

        for kind in ["ole-diagram", "ole-chart"] {
            let ole_xml = slide(
                &host(r#"<p:oleObj progId="MSGraph.Chart.8" r:id="rIdOle"/>"#),
                &timing(kind, 5, 3),
            );
            let ole = package(
                ole_xml,
                &[("rIdOle", OLE_REL[0], "../embeddings/oleObject1.bin", false)],
                &[("/ppt/embeddings/oleObject1.bin", OLE_CT)],
            );
            assert!(Sequence::from_package_slide(&ole, &PackURI::new(SLIDE).unwrap()).is_ok());
        }
    }

    #[test]
    fn rejects_external_missing_traversal_and_mismatched_chart_targets() {
        let xml = slide(
            &host(r#"<c:chart r:id="rIdChart"/>"#),
            &timing("chart", 5, 1),
        );
        let cases = [
            package(
                xml.clone(),
                &[(
                    "rIdChart",
                    CHART_REL[0],
                    "https://example.test/chart.xml",
                    true,
                )],
                &[],
            ),
            package(xml.clone(), &[], &[]),
            package(
                xml.clone(),
                &[("rIdChart", CHART_REL[0], "../charts/missing.xml", false)],
                &[],
            ),
            package(
                xml.clone(),
                &[("rIdChart", OLE_REL[0], "../charts/chart1.xml", false)],
                &[("/ppt/charts/chart1.xml", CHART_CT)],
            ),
            package(
                xml.clone(),
                &[("rIdChart", CHART_REL[0], "../charts/chart1.xml", false)],
                &[("/ppt/charts/chart1.xml", OLE_CT)],
            ),
            package(
                xml.clone(),
                &[(
                    "rIdChart",
                    CHART_REL[0],
                    "../../word/charts/chart1.xml",
                    false,
                )],
                &[("/word/charts/chart1.xml", CHART_CT)],
            ),
            package(
                xml,
                &[(
                    "rIdChart",
                    CHART_REL[0],
                    "../charts/chart1.xml#fragment",
                    false,
                )],
                &[("/ppt/charts/chart1.xml", CHART_CT)],
            ),
        ];
        for package in cases {
            assert!(Sequence::from_package_slide(&package, &PackURI::new(SLIDE).unwrap()).is_err());
        }
    }

    #[test]
    fn rejects_ambiguous_wrong_namespace_and_build_host_mismatches() {
        let duplicate = slide(
            &host(r#"<c:chart r:id="rIdChart"/><c:chart r:id="rIdOther"/>"#),
            &timing("chart", 5, 1),
        );
        let wrong_namespace = slide(
            &host(r#"<c:chart xmlns:x="urn:foreign" x:id="rIdChart"/>"#),
            &timing("chart", 5, 1),
        );
        let mismatch = slide(
            &host(r#"<c:chart r:id="rIdChart"/>"#),
            &timing("diagram", 5, 1),
        );
        for xml in [duplicate, wrong_namespace, mismatch] {
            let package = package(
                xml,
                &[
                    ("rIdChart", CHART_REL[0], "../charts/chart1.xml", false),
                    ("rIdOther", CHART_REL[0], "../charts/chart2.xml", false),
                ],
                &[
                    ("/ppt/charts/chart1.xml", CHART_CT),
                    ("/ppt/charts/chart2.xml", CHART_CT),
                ],
            );
            assert!(Sequence::from_package_slide(&package, &PackURI::new(SLIDE).unwrap()).is_err());
        }
    }

    #[test]
    fn rejects_duplicate_smartart_ids_wrong_parts_and_scan_boundaries() {
        let xml = slide(
            &host(r#"<dgm:relIds r:dm="same" r:lo="same" r:qs="same" r:cs="same"/>"#),
            &timing("diagram", 5, 1),
        );
        let package = package(
            xml,
            &[("same", DIAGRAM_DATA_REL[0], "../diagrams/data1.xml", false)],
            &[("/ppt/diagrams/data1.xml", DIAGRAM_DATA_CT)],
        );
        assert!(Sequence::from_package_slide(&package, &PackURI::new(SLIDE).unwrap()).is_err());

        let oversized = vec![b' '; MAX_SLIDE_XML + 1];
        let oversized_slide = PackURI::new(SLIDE).unwrap();
        let mut oversized_package = OpcPackage::new();
        oversized_package.add_part(Box::new(BlobPart::new(
            oversized_slide.clone(),
            SLIDE_CT.into(),
            oversized,
        )));
        assert!(Sequence::from_package_slide(&oversized_package, &oversized_slide).is_err());
    }

    #[test]
    fn validates_direct_mc_and_strict_chartex_producer_hosts() {
        let direct_xml = slide(
            &host_with_uri(r#"<c:chart r:id="rIdChartEx"/>"#, CHARTEX_URI),
            &timing("chart", 5, 1),
        );
        let direct = package(
            direct_xml,
            &[(
                "rIdChartEx",
                CHARTEX_REL[0],
                "../charts/chartEx1.xml",
                false,
            )],
            &[("/ppt/charts/chartEx1.xml", CHARTEX_CT)],
        );
        assert!(Sequence::from_package_slide(&direct, &PackURI::new(SLIDE).unwrap()).is_ok());

        let choice = host_with_uri(r#"<c:chart r:id="rIdChartEx"/>"#, CHARTEX_URI);
        let mc_host = format!(
            r#"<mc:AlternateContent xmlns:mc="{mc}" xmlns:cx="{cx}"><mc:Choice Requires="cx">{choice}</mc:Choice><mc:Fallback><p:pic/></mc:Fallback></mc:AlternateContent>"#,
            mc = std::str::from_utf8(MC_NS).unwrap(),
            cx = CHARTEX_URI,
        );
        let mc = package(
            slide(&mc_host, &timing("chart", 5, 2)),
            &[(
                "rIdChartEx",
                CHARTEX_REL[0],
                "../charts/chartEx2.xml",
                false,
            )],
            &[("/ppt/charts/chartEx2.xml", CHARTEX_CT)],
        );
        Sequence::from_package_slide(&mc, &PackURI::new(SLIDE).unwrap())
            .expect("Office-style ChartEx AlternateContent should validate");

        let strict_timing = timing("chart", 5, 3)
            .replace(
                std::str::from_utf8(P_NS).unwrap(),
                std::str::from_utf8(P_STRICT_NS).unwrap(),
            )
            .replace(
                std::str::from_utf8(R_NS).unwrap(),
                std::str::from_utf8(R_STRICT_NS).unwrap(),
            );
        let strict_xml = format!(
            r#"<p:sld xmlns:p="{p}" xmlns:a="{a}" xmlns:c="{c}" xmlns:r="{r}"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Host"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="{cx}"><c:chart r:id="rIdChartEx"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>{strict_timing}</p:sld>"#,
            p = std::str::from_utf8(P_STRICT_NS).unwrap(),
            a = std::str::from_utf8(A_STRICT_NS).unwrap(),
            c = std::str::from_utf8(C_STRICT_NS).unwrap(),
            r = std::str::from_utf8(R_STRICT_NS).unwrap(),
            cx = CHARTEX_URI,
        );
        let strict = package(
            strict_xml,
            &[(
                "rIdChartEx",
                CHARTEX_REL[0],
                "../charts/chartEx3.xml",
                false,
            )],
            &[("/ppt/charts/chartEx3.xml", CHARTEX_CT)],
        );
        Sequence::from_package_slide(&strict, &PackURI::new(SLIDE).unwrap())
            .expect("strict ChartEx host should validate");
    }

    #[test]
    fn rejects_chartex_relationship_content_type_and_uri_cross_wiring() {
        let chartex_xml = slide(
            &host_with_uri(r#"<c:chart r:id="rIdChart"/>"#, CHARTEX_URI),
            &timing("chart", 5, 1),
        );
        let classic_xml = slide(
            &host(r#"<c:chart r:id="rIdChart"/>"#),
            &timing("chart", 5, 1),
        );
        let vendor_xml = slide(
            &host_with_uri(r#"<c:chart r:id="rIdChart"/>"#, "urn:vendor:chartex"),
            &timing("chart", 5, 1),
        );
        let cases = [
            package(
                chartex_xml.clone(),
                &[("rIdChart", CHART_REL[0], "../charts/chart1.xml", false)],
                &[("/ppt/charts/chart1.xml", CHART_CT)],
            ),
            package(
                chartex_xml,
                &[("rIdChart", CHARTEX_REL[0], "../charts/chartEx1.xml", false)],
                &[("/ppt/charts/chartEx1.xml", CHART_CT)],
            ),
            package(
                classic_xml,
                &[("rIdChart", CHARTEX_REL[0], "../charts/chartEx1.xml", false)],
                &[("/ppt/charts/chartEx1.xml", CHARTEX_CT)],
            ),
            package(
                vendor_xml,
                &[("rIdChart", CHARTEX_REL[0], "../charts/chartEx1.xml", false)],
                &[("/ppt/charts/chartEx1.xml", CHARTEX_CT)],
            ),
        ];
        for package in cases {
            assert!(Sequence::from_package_slide(&package, &PackURI::new(SLIDE).unwrap()).is_err());
        }
    }

    #[test]
    fn rejects_vendor_versioned_and_malformed_chartex_alternate_content() {
        let choice = host_with_uri(r#"<c:chart r:id="rIdChartEx"/>"#, CHARTEX_URI);
        let mc = std::str::from_utf8(MC_NS).unwrap();
        let hostile = [
            format!(
                r#"<mc:AlternateContent xmlns:mc="{mc}" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"><mc:Choice Requires="cx2">{choice}</mc:Choice><mc:Fallback><p:pic/></mc:Fallback></mc:AlternateContent>"#
            ),
            format!(
                r#"<mc:AlternateContent xmlns:mc="{mc}" xmlns:cx="{CHARTEX_URI}" xmlns:v="urn:vendor"><mc:Choice Requires="cx v">{choice}</mc:Choice><mc:Fallback><p:pic/></mc:Fallback></mc:AlternateContent>"#
            ),
            format!(
                r#"<mc:AlternateContent xmlns:mc="{mc}" xmlns:cx="{CHARTEX_URI}"><mc:Choice>{choice}</mc:Choice></mc:AlternateContent>"#
            ),
            format!(
                r#"<mc:AlternateContent xmlns:mc="{mc}" xmlns:cx="{CHARTEX_URI}"><mc:Fallback><p:pic/></mc:Fallback><mc:Choice Requires="cx">{choice}</mc:Choice></mc:AlternateContent>"#
            ),
            format!(
                r#"<mc:AlternateContent xmlns:mc="{mc}" xmlns:cx="{CHARTEX_URI}"><mc:Choice Requires="cx"><mc:AlternateContent><mc:Choice Requires="cx">{choice}</mc:Choice></mc:AlternateContent></mc:Choice></mc:AlternateContent>"#
            ),
        ];
        for host in hostile {
            let package = package(
                slide(&host, &timing("chart", 5, 1)),
                &[(
                    "rIdChartEx",
                    CHARTEX_REL[0],
                    "../charts/chartEx1.xml",
                    false,
                )],
                &[("/ppt/charts/chartEx1.xml", CHARTEX_CT)],
            );
            assert!(Sequence::from_package_slide(&package, &PackURI::new(SLIDE).unwrap()).is_err());
        }
    }

    #[test]
    fn enforces_chartex_relationship_id_boundary() {
        for (length, valid) in [(255usize, true), (256usize, false)] {
            let id = "r".repeat(length);
            let xml = slide(
                &host_with_uri(&format!(r#"<c:chart r:id="{id}"/>"#), CHARTEX_URI),
                &timing("chart", 5, 1),
            );
            let package = package(
                xml,
                &[(&id, CHARTEX_REL[0], "../charts/chartEx1.xml", false)],
                &[("/ppt/charts/chartEx1.xml", CHARTEX_CT)],
            );
            assert_eq!(
                Sequence::from_package_slide(&package, &PackURI::new(SLIDE).unwrap()).is_ok(),
                valid,
            );
        }
    }
}
